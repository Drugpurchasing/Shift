import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, time
from openpyxl import Workbook
from github import Github, Auth 
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter
import random
from statistics import stdev
from io import BytesIO
from streamlit_gsheets import GSheetsConnection
# Library สำหรับเชื่อมต่อ Google Sheets (ไม่จำเป็นต้อง import GSheetsConnection โดยตรง)
# -- App Configuration --
st.set_page_config(page_title="Pharmacist Data Manager", layout="centered")
st.title("👩‍⚕️ Pharmacist Data Manager (Google Sheets)")

# ---------------- CORRECT ORDER ----------------
# 1. ESTABLISH THE CONNECTION FIRST
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as e:
    st.error("Failed to create Google Sheets connection. Check your secrets.toml file.")
    st.stop() # Stop the app if connection fails

# 2. DEFINE THE FUNCTION THAT USES THE CONNECTION
@st.cache_data(ttl=60)
def load_data(worksheet_name):
    """Loads data using the established 'conn' object."""
    try:
        data = conn.read(worksheet=worksheet_name, usecols=list(range(5)), header=0)
        data = data.dropna(how="all")
        return data
    except Exception as e:
        st.error(f"Failed to load data from worksheet '{worksheet_name}': {e}")
        return pd.DataFrame()

# -----------------------------------------------

# -- Main App --
st.header("Current Pharmacist Data")
st.write("This data is being read live from your Google Sheet.")

WORKSHEET_NAME = "Pharmacists" # ⚠️ Change this to your actual worksheet name

# 3. NOW, YOU CAN SAFELY CALL THE FUNCTION
df = load_data(WORKSHEET_NAME)

if not df.empty:
    st.dataframe(df)
else:
    st.warning("Could not load data or the worksheet is empty.")

# =========================================================================
# ================== PHARMACIST SCHEDULER CLASS (ฉบับปรับปรุง) =============
# =========================================================================
class PharmacistScheduler:
    """
    Pharmacy shift scheduler with optimization and Excel export.
    Designed for multi-constraint pharmacist roster planning.
    """
    W_CONSECUTIVE = 8
    W_HOURS = 4
    W_PREFERENCE = 4

    def __init__(self, dataframes: dict):
        self.pharmacists = {}
        self.shift_types = {}
        self.departments = {}
        self.pre_assignments = {}
        self.historical_scores = {}
        self.preference_multipliers = {}
        self.special_notes = {}
        self.shift_limits = {}
        self.problem_days = set()
        
        # Holiday data from sheets
        self.holidays = {'specific_dates': []}

        self.process_dataframes(dataframes)
        self.load_historical_scores(dataframes)
        self._calculate_preference_multipliers()

        self.night_shifts = {
            'I100-10', 'I100-12N', 'I400-12N', 'I400-10', 'O400ER-12N', 'O400ER-10'
        }
        
        for pharmacist in self.pharmacists:
            self.pharmacists[pharmacist]['shift_counts'] = {
                shift_type: 0 for shift_type in self.shift_types
            }
    
    def process_dataframes(self, dataframes: dict):
        pharmacists_df = dataframes.get('pharmacists')
        shifts_df = dataframes.get('shifts')
        departments_df = dataframes.get('departments')
        pre_assign_df = dataframes.get('pre_assignments')
        notes_df = dataframes.get('special_notes')
        limits_df = dataframes.get('shift_limits')

        # === Process Pharmacists ===
        self.pharmacists = {}
        for _, row in pharmacists_df.iterrows():
            name = row['Name']
            max_hours = row.get('Max Hours', 250)
            if pd.isna(max_hours) or max_hours == '' or max_hours is None:
                max_hours = 250
            else:
                max_hours = float(max_hours)
            self.pharmacists[name] = {
                'night_shift_count': 0,
                'skills': str(row['Skills']).split(','),
                'holidays': [date for date in str(row['Holidays']).split(',') if date != '1900-01-00' and date.strip() and date != 'nan'],
                'shift_counts': {},
                'preferences': {f'rank{i}': row[f'Rank{i}'] for i in range(1, 9)},
                'max_hours': max_hours
            }
        
        # === Process Shifts ===
        self.shift_types = {}
        for _, row in shifts_df.iterrows():
            shift_code = row['Shift Code']
            self.shift_types[shift_code] = {
                'description': row['Description'],
                'shift_type': row['Shift Type'],
                'start_time': str(row['Start Time']), # Ensure time is string
                'end_time': str(row['End Time']),     # Ensure time is string
                'hours': row['Hours'],
                'required_skills': str(row['Required Skills']).split(','),
                'restricted_next_shifts': str(row['Restricted Next Shifts']).split(',') if pd.notna(row['Restricted Next Shifts']) else [],
            }

        # === Process Departments ===
        self.departments = {}
        for _, row in departments_df.iterrows():
            department = row['Department']
            self.departments[department] = str(row['Shift Codes']).split(',')
        
        # === Process PreAssignments ===
        pre_assign_df['Date'] = pd.to_datetime(pre_assign_df['Date'], errors='coerce').dt.strftime('%Y-%m-%d')
        self.pre_assignments = {}
        for pharmacist, group in pre_assign_df.groupby('Pharmacist'):
            date_dict = {}
            for date, g in group.groupby('Date'):
                shifts = []
                for shift_str in g['Shift']:
                    shifts.extend([s.strip() for s in str(shift_str).split(',') if s.strip()])
                date_dict[date] = shifts
            self.pre_assignments[pharmacist] = date_dict

        # === Process SpecialNotes (if available) ===
        if notes_df is not None:
            # Assuming first column is the pharmacist's name, so reset index
            notes_df.reset_index(inplace=True)
            # The first column name might be 'index' or the original column name
            pharmacist_col_name = notes_df.columns[0]
            
            for _, row_data in notes_df.iterrows():
                pharmacist = row_data[pharmacist_col_name]
                if pharmacist in self.pharmacists:
                    for date_col, note in row_data.drop(pharmacist_col_name).items():
                        if pd.notna(note) and str(note).strip():
                            try:
                                date_str = pd.to_datetime(date_col).strftime('%Y-%m-%d')
                                if pharmacist not in self.special_notes:
                                    self.special_notes[pharmacist] = {}
                                self.special_notes[pharmacist][date_str] = str(note).strip()
                            except (ValueError, TypeError):
                                continue # Ignore columns that aren't valid dates

        # === Process ShiftLimits (if available) ===
        if limits_df is not None:
            for _, row in limits_df.iterrows():
                pharmacist = row['Pharmacist']
                category = row['ShiftCategory']
                max_count = row['MaxCount']
                if pharmacist in self.pharmacists:
                    if pharmacist not in self.shift_limits:
                        self.shift_limits[pharmacist] = {}
                    self.shift_limits[pharmacist][category] = int(max_count)

    def load_historical_scores(self, dataframes: dict):
        df = dataframes.get('historical_scores')
        if df is None:
            st.info("INFO: Sheet 'HistoricalScores' not found. Proceeding without historical data.")
            return

        if 'Pharmacist' in df.columns and 'Total Preference Score' in df.columns:
            for _, row in df.iterrows():
                pharmacist = row['Pharmacist']
                score = row['Total Preference Score']
                if pharmacist in self.pharmacists:
                    self.historical_scores[pharmacist] = score
        else:
            st.warning("WARNING: 'HistoricalScores' sheet found, but required columns ('Pharmacist', 'Total Preference Score') are missing.")

    # ... (ส่วนที่เหลือของคลาส PharmacistScheduler เหมือนเดิมทั้งหมด) ...
    def _pre_check_staffing_levels(self, year, month):
        st.write("\nRunning pre-check for staffing levels (including all shifts + 3 buffer)...")
        start_date = datetime(year, month, 1)
        end_date = datetime(year, month + 1, 1) - timedelta(days=1) if month < 12 else datetime(year, 12, 31)
        dates = pd.date_range(start_date, end_date)

        all_ok = True
        for date in dates:
            available_pharmacists_count = sum(1 for p_name, p_info in self.pharmacists.items()
                                              if date.strftime('%Y-%m-%d') not in p_info['holidays'])
            required_shifts_base = sum(1 for st in self.shift_types
                                       if self.is_shift_available_on_date(st, date))
            total_required_shifts_with_buffer = required_shifts_base + 3

            if available_pharmacists_count < total_required_shifts_with_buffer:
                all_ok = False
                self.problem_days.add(date)
                st.warning(f"WARNING: Potential shortage on {date.strftime('%Y-%m-%d')}. "
                      f"Available Pharmacists: {available_pharmacists_count}, "
                      f"Required Shifts (with +3 buffer): {total_required_shifts_with_buffer}")
        if all_ok:
            st.success("Pre-check complete. All days have sufficient staffing levels for the total workload.")
        else:
            st.warning("Pre-check complete. Identified days with potential staff shortages. These will be prioritized.")
        return not all_ok


    def _calculate_preference_multipliers(self):
        if not self.historical_scores:
            for pharmacist in self.pharmacists:
                self.preference_multipliers[pharmacist] = 1.0
            return
        min_score = min(self.historical_scores.values())
        max_score = max(self.historical_scores.values())
        if min_score == max_score:
            for pharmacist in self.pharmacists:
                self.preference_multipliers[pharmacist] = 1.0
            return
        for pharmacist, score in self.historical_scores.items():
            normalized_score = (score - min_score) / (max_score - min_score)
            min_multiplier = 0.7
            self.preference_multipliers[pharmacist] = min_multiplier + (1 - min_multiplier) * normalized_score
        for pharmacist in self.pharmacists:
            if pharmacist not in self.preference_multipliers:
                min_multiplier = 0.7
                self.preference_multipliers[pharmacist] = min_multiplier

    def convert_time_to_minutes(self, time_input):
        if isinstance(time_input, str):
            try:
                hours, minutes = map(int, time_input.split(':'))
            except ValueError: # Handle if format is not HH:MM
                return 0
        elif isinstance(time_input, time):
            hours, minutes = time_input.hour, time_input.minute
        else:
            return 0 # Default for unexpected types
        return hours * 60 + minutes

    def check_time_overlap(self, start1, end1, start2, end2):
        start1_mins = self.convert_time_to_minutes(start1)
        end1_mins = self.convert_time_to_minutes(end1)
        start2_mins = self.convert_time_to_minutes(start2)
        end2_mins = self.convert_time_to_minutes(end2)
        if end1_mins < start1_mins: end1_mins += 24 * 60
        if end2_mins < start2_mins: end2_mins += 24 * 60
        return start1_mins < end2_mins and end1_mins > start2_mins

    def check_mixing_expert_ratio_optimized(self, schedule_dict, date, current_shift=None, current_pharm=None):
        mixing_shifts = [p for s, p in schedule_dict[date].items()
                         if s.startswith('C8') and p not in ['NO SHIFT', 'UNASSIGNED', 'UNFILLED']]
        if current_shift and current_shift.startswith('C8') and current_pharm:
            mixing_shifts.append(current_pharm)
        if not mixing_shifts: return True
        total_mixing = len(mixing_shifts)
        expert_count = sum(1 for pharm in mixing_shifts
                           if pharm in self.pharmacists and 'mixing_expert' in self.pharmacists[pharm]['skills'])
        return expert_count >= (2 * total_mixing / 3)

    def count_consecutive_shifts(self, pharmacist, date, schedule, max_days=6):
        count = 0
        current_date = date - timedelta(days=1)
        for _ in range(max_days):
            if current_date in schedule.index and pharmacist in schedule.loc[current_date].values:
                count += 1
                current_date -= timedelta(days=1)
            else:
                break
        return count

    def is_holiday(self, date):
        return date.strftime('%Y-%m-%d') in self.holidays['specific_dates']

    def calculate_weekend_off_variance(self, schedule, year, month):
        weekend_off_counts = {p: 0 for p in self.pharmacists}
        start_date = datetime(year, month, 1)
        end_date = datetime(year, month + 1, 1) - timedelta(days=1) if month < 12 else datetime(year, 12, 31)
        for date in pd.date_range(start_date, end_date):
            if date.weekday() >= 5:
                working_on_weekend = {schedule.loc[date, shift] for shift in schedule.columns if schedule.loc[date, shift] not in ['NO SHIFT', 'UNFILLED', 'UNASSIGNED']}
                for p_name in self.pharmacists:
                    if p_name not in working_on_weekend:
                        weekend_off_counts[p_name] += 1
        if len(weekend_off_counts) > 1:
            return np.var(list(weekend_off_counts.values()))
        return 0

    def is_night_shift(self, shift_type):
        return shift_type in self.night_shifts

    def is_shift_available_on_date(self, shift_type, date):
        shift_info = self.shift_types[shift_type]
        is_holiday_date = self.is_holiday(date)
        is_saturday = date.weekday() == 5
        is_sunday = date.weekday() == 6
        if shift_info['shift_type'] == 'weekday': return not (is_holiday_date or is_saturday or is_sunday)
        elif shift_info['shift_type'] == 'saturday': return is_saturday and not is_holiday_date
        elif shift_info['shift_type'] == 'holiday': return is_holiday_date or is_saturday or is_sunday
        elif shift_info['shift_type'] == 'night': return True
        return False

    def get_department_from_shift(self, shift_type):
        if shift_type.startswith('I100'): return 'IPD100'
        elif shift_type.startswith('O100'): return 'OPD100'
        elif shift_type.startswith('Care'): return 'Care'
        elif shift_type.startswith('C8'): return 'Mixing'
        elif shift_type.startswith('I400'): return 'IPD400'
        elif shift_type.startswith('O400F1'): return 'OPD400F1'
        elif shift_type.startswith('O400F2'): return 'OPD400F2'
        elif shift_type.startswith('O400ER'): return 'ER'
        elif shift_type.startswith('ARI'): return 'ARI'
        return None

    def _get_shift_category(self, shift_type):
        if self.is_night_shift(shift_type):
            return 'Night'
        if shift_type.startswith('C8'):
            return 'Mixing'
        return None

    def get_night_shift_count(self, pharmacist):
        return self.pharmacists[pharmacist]['night_shift_count']

    def get_preference_score(self, pharmacist, shift_type):
        department = self.get_department_from_shift(shift_type)
        for rank in range(1, 9):
            if self.pharmacists[pharmacist]['preferences'][f'rank{rank}'] == department:
                return rank
        return 9

    def has_restricted_sequence_optimized(self, pharmacist, date, shift_type, schedule_dict):
        previous_date = date - timedelta(days=1)
        if previous_date in schedule_dict:
            for prev_shift, assigned_pharm in schedule_dict[previous_date].items():
                if assigned_pharm == pharmacist:
                    restricted = self.shift_types[prev_shift].get('restricted_next_shifts', [])
                    if shift_type in restricted: return True
        return False

    def has_overlapping_shift_optimized(self, pharmacist, date, new_shift_type, schedule_dict):
        if date not in schedule_dict: return False
        new_start = self.shift_types[new_shift_type]['start_time']
        new_end = self.shift_types[new_shift_type]['end_time']
        for existing_shift, assigned_pharm in schedule_dict[date].items():
            if assigned_pharm == pharmacist and existing_shift != new_shift_type:
                existing_start = self.shift_types[existing_shift]['start_time']
                existing_end = self.shift_types[existing_shift]['end_time']
                if self.check_time_overlap(new_start, new_end, existing_start, existing_end):
                    return True
        return False

    def has_nearby_night_shift_optimized(self, pharmacist, date, schedule_dict):
        for delta in [-2, -1, 1, 2]:
            check_date = date + timedelta(days=delta)
            if check_date in schedule_dict:
                for shift, assigned_pharm in schedule_dict[check_date].items():
                    if assigned_pharm == pharmacist and self.is_night_shift(shift):
                        return True
        return False

    def get_pharmacist_shifts(self, pharmacist, date, current_schedule):
        shifts = []
        if date in current_schedule.index:
            for shift_type in current_schedule.columns:
                if current_schedule.loc[date, shift_type] == pharmacist:
                    shifts.append(shift_type)
        return shifts

    def calculate_total_hours(self, pharmacist, schedule):
        total_hours = 0
        for date in schedule.index:
            for shift_type, assigned_pharm in schedule.loc[date].items():
                if assigned_pharm == pharmacist and shift_type in self.shift_types:
                    total_hours += self.shift_types[shift_type]['hours']
        return total_hours

    def _get_hour_imbalance_penalty(self, hours_dict):
        if not hours_dict or len(hours_dict) < 2:
            return 0
        hour_values = list(hours_dict.values())
        hour_stdev = stdev(hour_values)
        hour_range = max(hour_values) - min(hour_values)
        stdev_penalty = hour_stdev ** 2
        range_penalty = 0
        if hour_range > 10:
            range_penalty = (hour_range - 10) ** 2
        return stdev_penalty + range_penalty

    def calculate_schedule_metrics(self, schedule, year, month):
        hours = {p: self.calculate_total_hours(p, schedule) for p in self.pharmacists}
        night_counts = {p: self.pharmacists[p]['night_shift_count'] for p in self.pharmacists}
        weekend_off_var = self.calculate_weekend_off_variance(schedule, year, month)
        hour_penalty = self._get_hour_imbalance_penalty(hours)
        metrics = {
            'hour_imbalance_penalty': hour_penalty,
            'night_variance': np.var(list(night_counts.values())) if night_counts else 0,
            'preference_score': sum(self.calculate_preference_penalty(p, schedule) for p in self.pharmacists),
            'weekend_off_variance': weekend_off_var
        }
        if len(hours) > 1:
            metrics['hour_diff_for_logging'] = stdev(hours.values())
        else:
            metrics['hour_diff_for_logging'] = 0
        return metrics

    def generate_monthly_schedule_shuffled(self, year, month, shuffled_shifts=None, shuffled_pharmacists=None, iteration_num=1):
        start_date = datetime(year, month, 1)
        end_date = datetime(year + 1, 1, 1) - timedelta(days=1) if month == 12 else datetime(year, month + 1, 1) - timedelta(days=1)
        dates = pd.date_range(start_date, end_date)
        schedule_dict = {date: {shift: 'NO SHIFT' for shift in self.shift_types} for date in dates}
        pharmacist_hours = {p: 0 for p in self.pharmacists}
        pharmacist_consecutive_days = {p: 0 for p in self.pharmacists}

        if shuffled_shifts is None:
            shuffled_shifts = list(self.shift_types.keys())
            random.shuffle(shuffled_shifts)
        if shuffled_pharmacists is None:
            shuffled_pharmacists = list(self.pharmacists.keys())
            random.shuffle(shuffled_pharmacists)

        for pharmacist in self.pharmacists:
            self.pharmacists[pharmacist]['night_shift_count'] = 0
            self.pharmacists[pharmacist]['mixing_shift_count'] = 0
            self.pharmacists[pharmacist]['category_counts'] = {
                'Mixing': 0,
                'Night': 0
            }

        for pharmacist, assignments in self.pre_assignments.items():
            if pharmacist not in self.pharmacists: continue
            for date_str, shift_types in assignments.items():
                date = pd.to_datetime(date_str)
                if date not in schedule_dict: continue
                for shift_type in shift_types:
                    if shift_type in self.shift_types:
                        schedule_dict[date][shift_type] = pharmacist
                        self._update_shift_counts(pharmacist, shift_type)
                        pharmacist_hours[pharmacist] += self.shift_types[shift_type]['hours']

        all_dates = list(pd.date_range(start_date, end_date))
        problem_dates_sorted = sorted([d for d in all_dates if d in self.problem_days])
        other_dates_sorted = sorted([d for d in all_dates if d not in self.problem_days])
        processing_order_dates = problem_dates_sorted + other_dates_sorted
        unfilled_info = {'problem_days': [], 'other_days': []}
        night_shifts_ordered = [s for s in shuffled_shifts if self.is_night_shift(s)]
        mixing_shifts_ordered = [s for s in shuffled_shifts if s.startswith('C8') and not self.is_night_shift(s)]
        care_shifts_ordered = [s for s in shuffled_shifts if s.startswith('Care') and not self.is_night_shift(s) and not s.startswith('C8')]
        other_shifts_ordered = [s for s in shuffled_shifts if not self.is_night_shift(s) and not s.startswith('C8') and not s.startswith('Care')]
        standard_shift_order = night_shifts_ordered + mixing_shifts_ordered + care_shifts_ordered + other_shifts_ordered
        problem_day_shift_order = mixing_shifts_ordered + care_shifts_ordered + night_shifts_ordered + other_shifts_ordered

        for date in processing_order_dates:
            pharmacists_working_yesterday = set()
            previous_date = date - timedelta(days=1)
            if previous_date in schedule_dict:
                    pharmacists_working_yesterday = {p for p in schedule_dict[previous_date].values() if p in self.pharmacists}
            for p_name in self.pharmacists:
                if p_name in pharmacists_working_yesterday:
                    pharmacist_consecutive_days[p_name] += 1
                else:
                    pharmacist_consecutive_days[p_name] = 0

            is_day_before_problem_day = (date + timedelta(days=1)) in self.problem_days

            shifts_to_process = problem_day_shift_order if date in self.problem_days else standard_shift_order
            for shift_type in shifts_to_process:
                if schedule_dict[date][shift_type] not in ['NO SHIFT', 'UNASSIGNED', 'UNFILLED'] or not self.is_shift_available_on_date(shift_type, date):
                    continue
                available = self._get_available_pharmacists_optimized(shuffled_pharmacists, date, shift_type, schedule_dict, pharmacist_hours, pharmacist_consecutive_days)
                if available:
                    chosen = self._select_best_pharmacist(available, shift_type, date, is_day_before_problem_day)
                    pharmacist_to_assign = chosen['name']
                    schedule_dict[date][shift_type] = pharmacist_to_assign
                    self._update_shift_counts(pharmacist_to_assign, shift_type)
                    pharmacist_hours[pharmacist_to_assign] += self.shift_types[shift_type]['hours']
                else:
                    schedule_dict[date][shift_type] = 'UNFILLED'
                    if date in self.problem_days:
                        unfilled_info['problem_days'].append((date, shift_type))
                    else:
                        unfilled_info['other_days'].append((date, shift_type))
                        final_schedule = pd.DataFrame.from_dict(schedule_dict, orient='index')
                        final_schedule.fillna('NO SHIFT', inplace=True)
                        return final_schedule, unfilled_info
        final_schedule = pd.DataFrame.from_dict(schedule_dict, orient='index')
        final_schedule = final_schedule.reindex(columns=list(self.shift_types.keys()), fill_value='NO SHIFT')
        final_schedule.fillna('NO SHIFT', inplace=True)
        return final_schedule, unfilled_info

    def _update_shift_counts(self, pharmacist, shift_type):
        if self.is_night_shift(shift_type):
            self.pharmacists[pharmacist]['night_shift_count'] += 1
        if shift_type.startswith('C8'):
            self.pharmacists[pharmacist]['mixing_shift_count'] += 1
        category = self._get_shift_category(shift_type)
        if category and category in self.pharmacists[pharmacist]['category_counts']:
            self.pharmacists[pharmacist]['category_counts'][category] += 1

    def _get_available_pharmacists_optimized(self, pharmacists, date, shift_type, schedule_dict, current_hours_dict, consecutive_days_dict):
        available_pharmacists = []
        pharmacists_on_night_yesterday = set()
        previous_date = date - timedelta(days=1)
        if previous_date in schedule_dict:
            pharmacists_on_night_yesterday = {
                p for s, p in schedule_dict[previous_date].items()
                if p in self.pharmacists and self.is_night_shift(s)
            }
        for pharmacist in pharmacists:
            if date.strftime('%Y-%m-%d') in self.pharmacists[pharmacist]['holidays']: continue
            if self.has_overlapping_shift_optimized(pharmacist, date, shift_type, schedule_dict): continue
            if pharmacist in pharmacists_on_night_yesterday: continue
            p_skills = self.pharmacists[pharmacist]['skills']
            s_req_skills = self.shift_types[shift_type]['required_skills']
            if not all(skill.strip() in p_skills for skill in s_req_skills if skill.strip()): continue
            projected_hours = current_hours_dict[pharmacist] + self.shift_types[shift_type]['hours']
            if projected_hours > self.pharmacists[pharmacist].get('max_hours', 250): continue
            if self.has_restricted_sequence_optimized(pharmacist, date, shift_type, schedule_dict): continue
            category = self._get_shift_category(shift_type)
            if category:
                limit = self.shift_limits.get(pharmacist, {}).get(category)
                if limit is not None:
                    current_count = self.pharmacists[pharmacist]['category_counts'][category]
                    if current_count >= limit:
                        continue
            if self.is_night_shift(shift_type):
                if self.has_nearby_night_shift_optimized(pharmacist, date, schedule_dict): continue
                next_date = date + timedelta(days=1)
                if pharmacist in self.pre_assignments and next_date.strftime('%Y-%m-%d') in self.pre_assignments[pharmacist]: continue
            if shift_type.startswith('C8'):
                if not self.check_mixing_expert_ratio_optimized(schedule_dict, date, shift_type, pharmacist):
                    continue
            original_preference = self.get_preference_score(pharmacist, shift_type)
            multiplier = self.preference_multipliers.get(pharmacist, 1.0)
            pharmacist_data = {
                'name': pharmacist,
                'preference_score': original_preference * multiplier,
                'consecutive_days': consecutive_days_dict[pharmacist],
                'night_count': self.pharmacists[pharmacist]['night_shift_count'],
                'mixing_count': self.pharmacists[pharmacist]['mixing_shift_count'],
                'current_hours': current_hours_dict[pharmacist],
            }
            available_pharmacists.append(pharmacist_data)
        return available_pharmacists

    def _calculate_suitability_score(self, pharmacist_data):
        consecutive_penalty = self.W_CONSECUTIVE * (pharmacist_data['consecutive_days'] ** 2)
        hours_penalty = self.W_HOURS * pharmacist_data['current_hours']
        preference_penalty = self.W_PREFERENCE * pharmacist_data['preference_score']
        return consecutive_penalty + hours_penalty + preference_penalty

    def _select_best_pharmacist(self, available_pharmacists, shift_type, date, is_day_before_problem_day):
        if self.is_night_shift(shift_type) and is_day_before_problem_day:
            problem_day = date + timedelta(days=1)
            problem_day_str = problem_day.strftime('%Y-%m-%d')

            candidates_off_tomorrow = []
            for p_data in available_pharmacists:
                p_name = p_data['name']
                if problem_day_str in self.pharmacists[p_name]['holidays']:
                    candidates_off_tomorrow.append(p_data)

            if candidates_off_tomorrow:
                return min(candidates_off_tomorrow, key=lambda x: (x['night_count'], self._calculate_suitability_score(x)))

        if self.is_night_shift(shift_type):
            return min(available_pharmacists, key=lambda x: (x['night_count'], self._calculate_suitability_score(x)))
        elif shift_type.startswith('C8'):
            return min(available_pharmacists, key=lambda x: (x['mixing_count'], self._calculate_suitability_score(x)))
        else:
            return min(available_pharmacists, key=lambda x: self._calculate_suitability_score(x))

    def calculate_preference_penalty(self, pharmacist, schedule):
        penalty = 0
        for date in schedule.index:
            for shift_type, assigned_pharm in schedule.loc[date].items():
                if assigned_pharm == pharmacist:
                    penalty += self.get_preference_score(pharmacist, shift_type)
        return penalty

    def is_schedule_better(self, current_metrics, best_metrics):
        current_unfilled = current_metrics.get('unfilled_problem_shifts', float('inf'))
        best_unfilled = best_metrics.get('unfilled_problem_shifts', float('inf'))
        if current_unfilled < best_unfilled: return True
        if current_unfilled > best_unfilled: return False
        weights = {
            'preference_score': 1.0,
            'hour_imbalance_penalty': 25.0,
            'night_variance': 800.0,
            'weekend_off_variance': 1000.0
        }
        current_score = sum(weights[k] * current_metrics.get(k, 0) for k in weights)
        best_score = sum(weights[k] * best_metrics.get(k, 0) for k in weights)
        return current_score < best_score

    def optimize_schedule(self, year, month, iterations=10):
        best_schedule = None
        best_metrics = {'unfilled_problem_shifts': float('inf'), 'hour_imbalance_penalty': float('inf'), 'night_variance': float('inf'), 'preference_score': float('inf')}
        best_unfilled_info = {}
        self._pre_check_staffing_levels(year, month)
        
        iteration_placeholder = st.empty()
        log_placeholder = st.empty()
        
        for i in range(iterations):
            iteration_placeholder.text(f"--- Running Iteration {i+1}/{iterations} ---")
            current_schedule, unfilled_info = self.generate_monthly_schedule_shuffled(year, month, iteration_num=i+1)
            if unfilled_info['other_days']: continue
            metrics = self.calculate_schedule_metrics(current_schedule, year, month)
            metrics['unfilled_problem_shifts'] = len(unfilled_info['problem_days'])
            
            log_message = (f"Iteration {i+1} Results -> "
                  f"Unfilled (Problem Days): {metrics['unfilled_problem_shifts']} | "
                  f"Hour SD: {metrics.get('hour_diff_for_logging', 0):.2f} | "
                  f"Night Var: {metrics.get('night_variance', 0):.2f} | "
                  f"Pref Penalty: {metrics.get('preference_score', 0):.1f}")
            log_placeholder.info(log_message)

            if best_schedule is None or self.is_schedule_better(metrics, best_metrics):
                best_schedule = current_schedule.copy()
                best_metrics = metrics.copy()
                best_unfilled_info = unfilled_info.copy()
                log_placeholder.success(f"{log_message}\n*** Found a more balanced schedule! ***")

        if best_schedule is not None:
            st.success("Optimization complete! Final metrics for the best schedule found:")
            st.json({
                "Unfilled Shifts (Problem Days)": best_metrics.get('unfilled_problem_shifts', 0),
                "Hour SD": f"{best_metrics.get('hour_diff_for_logging', 0):.2f}",
                "Night Variance": f"{best_metrics.get('night_variance', 0):.2f}",
                "Preference Penalty": f"{best_metrics.get('preference_score', 0):.1f}"
            })
        else:
            st.error("Optimization failed to find any valid schedule.")
        return best_schedule, best_unfilled_info
    
    def export_to_excel(self, schedule, unfilled_info):
        wb = Workbook()
        ws = wb.active
        ws.title = 'Monthly Schedule'
        ws_daily = wb.create_sheet("Daily Summary")
        ws_daily_codes = wb.create_sheet("Daily Summary (Codes)")
        ws_pref = wb.create_sheet("Preference Scores")
        ws_negotiate = wb.create_sheet("Negotiation Suggestions")
        header_fill = PatternFill(start_color='FFD3D3D3', end_color='FFD3D3D3', fill_type='solid')
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        ws.cell(row=1, column=1, value='Date').fill = header_fill
        for col, shift_type in enumerate(self.shift_types, 2):
            cell = ws.cell(row=1, column=col, value=f"{self.shift_types[shift_type]['description']}\n({self.shift_types[shift_type]['hours']} hrs)")
            cell.fill, cell.font, cell.alignment = header_fill, Font(bold=True), Alignment(wrap_text=True)
        schedule.sort_index(inplace=True)
        for row, date in enumerate(schedule.index, 2):
            ws.cell(row=row, column=1, value=date.strftime('%Y-%m-%d'))
            is_holiday = self.is_holiday(date)
            is_weekend = date.weekday() >= 5
            for col, shift_type in enumerate(self.shift_types, 2):
                cell = ws.cell(row=row, column=col, value=schedule.loc[date, shift_type])
                cell.border = border
                if schedule.loc[date, shift_type] == 'NO SHIFT': cell.fill = PatternFill(start_color='FFCCCCCC', fill_type='solid')
                elif is_holiday: cell.fill = PatternFill(start_color='FFFFB6C1', fill_type='solid')
                elif is_weekend: cell.fill = PatternFill(start_color='FFFFE4E1', fill_type='solid')
                elif schedule.loc[date, shift_type] == 'UNFILLED': cell.fill = PatternFill(start_color='FFFFFF00', fill_type='solid')
        ws.column_dimensions['A'].width = 22
        for col in range(2, len(self.shift_types) + 2):
            ws.column_dimensions[get_column_letter(col)].width = 20
        self.create_schedule_summaries(ws, schedule)
        self.create_daily_summary(ws_daily, schedule)
        self.create_preference_score_summary(ws_pref, schedule)
        self.create_daily_summary_with_codes(ws_daily_codes, schedule)
        self.create_negotiation_summary(ws_negotiate, schedule)
        
        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        return buffer

    def create_negotiation_summary(self, ws, schedule):
        header_fill = PatternFill(start_color='FF4F81BD', end_color='FF4F81BD', fill_type='solid')
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        bold_white_font = Font(bold=True, color="FFFFFFFF")
        alignment = Alignment(wrap_text=True, vertical='top')
        headers = ["Date", "Unfilled Shift", "Suggested Negotiation Candidates (Ranked)"]
        for col, header_text in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header_text)
            cell.fill, cell.font, cell.border, cell.alignment = header_fill, bold_white_font, border, alignment
        unfilled_shifts = []
        for date in schedule.index:
            for shift_type, assigned_pharm in schedule.loc[date].items():
                if assigned_pharm == 'UNFILLED':
                    unfilled_shifts.append((date, shift_type))
        if not unfilled_shifts:
            ws.cell(row=2, column=1, value="No unfilled shifts in the final schedule.")
            return
        current_row = 2
        for date, shift_type in unfilled_shifts:
            required_skills = self.shift_types[shift_type].get('required_skills', [])
            all_candidates = []
            for p_name, p_info in self.pharmacists.items():
                if not all(skill.strip() in p_info['skills'] for skill in required_skills if skill.strip()): continue
                if len(self.get_pharmacist_shifts(p_name, date, schedule)) > 0: continue
                is_on_holiday = date.strftime('%Y-%m-%d') in p_info['holidays']
                pharmacist_data = {
                    'name': p_name,
                    'preference_score': self.get_preference_score(p_name, shift_type),
                    'consecutive_days': self.count_consecutive_shifts(p_name, date, schedule),
                    'current_hours': self.calculate_total_hours(p_name, schedule),
                }
                suitability_score = self._calculate_suitability_score(pharmacist_data)
                all_candidates.append({'name': p_name, 'is_on_holiday': is_on_holiday, 'score': suitability_score})
            sorted_candidates = sorted(all_candidates, key=lambda x: (x['is_on_holiday'], x['score']))
            suggestions_text = []
            for i, cand in enumerate(sorted_candidates[:3]):
                status = "(On Holiday)" if cand['is_on_holiday'] else "(Available)"
                suggestions_text.append(f"{i+1}. {cand['name']} {status}")
            final_text = "\n".join(suggestions_text) if suggestions_text else "No suitable candidate found"
            ws.cell(row=current_row, column=1, value=date.strftime('%Y-%m-%d')).border = border
            ws.cell(row=current_row, column=2, value=shift_type).border = border
            cell = ws.cell(row=current_row, column=3, value=final_text)
            cell.border, cell.alignment, ws.row_dimensions[current_row].height = border, alignment, 50
            current_row += 1
        ws.column_dimensions['A'].width, ws.column_dimensions['B'].width, ws.column_dimensions['C'].width = 15, 25, 45

    def create_schedule_summaries(self, ws, schedule):
        summary_row = len(schedule) + 3
        ws.cell(row=summary_row, column=1, value="Summary").font = Font(bold=True)
        hours_row = summary_row + 2
        ws.cell(row=hours_row, column=1, value="Working Hours Summary").font = Font(bold=True)
        for i, pharmacist in enumerate(self.pharmacists):
            hours = self.calculate_total_hours(pharmacist, schedule)
            ws.cell(row=hours_row + i + 1, column=1, value=pharmacist)
            ws.cell(row=hours_row + i + 1, column=2, value=f"Total Hours: {hours}")
        night_row = hours_row + len(self.pharmacists) + 2
        ws.cell(row=night_row, column=1, value="Night Shift Summary").font = Font(bold=True)
        for i, pharmacist in enumerate(self.pharmacists):
            row = night_row + i + 1
            ws.cell(row=row, column=1, value=pharmacist)
            ws.cell(row=row, column=2, value=f"Night Shifts: {self.pharmacists[pharmacist]['night_shift_count']}")
        shift_row = night_row + len(self.pharmacists) + 2
        ws.cell(row=shift_row, column=1, value="Shift Count Summary").font = Font(bold=True)
        shift_types_list = list(self.shift_types.keys())
        for col_idx, shift_type in enumerate(shift_types_list, 2):
            ws.cell(row=shift_row, column=col_idx, value=shift_type).font = Font(bold=True)
        for row_idx, pharmacist in enumerate(self.pharmacists, 1):
            row_num = shift_row + row_idx
            ws.cell(row=row_num, column=1, value=pharmacist)
            for col_idx, shift_type in enumerate(shift_types_list, 2):
                count = sum(1 for d in schedule.index if schedule.loc[d, shift_type] == pharmacist)
                ws.cell(row=row_num, column=col_idx, value=count)

    def _setup_daily_summary_styles(self):
        return {
            'header_fill': PatternFill(fill_type='solid', start_color='FFD3D3D3'),
            'weekend_fill': PatternFill(fill_type='solid', start_color='FFFFE4E1'),
            'holiday_fill': PatternFill(fill_type='solid', start_color='FFFFB6C1'),
            'holiday_empty_fill': PatternFill(fill_type='solid', start_color='FFFFFF00'),
            'off_fill': PatternFill(fill_type='solid', start_color='FFD3D3D3'),
            'border': Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin')),
            'fills': {p: PatternFill(fill_type='solid', start_color=c) for p, c in [
                ('I100', 'FF00B050'), ('O100', 'FF00B0F0'), ('Care', 'FFD40202'), ('C8', 'FFE6B8AF'),
                ('I400', 'FFFF00FF'), ('O400F1', 'FF0033CC'), ('O400F2', 'FFC78AF2'),
                ('O400ER', 'FFED7D31'), ('ARI', 'FF7030A0')]},
            'fonts': {
                'O400F1': Font(bold=True, color="FFFFFFFF"), 'ARI': Font(bold=True, color="FFFFFFFF"),
                'default': Font(bold=True), 'header': Font(bold=True)
            }
        }

    def create_daily_summary(self, ws, schedule):
        styles = self._setup_daily_summary_styles()
        ordered_pharmacists = [
            "ภญ.ประภัสสรา (มิ้น)", "ภญ.ฐิฏิการ (เอ้)", "ภก.บัณฑิตวงศ์ (แพท)", "ภก.ชานนท์ (บุ้ง)", "ภญ.กมลพรรณ (ใบเตย)", "ภญ.กนกพร (นุ้ย)",
            "ภก.เอกวรรณ (โม)", "ภญ.อาภาภัทร (มะปราง)", "ภก.ชวนันท์ (เท่ห์)", "ภญ.ธนพร (ฟ้า ธนพร)", "ภญ.วิลินดา (เชอร์รี่)", "ภญ.ชลนิชา (เฟื่อง)",
            "ภญ.ปริญญ์ (ขมิ้น)", "ภก.ธนภรณ์ (กิ๊ฟ)", "ภญ.ปุณยวีร์ (มิ้นท์)", "ภญ.อมลกานต์ (บอม)", "ภญ.อรรชนา (อ้อม)", "ภญ.ศศิวิมล (ฟิลด์)",
            "ภญ.วรรณิดา (ม่าน)", "ภญ.ปาณิศา (แบม)", "ภญ.จิรัชญา (ศิกานต์)", "ภญ.อภิชญา (น้ำตาล)", "ภญ.วรางคณา (ณา)", "ภญ.ดวงดาว (ปลา)",
            "ภญ.พรนภา (ผึ้ง)", "ภญ.ธนาภรณ์ (ลูกตาล)", "ภญ.วิลาสินี (เจ้นท์)", "ภญ.ภาวิตา (จูน)", "ภญ.ศิรดา (พลอย)", "ภญ.ศุภิสรา (แพร)",
            "ภญ.กันต์หทัย (ซีน)","ภญ.พัทธ์ธีรา (วิว)","ภญ.จุฑามาศ (กวาง)",'ภญ. ณัฐพร (แอม)'
        ]
        ws.cell(row=1, column=1, value='Pharmacist').fill = styles['header_fill']
        sorted_dates = sorted(schedule.index)
        for col, date in enumerate(sorted_dates, 2):
            cell = ws.cell(row=1, column=col, value=date.strftime('%d/%m'))
            cell.fill, cell.font = styles['header_fill'], styles['fonts']['header']
            if date.weekday() >= 5: cell.fill = styles['weekend_fill']
            if self.is_holiday(date): cell.fill = styles['holiday_fill']
        current_row = 2
        for pharmacist in ordered_pharmacists:
            if pharmacist not in self.pharmacists: continue
            ws.cell(row=current_row, column=1, value="").fill = styles['header_fill']
            ws.cell(row=current_row + 1, column=1, value=pharmacist).fill = styles['header_fill']
            ws.cell(row=current_row + 2, column=1, value="").fill = styles['header_fill']
            for col, date in enumerate(sorted_dates, 2):
                note_cell, cell1, cell2 = [ws.cell(row=current_row + r, column=col) for r in range(3)]
                all_cells = [note_cell, cell1, cell2]
                for cell in all_cells:
                    cell.border = styles['border']
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                date_str = date.strftime('%Y-%m-%d')
                note_text = self.special_notes.get(pharmacist, {}).get(date_str)
                shifts = self.get_pharmacist_shifts(pharmacist, date, schedule)
                is_personal_holiday = date_str in self.pharmacists[pharmacist]['holidays']
                is_public_holiday_or_weekend = self.is_holiday(date) or date.weekday() >= 5
                if is_personal_holiday:
                    cell2.value = 'X'
                    cell1.value = None
                    for cell in all_cells:
                        cell.fill = styles['off_fill']
                else:
                    if len(shifts) > 0:
                        shift = shifts[0]
                        cell2.value = f"{int(self.shift_types[shift]['hours'])}N" if self.is_night_shift(shift) else int(self.shift_types[shift]['hours'])
                        prefix = next((p for p in styles['fills'] if shift.startswith(p)), None)
                        if prefix:
                            fill_color = styles['fills'][prefix]
                            cell2.fill, cell2.font = fill_color, styles['fonts'].get(prefix, Font(bold=True))
                            if len(shifts) == 1: cell1.fill = fill_color
                    if len(shifts) > 1:
                        shift = shifts[1]
                        cell1.value = f"{int(self.shift_types[shift]['hours'])}N" if self.is_night_shift(shift) else int(self.shift_types[shift]['hours'])
                        prefix = next((p for p in styles['fills'] if shift.startswith(p)), None)
                        if prefix: cell1.fill, cell1.font = styles['fills'][prefix], styles['fonts'].get(prefix, Font(bold=True))
                    if is_public_holiday_or_weekend:
                        note_cell.fill = styles['holiday_empty_fill']
                        if not shifts:
                            if not cell1.value: cell1.fill = styles['holiday_empty_fill']
                            if not cell2.value: cell2.fill = styles['holiday_empty_fill']
                if note_text:
                    note_cell.value = note_text
                    note_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            current_row += 3
        total_row, unfilled_row = current_row + 1, current_row + 2
        ws.cell(row=total_row, column=1, value="Total Hours").fill = styles['header_fill']
        ws.cell(row=unfilled_row, column=1, value="Unfilled Shifts").fill = styles['header_fill']
        for col, date in enumerate(sorted_dates, 2):
            total_hours = sum(self.shift_types[st]['hours'] for st, p in schedule.loc[date].items() if p not in ['NO SHIFT', 'UNFILLED', 'UNASSIGNED'] and st in self.shift_types)
            unfilled_shifts = [st for st, p in schedule.loc[date].items() if p in ['UNFILLED', 'UNASSIGNED']]
            ws.cell(row=total_row, column=col, value=total_hours).border = styles['border']
            unfilled_cell = ws.cell(row=unfilled_row, column=col)
            unfilled_cell.border = styles['border']
            if unfilled_shifts:
                unfilled_cell.value, unfilled_cell.fill = "\n".join(unfilled_shifts), PatternFill(start_color='FFFFFF00', fill_type='solid')
            else:
                unfilled_cell.value = "0"
        ws.column_dimensions['A'].width = 25
        for col in range(2, len(schedule.index) + 3):
            ws.column_dimensions[get_column_letter(col)].width = 7

    def create_preference_score_summary(self, ws, schedule):
        header_fill = PatternFill(start_color='FFD3D3D3', end_color='FFD3D3D3', fill_type='solid')
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        bold_font = Font(bold=True)
        headers = ["Pharmacist", "Preference Score (%)", "Total Shifts Worked"]
        for col, header_text in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header_text)
            cell.fill, cell.font, cell.border = header_fill, bold_font, border
        preference_scores = self.calculate_pharmacist_preference_scores(schedule)
        pharmacist_list = sorted(self.pharmacists.keys())
        for row, pharmacist in enumerate(pharmacist_list, 2):
            total_shifts = sum(1 for date in schedule.index for p in schedule.loc[date] if p == pharmacist)
            score = preference_scores.get(pharmacist, 0)
            ws.cell(row=row, column=1, value=pharmacist).border = border
            score_cell = ws.cell(row=row, column=2, value=score)
            score_cell.border = border
            score_cell.number_format = '0.00"%"'
            ws.cell(row=row, column=3, value=total_shifts).border = border
        ws.column_dimensions['A'].width, ws.column_dimensions['B'].width, ws.column_dimensions['C'].width = 30, 25, 25

    def create_daily_summary_with_codes(self, ws, schedule):
        styles = self._setup_daily_summary_styles()
        ordered_pharmacists = [
            "ภญ.ประภัสสรา (มิ้น)", "ภญ.ฐิฏิการ (เอ้)", "ภก.บัณฑิตวงศ์ (แพท)", "ภก.ชานนท์ (บุ้ง)", "ภญ.กมลพรรณ (ใบเตย)", "ภญ.กนกพร (นุ้ย)",
            "ภก.เอกวรรณ (โม)", "ภญ.อาภาภัทร (มะปราง)", "ภก.ชวนันท์ (เท่ห์)", "ภญ.ธนพร (ฟ้า ธนพร)", "ภญ.วิลินดา (เชอร์รี่)", "ภญ.ชลนิชา (เฟื่อง)",
            "ภญ.ปริญญ์ (ขมิ้น)", "ภก.ธนภรณ์ (กิ๊ฟ)", "ภญ.ปุณยวีร์ (มิ้นท์)", "ภญ.อมลกานต์ (บอม)", "ภญ.อรรชนา (อ้อม)", "ภญ.ศศิวิมล (ฟิลด์)",
            "ภญ.วรรณิดา (ม่าน)", "ภญ.ปาณิศา (แบม)", "ภญ.จิรัชญา (ศิกานต์)", "ภญ.อภิชญา (น้ำตาล)", "ภญ.วรางคณา (ณา)", "ภญ.ดวงดาว (ปลา)",
            "ภญ.พรนภา (ผึ้ง)", "ภญ.ธนาภรณ์ (ลูกตาล)", "ภญ.วิลาสินี (เจ้นท์)", "ภญ.ภาวิตา (จูน)", "ภญ.ศิรดา (พลอย)", "ภญ.ศุภิสรา (แพร)",
            "ภญ.กันต์หทัย (ซีน)","ภญ.พัทธ์ธีรา (วิว)","ภญ.จุฑามาศ (กวาง)",'ภญ. ณัฐพร (แอม)'
        ]
        ws.cell(row=1, column=1, value='Pharmacist').fill = styles['header_fill']
        sorted_dates = sorted(schedule.index)
        for col, date in enumerate(sorted_dates, 2):
            cell = ws.cell(row=1, column=col, value=date.strftime('%d/%m'))
            cell.fill, cell.font = styles['header_fill'], styles['fonts']['header']
            if date.weekday() >= 5: cell.fill = styles['weekend_fill']
            if self.is_holiday(date): cell.fill = styles['holiday_fill']
        current_row = 2
        for pharmacist in ordered_pharmacists:
            if pharmacist not in self.pharmacists: continue
            ws.cell(row=current_row, column=1, value="").fill = styles['header_fill']
            ws.cell(row=current_row + 1, column=1, value=pharmacist).fill = styles['header_fill']
            ws.cell(row=current_row + 2, column=1, value="").fill = styles['header_fill']
            for col, date in enumerate(sorted_dates, 2):
                note_cell, cell1, cell2 = [ws.cell(row=current_row + r, column=col) for r in range(3)]
                all_cells = [note_cell, cell1, cell2]
                for cell in all_cells:
                    cell.border = styles['border']
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if cell != note_cell: cell.font = Font(bold=True, size=9)
                date_str = date.strftime('%Y-%m-%d')
                note_text = self.special_notes.get(pharmacist, {}).get(date_str)
                shifts = self.get_pharmacist_shifts(pharmacist, date, schedule)
                is_personal_holiday = date_str in self.pharmacists[pharmacist]['holidays']
                is_public_holiday_or_weekend = self.is_holiday(date) or date.weekday() >= 5
                if is_personal_holiday:
                    cell2.value = 'OFF'
                    cell1.value = None
                    for cell in all_cells:
                        cell.fill = styles['off_fill']
                else:
                    if len(shifts) > 0:
                        shift_code, cell = shifts[0], cell2
                        cell.value = shift_code
                        prefix = next((p for p in styles['fills'] if shift_code.startswith(p)), None)
                        if prefix:
                            fill_color = styles['fills'][prefix]
                            cell.fill, cell.font = fill_color, styles['fonts'].get(prefix, styles['fonts']['default'])
                            if len(shifts) == 1: cell1.fill = fill_color
                    if len(shifts) > 1:
                        shift_code, cell = shifts[1], cell1
                        cell.value = shift_code
                        prefix = next((p for p in styles['fills'] if shift_code.startswith(p)), None)
                        if prefix: cell.fill, cell.font = styles['fills'][prefix], styles['fonts'].get(prefix, styles['fonts']['default'])
                    if is_public_holiday_or_weekend:
                        note_cell.fill = styles['holiday_empty_fill']
                        if not shifts:
                            if not cell1.value: cell1.fill = styles['holiday_empty_fill']
                            if not cell2.value: cell2.fill = styles['holiday_empty_fill']
                if note_text:
                    note_cell.value = note_text
                    note_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            current_row += 3
        total_row, unfilled_row = current_row + 1, current_row + 2
        ws.cell(row=total_row, column=1, value="Total Hours").fill = styles['header_fill']
        ws.cell(row=unfilled_row, column=1, value="Unfilled Shifts").fill = styles['header_fill']
        for col, date in enumerate(sorted_dates, 2):
            total_hours = sum(self.shift_types[st]['hours'] for st, p in schedule.loc[date].items() if p not in ['NO SHIFT', 'UNFILLED', 'UNASSIGNED'] and st in self.shift_types)
            unfilled_shifts = [st for st, p in schedule.loc[date].items() if p in ['UNFILLED', 'UNASSIGNED']]
            ws.cell(row=total_row, column=col, value=total_hours).border = styles['border']
            unfilled_cell = ws.cell(row=unfilled_row, column=col)
            unfilled_cell.border = styles['border']
            if unfilled_shifts:
                unfilled_cell.value, unfilled_cell.fill = "\n".join(unfilled_shifts), PatternFill(start_color='FFFFFF00', fill_type='solid')
            else:
                unfilled_cell.value = "0"
        ws.column_dimensions['A'].width = 25
        for col in range(2, len(schedule.index) + 2):
            ws.column_dimensions[get_column_letter(col)].width = 15

    def calculate_pharmacist_preference_scores(self, schedule):
        scores = {}
        MAX_POINTS_PER_SHIFT = 8
        for pharmacist in self.pharmacists:
            total_achieved_points = 0
            total_shifts_worked = 0
            for date in schedule.index:
                for shift_type, assigned_pharm in schedule.loc[date].items():
                    if assigned_pharm == pharmacist:
                        total_shifts_worked += 1
                        rank = self.get_preference_score(pharmacist, shift_type)
                        points = max(0, 9 - rank)
                        total_achieved_points += points
            if total_shifts_worked == 0:
                scores[pharmacist] = 0
            else:
                max_possible_points = total_shifts_worked * MAX_POINTS_PER_SHIFT
                if max_possible_points == 0:
                    scores[pharmacist] = 0
                else:
                    percentage_score = (total_achieved_points / max_possible_points) * 100
                    scores[pharmacist] = percentage_score
        return scores

# =========================================================================
# ================== STREAMLIT APPLICATION UI =============================
# =========================================================================

st.set_page_config(page_title="Pharmacist Scheduler", layout="wide")
st.title("👩‍⚕️ Pharmacist Shift Scheduler")
st.info("ตั้งค่าเดือนและปีที่ต้องการจัดตารางในแถบด้านข้าง แล้วกดปุ่ม 'Generate Schedule'")


# --- Sidebar for Inputs ---
with st.sidebar:
    st.header("⚙️ Configuration")
    
    # รับค่าปี เดือน และจำนวน Iterations จากผู้ใช้
    current_year = datetime.now().year
    year = st.number_input("Year", min_value=current_year, max_value=current_year + 5, value=current_year)
    month = st.selectbox("Month", options=range(1, 13), format_func=lambda x: datetime(year, x, 1).strftime("%B"), index=datetime.now().month -1)
    iterations = st.slider("Optimization Iterations", min_value=1, max_value=500, value=50, help="ยิ่งค่าสูง อาจจะได้ตารางที่ดีขึ้น แต่ใช้เวลาประมวลผลนานขึ้น")

    generate_button = st.button("Generate Schedule", type="primary", use_container_width=True)


# --- Main Area for Outputs ---
if generate_button:
    try:
        with st.spinner(f"กำลังโหลดข้อมูล Excel จาก GitHub และจัดตารางสำหรับเดือน {datetime(year, month, 1).strftime('%B %Y')}..."):
            
            # 1. ระบุ Raw URL ของไฟล์ Excel บน GitHub
            # ⚠️⚠️⚠️ สำคัญ: เปลี่ยน URL นี้เป็น Raw URL ของไฟล์ Excel ของคุณ ⚠️⚠️⚠️
            excel_url = "https://raw.githubusercontent.com/Drugpurchasing/Shift/main/pharmacist_schedule.xlsx"

            st.write("Reading data from Excel file on GitHub...")

            # 2. อ่านทุกชีตจากไฟล์ Excel มาเก็บใน Dictionary
            # ใช้ try-except เพื่อดักจับข้อผิดพลาดหากไม่สามารถโหลดไฟล์ได้ (เช่น URL ผิด, ไฟล์ไม่มีอยู่)
            try:
                # การตั้งค่า sheet_name=None จะทำให้ pandas อ่านทุกชีตในไฟล์มาเป็น dict
                all_sheets_dict = pd.read_excel(excel_url, sheet_name=None)
                st.success("Successfully read data from GitHub Excel file!")
            except Exception as e:
                st.error(f"ไม่สามารถอ่านไฟล์ Excel จาก GitHub ได้ กรุณาตรวจสอบ Raw URL อีกครั้ง. Error: {e}")
                st.stop() # หยุดการทำงานของแอปถ้าโหลดไฟล์ไม่ได้

            # 3. ฟังก์ชันสำหรับดึง DataFrame ของแต่ละชีตออกจาก dict ที่โหลดมา
            #    พร้อมแจ้งเตือนหากไม่มีชีตนั้นๆ (สำหรับชีตที่ไม่บังคับ)
            def get_df_from_dict(sheet_name):
                df = all_sheets_dict.get(sheet_name)
                if df is None:
                    # แจ้งเตือนสำหรับชีตที่ไม่บังคับ (Optional)
                    st.warning(f"Sheet '{sheet_name}' not found in the Excel file. Proceeding without it.")
                return df

            # === ดึงข้อมูลแต่ละชีตออกมาใส่ตัวแปร ===
            pharmacists_df = get_df_from_dict("Pharmacists")
            shifts_df = get_df_from_dict("Shifts")
            departments_df = get_df_from_dict("Departments")
            pre_assignments_df = get_df_from_dict("PreAssignments")
            
            # ชีตที่ไม่บังคับ (Optional Sheets)
            historical_scores_df = get_df_from_dict("HistoricalScores")
            special_notes_df = get_df_from_dict("SpecialNotes")
            shift_limits_df = get_df_from_dict("ShiftLimits")
            holiday_df = get_df_from_dict("Holiday")
            holiday2_df = get_df_from_dict("Holiday 2") # ชื่อชีตอาจมีเว้นวรรค
            prefre_df = get_df_from_dict("Prefre")
            
            # === ตรวจสอบว่าชีตที่จำเป็น (Required) มีอยู่จริงหรือไม่ ===
            required_sheets = {
                "Pharmacists": pharmacists_df, 
                "Shifts": shifts_df, 
                "Departments": departments_df, 
                "PreAssignments": pre_assignments_df
            }
            missing_sheets = [name for name, df in required_sheets.items() if df is None]
            if missing_sheets:
                st.error(f"Required sheets are missing from the Excel file: {', '.join(missing_sheets)}")
                st.stop()

            # 4. เตรียมข้อมูลทั้งหมดใส่ dict เพื่อส่งให้ Scheduler (โครงสร้างเหมือนเดิม)
            all_dataframes = {
                "pharmacists": pharmacists_df,
                "shifts": shifts_df,
                "departments": departments_df,
                "pre_assignments": pre_assignments_df,
                "historical_scores": historical_scores_df,
                "special_notes": special_notes_df,
                "shift_limits": shift_limits_df,
                "holiday": holiday_df,
                "holiday2": holiday2_df,
                "prefre": prefre_df,
            }

            # 5. เริ่มต้น Scheduler และรันการ Optimize (เหมือนเดิม)
            scheduler = PharmacistScheduler(dataframes=all_dataframes)
            best_schedule, best_unfilled_info = scheduler.optimize_schedule(year, month, iterations)

        # (ส่วนที่เหลือของโค้ดในการแสดงผลลัพธ์เหมือนเดิมทั้งหมด)
        if best_schedule is not None:
            st.header("✅ Optimization Complete")
            
            # เก็บผลลัพธ์ไว้ใน session state เพื่อให้สามารถดาวน์โหลดได้
            st.session_state['best_schedule'] = best_schedule
            st.session_state['best_unfilled_info'] = best_unfilled_info
            st.session_state['scheduler_instance'] = scheduler
            st.session_state['output_filename'] = f'Pharmacist_Schedule_{year}_{month}.xlsx'
        else:
            st.error("ไม่สามารถสร้างตารางที่เหมาะสมได้ กรุณาตรวจสอบข้อมูลในไฟล์ Excel")

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดระหว่างการสร้างตาราง: {e}")
        st.exception(e)

# แสดงผลลัพธ์ถ้ามีข้อมูลอยู่ใน session state
if 'best_schedule' in st.session_state:
    st.header("📊 Results")
    
    # สร้างไฟล์ Excel ในหน่วยความจำเพื่อรอการดาวน์โหลด
    excel_buffer = st.session_state['scheduler_instance'].export_to_excel(
        st.session_state['best_schedule'],
        st.session_state['best_unfilled_info']
    )
    
    st.download_button(
        label="📥 Download Full Schedule (Excel)",
        data=excel_buffer,
        file_name=st.session_state['output_filename'],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    st.subheader("Generated Schedule Preview")
    st.dataframe(st.session_state['best_schedule'])













