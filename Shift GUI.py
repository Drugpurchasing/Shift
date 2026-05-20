import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, time
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter
import random
from statistics import stdev
import io
import os

# ตั้งค่าหน้าตาของ Streamlit App ให้เป็นแบบกว้าง (Wide Layout)
st.set_page_config(
    page_title="Intelligent Pharmacy Scheduling Support System",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# คอนฟิกูเรชันพื้นฐานของชนิดเวร
SCHEDULE_SOURCES = {
    "1": {
        "label": "เภสัชกร",
        "employee_sheet_name": "employee",
        "output_prefix": "Pharmacist_Schedule",
    },
    "2": {
        "label": "ผู้ช่วยเภสัชกร",
        "employee_sheet_name": "employee",
        "output_prefix": "Assistant_Pharmacist_Schedule",
    },
}

class PharmacistScheduler:
    """
    Pharmacy shift scheduler with optimization and Excel export.
    Designed for multi-constraint pharmacist roster planning.
    """
    W_CONSECUTIVE = 10
    W_HOURS = 4
    W_PREFERENCE = 6
    MAX_CONSECUTIVE_DAYS = 3

    MIN_WEEKEND_OFF_DAYS = 4
    W_SHIFT_PACING = 900
    W_MONTH_SEGMENT_BALANCE = 180
    W_WEEKEND_OFF_PROTECTION = 2200

    def __init__(self, excel_file, employee_sheet_name='employee', staff_type='เภสัชกร'):
        self.pharmacists = {}
        self.shift_types = {}
        self.departments = {}
        self.pre_assignments = {}
        self.historical_scores = {}
        self.preference_multipliers = {}
        self.special_notes = {}
        self.shift_limits = {}
        self.employee_sheet_name = employee_sheet_name
        self.staff_type = staff_type
        self.employee_order = []
        self.no_preference_staff = set()
        self.excel_file = excel_file
        self.problem_days = set()
        self.run_logs = []
        self.run_config = {}
        self.soul_mates = {
            'ภก.ชานนท์ (บุ้ง)': 'ภญ.อาภาภัทร (มะปราง)'
        }

        self.read_data_from_excel(self.excel_file)
        self.load_historical_scores()
        self._calculate_preference_multipliers()

        self.night_shifts = {
            'I100-10', 'I100-12N', 'I400-12N', 'I400-10', 'O400ER-12N', 'O400ER-10'
        }
        self.holidays = {
            'specific_dates': ['2026-06-03','2026-06-01']
        }
        for pharmacist in self.pharmacists:
            self.pharmacists[pharmacist]['shift_counts'] = {
                shift_type: 0 for shift_type in self.shift_types
            }

    def _pre_check_staffing_levels(self, year, month):
        start_date = datetime(year, month, 1)
        end_date = datetime(year, month + 1, 1) - timedelta(days=1) if month < 12 else datetime(year, 12, 31)
        dates = pd.date_range(start_date, end_date)

        all_ok = True
        for date in dates:
            available_pharmacists_count = sum(1 for p_name, p_info in self.pharmacists.items()
                                              if date.strftime('%Y-%m-%d') not in p_info['holidays'])
            required_shifts_base = sum(1 for st in self.shift_types
                                       if self.is_shift_available_on_date(st, date))
            total_required_shifts_with_buffer = required_shifts_base + 5

            if available_pharmacists_count < total_required_shifts_with_buffer:
                all_ok = False
                self.problem_days.add(date)
        return not all_ok

    def load_historical_scores(self):
        try:
            df = pd.read_excel(self.excel_file, sheet_name='HistoricalScores')
            if 'Pharmacist' in df.columns and 'Total Preference Score' in df.columns:
                for _, row in df.iterrows():
                    pharmacist = row['Pharmacist']
                    score = row['Total Preference Score']
                    if pharmacist in self.pharmacists:
                        self.historical_scores[pharmacist] = score
        except Exception:
            pass

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

    def _normalize_preference_value(self, value):
        if pd.isna(value): return None
        cleaned = str(value).strip()
        if cleaned == "" or cleaned.lower() == "none" or cleaned.lower() == "nan": return None
        return cleaned

    def _has_no_preferences(self, preferences):
        return all(self._normalize_preference_value(value) is None for value in preferences.values())

    def get_ordered_employees(self):
        ordered = [name for name in self.employee_order if name in self.pharmacists]
        missing = [name for name in self.pharmacists if name not in ordered]
        return ordered + missing

    def get_total_shift_count_in_schedule_dict(self, pharmacist, schedule_dict):
        count = 0
        for shifts in schedule_dict.values():
            for assigned in shifts.values():
                if assigned == pharmacist: count += 1
        return count

    def get_unique_departments_worked(self, pharmacist, schedule_dict):
        departments = set()
        for shifts in schedule_dict.values():
            for shift_type, assigned in shifts.items():
                if assigned == pharmacist:
                    department = self.get_department_from_shift(shift_type)
                    if department: departments.add(department)
        return departments

    def get_average_monthly_shift_target(self, year, month):
        cache_key = (year, month)
        if not hasattr(self, '_monthly_shift_target_cache'):
            self._monthly_shift_target_cache = {}
        if cache_key in self._monthly_shift_target_cache:
            return self._monthly_shift_target_cache[cache_key]

        start_date = datetime(year, month, 1)
        end_date = datetime(year + 1, 1, 1) - timedelta(days=1) if month == 12 else datetime(year, month + 1, 1) - timedelta(days=1)
        total_open_shifts = 0
        for date in pd.date_range(start_date, end_date):
            for shift_type in self.shift_types:
                if self.is_shift_available_on_date(shift_type, date): total_open_shifts += 1

        staff_count = max(len(self.pharmacists), 1)
        target = max(total_open_shifts / staff_count, 1)
        self._monthly_shift_target_cache[cache_key] = target
        return target

    def get_month_segment(self, date):
        days_in_month = pd.Timestamp(date).days_in_month
        if date.day <= days_in_month / 3: return 'early'
        if date.day <= (2 * days_in_month) / 3: return 'middle'
        return 'late'

    def get_month_segment_shift_count(self, pharmacist, date, schedule_dict):
        target_segment = self.get_month_segment(date)
        count = 0
        for d, shifts in schedule_dict.items():
            if self.get_month_segment(d) != target_segment: continue
            if pharmacist in shifts.values(): count += 1
        return count

    def get_total_weekend_days_in_schedule(self, schedule_dict):
        return sum(1 for d in schedule_dict if d.weekday() >= 5)

    def get_weekend_days_worked_until(self, pharmacist, schedule_dict):
        count = 0
        for d, shifts in schedule_dict.items():
            if d.weekday() >= 5 and pharmacist in shifts.values(): count += 1
        return count

    def calculate_weekend_min_off_violations(self, schedule):
        weekend_dates = [d for d in schedule.index if d.weekday() >= 5]
        violations = {}
        for pharmacist in self.pharmacists:
            off_count = 0
            for d in weekend_dates:
                if pharmacist not in schedule.loc[d].values: off_count += 1
            shortfall = max(0, self.MIN_WEEKEND_OFF_DAYS - off_count)
            if shortfall > 0: violations[pharmacist] = shortfall
        return violations

    def calculate_month_segment_variance(self, schedule):
        variances = []
        for pharmacist in self.pharmacists:
            counts = {'early': 0, 'middle': 0, 'late': 0}
            for d in schedule.index:
                if pharmacist in schedule.loc[d].values:
                    counts[self.get_month_segment(d)] += 1
            variances.append(np.var(list(counts.values())))
        return float(np.mean(variances)) if variances else 0

    def read_data_from_excel(self, file_path):
        skill_groups_map = {}
        try:
            subset_df = pd.read_excel(file_path, sheet_name='Skill subset')
            if 'Group Name' in subset_df.columns and 'Skills' in subset_df.columns:
                for _, row in subset_df.iterrows():
                    group_name = str(row['Group Name']).strip()
                    if group_name and group_name != 'nan':
                        skills_list = [s.strip() for s in str(row['Skills']).split(',') if s.strip() and s.strip() != 'nan']
                        skill_groups_map[group_name] = skills_list
        except Exception:
            pass

        pharmacists_df = pd.read_excel(file_path, sheet_name=self.employee_sheet_name)
        self.pharmacists = {}
        self.employee_order = []
        self.no_preference_staff = set()
        for _, row in pharmacists_df.iterrows():
            name = str(row['Name']).strip()
            if not name or name.lower() == 'nan': continue
            self.employee_order.append(name)
            max_hours = row.get('Max Hours', 250)
            if pd.isna(max_hours) or max_hours == '' or max_hours is None: max_hours = 250
            else: max_hours = float(max_hours)

            raw_skills = str(row['Skills']).split(',')
            expanded_skills = set()
            for s in raw_skills:
                s_clean = s.strip()
                if s_clean in skill_groups_map: expanded_skills.update(skill_groups_map[s_clean])
                elif s_clean and s_clean != 'nan': expanded_skills.add(s_clean)

            preferences = {f'rank{i}': self._normalize_preference_value(row.get(f'Rank{i}')) for i in range(1, 9)}
            no_preference = self._has_no_preferences(preferences)
            if no_preference: self.no_preference_staff.add(name)

            self.pharmacists[name] = {
                'night_shift_count': 0,
                'skills': list(expanded_skills),
                'holidays': [date for date in str(row.get('Holidays', '')).split(',') if date != '1900-01-00' and date.strip() and date != 'nan'],
                'shift_counts': {},
                'preferences': preferences,
                'no_preference': no_preference,
                'max_hours': max_hours
            }

        shifts_df = pd.read_excel(file_path, sheet_name='Shifts')
        self.shift_types = {}
        for _, row in shifts_df.iterrows():
            shift_code = row['Shift Code']
            self.shift_types[shift_code] = {
                'description': row['Description'],
                'shift_type': row['Shift Type'],
                'start_time': row['Start Time'],
                'end_time': row['End Time'],
                'hours': row['Hours'],
                'required_skills': row['Required Skills'].split(','),
                'restricted_next_shifts': row['Restricted Next Shifts'].split(',') if pd.notna(row['Restricted Next Shifts']) else [],
            }

        departments_df = pd.read_excel(file_path, sheet_name='Departments')
        self.departments = {}
        for _, row in departments_df.iterrows():
            department = row['Department']
            self.departments[department] = row['Shift Codes'].split(',')

        pre_assign_df = pd.read_excel(file_path, sheet_name='PreAssignments')
        pre_assign_df['Date'] = pd.to_datetime(pre_assign_df['Date']).dt.strftime('%Y-%m-%d')
        self.pre_assignments = {}
        for pharmacist, group in pre_assign_df.groupby('Pharmacist'):
            date_dict = {}
            for date, g in group.groupby('Date'):
                shifts = []
                for shift_str in g['Shift']:
                    shifts.extend([s.strip() for s in str(shift_str).split(',') if s.strip()])
                date_dict[date] = shifts
            self.pre_assignments[pharmacist] = date_dict

        try:
            notes_df = pd.read_excel(file_path, sheet_name='SpecialNotes', index_col=0)
            for pharmacist, row_data in notes_df.iterrows():
                if pharmacist in self.pharmacists:
                    for date_col, note in row_data.items():
                        if pd.notna(note) and str(note).strip():
                            date_str = pd.to_datetime(date_col).strftime('%Y-%m-%d')
                            if pharmacist not in self.special_notes: self.special_notes[pharmacist] = {}
                            self.special_notes[pharmacist][date_str] = str(note).strip()
        except Exception:
            pass

        try:
            limits_df = pd.read_excel(file_path, sheet_name='ShiftLimits')
            for _, row in limits_df.iterrows():
                pharmacist = row['Pharmacist']
                category = row['ShiftCategory']
                max_count = row['MaxCount']
                if pharmacist in self.pharmacists:
                    if pharmacist not in self.shift_limits: self.shift_limits[pharmacist] = {}
                    self.shift_limits[pharmacist][category] = int(max_count)
        except Exception:
            pass

        self.min_shift_requirements = {}
        try:
            min_req_df = pd.read_excel(file_path, sheet_name='MinShiftRequirements')
            for _, row in min_req_df.iterrows():
                pharmacist = str(row['Pharmacist']).strip()
                department = str(row['Department']).strip()
                min_count  = int(row['MinCount'])
                if pharmacist in self.pharmacists:
                    if pharmacist not in self.min_shift_requirements: self.min_shift_requirements[pharmacist] = {}
                    self.min_shift_requirements[pharmacist][department] = min_count
        except Exception:
            self.min_shift_requirements = {}

    def convert_time_to_minutes(self, time_input):
        if isinstance(time_input, str): hours, minutes = map(int, time_input.split(':'))
        elif isinstance(time_input, time): hours, minutes = time_input.hour, time_input.minute
        else: raise ValueError("Invalid input type.")
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
        mixing_shifts = [p for s, p in schedule_dict[date].items() if s.startswith('C8') and p not in ['NO SHIFT', 'UNASSIGNED', 'UNFILLED']]
        if current_shift and current_shift.startswith('C8') and current_pharm: mixing_shifts.append(current_pharm)
        if not mixing_shifts: return True
        total_mixing = len(mixing_shifts)
        expert_count = sum(1 for pharm in mixing_shifts if pharm in self.pharmacists and 'mixing_expert' in self.pharmacists[pharm]['skills'])
        return expert_count >= (2 * total_mixing / 3)

    def count_consecutive_shifts(self, pharmacist, date, schedule, max_days=6):
        count = 0
        current_date = date - timedelta(days=1)
        for _ in range(max_days):
            if current_date in schedule.index and pharmacist in schedule.loc[current_date].values:
                count += 1
                current_date -= timedelta(days=1)
            else: break
        return count

    def is_holiday(self, date):
        return date.strftime('%Y-%m-%d') in self.holidays['specific_dates']

    def get_dept_shift_count(self, pharmacist, department, schedule_dict):
        count = 0
        for date, shifts in schedule_dict.items():
            for shift_type, assigned in shifts.items():
                if assigned == pharmacist:
                    if self.get_department_from_shift(shift_type) == department: count += 1
        return count

    def _needs_min_shift(self, pharmacist, shift_type, schedule_dict):
        dept = self.get_department_from_shift(shift_type)
        if not dept: return False
        min_req = self.min_shift_requirements.get(pharmacist, {}).get(dept, 0)
        if min_req == 0: return False
        current_count = self.get_dept_shift_count(pharmacist, dept, schedule_dict)
        return current_count < min_req

    def calculate_weekend_off_variance(self, schedule, year, month):
        weekend_off_counts = {p: 0 for p in self.pharmacists}
        start_date = datetime(year, month, 1)
        end_date = datetime(year, month + 1, 1) - timedelta(days=1) if month < 12 else datetime(year, 12, 31)
        for date in pd.date_range(start_date, end_date):
            if date.weekday() >= 5:
                working_on_weekend = {schedule.loc[date, shift] for shift in schedule.columns if schedule.loc[date, shift] not in ['NO SHIFT', 'UNFILLED', 'UNASSIGNED']}
                for p_name in self.pharmacists:
                    if p_name not in working_on_weekend: weekend_off_counts[p_name] += 1
        if len(weekend_off_counts) > 1: return np.var(list(weekend_off_counts.values()))
        return 0

    def is_night_shift(self, shift_type):
        return shift_type in self.night_shifts

    def _get_holiday_blocks(self, year, month):
        start_date = datetime(year, month, 1)
        end_date = datetime(year, month + 1, 1) - timedelta(days=1) if month < 12 else datetime(year, 12, 31)
        specific_holidays = set(datetime.strptime(d, '%Y-%m-%d') for d in self.holidays['specific_dates'])

        all_off_days = set()
        current = start_date - timedelta(days=7)
        scan_end = end_date + timedelta(days=7)
        d = current
        while d <= scan_end:
            if d.weekday() >= 5 or d in specific_holidays: all_off_days.add(d)
            d += timedelta(days=1)

        sorted_days = sorted(all_off_days)
        blocks = []
        if not sorted_days: return {}

        current_block = [sorted_days[0]]
        for day in sorted_days[1:]:
            if (day - current_block[-1]).days == 1: current_block.append(day)
            else:
                blocks.append(current_block)
                current_block = [day]
        blocks.append(current_block)

        result = {}
        for block in blocks:
            has_weekend = any(d.weekday() >= 5 for d in block)
            if has_weekend and len(block) >= 1:
                last_day = block[-1]
                result[last_day] = set(block)
        return result

    def is_shift_available_on_date(self, shift_type, date):
        shift_info = self.shift_types[shift_type]
        is_holiday_date = self.is_holiday(date)
        weekday_num = date.weekday()
        is_saturday = (weekday_num == 5)
        is_sunday   = (weekday_num == 6)
        is_mon_to_thu = (0 <= weekday_num <= 3)
        s_type = str(shift_info['shift_type']).strip().lower()

        if s_type == 'weekday': return not (is_holiday_date or is_saturday or is_sunday)
        elif s_type == 'saturday': return is_saturday and not is_holiday_date
        elif s_type == 'sat-holiday': return is_saturday or is_holiday_date
        elif s_type == 'sunday': return is_sunday and not is_holiday_date
        elif s_type == 'weekend': return is_saturday or is_sunday or is_holiday_date
        elif s_type == 'sun-holiday': return is_sunday and is_holiday_date
        elif s_type in ['mon-thu', 'วันจันทร์-พฤหัส', 'วันจันทร์-พฤหัสบดี']: return is_mon_to_thu and not is_holiday_date
        elif s_type == 'holiday': return is_holiday_date and not is_saturday and not is_sunday
        elif s_type == 'last-day-holiday': return date in self._get_holiday_blocks(date.year, date.month)
        elif s_type == 'night': return True
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
        elif shift_type.startswith('Refill'): return 'Refill'
        return None

    def _get_shift_category(self, shift_type):
        if self.is_night_shift(shift_type): return 'Night'
        if shift_type.startswith('C8'): return 'Mixing'
        return None

    def get_preference_score(self, pharmacist, shift_type):
        p_skills = [skill.strip().lower() for skill in self.pharmacists[pharmacist]['skills']]
        if 'junior' in p_skills: return 1
        if self.pharmacists[pharmacist].get('no_preference', False): return 5
        department = self.get_department_from_shift(shift_type)
        for rank in range(1, 9):
            if self.pharmacists[pharmacist]['preferences'].get(f'rank{rank}') == department: return rank
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
                if self.check_time_overlap(new_start, new_end, existing_start, existing_end): return True
        return False

    def has_nearby_night_shift_optimized(self, pharmacist, date, schedule_dict):
        for delta in [-2, -1, 1, 2]:
            check_date = date + timedelta(days=delta)
            if check_date in schedule_dict:
                for shift, assigned_pharm in schedule_dict[check_date].items():
                    if assigned_pharm == pharmacist and self.is_night_shift(shift): return True
        return False

    def get_pharmacist_shifts(self, pharmacist, date, current_schedule):
        shifts = []
        if date in current_schedule.index:
            for shift_type in current_schedule.columns:
                if current_schedule.loc[date, shift_type] == pharmacist: shifts.append(shift_type)
        return shifts

    def calculate_total_hours(self, pharmacist, schedule):
        total_hours = 0
        for date in schedule.index:
            for shift_type, assigned_pharm in schedule.loc[date].items():
                if assigned_pharm == pharmacist and shift_type in self.shift_types:
                    total_hours += self.shift_types[shift_type]['hours']
        return total_hours

    def _get_hour_imbalance_penalty(self, hours_dict):
        if not hours_dict or len(hours_dict) < 2: return 0
        hour_values = list(hours_dict.values())
        hour_stdev = stdev(hour_values)
        hour_range = max(hour_values) - min(hour_values)
        stdev_penalty = hour_stdev ** 2
        range_penalty = (hour_range - 10) ** 2 if hour_range > 10 else 0
        return stdev_penalty + range_penalty

    def calculate_schedule_metrics(self, schedule, year, month):
        hours = {p: self.calculate_total_hours(p, schedule) for p in self.pharmacists}
        night_counts = {p: self.pharmacists[p]['night_shift_count'] for p in self.pharmacists}
        weekend_off_var = self.calculate_weekend_off_variance(schedule, year, month)
        hour_penalty = self._get_hour_imbalance_penalty(hours)

        pref_percentages = self.calculate_pharmacist_preference_scores(schedule)
        pref_variance = np.var(list(pref_percentages.values())) if len(pref_percentages) > 1 else 0
        weekend_min_off_violations = self.calculate_weekend_min_off_violations(schedule)

        metrics = {
            'hour_imbalance_penalty': hour_penalty,
            'night_variance': np.var(list(night_counts.values())) if night_counts else 0,
            'preference_score': sum(self.calculate_preference_penalty(p, schedule) for p in self.pharmacists),
            'preference_variance': pref_variance,
            'weekend_off_variance': weekend_off_var,
            'weekend_min_off_shortfall': sum(weekend_min_off_violations.values()),
            'month_segment_variance': self.calculate_month_segment_variance(schedule),
        }
        metrics['hour_diff_for_logging'] = stdev(hours.values()) if len(hours) > 1 else 0
        return metrics

    def _log_schedule_event(self, event_type, message, **kwargs):
        log_row = {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Event Type": event_type,
            "Message": message,
        }
        for key, value in kwargs.items(): log_row[key] = value
        self.run_logs.append(log_row)

    def _reset_runtime_shift_counters(self):
        for pharmacist in self.pharmacists:
            self.pharmacists[pharmacist]['night_shift_count'] = 0
            self.pharmacists[pharmacist]['mixing_shift_count'] = 0
            self.pharmacists[pharmacist]['care_shift_count'] = 0
            self.pharmacists[pharmacist]['category_counts'] = {'Mixing': 0, 'Night': 0}

    def _has_required_skills_for_shift(self, pharmacist, shift_type):
        p_skills = {str(skill).strip().lower() for skill in self.pharmacists.get(pharmacist, {}).get('skills', []) if str(skill).strip() and str(skill).strip().lower() != 'nan'}
        required_skills = {str(skill).strip().lower() for skill in self.shift_types.get(shift_type, {}).get('required_skills', []) if str(skill).strip() and str(skill).strip().lower() != 'nan'}
        return required_skills.issubset(p_skills)

    def _get_true_random_candidates(self, staff_pool, date, shift_type, schedule_dict):
        date_str = pd.to_datetime(date).strftime('%Y-%m-%d')
        candidates = []
        for pharmacist in staff_pool:
            if pharmacist not in self.pharmacists: continue
            if pharmacist in schedule_dict[date].values(): continue
            if date_str in self.pharmacists[pharmacist].get('holidays', []): continue
            if not self._has_required_skills_for_shift(pharmacist, shift_type): continue
            candidates.append(pharmacist)
        return candidates

    def _select_fair_random_candidate(self, candidates, pharmacist_hours, fairness_buffer_hours=8):
        if not candidates: return None, []
        min_hours = min(pharmacist_hours.get(p, 0) for p in candidates)
        fair_pool = [p for p in candidates if pharmacist_hours.get(p, 0) <= min_hours + fairness_buffer_hours]
        if not fair_pool: fair_pool = candidates
        return random.choice(fair_pool), fair_pool

    def generate_monthly_schedule_true_random(self, year, month, iteration_num=1):
        self._reset_runtime_shift_counters()
        start_date = datetime(year, month, 1)
        end_date = datetime(year + 1, 1, 1) - timedelta(days=1) if month == 12 else datetime(year, month + 1, 1) - timedelta(days=1)
        dates = pd.date_range(start_date, end_date)

        schedule_dict = {date: {shift: 'NO SHIFT' for shift in self.shift_types} for date in dates}
        pharmacist_hours = {p: 0 for p in self.pharmacists}
        staff_pool = list(self.pharmacists.keys())
        unfilled_info = {'problem_days': [], 'other_days': []}

        for pharmacist, assignments in self.pre_assignments.items():
            if pharmacist not in self.pharmacists: continue
            for date_str, shift_types in assignments.items():
                date = pd.to_datetime(date_str)
                if date not in schedule_dict: continue
                for shift_type in shift_types:
                    if shift_type not in self.shift_types: continue
                    schedule_dict[date][shift_type] = pharmacist
                    self._update_shift_counts(pharmacist, shift_type)
                    pharmacist_hours[pharmacist] += self.shift_types[shift_type]['hours']

        for date in dates:
            shifts_to_process = list(self.shift_types.keys())
            random.shuffle(shifts_to_process)
            for shift_type in shifts_to_process:
                if not self.is_shift_available_on_date(shift_type, date):
                    schedule_dict[date][shift_type] = 'NO SHIFT'
                    continue
                if schedule_dict[date][shift_type] not in ['NO SHIFT', 'UNASSIGNED', 'UNFILLED']: continue

                candidates = self._get_true_random_candidates(staff_pool, date, shift_type, schedule_dict)
                if not candidates:
                    schedule_dict[date][shift_type] = 'UNFILLED'
                    if date in self.problem_days: unfilled_info['problem_days'].append((date, shift_type))
                    else: unfilled_info['other_days'].append((date, shift_type))
                    continue

                chosen, fair_pool = self._select_fair_random_candidate(candidates, pharmacist_hours, fairness_buffer_hours=8)
                if chosen is None:
                    schedule_dict[date][shift_type] = 'UNFILLED'
                    if date in self.problem_days: unfilled_info['problem_days'].append((date, shift_type))
                    else: unfilled_info['other_days'].append((date, shift_type))
                    continue

                schedule_dict[date][shift_type] = chosen
                self._update_shift_counts(chosen, shift_type)
                pharmacist_hours[chosen] += self.shift_types[shift_type]['hours']

        final_schedule = pd.DataFrame.from_dict(schedule_dict, orient='index')
        final_schedule = final_schedule.reindex(columns=list(self.shift_types.keys()), fill_value='NO SHIFT')
        return final_schedule, unfilled_info

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

        self._reset_runtime_shift_counters()

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
        processing_order_dates = sorted([d for d in all_dates if d in self.problem_days]) + sorted([d for d in all_dates if d not in self.problem_days])
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
                if p_name in pharmacists_working_yesterday: pharmacist_consecutive_days[p_name] += 1
                else: pharmacist_consecutive_days[p_name] = 0

            is_day_before_problem_day = (date + timedelta(days=1)) in self.problem_days
            shifts_to_process = problem_day_shift_order if date in self.problem_days else standard_shift_order
            for shift_type in shifts_to_process:
                if schedule_dict[date][shift_type] not in ['NO SHIFT', 'UNASSIGNED', 'UNFILLED'] or not self.is_shift_available_on_date(shift_type, date): continue
                available = self._get_available_pharmacists_optimized(shuffled_pharmacists, date, shift_type, schedule_dict, pharmacist_hours, pharmacist_consecutive_days)
                if available:
                    chosen = self._select_best_pharmacist(available, shift_type, date, is_day_before_problem_day)
                    pharmacist_to_assign = chosen['name']
                    schedule_dict[date][shift_type] = pharmacist_to_assign
                    self._update_shift_counts(pharmacist_to_assign, shift_type)
                    pharmacist_hours[pharmacist_to_assign] += self.shift_types[shift_type]['hours']
                else:
                    rescued = self._try_rescue_assign_with_swap(shuffled_pharmacists, date, shift_type, schedule_dict, pharmacist_hours, pharmacist_consecutive_days, is_day_before_problem_day)
                    if not rescued:
                        schedule_dict[date][shift_type] = 'UNFILLED'
                        if date in self.problem_days: unfilled_info['problem_days'].append((date, shift_type))
                        else: unfilled_info['other_days'].append((date, shift_type))

        final_schedule = pd.DataFrame.from_dict(schedule_dict, orient='index')
        final_schedule = final_schedule.reindex(columns=list(self.shift_types.keys()), fill_value='NO SHIFT')
        return final_schedule, unfilled_info

    def _update_shift_counts(self, pharmacist, shift_type):
        if self.is_night_shift(shift_type): self.pharmacists[pharmacist]['night_shift_count'] += 1
        if shift_type.startswith('C8'): self.pharmacists[pharmacist]['mixing_shift_count'] += 1
        if shift_type.startswith('Care'): self.pharmacists[pharmacist]['care_shift_count'] += 1
        category = self._get_shift_category(shift_type)
        if category and category in self.pharmacists[pharmacist]['category_counts']:
            self.pharmacists[pharmacist]['category_counts'][category] += 1

    def _revert_shift_counts(self, pharmacist, shift_type):
        if pharmacist not in self.pharmacists or shift_type not in self.shift_types: return
        if self.is_night_shift(shift_type): self.pharmacists[pharmacist]['night_shift_count'] = max(0, self.pharmacists[pharmacist].get('night_shift_count', 0) - 1)
        if shift_type.startswith('C8'): self.pharmacists[pharmacist]['mixing_shift_count'] = max(0, self.pharmacists[pharmacist].get('mixing_shift_count', 0) - 1)
        if shift_type.startswith('Care'): self.pharmacists[pharmacist]['care_shift_count'] = max(0, self.pharmacists[pharmacist].get('care_shift_count', 0) - 1)
        category = self._get_shift_category(shift_type)
        if category and category in self.pharmacists[pharmacist].get('category_counts', {}):
            self.pharmacists[pharmacist]['category_counts'][category] = max(0, self.pharmacists[pharmacist]['category_counts'].get(category, 0) - 1)

    def _is_preassigned_shift(self, pharmacist, date, shift_type):
        date_str = pd.to_datetime(date).strftime('%Y-%m-%d')
        return (pharmacist in self.pre_assignments and date_str in self.pre_assignments[pharmacist] and shift_type in self.pre_assignments[pharmacist][date_str])

    def _is_skill_scarce_shift(self, shift_type):
        required_skills = [str(s).strip().lower() for s in self.shift_types.get(shift_type, {}).get('required_skills', []) if str(s).strip()]
        return bool(required_skills) or shift_type.startswith('C8') or shift_type.startswith('Care')

    def _try_rescue_assign_with_swap(self, pharmacists, date, target_shift, schedule_dict, pharmacist_hours, pharmacist_consecutive_days, is_day_before_problem_day=False):
        if not self._is_skill_scarce_shift(target_shift): return False
        if not self.is_shift_available_on_date(target_shift, date): return False

        required_skills = [str(s).strip().lower() for s in self.shift_types[target_shift].get('required_skills', []) if str(s).strip()]
        rescue_options = []

        for old_shift, donor in list(schedule_dict[date].items()):
            if donor not in self.pharmacists or old_shift == target_shift: continue
            if schedule_dict[date].get(target_shift) not in ['NO SHIFT', 'UNASSIGNED', 'UNFILLED']: continue
            if self._is_preassigned_shift(donor, date, old_shift) or self.is_night_shift(old_shift): continue

            donor_skills = [s.strip().lower() for s in self.pharmacists[donor].get('skills', [])]
            if required_skills and not all(skill in donor_skills for skill in required_skills): continue

            original_old_assignee = schedule_dict[date][old_shift]
            original_target_assignee = schedule_dict[date][target_shift]
            schedule_dict[date][old_shift] = 'NO SHIFT'
            schedule_dict[date][target_shift] = 'NO SHIFT'

            temp_hours = pharmacist_hours.copy()
            temp_hours[donor] = max(0, temp_hours.get(donor, 0) - self.shift_types[old_shift]['hours'])

            donor_available = self._get_available_pharmacists_optimized([donor], date, target_shift, schedule_dict, temp_hours, pharmacist_consecutive_days)
            if donor_available:
                replacement_pool = [p for p in pharmacists if p != donor]
                replacement_candidates = self._get_available_pharmacists_optimized(replacement_pool, date, old_shift, schedule_dict, temp_hours, pharmacist_consecutive_days)
                if replacement_candidates:
                    replacement = self._select_best_pharmacist(replacement_candidates, old_shift, date, is_day_before_problem_day)
                    old_shift_required_skills = [str(s).strip().lower() for s in self.shift_types[old_shift].get('required_skills', []) if str(s).strip()]
                    old_shift_scarcity_penalty = 1000 if self._is_skill_scarce_shift(old_shift) else 0
                    old_shift_skill_penalty = len(old_shift_required_skills) * 200
                    total_score = old_shift_scarcity_penalty + old_shift_skill_penalty + self._calculate_suitability_score(donor_available[0]) + self._calculate_suitability_score(replacement)

                    rescue_options.append({
                        'score': total_score, 'donor': donor, 'donor_from_shift': old_shift,
                        'target_shift': target_shift, 'replacement': replacement['name'], 'replacement_to_shift': old_shift,
                    })

            schedule_dict[date][old_shift] = original_old_assignee
            schedule_dict[date][target_shift] = original_target_assignee

        if not rescue_options: return False
        best = min(rescue_options, key=lambda x: x['score'])
        donor, old_shift, replacement = best['donor'], best['donor_from_shift'], best['replacement']

        schedule_dict[date][old_shift] = replacement
        schedule_dict[date][target_shift] = donor

        self._revert_shift_counts(donor, old_shift)
        pharmacist_hours[donor] = max(0, pharmacist_hours.get(donor, 0) - self.shift_types[old_shift]['hours'])
        self._update_shift_counts(donor, best['target_shift'])
        pharmacist_hours[donor] += self.shift_types[best['target_shift']]['hours']
        self._update_shift_counts(replacement, old_shift)
        pharmacist_hours[replacement] += self.shift_types[old_shift]['hours']
        return True

    def _get_available_pharmacists_optimized(self, pharmacists, date, shift_type, schedule_dict, current_hours_dict, consecutive_days_dict):
        available_pharmacists = []
        pharmacists_on_night_yesterday = set()
        previous_date = date - timedelta(days=1)
        if previous_date in schedule_dict:
            pharmacists_on_night_yesterday = {p for s, p in schedule_dict[previous_date].items() if p in self.pharmacists and self.is_night_shift(s)}

        new_start = self.shift_types[shift_type]['start_time']
        new_end = self.shift_types[shift_type]['end_time']
        new_dept = self.get_department_from_shift(shift_type)

        for pharmacist in pharmacists:
            if date.strftime('%Y-%m-%d') in self.pharmacists[pharmacist]['holidays']: continue
            if self.has_overlapping_shift_optimized(pharmacist, date, shift_type, schedule_dict): continue
            if pharmacist in pharmacists_on_night_yesterday: continue
            p_skills = [skill.strip().lower() for skill in self.pharmacists[pharmacist]['skills']]
            s_req_skills = [skill.strip().lower() for skill in self.shift_types[shift_type]['required_skills'] if skill.strip()]
            if not all(skill in p_skills for skill in s_req_skills): continue

            projected_hours = current_hours_dict[pharmacist] + self.shift_types[shift_type]['hours']
            if projected_hours > self.pharmacists[pharmacist].get('max_hours', 250): continue
            if self.has_restricted_sequence_optimized(pharmacist, date, shift_type, schedule_dict): continue

            is_junior = 'junior' in p_skills
            junior_conflict = False
            if is_junior:
                current_juniors_in_dept = 0
                total_dept_shifts_at_time = 0
                for s_type, s_info in self.shift_types.items():
                    if self.get_department_from_shift(s_type) == new_dept and self.is_shift_available_on_date(s_type, date):
                        if self.check_time_overlap(new_start, new_end, s_info['start_time'], s_info['end_time']): total_dept_shifts_at_time += 1

                for existing_shift, assigned_pharm in schedule_dict[date].items():
                    if assigned_pharm in self.pharmacists:
                        if 'junior' in [s.strip().lower() for s in self.pharmacists[assigned_pharm]['skills']]:
                            if new_dept == self.get_department_from_shift(existing_shift):
                                if self.check_time_overlap(new_start, new_end, self.shift_types[existing_shift]['start_time'], self.shift_types[existing_shift]['end_time']):
                                    current_juniors_in_dept += 1

                max_juniors_allowed = 2 if total_dept_shifts_at_time >= 4 or shift_type in ['O400F2-8/1', 'O400F2-8/2', 'O400F2-8/3', 'Care/1', 'Care/2'] else 1
                if current_juniors_in_dept + 1 > max_juniors_allowed: junior_conflict = True

                if not junior_conflict and shift_type in ('O400F2-6', 'O400ER-6'):
                    other_shift = 'O400ER-6' if shift_type == 'O400F2-6' else 'O400F2-6'
                    assigned_pharm = schedule_dict[date].get(other_shift)
                    if assigned_pharm in self.pharmacists and 'junior' in [s.strip().lower() for s in self.pharmacists[assigned_pharm]['skills']]: junior_conflict = True

                if not junior_conflict and shift_type in ('I100-6', 'I400-6'):
                    other_shift = 'I400-6' if shift_type == 'I100-6' else 'I100-6'
                    assigned_pharm = schedule_dict[date].get(other_shift)
                    if assigned_pharm in self.pharmacists and 'junior' in [s.strip().lower() for s in self.pharmacists[assigned_pharm]['skills']]: junior_conflict = True

            if junior_conflict: continue

            category = self._get_shift_category(shift_type)
            if category:
                limit = self.shift_limits.get(pharmacist, {}).get(category)
                if limit is not None and self.pharmacists[pharmacist]['category_counts'][category] >= limit: continue

            if self.is_night_shift(shift_type):
                if self.has_nearby_night_shift_optimized(pharmacist, date, schedule_dict): continue
                if pharmacist in self.pre_assignments and (date + timedelta(days=1)).strftime('%Y-%m-%d') in self.pre_assignments[pharmacist]: continue

            if shift_type.startswith('C8') and not self.check_mixing_expert_ratio_optimized(schedule_dict, date, shift_type, pharmacist): continue

            current_streak = self.get_dynamic_consecutive_days(pharmacist, date, schedule_dict)
            if current_streak >= self.MAX_CONSECUTIVE_DAYS: continue

            original_preference = self.get_preference_score(pharmacist, shift_type)
            multiplier = self.preference_multipliers.get(pharmacist, 1.0)
            current_hrs = current_hours_dict[pharmacist]
            days_in_month = pd.Timestamp(date).days_in_month
            time_elapsed_pct = date.day / days_in_month
            hours_used_pct = current_hrs / max_hours if max_hours > 0 else 1.0

            is_weekend = date.weekday() >= 5
            weekend_days_worked = self.get_weekend_days_worked(pharmacist, schedule_dict) if is_weekend else 0

            soulmate = self.soul_mates.get(pharmacist)
            soulmate_working_today = soulmate in schedule_dict[date].values() if soulmate else False
            mate_on_holiday = (soulmate in self.pharmacists and date.strftime('%Y-%m-%d') in self.pharmacists[soulmate]['holidays']) if soulmate else False

            pharmacist_data = {
                'name': pharmacist, 'preference_score': original_preference * multiplier, 'consecutive_days': current_streak,
                'night_count': self.pharmacists[pharmacist]['night_shift_count'], 'mixing_count': self.pharmacists[pharmacist]['mixing_shift_count'],
                'care_count': self.pharmacists[pharmacist].get('care_shift_count', 0), 'current_hours': current_hrs, 'max_hours': max_hours,
                'time_elapsed_pct': time_elapsed_pct, 'hours_used_pct': hours_used_pct, 'is_weekend': is_weekend, 'weekend_days_worked': weekend_days_worked,
                'has_soulmate': bool(soulmate), 'soulmate_working_today': soulmate_working_today, 'mate_on_holiday': mate_on_holiday,
                'needs_min_shift': self._needs_min_shift(pharmacist, shift_type, schedule_dict), 'no_preference': self.pharmacists[pharmacist].get('no_preference', False),
                'department_count': self.get_dept_shift_count(pharmacist, new_dept, schedule_dict) if new_dept else 0,
                'total_shift_count': self.get_total_shift_count_in_schedule_dict(pharmacist, schedule_dict),
                'has_worked_this_department': (new_dept in self.get_unique_departments_worked(pharmacist, schedule_dict)) if new_dept else True,
                'average_monthly_shift_target': self.get_average_monthly_shift_target(date.year, date.month),
                'month_segment_shift_count': self.get_month_segment_shift_count(pharmacist, date, schedule_dict),
                'total_weekend_days': self.get_total_weekend_days_in_schedule(schedule_dict),
                'weekend_days_worked_before': self.get_weekend_days_worked_until(pharmacist, schedule_dict),
            }
            available_pharmacists.append(pharmacist_data)
        return available_pharmacists

    def validate_min_shift_requirements(self, schedule):
        violations = []
        schedule_dict = {date: schedule.loc[date].to_dict() for date in schedule.index}
        for pharmacist, dept_reqs in self.min_shift_requirements.items():
            for department, min_count in dept_reqs.items():
                actual = self.get_dept_shift_count(pharmacist, department, schedule_dict)
                if actual < min_count:
                    violations.append({'pharmacist': pharmacist, 'department': department, 'required': min_count, 'actual': actual, 'shortfall': min_count - actual})
        return violations

    def _calculate_suitability_score(self, pharmacist_data):
        consecutive_penalty = self.W_CONSECUTIVE * (pharmacist_data['consecutive_days'] ** 2)
        hours_penalty = self.W_HOURS * (pharmacist_data['hours_used_pct'] * 100)
        preference_penalty = self.W_PREFERENCE * pharmacist_data['preference_score']
        min_shift_bonus = -200 if pharmacist_data.get('needs_min_shift', False) else 0

        no_preference_department_balance_penalty = 0
        if pharmacist_data.get('no_preference', False):
            no_preference_department_balance_penalty = (pharmacist_data.get('department_count', 0) * 200) + (pharmacist_data.get('total_shift_count', 0) * 3)
            if not pharmacist_data.get('has_worked_this_department', True): no_preference_department_balance_penalty -= 300

        pacing_penalty = 500 * (pharmacist_data['hours_used_pct'] - pharmacist_data['time_elapsed_pct']) if pharmacist_data['hours_used_pct'] > pharmacist_data['time_elapsed_pct'] else 0

        target_monthly_shifts = pharmacist_data.get('average_monthly_shift_target', 1)
        projected_shift_pct = (pharmacist_data.get('total_shift_count', 0) + 1) / target_monthly_shifts if target_monthly_shifts > 0 else 1
        month_progress_pct = pharmacist_data.get('time_elapsed_pct', 1)
        allowed_pacing_buffer = 0.12
        shift_pacing_penalty = self.W_SHIFT_PACING * ((projected_shift_pct - month_progress_pct - allowed_pacing_buffer) ** 2) if projected_shift_pct > month_progress_pct + allowed_pacing_buffer else 0

        month_segment_penalty = self.W_MONTH_SEGMENT_BALANCE * (pharmacist_data.get('month_segment_shift_count', 0) ** 2)

        weekend_off_protection_penalty = 0
        if pharmacist_data.get('is_weekend', False):
            max_weekend_work_days = max(pharmacist_data.get('total_weekend_days', 0) - self.MIN_WEEKEND_OFF_DAYS, 0)
            projected_weekend_work_days = pharmacist_data.get('weekend_days_worked_before', 0) + 1
            if projected_weekend_work_days > max_weekend_work_days:
                weekend_off_protection_penalty = self.W_WEEKEND_OFF_PROTECTION * ((projected_weekend_work_days - max_weekend_work_days) ** 2)

        return consecutive_penalty + hours_penalty + preference_penalty + min_shift_bonus + no_preference_department_balance_penalty + pacing_penalty + shift_pacing_penalty + month_segment_penalty + weekend_off_protection_penalty

    def _select_best_pharmacist(self, available_pharmacists, shift_type, date, is_day_before_problem_day):
        if self.is_night_shift(shift_type) and is_day_before_problem_day:
            problem_day_str = (date + timedelta(days=1)).strftime('%Y-%m-%d')
            candidates_off_tomorrow = [p for p in available_pharmacists if problem_day_str in self.pharmacists[p['name']]['holidays']]
            if candidates_off_tomorrow: return min(candidates_off_tomorrow, key=lambda x: (x['night_count'], self._calculate_suitability_score(x)))

        if self.is_night_shift(shift_type): return min(available_pharmacists, key=lambda x: (x['night_count'], self._calculate_suitability_score(x)))
        elif shift_type.startswith('C8'): return min(available_pharmacists, key=lambda x: (x['mixing_count'], self._calculate_suitability_score(x)))
        elif shift_type.startswith('Care'): return min(available_pharmacists, key=lambda x: (x['care_count'], self._calculate_suitability_score(x)))
        return min(available_pharmacists, key=lambda x: self._calculate_suitability_score(x))

    def calculate_preference_penalty(self, pharmacist, schedule):
        penalty = 0
        for date in schedule.index:
            for shift_type, assigned_pharm in schedule.loc[date].items():
                if assigned_pharm == pharmacist: penalty += self.get_preference_score(pharmacist, shift_type)
        return penalty

    def get_dynamic_consecutive_days(self, pharmacist, date, schedule_dict):
        streak = 0
        curr_date = date - timedelta(days=1)
        while curr_date in schedule_dict:
            if any(p == pharmacist for p in schedule_dict[curr_date].values() if p not in ['NO SHIFT', 'UNFILLED', 'UNASSIGNED']):
                streak += 1
                curr_date -= timedelta(days=1)
            else: break
        curr_date = date + timedelta(days=1)
        while curr_date in schedule_dict:
            if any(p == pharmacist for p in schedule_dict[curr_date].values() if p not in ['NO SHIFT', 'UNFILLED', 'UNASSIGNED']):
                streak += 1
                curr_date += timedelta(days=1)
            else: break
        return streak

    def get_weekend_days_worked(self, pharmacist, schedule_dict):
        return sum(1 for d, shifts in schedule_dict.items() if d.weekday() >= 5 and pharmacist in shifts.values())

    def is_schedule_better(self, current_metrics, best_metrics):
        current_unfilled = current_metrics.get('unfilled_problem_shifts', float('inf'))
        best_unfilled = best_metrics.get('unfilled_problem_shifts', float('inf'))
        if current_unfilled < best_unfilled: return True
        if current_unfilled > best_unfilled: return False
        weights = {'preference_score': 1.0, 'preference_variance': 50.0, 'hour_imbalance_penalty': 25.0, 'night_variance': 800.0, 'weekend_off_variance': 1000.0, 'weekend_min_off_shortfall': 5000.0, 'month_segment_variance': 2500.0}
        return sum(weights[k] * current_metrics.get(k, 0) for k in weights) < sum(weights[k] * best_metrics.get(k, 0) for k in weights)

    def optimize_schedule(self, year, month, iterations=10, true_random_override=False, enable_run_log=False):
        self.run_logs = []
        self.run_config = {"Year": year, "Month": month, "Iterations": iterations, "True Random Override": true_random_override, "Enable Run Log": enable_run_log, "Staff Type": self.staff_type, "Run At": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
        self._log_schedule_event("RUN_START", "Start schedule generation", Year=year, Month=month, Iterations=iterations, TrueRandomOverride=true_random_override, EnableRunLog=enable_run_log)

        # สร้าง UI progress bar ใน Streamlit
        progress_bar = st.progress(0)
        status_text = st.empty()

        if true_random_override:
            best_schedule, best_unfilled_info = None, {}
            best_metrics = {'unfilled_problem_shifts': float('inf'), 'hour_imbalance_penalty': float('inf'), 'hour_diff_for_logging': float('inf')}
            for i in range(iterations):
                status_text.text(f"กำลังรันโหมด สุ่มอิสระ รอบที่ {i + 1}/{iterations}...")
                current_schedule, unfilled_info = self.generate_monthly_schedule_true_random(year, month, iteration_num=i + 1)
                metrics = self.calculate_schedule_metrics(current_schedule, year, month)
                metrics['unfilled_problem_shifts'] = len(unfilled_info.get('problem_days', [])) + len(unfilled_info.get('other_days', []))
                
                current_key = (metrics['unfilled_problem_shifts'], metrics.get('hour_imbalance_penalty', float('inf')), metrics.get('hour_diff_for_logging', float('inf')))
                best_key = (best_metrics.get('unfilled_problem_shifts', float('inf')), best_metrics.get('hour_imbalance_penalty', float('inf')), best_metrics.get('hour_diff_for_logging', float('inf')))
                if best_schedule is None or current_key < best_key:
                    best_schedule = current_schedule.copy()
                    best_unfilled_info = unfilled_info.copy()
                    best_metrics = metrics.copy()
                progress_bar.progress((i + 1) / iterations)
            status_text.text("จัดตารางเสร็จสิ้น!")
            return best_schedule, best_unfilled_info

        best_schedule, best_unfilled_info = None, {}
        best_metrics = {'unfilled_problem_shifts': float('inf'), 'hour_imbalance_penalty': float('inf'), 'night_variance': float('inf'), 'preference_score': float('inf')}
        self._pre_check_staffing_levels(year, month)

        for i in range(iterations):
            status_text.text(f"กำลังคำนวณและค้นหาความสมดุล รอบที่ {i + 1}/{iterations}...")
            current_schedule, unfilled_info = self.generate_monthly_schedule_shuffled(year, month, iteration_num=i + 1)
            metrics = self.calculate_schedule_metrics(current_schedule, year, month)
            metrics['unfilled_problem_shifts'] = len(unfilled_info['problem_days']) + len(unfilled_info['other_days'])

            if best_schedule is None or self.is_schedule_better(metrics, best_metrics):
                best_schedule = current_schedule.copy()
                best_metrics = metrics.copy()
                best_unfilled_info = unfilled_info.copy()
            progress_bar.progress((i + 1) / iterations)
            
        status_text.text("คำนวณหาตารางที่เหมาะสมที่สุดเรียบร้อยแล้ว!")
        return best_schedule, best_unfilled_info

    def export_to_excel_in_memory(self, schedule, unfilled_info, enable_run_log=False):
        """ ส่งออกข้อมูล Excel เป็น Bytes IO เพื่อเปิดให้ดาวน์โหลดบน Streamlit """
        wb = Workbook()
        ws = wb.active
        ws.title = 'Monthly Schedule'
        ws_daily = wb.create_sheet("Daily Summary")
        ws_daily_codes = wb.create_sheet("Daily Summary (Codes)")
        ws_pref = wb.create_sheet("Preference Scores")
        ws_negotiate = wb.create_sheet("Negotiation Suggestions")
        ws_signature = wb.create_sheet("Signature Sheet")
        ws_min_req = wb.create_sheet("Min Req Violations")
        ws_run_logs = wb.create_sheet("Run Logs") if enable_run_log else None
        
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
        self.create_signature_sheet(ws_signature, schedule)

        violations = self.validate_min_shift_requirements(schedule)
        headers = ["Pharmacist", "Department", "Required", "Actual", "Shortfall"]
        min_header_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
        
        for col, h in enumerate(headers, 1):
            cell = ws_min_req.cell(row=1, column=col, value=h)
            cell.fill, cell.font, cell.border = min_header_fill, Font(bold=True, color="FFFFFFFF"), border

        if not violations:
            ws_min_req.cell(row=2, column=1, value="✅ All minimum shift requirements satisfied.")
        else:
            for row, v in enumerate(violations, 2):
                for col, val in enumerate([v['pharmacist'], v['department'], v['required'], v['actual'], v['shortfall']], 1):
                    cell = ws_min_req.cell(row=row, column=col, value=val)
                    cell.border = border
                    if v['shortfall'] > 0: cell.fill = PatternFill(start_color='FFFFF2CC', fill_type='solid')

        ws_min_req.column_dimensions['A'].width = 35
        for col in ['B','C','D','E']: ws_min_req.column_dimensions[col].width = 15

        if enable_run_log and ws_run_logs is not None:
            ws_run_logs.cell(row=1, column=1, value="Run Configuration").font = Font(bold=True)
            config_row = 2
            for key, value in self.run_config.items():
                ws_run_logs.cell(row=config_row, column=1, value=key)
                ws_run_logs.cell(row=config_row, column=2, value=str(value))
                config_row += 1
            if self.run_logs:
                all_keys = list(self.run_logs[0].keys())
                header_row = config_row + 2
                for col_idx, key in enumerate(all_keys, 1):
                    cell = ws_run_logs.cell(row=header_row, column=col_idx, value=key)
                    cell.font, cell.fill = Font(bold=True), header_fill
                for row_idx, log in enumerate(self.run_logs, header_row + 1):
                    for col_idx, key in enumerate(all_keys, 1):
                        ws_run_logs.cell(row=row_idx, column=col_idx, value=str(log.get(key, "")))

        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()

    def create_signature_sheet(self, ws, schedule):
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        header_fill = PatternFill(start_color='FFD3D3D3', fill_type='solid')
        x_fill = PatternFill(start_color='FFD3D3D3', fill_type='solid')
        white_fill = PatternFill(start_color='FFFFFFFF', fill_type='solid')

        shift_colors = {'I100': 'FF00B050', 'O100': 'FF00B0F0', 'Care': 'FFD40202', 'C8': 'FFE6B8AF', 'I400': 'FFFF00FF', 'O400F1': 'FF0033CC', 'O400F2': 'FFC78AF2', 'O400ER': 'FFED7D31', 'ARI': 'FF7030A0', 'Refill': 'FF741b47'}
        ws.cell(row=1, column=1, value='Shift / Date').fill = header_fill
        ws.cell(row=1, column=1).border, ws.cell(row=1, column=1).font = border, Font(bold=True)

        sorted_dates = sorted(schedule.index)
        for col, date in enumerate(sorted_dates, 2):
            cell = ws.cell(row=1, column=col, value=date.strftime('%d/%m'))
            cell.fill, cell.border, cell.font = header_fill, border, Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')

        for row, shift_type in enumerate(self.shift_types.keys(), 2):
            shift_info = self.shift_types[shift_type]
            shift_desc = f"{shift_info['description']} ({int(shift_info['hours'])} ชม.)\n({shift_info['start_time']} - {shift_info['end_time']})"
            row_color, font_color = 'FFFFFFFF', 'FF000000'

            for prefix in sorted(shift_colors.keys(), key=len, reverse=True):
                if shift_type.startswith(prefix):
                    row_color = shift_colors[prefix]
                    if prefix in ['Care', 'O400F1', 'ARI','Refill']: font_color = 'FFFFFFFF'
                    break

            cell_first = ws.cell(row=row, column=1, value=shift_desc)
            cell_first.fill, cell_first.border, cell_first.font = PatternFill(start_color=row_color, fill_type='solid'), border, Font(color=font_color, bold=True)
            cell_first.alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')

            for col, date in enumerate(sorted_dates, 2):
                cell = ws.cell(row=row, column=col)
                cell.border, cell.alignment = border, Alignment(horizontal='center', vertical='center')
                if schedule.loc[date, shift_type] == 'NO SHIFT':
                    cell.value, cell.fill = 'X', x_fill
                else:
                    cell.value, cell.fill = '', white_fill

        ws.column_dimensions['A'].width = 40
        for col in range(2, len(sorted_dates) + 2): ws.column_dimensions[get_column_letter(col)].width = 7

    def create_negotiation_summary(self, ws, schedule):
        header_fill = PatternFill(start_color='FF4F81BD', fill_type='solid')
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        headers = ["Date", "Unfilled Shift", "Suggested Negotiation Candidates (Ranked)"]
        for col, header_text in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header_text)
            cell.fill, cell.font, cell.border = header_fill, Font(bold=True, color="FFFFFFFF"), border
            
        unfilled_shifts = [(d, st) for d in schedule.index for st, p in schedule.loc[d].items() if p == 'UNFILLED']
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
                
                max_hrs = p_info.get('max_hours', 250)
                current_hrs = self.calculate_total_hours(p_name, schedule)
                hours_used_pct = current_hrs / max_hrs if max_hrs > 0 else 1.0
                
                pharmacist_data = {
                    'name': p_name, 'preference_score': self.get_preference_score(p_name, shift_type),
                    'consecutive_days': self.count_consecutive_shifts(p_name, date, schedule), 'night_count': p_info.get('night_shift_count', 0),
                    'mixing_count': p_info.get('mixing_shift_count', 0), 'current_hours': current_hrs, 'max_hours': max_hrs,
                    'time_elapsed_pct': date.day / pd.Timestamp(date).days_in_month, 'hours_used_pct': hours_used_pct
                }
                all_candidates.append({'name': p_name, 'is_on_holiday': is_on_holiday, 'score': self._calculate_suitability_score(pharmacist_data)})
                
            sorted_candidates = sorted(all_candidates, key=lambda x: (x['is_on_holiday'], x['score']))
            suggestions_text = [f"{i+1}. {cand['name']} {'(On Holiday)' if cand['is_on_holiday'] else '(Available)'}" for i, cand in enumerate(sorted_candidates[:3])]
            
            ws.cell(row=current_row, column=1, value=date.strftime('%Y-%m-%d')).border = border
            ws.cell(row=current_row, column=2, value=shift_type).border = border
            cell = ws.cell(row=current_row, column=3, value="\n".join(suggestions_text) if suggestions_text else "No suitable candidate found")
            cell.border, cell.alignment = border, Alignment(wrap_text=True, vertical='top')
            ws.row_dimensions[current_row].height = 50
            current_row += 1
        ws.column_dimensions['A'].width, ws.column_dimensions['B'].width, ws.column_dimensions['C'].width = 15, 25, 45

    def create_schedule_summaries(self, ws, schedule):
        summary_row = len(schedule) + 3
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
        for col_idx, shift_type in enumerate(shift_types_list, 2): ws.cell(row=shift_row, column=col_idx, value=shift_type).font = Font(bold=True)
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
            'fills': {p: PatternFill(fill_type='solid', start_color=c) for p, c in [('I100', 'FF00B050'), ('O100', 'FF00B0F0'), ('Care', 'FFD40202'), ('C8', 'FFE6B8AF'), ('I400', 'FFFF00FF'), ('O400F1', 'FF0033CC'), ('O400F2', 'FFC78AF2'), ('O400ER', 'FFED7D31'), ('ARI', 'FF7030A0'), ('Refill', 'FF741b47')]},
            'fonts': {'O400F1': Font(bold=True, color="FFFFFFFF"), 'ARI': Font(bold=True, color="FFFFFFFF"), 'Refill': Font(bold=True, color="FFFFFFFF"), 'default': Font(bold=True), 'header': Font(bold=True)}
        }

    def create_daily_summary(self, ws, schedule):
        styles = self._setup_daily_summary_styles()
        ordered_pharmacists = self.get_ordered_employees()
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
            for r in range(3): ws.cell(row=current_row + r, column=1, value="" if r != 1 else pharmacist).fill = styles['header_fill']

            for col, date in enumerate(sorted_dates, 2):
                note_cell, cell1, cell2 = [ws.cell(row=current_row + r, column=col) for r in range(3)]
                all_cells = [note_cell, cell1, cell2]
                for cell in all_cells:
                    cell.border, cell.alignment = styles['border'], Alignment(horizontal="center", vertical="center")

                date_str = date.strftime('%Y-%m-%d')
                shifts = self.get_pharmacist_shifts(pharmacist, date, schedule)
                
                if date_str in self.pharmacists[pharmacist]['holidays']:
                    cell2.value = 'X'
                    for cell in all_cells: cell.fill = styles['off_fill']
                else:
                    if len(shifts) > 0:
                        shift = shifts[0]
                        cell2.value = f"{int(self.shift_types[shift]['hours'])}N" if self.is_night_shift(shift) else int(self.shift_types[shift]['hours'])
                        prefix = next((p for p in styles['fills'] if shift.startswith(p)), None)
                        if prefix:
                            cell2.fill, cell2.font = styles['fills'][prefix], styles['fonts'].get(prefix, Font(bold=True))
                            if len(shifts) == 1: cell1.fill = styles['fills'][prefix]

                    if len(shifts) > 1:
                        shift = shifts[1]
                        cell1.value = f"{int(self.shift_types[shift]['hours'])}N" if self.is_night_shift(shift) else int(self.shift_types[shift]['hours'])
                        prefix = next((p for p in styles['fills'] if shift.startswith(p)), None)
                        if prefix: cell1.fill, cell1.font = styles['fills'][prefix], styles['fonts'].get(prefix, Font(bold=True))

                    if self.is_holiday(date) or date.weekday() >= 5:
                        note_cell.fill = styles['holiday_empty_fill']
                        if not shifts:
                            cell1.fill = styles['holiday_empty_fill']
                            cell2.fill = styles['holiday_empty_fill']

                note_text = self.special_notes.get(pharmacist, {}).get(date_str)
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
            else: unfilled_cell.value = "0"
            
        ws.column_dimensions['A'].width = 25
        for col in range(2, len(schedule.index) + 3): ws.column_dimensions[get_column_letter(col)].width = 7

    def create_preference_score_summary(self, ws, schedule):
        header_fill = PatternFill(start_color='FFD3D3D3', fill_type='solid')
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        headers = ["Pharmacist", "Preference Score (%)", "Total Shifts Worked"]
        for col, header_text in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header_text)
            cell.fill, cell.font, cell.border = header_fill, Font(bold=True), border
        preference_scores = self.calculate_pharmacist_preference_scores(schedule)
        for row, pharmacist in enumerate(sorted(self.pharmacists.keys()), 2):
            total_shifts = sum(1 for date in schedule.index for p in schedule.loc[date] if p == pharmacist)
            ws.cell(row=row, column=1, value=pharmacist).border = border
            score_cell = ws.cell(row=row, column=2, value=preference_scores.get(pharmacist, 0))
            score_cell.border, score_cell.number_format = border, '0.00"%"'
            ws.cell(row=row, column=3, value=total_shifts).border = border
        ws.column_dimensions['A'].width, ws.column_dimensions['B'].width, ws.column_dimensions['C'].width = 30, 25, 25

    def create_daily_summary_with_codes(self, ws, schedule):
        styles = self._setup_daily_summary_styles()
        ordered_pharmacists = self.get_ordered_employees()
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
            for r in range(3): ws.cell(row=current_row + r, column=1, value="" if r != 1 else pharmacist).fill = styles['header_fill']

            for col, date in enumerate(sorted_dates, 2):
                note_cell, cell1, cell2 = [ws.cell(row=current_row + r, column=col) for r in range(3)]
                all_cells = [note_cell, cell1, cell2]
                for cell in all_cells:
                    cell.border, cell.alignment = styles['border'], Alignment(horizontal="center", vertical="center")
                    if cell != note_cell: cell.font = Font(bold=True, size=9)

                date_str = date.strftime('%Y-%m-%d')
                shifts = self.get_pharmacist_shifts(pharmacist, date, schedule)
                
                if date_str in self.pharmacists[pharmacist]['holidays']:
                    cell2.value = 'OFF'
                    for cell in all_cells: cell.fill = styles['off_fill']
                else:
                    if len(shifts) > 0:
                        shift_code = shifts[0]
                        cell2.value = shift_code
                        prefix = next((p for p in styles['fills'] if shift_code.startswith(p)), None)
                        if prefix:
                            cell2.fill, cell2.font = styles['fills'][prefix], styles['fonts'].get(prefix, styles['fonts']['default'])
                            if len(shifts) == 1: cell1.fill = styles['fills'][prefix]

                    if len(shifts) > 1:
                        shift_code = shifts[1]
                        cell1.value = shift_code
                        prefix = next((p for p in styles['fills'] if shift_code.startswith(p)), None)
                        if prefix: cell1.fill, cell1.font = styles['fills'][prefix], styles['fonts'].get(prefix, styles['fonts']['default'])

                    if self.is_holiday(date) or date.weekday() >= 5:
                        note_cell.fill = styles['holiday_empty_fill']
                        if not shifts:
                            cell1.fill = styles['holiday_empty_fill']
                            cell2.fill = styles['holiday_empty_fill']

                note_text = self.special_notes.get(pharmacist, {}).get(date_str)
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
            else: unfilled_cell.value = "0"
            
        ws.column_dimensions['A'].width = 25
        for col in range(2, len(schedule.index) + 2): ws.column_dimensions[get_column_letter(col)].width = 15

    def calculate_pharmacist_preference_scores(self, schedule):
        scores = {}
        MAX_POINTS_PER_SHIFT = 8
        for pharmacist in self.pharmacists:
            total_achieved_points, total_shifts_worked = 0, 0
            for date in schedule.index:
                for shift_type, assigned_pharm in schedule.loc[date].items():
                    if assigned_pharm == pharmacist:
                        total_shifts_worked += 1
                        rank = self.get_preference_score(pharmacist, shift_type)
                        total_achieved_points += max(0, 9 - rank)
            if total_shifts_worked == 0: scores[pharmacist] = 0
            else:
                max_possible_points = total_shifts_worked * MAX_POINTS_PER_SHIFT
                scores[pharmacist] = (total_achieved_points / max_possible_points) * 100 if max_possible_points > 0 else 0
        return scores


# ==========================================
# STREAMLIT UI VIEW AND EXECUTION CONTROLLER
# ==========================================

st.title("🏥 Intelligent Pharmacy Scheduling Support System")
st.caption("ระบบสนับสนุนการจัดตารางเวรเภสัชกรและผู้ช่วยเภสัชกรอัจฉริยะ (เวอร์ชันพัฒนาต่อยอดประมวลผลผ่านเว็บแอปพลิเคชัน)")

# ----------------- SIDEBAR -----------------
st.sidebar.header("⚙️ ตั้งค่าการประมวลผล")

# 1. เลือกประเภทตาราง
schedule_key = st.sidebar.selectbox("เลือกชนิดการรันข้อมูล", list(SCHEDULE_SOURCES.keys()), format_func=lambda x: SCHEDULE_SOURCES[x]['label'])
selected_source = SCHEDULE_SOURCES[schedule_key]

# 2. ตั้งค่าวันเวลาและรอบ
year = st.sidebar.number_input("ปี ค.ศ. ที่ต้องการรัน", min_value=2000, max_value=2100, value=2026)
month = st.sidebar.slider("เดือนที่ต้องการรัน", min_value=1, max_value=12, value=6)
iterations = st.sidebar.number_input("จำนวนรอบการประมวลผล (Iterations)", min_value=1, max_value=500, value=20)

# 3. แอดวานซ์ออปชัน
st.sidebar.markdown("---")
st.sidebar.subheader("🛠️ ตัวเลือกขั้นสูง")
enable_run_log = st.sidebar.toggle("บันทึกประวัติการตัดสินใจ (Run Logs)", value=False)
true_random_override = st.sidebar.toggle("โหมดสุ่มอิสระ (True Random Override)", value=False)

# 4. อัปโหลดไฟล์แทนโมเดล Colab Drive
st.sidebar.markdown("---")
st.sidebar.subheader("📂 นำเข้าข้อมูลดิบ")
uploaded_file = st.sidebar.file_uploader("อัปโหลดไฟล์ Excel ต้นทางของคุณที่นี่ (.xlsx)", type=["xlsx"])

# ----------------- MAIN VIEW -----------------
if uploaded_file is not None:
    st.info("💡 อ่านข้อมูลโครงสร้างจากไฟล์สำเร็จ กรุณากดปุ่มเพื่อเริ่มคำนวณ")
    
    if st.button("🚀 เริ่มต้นคำนวณและประมวลผลตารางเวร", type="primary"):
        with st.spinner("ระบบกำลังจำลองการจัดตารางเวรและค้นหาโครงสร้างที่เสถียรที่สุด..."):
            try:
                # สร้างอินสแตนซ์ของคลาสจัดตาราง
                scheduler = PharmacistScheduler(
                    excel_file=uploaded_file,
                    employee_sheet_name=selected_source['employee_sheet_name'],
                    staff_type=selected_source['label']
                )

                # สั่งรันโมเดลคำนวณหาคำตอบที่ดีที่สุด
                best_schedule, best_unfilled_info = scheduler.optimize_schedule(
                    year=year,
                    month=month,
                    iterations=iterations,
                    true_random_override=true_random_override,
                    enable_run_log=enable_run_log
                )

                if best_schedule is not None:
                    st.success("🎉 การประมวลผลสำเร็จ! ตารางเวรที่ได้มีดัชนีคะแนนความเหมาะสมสูงที่สุด")

                    # คำนวณสถิติส่งท้ายสำหรับแสดงผลแบบ Dashboard
                    metrics = scheduler.calculate_schedule_metrics(best_schedule, year, month)
                    total_unfilled = len(best_unfilled_info.get('problem_days', [])) + len(best_unfilled_info.get('other_days', []))

                    # 🏥 DASHBOARD METRICS VIEW
                    st.subheader("📊 ดัชนีชี้วัดความสมดุลของตาราง (Schedule Metrics)")
                    m_col1, m_col2, m_col3, m_col4 = st.columns(4)
                    m_col1.metric("เวรว่างที่ยังไม่มีคนลง (Unfilled)", f"{total_unfilled} เวร", delta=f"{total_unfilled} Target" if total_unfilled > 0 else "สมบูรณ์ 100%", delta_color="inverse")
                    m_col2.metric("ความเบี่ยงเบนชั่วโมงทำงาน (Hour SD)", f"{metrics['hour_diff_for_logging']:.2f} ชม.")
                    m_col3.metric("ความแปรปรวนเวรดึก (Night Var)", f"{metrics['night_variance']:.2f}")
                    m_col4.metric("คะแนนความพึงพอใจโดยรวม", f"{metrics['preference_score']:.1f}")

                    # 🗂️ STREAMLIT TABS FOR DATA DISPLAY
                    st.markdown("---")
                    st.subheader("📋 รายละเอียดตารางและบทสรุปรายบุคคล")
                    
                    tab1, tab2, tab3, tab4 = st.tabs(["🗓️ ตารางเวรรวมรายวัน", "👤 ชั่วโมงทำงานสะสม", "🎯 คะแนนความพึงพอใจ (%)", "⚠️ ข้อผิดพลาดเกณฑ์ขั้นต่ำ"])
                    
                    with tab1:
                        st.dataframe(best_schedule.rename(index=lambda x: x.strftime('%Y-%m-%d')), use_container_width=True)
                        
                    with tab2:
                        hours_data = [{"Pharmacist": p, "Total Hours": scheduler.calculate_total_hours(p, best_schedule)} for p in scheduler.pharmacists]
                        st.dataframe(pd.DataFrame(hours_data), use_container_width=True)
                        
                    with tab3:
                        pref_scores = scheduler.calculate_pharmacist_preference_scores(best_schedule)
                        pref_data = [{"Pharmacist": p, "Preference Score (%)": f"{score:.2f}%"} for p, score in pref_scores.items()]
                        st.dataframe(pd.DataFrame(pref_data), use_container_width=True)
                        
                    with tab4:
                        violations = scheduler.validate_min_shift_requirements(best_schedule)
                        if not violations:
                            st.balloons()
                            st.success("✅ ผ่านเกณฑ์ขั้นต่ำ! ไม่พบผู้ปฏิบัติงานคนใดได้รับเวรต่ำกว่าเกณฑ์ MinShiftRequirements ที่ระบุ")
                        else:
                            st.warning("⚠️ พบข้อขัดแย้งตามกฎของเกณฑ์ขั้นต่ำในบางบุคคล:")
                            st.dataframe(pd.DataFrame(violations), use_container_width=True)

                    # 💾 DOWNLOAD EXCEL BUTTON
                    st.markdown("---")
                    st.subheader("💾 ส่งออกไฟล์ผลลัพธ์")
                    
                    # แปลงไฟล์ในหน่วยความจำเป็น bytes stream
                    excel_bytes = scheduler.export_to_excel_in_memory(best_schedule, best_unfilled_info, enable_run_log=enable_run_log)
                    
                    mode_suffix = "TRUE_RANDOM" if true_random_override else "OPTIMIZED"
                    filename = f"{selected_source['output_prefix']}_{year}_{month:02d}_{mode_suffix}.xlsx"
                    
                    st.download_button(
                        label="📥 ดาวน์โหลดไฟล์ตารางเวรสำหรับพิมพ์และใช้งาน (.xlsx)",
                        data=excel_bytes,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                else:
                    st.error("❌ ระบบไม่สามารถประมวลผลโครงสร้างตารางที่สอดคล้องกับเงื่อนไขบังคับได้ กรุณาเพิ่มจำนวนรอบหรือตรวจสอบวันลาหยุด")

            except Exception as e:
                st.error(f"🚨 เกิดข้อผิดพลาดในการรันแอปพลิเคชัน: {str(e)}")
else:
    # หน้าต้อนรับเบื้องต้นเมื่อยังไม่ได้อัปโหลดไฟล์
    st.markdown("---")
    st.info("👋 ยินดีต้อนรับสู่ระบบจัดตารางอัจฉริยะ! กรุณาเตรียมไฟล์เอกสารข้อมูลบุคลากร เวร แผนก และเงื่อนไขต่างๆ ให้พร้อม จากนั้นอัปโหลดไฟล์ผ่านแถบด้านซ้ายมือเพื่อเริ่มต้นกระบวนการ")