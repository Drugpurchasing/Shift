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


# --- The PharmacistScheduler Class ---
# ไม่มีการเปลี่ยนแปลงตรรกะหลักของคลาสนี้
# --- The PharmacistScheduler Class ---
# (MODIFIED to include the 'junior' skill constraint)
class PharmacistScheduler:
    """
    Pharmacy shift scheduler with optimization and Excel export.
    Designed for multi-constraint pharmacist roster planning.
    """
    W_CONSECUTIVE = 8
    W_HOURS = 16
    W_PREFERENCE = 4
    W_WEEKEND_OFF = 6
    W_JUNIOR_BONUS = 0

    def __init__(self, excel_file_path, logger=print, progress_bar=None):
        self.logger = logger
        self.progress_bar = progress_bar
        self.pharmacists = {}
        self.shift_types = {}
        self.departments = {}
        self.pre_assignments = {}
        self.historical_scores = {}
        self.preference_multipliers = {}
        self.special_notes = {}
        self.shift_limits = {}
        self.excel_file_path = excel_file_path
        self.problem_days = set()

        self.read_data_from_excel(self.excel_file_path)
        self.load_historical_scores()
        self._calculate_preference_multipliers()

        self.night_shifts = {
            'I100-10', 'I100-12N', 'I400-12N', 'I400-10', 'O400ER-12N', 'O400ER-10'
        }
        self.holidays = {
            'specific_dates': ['2025-12-05','2025-12-10', '2025-12-31']
        }
        for pharmacist in self.pharmacists:
            self.pharmacists[pharmacist]['shift_counts'] = {
                shift_type: 0 for shift_type in self.shift_types
            }

    def _update_progress(self, value, text):
        if self.progress_bar:
            self.progress_bar.progress(value, text=text)

    def read_data_from_excel(self, file_path_or_url):
        try:
            self._update_progress(0, "กำลังโหลดข้อมูล: เภสัชกร...")
            pharmacists_df = pd.read_excel(file_path_or_url, sheet_name='Pharmacists', engine='openpyxl')
            self.pharmacists = {}
            for _, row in pharmacists_df.iterrows():
                name = row['Name']
                max_hours = row.get('Max Hours', 250)
                if pd.isna(max_hours) or max_hours == '' or max_hours is None:
                    max_hours = 250
                else:
                    max_hours = float(max_hours)

                # Make sure skills are clean strings
                skills_raw = str(row['Skills']).split(',')
                skills_clean = [skill.strip() for skill in skills_raw if skill.strip()]

                # <<< MODIFICATION START: Check for 'None' preferences >>>
                preferences = {}
                all_prefs_are_none = True
                for i in range(1, 9):
                    rank_val = row.get(f'Rank{i}')
                    # Check if value is None, NaN, or an empty string
                    if pd.isna(rank_val) or str(rank_val).strip().lower() in ['', 'none', 'nan']:
                        preferences[f'rank{i}'] = None # Standardize to None
                    else:
                        preferences[f'rank{i}'] = str(rank_val).strip()
                        all_prefs_are_none = False # Found a real preference
                # <<< MODIFICATION END >>>

                self.pharmacists[name] = {
                    'night_shift_count': 0,
                    'skills': skills_clean,
                    'holidays': [date for date in str(row['Holidays']).split(',') if
                                 date != '1900-01-00' and date.strip() and date != 'nan'],
                    'shift_counts': {},
                    'preferences': preferences, # <<< MODIFIED
                    'max_hours': max_hours,
                    'has_preferences': not all_prefs_are_none # <<< NEW FLAG
                }

            self._update_progress(20, "กำลังโหลดข้อมูล: ประเภทเวร...")
            shifts_df = pd.read_excel(file_path_or_url, sheet_name='Shifts', engine='openpyxl')
            self.shift_types = {}
            for _, row in shifts_df.iterrows():
                shift_code = row['Shift Code']
                self.shift_types[shift_code] = {
                    'description': row['Description'],
                    'shift_type': row['Shift Type'],
                    'start_time': row['Start Time'],
                    'end_time': row['End Time'],
                    'hours': row['Hours'],
                    'required_skills': str(row['Required Skills']).split(','),
                    'restricted_next_shifts': str(row['Restricted Next Shifts']).split(',') if pd.notna(
                        row['Restricted Next Shifts']) else [],
                }

            self._update_progress(40, "กำลังโหลดข้อมูล: แผนกและเวรที่กำหนดล่วงหน้า...")
            departments_df = pd.read_excel(file_path_or_url, sheet_name='Departments', engine='openpyxl')
            self.departments = {}
            for _, row in departments_df.iterrows():
                department = row['Department']
                self.departments[department] = str(row['Shift Codes']).split(',')

            pre_assign_df = pd.read_excel(file_path_or_url, sheet_name='PreAssignments', engine='openpyxl')
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

            self._update_progress(60, "กำลังโหลดข้อมูล: หมายเหตุพิเศษ...")
            notes_df = pd.read_excel(file_path_or_url, sheet_name='SpecialNotes', index_col=0, engine='openpyxl')
            for pharmacist, row_data in notes_df.iterrows():
                if pharmacist in self.pharmacists:
                    for date_col, note in row_data.items():
                        if pd.notna(note) and str(note).strip():
                            date_str = pd.to_datetime(date_col).strftime('%Y-%m-%d')
                            if pharmacist not in self.special_notes:
                                self.special_notes[pharmacist] = {}
                            self.special_notes[pharmacist][date_str] = str(note).strip()
        except Exception:
            pass  # Ignore if sheets don't exist

        try:
            self._update_progress(80, "กำลังโหลดข้อมูล: ข้อจำกัดเวร...")
            limits_df = pd.read_excel(file_path_or_url, sheet_name='ShiftLimits', engine='openpyxl')
            for _, row in limits_df.iterrows():
                pharmacist = row['Pharmacist']
                category = row['ShiftCategory']
                max_count = row['MaxCount']
                if pharmacist in self.pharmacists:
                    if pharmacist not in self.shift_limits:
                        self.shift_limits[pharmacist] = {}
                    self.shift_limits[pharmacist][category] = int(max_count)
        except Exception:
            pass

    def load_historical_scores(self):
        try:
            self._update_progress(90, "กำลังโหลดข้อมูล: คะแนนย้อนหลัง...")
            df = pd.read_excel(self.excel_file_path, sheet_name='HistoricalScores', engine='openpyxl')
            if 'Pharmacist' in df.columns and 'Total Preference Score' in df.columns:
                for _, row in df.iterrows():
                    pharmacist = row['Pharmacist']
                    score = row['Total Preference Score']
                    if pharmacist in self.pharmacists:
                        self.historical_scores[pharmacist] = score
        except Exception:
            self.logger("ไม่พบข้อมูลคะแนนย้อนหลัง (HistoricalScores)")

    def _pre_check_staffing_levels(self, year, month):
        self.logger("กำลังตรวจสอบจำนวนพนักงานเปรียบเทียบกับจำนวนเวรที่ต้องการ...")
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
                self.logger(f"⚠️ คำเตือน: วันที่ {date.strftime('%Y-%m-%d')} อาจมีเภสัชกรไม่พอ")
        if all_ok:
            self.logger("✅ จำนวนพนักงานเพียงพอสำหรับทุกวัน")
        return not all_ok

    def optimize_schedule(self, year, month, iterations, progress_bar):
        best_schedule = None
        best_metrics = {'unfilled_problem_shifts': float('inf'), 'hour_imbalance_penalty': float('inf'),
                        'night_variance': float('inf'), 'preference_score': float('inf')}
        best_unfilled_info = {}
        self._pre_check_staffing_levels(year, month)
        self.logger(f"กำลังเริ่มการคำนวณ {iterations} รอบ...")
        for i in range(iterations):
            current_schedule, unfilled_info = self.generate_monthly_schedule_shuffled(year, month, progress_bar,
                                                                                      iteration_num=i + 1)
            if unfilled_info['other_days']: continue
            metrics = self.calculate_schedule_metrics(current_schedule, year, month)
            metrics['unfilled_problem_shifts'] = len(unfilled_info['problem_days'])
            if best_schedule is None or self.is_schedule_better(metrics, best_metrics):
                best_schedule = current_schedule.copy()
                best_metrics = metrics.copy()
                best_unfilled_info = unfilled_info.copy()
                self.logger(f"รอบที่ {i + 1}: ✅ พบตารางที่ดีกว่าเดิม!")
        if best_schedule is not None:
            self.logger("\n🎉 คำนวณเสร็จสิ้น! พบตารางที่ดีที่สุดแล้ว")
        else:
            self.logger("\n❌ ไม่สามารถหาตารางที่เหมาะสมได้")
        return best_schedule, best_unfilled_info

    def optimize_schedule_for_dates(self, dates_to_schedule, iterations, progress_bar):
        best_schedule = None
        best_metrics = {'unfilled_problem_shifts': float('inf'), 'preference_score': float('inf')}
        best_unfilled_info = {}
        self._pre_check_staffing_for_dates(dates_to_schedule)
        self.logger(f"\nกำลังเริ่มการคำนวณ {iterations} รอบ สำหรับวันที่เลือก...")
        for i in range(iterations):
            current_schedule, unfilled_info = self.generate_schedule_for_dates(dates_to_schedule, progress_bar,
                                                                               iteration_num=i + 1)
            metrics = self.calculate_metrics_for_schedule(current_schedule)
            metrics['unfilled_problem_shifts'] = len(unfilled_info['problem_days']) + len(unfilled_info['other_days'])
            if best_schedule is None or self.is_schedule_better(metrics, best_metrics):
                best_schedule = current_schedule.copy()
                best_metrics = metrics.copy()
                best_unfilled_info = unfilled_info.copy()
                self.logger(f"รอบที่ {i + 1}: ✅ พบตารางที่ดีกว่าเดิม!")
        if best_schedule is not None:
            self.logger("\n🎉 คำนวณเสร็จสิ้น! พบตารางที่ดีที่สุดแล้ว")
        else:
            self.logger("\n❌ ไม่สามารถหาตารางที่เหมาะสมได้")
        return best_schedule, best_unfilled_info

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
            hours, minutes = map(int, time_input.split(':'))
        elif isinstance(time_input, time):
            hours, minutes = time_input.hour, time_input.minute
        else:
            raise ValueError("Invalid input type. Expected string (HH:MM) or datetime.time object.")
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
                working_on_weekend = {schedule.loc[date, shift] for shift in schedule.columns if
                                      schedule.loc[date, shift] not in ['NO SHIFT', 'UNFILLED', 'UNASSIGNED']}
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
        
        # <<< MODIFICATION START >>>
        weekday = date.weekday() # Monday is 0, Sunday is 6
        is_friday = (weekday == 4)
        is_saturday = (weekday == 5)
        is_sunday = (weekday == 6)

        shift_category = shift_info['shift_type']

        if shift_category == 'weekday':
            # Mon-Fri (0-4), excluding holidays
            return not (is_holiday_date or is_saturday or is_sunday)
            
        elif shift_category == 'วันจันทร์-พฤหัส':
            # Mon-Thu (0-3), excluding holidays, Fri, Sat, Sun
            return not (is_holiday_date or is_friday or is_saturday or is_sunday)
            
        elif shift_category == 'saturday':
            # Only Saturday, excluding holidays
            return is_saturday and not is_holiday_date
            
        elif shift_category == 'holiday':
            # Public holidays, Saturdays, or Sundays
            return is_holiday_date or is_saturday or is_sunday
            
        elif shift_category == 'night':
            # Any day
            return True
            
        return False # Default for unknown types
        # <<< MODIFICATION END >>>

    def get_department_from_shift(self, shift_type):
        if shift_type.startswith('I100'):
            return 'IPD100'
        elif shift_type.startswith('O100'):
            return 'OPD100'
        elif shift_type.startswith('Care'):
            return 'Care'
        elif shift_type.startswith('C8'):
            return 'Mixing'
        elif shift_type.startswith('I400'):
            return 'IPD400'
        elif shift_type.startswith('O400F1'):
            return 'OPD400F1'
        elif shift_type.startswith('O400F2'):
            return 'OPD400F2'
        elif shift_type.startswith('O400ER'):
            return 'ER'
        elif shift_type.startswith('ARI'):
            return 'ARI'
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
        # <<< MODIFICATION START: Handle 'No Preference' >>>
        pharmacist_info = self.pharmacists[pharmacist]

        # Check the new flag. If they have no preferences, return a neutral score.
        # 5 is neutral (average of 1-8 is 4.5, 9 is unranked, so 5 is a good mid-point penalty).
        if not pharmacist_info.get('has_preferences', True):
            return 5 # Neutral preference score
        # <<< MODIFICATION END >>>

        department = self.get_department_from_shift(shift_type)
        for rank in range(1, 9):
            # Use .get() to safely access the rank value, which might be None
            if pharmacist_info['preferences'].get(f'rank{rank}') == department:
                return rank
        
        # Not found in ranks 1-8, return the "unranked" penalty
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

    # <<< NEW METHOD START >>>
    def is_another_junior_in_same_department(self, candidate_pharmacist, date, new_shift_type, schedule_dict):
        """
        ตรวจสอบว่าการ assign เภสัชกร junior
        จะทำให้มี junior 2 คนทำงานใน "แผนกเดียวกัน" (ตาม prefix ของเวร)
        ในวันเดียวกันหรือไม่
        """
        # 1. กฎนี้ใช้เฉพาะเมื่อคนที่กำลังจะ assign เป็น junior
        if 'junior' not in self.pharmacists[candidate_pharmacist]['skills']:
            return False

        # 2. หาว่าเวรใหม่นี้ อยู่แผนกอะไร (เช่น IPD100, Mixing)
        new_dept = self.get_department_from_shift(new_shift_type)

        # ถ้าเวรใหม่ไม่สังกัดแผนก (เช่น Care) ให้ถือว่ากฎนี้ไม่ทำงาน
        if new_dept is None:
            return False

        # 3. ตรวจสอบเวรทั้งหมดที่จัดไปแล้วในวันนั้น
        if date in schedule_dict:
            for existing_shift, assigned_pharmacist in schedule_dict[date].items():
                
                # ข้ามเวรที่ยังว่าง
                if assigned_pharmacist in ['NO SHIFT', 'UNASSIGNED', 'UNFILLED']:
                    continue
                
                # ไม่ต้องเทียบกับตัวเอง (กรณีมีเวร pre-assign หลายอัน)
                if assigned_pharmacist == candidate_pharmacist:
                    continue

                # 4. ตรวจสอบว่าคนที่ถูกจัดไปแล้ว เป็น junior หรือไม่
                if 'junior' in self.pharmacists[assigned_pharmacist]['skills']:
                    
                    # 5. ถ้าเป็น junior, หาว่าเขาอยู่แผนกอะไร
                    existing_dept = self.get_department_from_shift(existing_shift)
                    
                    # 6. ถ้าแผนกตรงกัน = เจอปัญหา
                    if existing_dept == new_dept:
                        # Conflict: มี junior ทำงานในแผนกนี้แล้ว
                        return True

        # ไม่พบปัญหา
        return False

    # <<< NEW METHOD END >>>

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

    def generate_monthly_schedule_shuffled(self, year, month, progress_bar, shuffled_shifts=None,
                                           shuffled_pharmacists=None, iteration_num=1):
        start_date = datetime(year, month, 1)
        end_date = datetime(year + 1, 1, 1) - timedelta(days=1) if month == 12 else datetime(year, month + 1,
                                                                                             1) - timedelta(days=1)
        dates = pd.date_range(start_date, end_date)
        schedule_dict = {date: {shift: 'NO SHIFT' for shift in self.shift_types} for date in dates}
        pharmacist_hours = {p: 0 for p in self.pharmacists}
        pharmacist_consecutive_days = {p: 0 for p in self.pharmacists}
        pharmacist_weekend_work_count = {p: 0 for p in self.pharmacists}

        if shuffled_shifts is None:
            shuffled_shifts = list(self.shift_types.keys())
            random.shuffle(shuffled_shifts)
        if shuffled_pharmacists is None:
            shuffled_pharmacists = list(self.pharmacists.keys())
            random.shuffle(shuffled_pharmacists)
        for pharmacist in self.pharmacists:
            self.pharmacists[pharmacist]['night_shift_count'] = 0
            self.pharmacists[pharmacist]['mixing_shift_count'] = 0
            self.pharmacists[pharmacist]['category_counts'] = {'Mixing': 0, 'Night': 0}
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
        # <<< MODIFICATION START: Randomize 'other_dates' processing order >>>
        problem_dates_sorted = sorted([d for d in all_dates if d in self.problem_days])
        
        # ดึงวันที่เหลือออกมา
        other_dates = [d for d in all_dates if d not in self.problem_days]
        
        # สุ่มลำดับวันที่เหลือ
        random.shuffle(other_dates) 
        
        # วันที่มีปัญหา (เรียงลำดับ) จะถูกจัดก่อนเสมอ ตามด้วยวันที่เหลือ (สุ่มลำดับ)
        processing_order_dates = problem_dates_sorted + other_dates
        unfilled_info = {'problem_days': [], 'other_days': []}
        night_shifts_ordered = [s for s in shuffled_shifts if self.is_night_shift(s)]
        mixing_shifts_ordered = [s for s in shuffled_shifts if s.startswith('C8') and not self.is_night_shift(s)]
        care_shifts_ordered = [s for s in shuffled_shifts if
                               s.startswith('Care') and not self.is_night_shift(s) and not s.startswith('C8')]
        other_shifts_ordered = [s for s in shuffled_shifts if
                                not self.is_night_shift(s) and not s.startswith('C8') and not s.startswith('Care')]
        standard_shift_order = night_shifts_ordered + mixing_shifts_ordered + care_shifts_ordered + other_shifts_ordered
        problem_day_shift_order = mixing_shifts_ordered + care_shifts_ordered + night_shifts_ordered + other_shifts_ordered
        total_dates = len(processing_order_dates)
        for i, date in enumerate(processing_order_dates):
            if progress_bar:
                progress_text = f"รอบที่ {iteration_num}: กำลังจัดเวรวันที่ {date.strftime('%d/%m')}"
                progress_bar.progress((i + 1) / total_dates, text=progress_text)
            pharmacists_working_yesterday = set()
            previous_date = date - timedelta(days=1)
            if previous_date in schedule_dict:
                pharmacists_working_yesterday = {p for p in schedule_dict[previous_date].values() if
                                                 p in self.pharmacists}
            for p_name in self.pharmacists:
                if p_name in pharmacists_working_yesterday:
                    pharmacist_consecutive_days[p_name] += 1
                else:
                    pharmacist_consecutive_days[p_name] = 0
            is_day_before_problem_day = (date + timedelta(days=1)) in self.problem_days
            shifts_to_process = problem_day_shift_order if date in self.problem_days else standard_shift_order
            for shift_type in shifts_to_process:
                if schedule_dict[date][shift_type] not in ['NO SHIFT', 'UNASSIGNED',
                                                           'UNFILLED'] or not self.is_shift_available_on_date(
                    shift_type, date):
                    continue
                available = self._get_available_pharmacists_optimized(shuffled_pharmacists, date, shift_type,
                                                                    schedule_dict, pharmacist_hours,
                                                                    pharmacist_consecutive_days,
                                                                    pharmacist_weekend_work_count)
                
                if available:
                    chosen = self._select_best_pharmacist(available, shift_type, date, is_day_before_problem_day)
                    pharmacist_to_assign = chosen['name']
                    schedule_dict[date][shift_type] = pharmacist_to_assign
                    self._update_shift_counts(pharmacist_to_assign, shift_type)
                    pharmacist_hours[pharmacist_to_assign] += self.shift_types[shift_type]['hours']
                    if date.weekday() >= 5: # 5 = Saturday, 6 = Sunday
                        pharmacist_weekend_work_count[pharmacist_to_assign] += 1
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

    def _get_available_pharmacists_optimized(self, pharmacists, date, shift_type, schedule_dict, current_hours_dict,
                                             consecutive_days_dict, pharmacist_weekend_work_count):
        available_pharmacists = []
        pharmacists_on_night_yesterday = set()
        previous_date = date - timedelta(days=1)
        if previous_date in schedule_dict:
            pharmacists_on_night_yesterday = {p for s, p in schedule_dict[previous_date].items() if
                                              p in self.pharmacists and self.is_night_shift(s)}
        for pharmacist in pharmacists:
            if date.strftime('%Y-%m-%d') in self.pharmacists[pharmacist]['holidays']: continue
            if self.has_overlapping_shift_optimized(pharmacist, date, shift_type, schedule_dict): continue

            # <<< MODIFICATION: INTEGRATE NEW JUNIOR CHECK >>>
            if self.is_another_junior_in_same_department(pharmacist, date, shift_type, schedule_dict): continue

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
                
                # --- START MODIFICATION ---
                next_date = date + timedelta(days=1)
                
                # ตรวจสอบว่า "วันพรุ่งนี้" อยู่ในตารางที่เรากำลังจัดหรือไม่
                if next_date in schedule_dict:
                    
                    # ค้นหาว่าเภสัชกรคนนี้ "มีเวรใดๆ ก็ตาม" ที่ถูกจัดลงในวันพรุ่งนี้แล้วหรือยัง
                    # (ซึ่งอาจเกิดจาก pre-assignment หรือการสุ่มวันที่มาจัดก่อน)
                    pharmacists_working_tomorrow = {p for s, p in schedule_dict[next_date].items() if p in self.pharmacists}
                    
                    if pharmacist in pharmacists_working_tomorrow:
                        # ถ้ามีเวรในวันพรุ่งนี้แล้ว -> ห้ามลงเวรดึกวันนี้
                        continue 
                # --- END MODIFICATION ---
            if shift_type.startswith('C8'):
                if not self.check_mixing_expert_ratio_optimized(schedule_dict, date, shift_type, pharmacist):
                    continue
            is_junior = 'junior' in self.pharmacists[pharmacist]['skills']
            original_preference = self.get_preference_score(pharmacist, shift_type)
            multiplier = self.preference_multipliers.get(pharmacist, 1.0)
            pharmacist_data = {'name': pharmacist, 'preference_score': original_preference * multiplier,
                               'consecutive_days': consecutive_days_dict[pharmacist],
                               'night_count': self.pharmacists[pharmacist]['night_shift_count'],
                               'mixing_count': self.pharmacists[pharmacist]['mixing_shift_count'],
                               'current_hours': current_hours_dict[pharmacist], 
                               'weekend_work_count': pharmacist_weekend_work_count[pharmacist],
                               'is_junior': is_junior
                               }
            available_pharmacists.append(pharmacist_data)
        return available_pharmacists

    def _calculate_suitability_score(self, pharmacist_data, is_weekend_shift): # <<< MODIFIED SIGNATURE
        consecutive_penalty = self.W_CONSECUTIVE * (pharmacist_data['consecutive_days'] ** 2)
        hours_penalty = self.W_HOURS * pharmacist_data['current_hours']
        preference_penalty = self.W_PREFERENCE * pharmacist_data['preference_score']

        weekend_penalty = 0
        if is_weekend_shift:
            weekend_penalty = self.W_WEEKEND_OFF * (pharmacist_data['weekend_work_count'] ** 2)

        # <<< NEW: Apply a bonus (penalty reduction) if the pharmacist is a junior >>>
        junior_bonus = 0
        if pharmacist_data.get('is_junior', False):
            junior_bonus = -self.W_JUNIOR_BONUS # ลบโบนัสออกจากคะแนน (คะแนนน้อย = ดี)

        return consecutive_penalty + hours_penalty + preference_penalty + weekend_penalty + junior_bonus

        return consecutive_penalty + hours_penalty + preference_penalty + weekend_penalty
    def _select_best_pharmacist(self, available_pharmacists, shift_type, date, is_day_before_problem_day):

        is_weekend_shift = date.weekday() >= 5

        if self.is_night_shift(shift_type) and is_day_before_problem_day:
            problem_day = date + timedelta(days=1)
            problem_day_str = problem_day.strftime('%Y-%m-%d')
            candidates_off_tomorrow = []
            for p_data in available_pharmacists:
                p_name = p_data['name']
                if problem_day_str in self.pharmacists[p_name]['holidays']:
                    candidates_off_tomorrow.append(p_data)
            if candidates_off_tomorrow:
                return min(candidates_off_tomorrow,
                           key=lambda x: (x['night_count'], self._calculate_suitability_score(x, is_weekend_shift))) # <<< MODIFIED CALL
        if self.is_night_shift(shift_type):
            return min(available_pharmacists, key=lambda x: (x['night_count'], self._calculate_suitability_score(x, is_weekend_shift))) # <<< MODIFIED CALL
        elif shift_type.startswith('C8'):
            return min(available_pharmacists, key=lambda x: (x['mixing_count'], self._calculate_suitability_score(x, is_weekend_shift))) # <<< MODIFIED CALL
        else:
            return min(available_pharmacists, key=lambda x: self._calculate_suitability_score(x, is_weekend_shift)) # <<< MODIFIED CALL

    def calculate_preference_penalty(self, pharmacist, schedule):
        penalty = 0
        for date in schedule.index:
            for shift_type, assigned_pharm in schedule.loc[date].items():
                if assigned_pharm == pharmacist:
                    penalty += self.get_preference_score(pharmacist, shift_type)
        return penalty

    def is_schedule_better(self, current_metrics, best_metrics):
        # --- Priority 1: Unfilled Shifts (เหมือนเดิม) ---
        current_unfilled = current_metrics.get('unfilled_problem_shifts', float('inf'))
        best_unfilled = best_metrics.get('unfilled_problem_shifts', float('inf'))
        
        if current_unfilled < best_unfilled: 
            return True
        if current_unfilled > best_unfilled: 
            return False

        # --- Priority 2: Hour Imbalance (แก้ไขใหม่ตามโจทย์) ---
        # ถ้าเวรว่างเท่ากัน ให้มาเช็คที่ชั่วโมงการทำงาน
        current_hour_penalty = current_metrics.get('hour_imbalance_penalty', float('inf'))
        best_hour_penalty = best_metrics.get('hour_imbalance_penalty', float('inf'))

        if current_hour_penalty < best_hour_penalty:
            return True
        if current_hour_penalty > best_hour_penalty:
            return False

        # --- Priority 3: Combined Score of Other Metrics (Fallback) ---
        # ถ้าเวรว่าง และ ชั่วโมงการทำงาน ยังเท่ากัน
        # ค่อยใช้คะแนนที่เหลือ (ความชอบ, เวรดึก, เวรเสาร์-อาทิตย์) เป็นตัวตัดสิน
        weights = {
            'preference_score': 1.0, 
            'night_variance': 800.0,
            'weekend_off_variance': 1000.0
        }
        
        current_score = sum(weights[k] * current_metrics.get(k, 0) for k in weights)
        best_score = sum(weights[k] * best_metrics.get(k, 0) for k in weights)
        
        return current_score < best_score

    # ... The rest of the PharmacistScheduler class (export functions, etc.) remains unchanged ...
    # (The following functions are unchanged but included for completeness of the class)
    def export_to_excel(self, schedule, unfilled_info):
        wb = Workbook()
        ws = wb.active
        ws.title = 'Monthly Schedule'
        ws_daily = wb.create_sheet("Daily Summary")
        ws_daily_codes = wb.create_sheet("Daily Summary (Codes)")
        ws_pref = wb.create_sheet("Preference Scores")
        ws_negotiate = wb.create_sheet("Negotiation Suggestions")
        header_fill = PatternFill(start_color='FFD3D3D3', end_color='FFD3D3D3', fill_type='solid')
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'),
                        bottom=Side(style='thin'))
        ws.cell(row=1, column=1, value='Date').fill = header_fill
        for col, shift_type in enumerate(self.shift_types, 2):
            cell = ws.cell(row=1, column=col,
                           value=f"{self.shift_types[shift_type]['description']}\n({self.shift_types[shift_type]['hours']} hrs)")
            cell.fill, cell.font, cell.alignment = header_fill, Font(bold=True), Alignment(wrap_text=True)
        schedule.sort_index(inplace=True)
        for row, date in enumerate(schedule.index, 2):
            ws.cell(row=row, column=1, value=date.strftime('%Y-%m-%d'))
            is_holiday = self.is_holiday(date)
            is_weekend = date.weekday() >= 5
            for col, shift_type in enumerate(self.shift_types, 2):
                cell = ws.cell(row=row, column=col, value=schedule.loc[date, shift_type])
                cell.border = border
                if schedule.loc[date, shift_type] == 'NO SHIFT':
                    cell.fill = PatternFill(start_color='FFCCCCCC', fill_type='solid')
                elif is_holiday:
                    cell.fill = PatternFill(start_color='FFFFB6C1', fill_type='solid')
                elif is_weekend:
                    cell.fill = PatternFill(start_color='FFFFE4E1', fill_type='solid')
                elif schedule.loc[date, shift_type] == 'UNFILLED':
                    cell.fill = PatternFill(start_color='FFFFFF00', fill_type='solid')
        ws.column_dimensions['A'].width = 22
        for col in range(2, len(self.shift_types) + 2):
            ws.column_dimensions[get_column_letter(col)].width = 20
        self.create_schedule_summaries(ws, schedule)
        self.create_daily_summary(ws_daily, schedule)
        self.create_preference_score_summary(ws_pref, schedule)
        self.create_daily_summary_with_codes(ws_daily_codes, schedule)
        self.create_negotiation_summary(ws_negotiate, schedule)
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        return buffer

    def create_negotiation_summary(self, ws, schedule):
        header_fill = PatternFill(start_color='FF4F81BD', end_color='FF4F81BD', fill_type='solid')
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'),
                        bottom=Side(style='thin'))
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

            is_weekend_shift = date.weekday() >= 5
            
            for p_name, p_info in self.pharmacists.items():
                if not all(skill.strip() in p_info['skills'] for skill in required_skills if skill.strip()): continue
                if len(self.get_pharmacist_shifts(p_name, date, schedule)) > 0: continue
                is_on_holiday = date.strftime('%Y-%m-%d') in p_info['holidays']
                pharmacist_data = {
                    'name': p_name, 
                    'preference_score': self.get_preference_score(p_name, shift_type),
                    'consecutive_days': self.count_consecutive_shifts(p_name, date, schedule),
                    'current_hours': self.calculate_total_hours(p_name, schedule),
                    'weekend_work_count': 0 
                }
                suitability_score = self._calculate_suitability_score(pharmacist_data, is_weekend_shift)
                all_candidates.append({'name': p_name, 'is_on_holiday': is_on_holiday, 'score': suitability_score})
            sorted_candidates = sorted(all_candidates, key=lambda x: (x['is_on_holiday'], x['score']))
            suggestions_text = []
            for i, cand in enumerate(sorted_candidates[:3]):
                status = "(On Holiday)" if cand['is_on_holiday'] else "(Available)"
                suggestions_text.append(f"{i + 1}. {cand['name']} {status}")
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
            'border': Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'),
                             bottom=Side(style='thin')),
            'fills': {p: PatternFill(fill_type='solid', start_color=c) for p, c in
                      [('I100', 'FF00B050'), ('O100', 'FF00B0F0'), ('Care', 'FFD40202'), ('C8', 'FFE6B8AF'),
                       ('I400', 'FFFF00FF'), ('O400F1', 'FF0033CC'), ('O400F2', 'FFC78AF2'), ('O400ER', 'FFED7D31'),
                       ('ARI', 'FF7030A0')]},
            'fonts': {'O400F1': Font(bold=True, color="FFFFFFFF"), 'ARI': Font(bold=True, color="FFFFFFFF"),
                      'default': Font(bold=True), 'header': Font(bold=True)}
        }

    def create_daily_summary(self, ws, schedule):
        styles = self._setup_daily_summary_styles()
        ordered_pharmacists = ["ภญ.ประภัสสรา (มิ้น)", "ภญ.ฐิฏิการ (เอ้)", "ภก.บัณฑิตวงศ์ (แพท)", "ภก.ชานนท์ (บุ้ง)",
                                "ภญ.กมลพรรณ (ใบเตย)", "ภญ.กนกพร (นุ้ย)", "ภก.เอกวรรณ (โม)", "ภญ.อาภาภัทร (มะปราง)",
                                "ภก.ชวนันท์ (เท่ห์)", "ภญ.ธนพร (ฟ้า ธนพร)", "ภญ.วิลินดา (เชอร์รี่)",
                                "ภญ.ชลนิชา (เฟื่อง)", "ภญ.ปริญญ์ (ขมิ้น)", "ภก.ธนภรณ์ (กิ๊ฟ)", "ภญ.ปุณยวีร์ (มิ้นท์)",
                                "ภญ.อมลกานต์ (บอม)", "ภญ.อรรชนา (อ้อม)", "ภญ.ศศิวิมล (ฟิลด์)", 'ภญ. ณัฐพร (แอม)', "ภญ.วรรณิดา (ม่าน)",
                                "ภญ.ปาณิศา (แบม)", "ภญ.จิรัชญา (ศิกานต์)", "ภญ.อภิชญา (น้ำตาล)", "ภญ.วรางคณา (ณา)",
                                "ภญ.ดวงดาว (ปลา)", "ภญ.พรนภา (ผึ้ง)", "ภญ.ธนาภรณ์ (ลูกตาล)", "ภญ.วิลาสินี (เจ้นท์)",
                                "ภญ.ภาวิตา (จูน)", "ภญ.ศิรดา (พลอย)", "ภญ.ศุภิสรา (แพร)", "ภญ.กันต์หทัย (ซีน)",
                                "ภญ.พัทธ์ธีรา (วิว)", "ภญ.จุฑามาศ (กวาง)","ภญ.ณัฐนันท์ (มิ้น)","ภญ.สุภนิดา (แนนซ์)","ภญ.ชัญญาภัค (แมงปอ)","ภญ.ธมนวรรณ (ออฟ)","ภญ.วรนิษฐา (เมษา)","ภญ.ณัชชา (เอิง)","ภญ.พนิตตา (กวาง)","ภญ.ศวิตา (มิว)","ภญ.อภิญญา (อุ๋มอิ๋ม)"]
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
                        # <<< MODIFICATION START >>>
                        shift_info = self.shift_types[shift]
                        hours_val = int(shift_info['hours'])
                        if self.is_night_shift(shift):
                            cell2.value = f"{hours_val}N"
                        elif shift_info['description'] == "10.00-18.00น" or shift == "I400-8W/2":
                            cell2.value = f"{hours_val}*"
                        else:
                            cell2.value = hours_val
                        # <<< MODIFICATION END >>>
                        
                        prefix = next((p for p in styles['fills'] if shift.startswith(p)), None)
                        if prefix:
                            fill_color = styles['fills'][prefix]
                            cell2.fill, cell2.font = fill_color, styles['fonts'].get(prefix, Font(bold=True))
                            if len(shifts) == 1: cell1.fill = fill_color
                    if len(shifts) > 1:
                        shift = shifts[1]
                        # <<< MODIFICATION START >>>
                        shift_info = self.shift_types[shift]
                        hours_val = int(shift_info['hours'])
                        if self.is_night_shift(shift):
                            cell1.value = f"{hours_val}N"
                        elif shift_info['description'] == "10.00-18.00น" or shift == "I400-8W/2":
                            cell1.value = f"{hours_val}*"
                        else:
                            cell1.value = hours_val
                        # <<< MODIFICATION END >>>

                        prefix = next((p for p in styles['fills'] if shift.startswith(p)), None)
                        if prefix: cell1.fill, cell1.font = styles['fills'][prefix], styles['fonts'].get(prefix, Font(
                            bold=True))
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
            total_hours = sum(self.shift_types[st]['hours'] for st, p in schedule.loc[date].items() if
                              p not in ['NO SHIFT', 'UNFILLED', 'UNASSIGNED'] and st in self.shift_types)
            unfilled_shifts = [st for st, p in schedule.loc[date].items() if p in ['UNFILLED', 'UNASSIGNED']]
            ws.cell(row=total_row, column=col, value=total_hours).border = styles['border']
            unfilled_cell = ws.cell(row=unfilled_row, column=col)
            unfilled_cell.border = styles['border']
            if unfilled_shifts:
                unfilled_cell.value, unfilled_cell.fill = "\n".join(unfilled_shifts), PatternFill(
                    start_color='FFFFFF00', fill_type='solid')
            else:
                unfilled_cell.value = "0"
        ws.column_dimensions['A'].width = 25
        for col in range(2, len(schedule.index) + 3):
            ws.column_dimensions[get_column_letter(col)].width = 7

    def create_preference_score_summary(self, ws, schedule):
        header_fill = PatternFill(start_color='FFD3D3D3', end_color='FFD3D3D3', fill_type='solid')
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'),
                        bottom=Side(style='thin'))
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
        ordered_pharmacists = ["ภญ.ประภัสสรา (มิ้น)", "ภญ.ฐิฏิการ (เอ้)", "ภก.บัณฑิตวงศ์ (แพท)", "ภก.ชานนท์ (บุ้ง)",
                                "ภญ.กมลพรรณ (ใบเตย)", "ภญ.กนกพร (นุ้ย)", "ภก.เอกวรรณ (โม)", "ภญ.อาภาภัทร (มะปราง)",
                                "ภก.ชวนันท์ (เท่ห์)", "ภญ.ธนพร (ฟ้า ธนพร)", "ภญ.วิลินดา (เชอร์รี่)",
                                "ภญ.ชลนิชา (เฟื่อง)", "ภญ.ปริญญ์ (ขมิ้น)", "ภก.ธนภรณ์ (กิ๊ฟ)", "ภญ.ปุณยวีร์ (มิ้นท์)",
                                "ภญ.อมลกานต์ (บอม)", "ภญ.อรรชนา (อ้อม)", "ภญ.ศศิวิมล (ฟิลด์)", 'ภญ. ณัฐพร (แอม)', "ภญ.วรรณิดา (ม่าน)",
                                "ภญ.ปาณิศา (แบม)", "ภญ.จิรัชญา (ศิกานต์)", "ภญ.อภิชญา (น้ำตาล)", "ภญ.วรางคณา (ณา)",
                                "ภญ.ดวงดาว (ปลา)", "ภญ.พรนภา (ผึ้ง)", "ภญ.ธนาภรณ์ (ลูกตาล)", "ภญ.วิลาสินี (เจ้นท์)",
                                "ภญ.ภาวิตา (จูน)", "ภญ.ศิรดา (พลอย)", "ภญ.ศุภิสรา (แพร)", "ภญ.กันต์หทัย (ซีน)",
                                "ภญ.พัทธ์ธีรา (วิว)", "ภญ.จุฑามาศ (กวาง)","ภญ.ณัฐนันท์ (มิ้น)","ภญ.สุภนิดา (แนนซ์)","ภญ.ชัญญาภัค (แมงปอ)","ภญ.ธมนวรรณ (ออฟ)","ภญ.วรนิษฐา (เมษา)","ภญ.ณัชชา (เอิง)","ภญ.พนิตตา (กวาง)","ภญ.ศวิตา (มิว)","ภญ.อภิญญา (อุ๋มอิ๋ม)"]
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
                        if prefix: cell.fill, cell.font = styles['fills'][prefix], styles['fonts'].get(prefix,
                                                                                                       styles['fonts'][
                                                                                                           'default'])
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
            total_hours = sum(self.shift_types[st]['hours'] for st, p in schedule.loc[date].items() if
                              p not in ['NO SHIFT', 'UNFILLED', 'UNASSIGNED'] and st in self.shift_types)
            unfilled_shifts = [st for st, p in schedule.loc[date].items() if p in ['UNFILLED', 'UNASSIGNED']]
            ws.cell(row=total_row, column=col, value=total_hours).border = styles['border']
            unfilled_cell = ws.cell(row=unfilled_row, column=col)
            unfilled_cell.border = styles['border']
            if unfilled_shifts:
                unfilled_cell.value, unfilled_cell.fill = "\n".join(unfilled_shifts), PatternFill(
                    start_color='FFFFFF00', fill_type='solid')
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

    def _pre_check_staffing_for_dates(self, dates_to_schedule):
        self.logger("กำลังตรวจสอบจำนวนพนักงานเปรียบเทียบกับจำนวนเวรที่ต้องการ...")
        all_ok = True
        for date in dates_to_schedule:
            available_pharmacists_count = sum(1 for p_name, p_info in self.pharmacists.items()
                                              if date.strftime('%Y-%m-%d') not in p_info['holidays'])
            required_shifts_base = sum(1 for st in self.shift_types
                                       if self.is_shift_available_on_date(st, date))
            total_required_shifts_with_buffer = required_shifts_base + 3
            if available_pharmacists_count < total_required_shifts_with_buffer:
                all_ok = False
                self.problem_days.add(date)
                self.logger(f"⚠️ คำเตือน: วันที่ {date.strftime('%Y-%m-%d')} อาจมีเภสัชกรไม่พอ")
        if all_ok:
            self.logger("✅ จำนวนพนักงานเพียงพอสำหรับทุกวัน")
        return not all_ok

    def calculate_weekend_off_variance_for_dates(self, schedule):
        weekend_off_counts = {p: 0 for p in self.pharmacists}
        for date in schedule.index:
            if date.weekday() >= 5:
                working_on_weekend = {schedule.loc[date, shift] for shift in schedule.columns if
                                      schedule.loc[date, shift] not in ['NO SHIFT', 'UNFILLED', 'UNASSIGNED']}
                for p_name in self.pharmacists:
                    if p_name not in working_on_weekend:
                        weekend_off_counts[p_name] += 1
        if len(weekend_off_counts) > 1:
            return np.var(list(weekend_off_counts.values()))
        return 0

    def calculate_metrics_for_schedule(self, schedule):
        hours = {p: self.calculate_total_hours(p, schedule) for p in self.pharmacists}
        night_counts = {p: self.pharmacists[p]['night_shift_count'] for p in self.pharmacists}
        weekend_off_var = self.calculate_weekend_off_variance_for_dates(schedule)
        hour_penalty = self._get_hour_imbalance_penalty(hours)
        metrics = {
            'hour_imbalance_penalty': hour_penalty,
            'night_variance': np.var(list(night_counts.values())) if night_counts else 0,
            'preference_score': sum(self.calculate_preference_penalty(p, schedule) for p in self.pharmacists),
            'weekend_off_variance': weekend_off_var
        }
        if len(hours) > 1 and len(hours.values()) > 1:
            metrics['hour_diff_for_logging'] = stdev(hours.values())
        else:
            metrics['hour_diff_for_logging'] = 0
        return metrics

    def generate_schedule_for_dates(self, dates_to_schedule, progress_bar, iteration_num=1):
        schedule_dict = {date: {shift: 'NO SHIFT' for shift in self.shift_types} for date in dates_to_schedule}
        pharmacist_hours = {p: 0 for p in self.pharmacists}
        pharmacist_consecutive_days = {p: 0 for p in self.pharmacists}
        
        # --- 1. เพิ่มตัวแปรนับวันหยุดสุดสัปดาห์ ---
        pharmacist_weekend_work_count = {p: 0 for p in self.pharmacists} 

        shuffled_shifts = list(self.shift_types.keys())
        random.shuffle(shuffled_shifts)
        shuffled_pharmacists = list(self.pharmacists.keys())
        random.shuffle(shuffled_pharmacists)
        
        for pharmacist in self.pharmacists:
            self.pharmacists[pharmacist]['night_shift_count'] = 0
            self.pharmacists[pharmacist]['mixing_shift_count'] = 0
            self.pharmacists[pharmacist]['category_counts'] = {'Mixing': 0, 'Night': 0}
            
        for pharmacist, assignments in self.pre_assignments.items():
            if pharmacist not in self.pharmacists: continue
            for date_str, shift_types in assignments.items():
                date = pd.to_datetime(date_str).to_pydatetime().date()
                matching_date = next((dt for dt in dates_to_schedule if dt.date() == date), None)
                if matching_date is None: continue
                for shift_type in shift_types:
                    if shift_type in self.shift_types:
                        schedule_dict[matching_date][shift_type] = pharmacist
                        self._update_shift_counts(pharmacist, shift_type)
                        pharmacist_hours[pharmacist] += self.shift_types[shift_type]['hours']
                        
        all_dates = sorted(list(dates_to_schedule)) 
        
        problem_dates_sorted = sorted([d for d in all_dates if d in self.problem_days])
        other_dates = [d for d in all_dates if d not in self.problem_days]
        random.shuffle(other_dates)
        
        processing_order_dates = problem_dates_sorted + other_dates

        unfilled_info = {'problem_days': [], 'other_days': []}

        night_shifts_ordered = [s for s in shuffled_shifts if self.is_night_shift(s)]
        mixing_shifts_ordered = [s for s in shuffled_shifts if s.startswith('C8') and not self.is_night_shift(s)]
        care_shifts_ordered = [s for s in shuffled_shifts if
                               s.startswith('Care') and not self.is_night_shift(s) and not s.startswith('C8')]
        other_shifts_ordered = [s for s in shuffled_shifts if
                                not self.is_night_shift(s) and not s.startswith('C8') and not s.startswith('Care')]
        standard_shift_order = night_shifts_ordered + mixing_shifts_ordered + care_shifts_ordered + other_shifts_ordered
        problem_day_shift_order = mixing_shifts_ordered + care_shifts_ordered + night_shifts_ordered + other_shifts_ordered
        
        total_dates = len(processing_order_dates)
        
        for i, date in enumerate(processing_order_dates):
            if progress_bar:
                progress_text = f"รอบที่ {iteration_num}: กำลังจัดเวรวันที่ {date.strftime('%d/%m')}"
                progress_bar.progress((i + 1) / total_dates, text=progress_text)
            
            previous_date = date - timedelta(days=1)
            if previous_date in schedule_dict:
                pharmacists_working_yesterday = {p for p in schedule_dict[previous_date].values() if
                                                 p in self.pharmacists}
                for p_name in self.pharmacists:
                    if p_name in pharmacists_working_yesterday:
                        pharmacist_consecutive_days[p_name] += 1
                    else:
                        pharmacist_consecutive_days[p_name] = 0
            else:
                for p_name in self.pharmacists:
                    pharmacist_consecutive_days[p_name] = 0
            
            is_day_before_problem_day = (date + timedelta(days=1)) in self.problem_days
            shifts_to_process = problem_day_shift_order if date in self.problem_days else standard_shift_order
            
            for shift_type in shifts_to_process:
                if schedule_dict[date][shift_type] != 'NO SHIFT' or not self.is_shift_available_on_date(shift_type, date):
                    continue
                
                # --- 2. ส่ง pharmacist_weekend_work_count เข้าไปในฟังก์ชัน ---
                available = self._get_available_pharmacists_optimized(shuffled_pharmacists, date, shift_type,
                                                                      schedule_dict, pharmacist_hours,
                                                                      pharmacist_consecutive_days,
                                                                      pharmacist_weekend_work_count)
                if available:
                    chosen = self._select_best_pharmacist(available, shift_type, date, is_day_before_problem_day)
                    pharmacist_to_assign = chosen['name']
                    schedule_dict[date][shift_type] = pharmacist_to_assign
                    self._update_shift_counts(pharmacist_to_assign, shift_type)
                    pharmacist_hours[pharmacist_to_assign] += self.shift_types[shift_type]['hours']
                    
                    # --- 3. อัปเดตค่าเมื่อทำงานเสาร์อาทิตย์ ---
                    if date.weekday() >= 5: # 5 = Saturday, 6 = Sunday
                        pharmacist_weekend_work_count[pharmacist_to_assign] += 1

                else:
                    schedule_dict[date][shift_type] = 'UNFILLED'
                    if date in self.problem_days:
                        unfilled_info['problem_days'].append((date, shift_type))
                    else:
                        unfilled_info['other_days'].append((date, shift_type))
                        
        final_schedule = pd.DataFrame.from_dict(schedule_dict, orient='index')
        final_schedule.fillna('NO SHIFT', inplace=True)
        return final_schedule, unfilled_info


# --- Pharmacy Assistant Scheduler Class ---
# มีการปรับแก้เล็กน้อยเพื่อทำงานร่วมกับ Streamlit (logger, progress_bar, export)
# แต่ไม่มีการเปลี่ยนแปลงตรรกะหลักในการจัดเวร
class AssistantScheduler:
    """
    Scheduler for Pharmacy Assistants with a multi-pass system, pre-scheduling checks,
    prioritized filling for difficult days (with mixing shifts first), and long-term fairness tracking.
    """

    def __init__(self, excel_file_path, logger=print, progress_bar=None):
        self.excel_file_path = excel_file_path
        self.logger = logger
        self.progress_bar = progress_bar
        self.assistants = {}
        self.shift_types = {}
        self.departments = {}
        self.pre_assignments = {}
        self.special_notes = {}
        self.read_data_from_excel(excel_file_path)
        self.night_shifts = {'I100-16', 'I100-12N', 'I400-12N', 'I400-16', 'O400ER-12N', 'O400ER-16'}
        self.holidays = {'specific_dates': ['2025-12-05','2025-12-10', '2025-12-31']}

    def _update_progress(self, value, text):
        if self.progress_bar:
            self.progress_bar.progress(value, text=text)

    def read_data_from_excel(self, file_path):
        self._update_progress(0, "กำลังโหลดข้อมูลผู้ช่วย: แผนก...")
        departments_df = pd.read_excel(file_path, sheet_name='Departments')
        self.departments = {row['Department']: row['Shift Codes'].split(',') for _, row in departments_df.iterrows()}

        self._update_progress(20, "กำลังโหลดข้อมูลผู้ช่วย: ประเภทเวร...")
        shifts_df = pd.read_excel(file_path, sheet_name='Shifts')
        self.shift_types = {row['Shift Code']: {
            'description': row['Description'], 'shift_type': row['Shift Type'],
            'start_time': row['Start Time'], 'end_time': row['End Time'], 'hours': row['Hours'],
            'required_skills': str(row['Required Skills']).split(','),
            'restricted_next_shifts': str(row['Restricted Next Shifts']).split(',') if pd.notna(
                row['Restricted Next Shifts']) else []
        } for _, row in shifts_df.iterrows()}

        self._update_progress(40, "กำลังโหลดข้อมูลผู้ช่วย: รายชื่อ...")
        assistants_df = pd.read_excel(file_path, sheet_name='Assistants')
        for _, row in assistants_df.iterrows():
            name = row['Name']
            max_hours = row.get('Max Hours', 250)

            department_counts = {}
            for dept in self.departments.keys():
                prev_count_col = f"Prev_{dept}"
                prev_count_val = row.get(prev_count_col)
                department_counts[dept] = int(prev_count_val) if pd.notna(prev_count_val) else 0

            specific_shift_counts = {}
            for sc in self.shift_types.keys():
                prev_sc_col = f"Prev_{sc}"
                prev_sc_val = row.get(prev_sc_col)
                specific_shift_counts[sc] = int(prev_sc_val) if pd.notna(prev_sc_val) else 0

            prev_nights_val = row.get('Prev_Night_Shifts')
            base_night_shifts = int(prev_nights_val) if pd.notna(prev_nights_val) else 0

            prev_hours_val = row.get('Prev_Hours')
            base_hours = float(prev_hours_val) if pd.notna(prev_hours_val) else 0.0

            self.assistants[name] = {
                'skills': str(row['Skills']).split(','),
                'holidays': [date for date in str(row['Holidays']).split(',') if date and date != '1900-01-00'],
                'max_hours': float(max_hours) if pd.notna(max_hours) else 250,
                'department_counts': department_counts,
                'night_shift_count': base_night_shifts,
                'total_hours': base_hours,
                'shift_counts': {st: 0 for st in self.shift_types.keys()},
                'specific_shift_counts': specific_shift_counts
            }

        self._update_progress(80, "กำลังโหลดข้อมูลผู้ช่วย: เวรที่กำหนดล่วงหน้า...")
        pre_assign_df = pd.read_excel(file_path, sheet_name='PreAssignments')
        if not pre_assign_df.empty:
            pre_assign_df['Date'] = pd.to_datetime(pre_assign_df['Date']).dt.strftime('%Y-%m-%d')
            grouping_col = pre_assign_df.columns[0]
            for name, group in pre_assign_df.groupby(grouping_col):
                if name in self.assistants:
                    self.pre_assignments[name] = {
                        date: [s.strip() for shift_str in g['Shift'] for s in str(shift_str).split(',')] for date, g in
                        group.groupby('Date')}
        else:
            self.logger("INFO: ไม่พบข้อมูลในชีท 'PreAssignments'")

        try:
            self._update_progress(90, "กำลังโหลดข้อมูลผู้ช่วย: หมายเหตุพิเศษ...")
            notes_df = pd.read_excel(file_path, sheet_name='SpecialNotes', index_col=0)
            for assistant_name, row_data in notes_df.iterrows():
                if assistant_name in self.assistants:
                    for date_col, note in row_data.items():
                        if pd.notna(note) and str(note).strip():
                            date_str = pd.to_datetime(date_col).strftime('%Y-%m-%d')
                            if assistant_name not in self.special_notes:
                                self.special_notes[assistant_name] = {}
                            self.special_notes[assistant_name][date_str] = str(note).strip()
        except ValueError:
            self.logger("INFO: ไม่พบชีท 'SpecialNotes'")
        except Exception as e:
            self.logger(f"เกิดข้อผิดพลาดในการโหลด SpecialNotes: {e}")

    def _pre_schedule_staffing_check(self, dates):
        self.logger("\n--- กำลังตรวจสอบจำนวนเจ้าหน้าที่ล่วงหน้า ---")
        problematic_dates = []
        for date in dates:
            date_str = date.strftime('%Y-%m-%d')
            shifts_on_day = [st for st in self.shift_types if self.is_shift_available_on_date(st, date)]
            num_shifts = len(shifts_on_day)
            available_assistants = [name for name, data in self.assistants.items() if date_str not in data['holidays']]
            num_available = len(available_assistants)
            if num_available < num_shifts + 3:
                problematic_dates.append(pd.Timestamp(date.date()))
                self.logger(f"⚠️ คำเตือน: อาจมีเจ้าหน้าที่ไม่พอในวันที่ {date_str}. "
                            f"จำนวนเวร: {num_shifts}, เจ้าหน้าที่ที่มาทำงาน: {num_available}. (ต้องการอย่างน้อย {num_shifts + 3})")
        if not problematic_dates:
            self.logger("✅ จำนวนเจ้าหน้าที่เพียงพอสำหรับทุกวัน")
        self.logger("-------------------------------------------\n")
        return problematic_dates

    def _recalculate_counts(self, schedule):
        for name in self.assistants:
            self.assistants[name]['department_counts'] = self.assistants[name]['initial_department_counts'].copy()
            self.assistants[name]['night_shift_count'] = self.assistants[name]['initial_night_count']
            self.assistants[name]['total_hours'] = self.assistants[name]['initial_hours']
            self.assistants[name]['shift_counts'] = {st: 0 for st in self.shift_types.keys()}
            self.assistants[name]['specific_shift_counts'] = self.assistants[name][
                'initial_specific_shift_counts'].copy()
            # <<< NEW: Reset weekend work count >>>
            self.assistants[name]['weekend_work_count'] = 0

        for date in schedule.index:

            is_weekend = date.weekday() >= 5 # 5 = Sat, 6 = Sun

            for shift_type, assistant_name in schedule.loc[date].items():
                if assistant_name in self.assistants:
                    self.assistants[assistant_name]['shift_counts'][shift_type] += 1
                    self.assistants[assistant_name]['total_hours'] += self.shift_types[shift_type]['hours']
                    if shift_type in self.assistants[assistant_name]['specific_shift_counts']:
                        self.assistants[assistant_name]['specific_shift_counts'][shift_type] += 1
                    if self.is_night_shift(shift_type):
                        self.assistants[assistant_name]['night_shift_count'] += 1
                    department = self.get_department_from_shift(shift_type)
                    if department:
                        self.assistants[assistant_name]['department_counts'][department] += 1
                    if is_weekend:
                        self.assistants[assistant_name]['weekend_work_count'] += 1

    def _get_available_assistants(self, assistants, date, shift_type, schedule):
        available = []
        for assistant in assistants:
            if self.check_assistant_availability(assistant, date, shift_type, schedule):
                projected_hours = self.assistants[assistant]['total_hours'] + self.shift_types[shift_type]['hours']
                if projected_hours <= self.assistants[assistant].get('max_hours', 250):
                    available.append({
                        'name': assistant,
                        'consecutive_days': self.count_consecutive_shifts(assistant, date, schedule),
                        'night_count': self.assistants[assistant]['night_shift_count'],
                        'current_hours': self.assistants[assistant]['total_hours'],
                        'department_counts': self.assistants[assistant]['department_counts'],
                        'specific_shift_counts': self.assistants[assistant]['specific_shift_counts'],
                        # <<< NEW: Add current weekend work count to data packet >>>
                        'weekend_work_count': self.assistants[assistant]['weekend_work_count']
                    })
        return available

    def _select_best_assistant(self, available_assistants, shift_type, date, problematic_dates):
        all_current_hours = [data['total_hours'] for data in self.assistants.values()]
        min_hours_among_all = min(all_current_hours) if all_current_hours else 0

        shift_hours = self.shift_types[shift_type]['hours']

        is_weekend_shift = date.weekday() >= 5
        W_WEEKEND_PENALTY_FACTOR = 10  # Penalty weight for sorting

        for a_data in available_assistants:
            projected_hours = a_data['current_hours'] + shift_hours
            a_data['violates_16h_rule'] = (projected_hours - min_hours_among_all) > 16

            weekend_penalty = 0
            if is_weekend_shift:
                # Same logic: penalty scales quadratically with weekend days already worked
                weekend_penalty = (a_data['weekend_work_count'] ** 2) * W_WEEKEND_PENALTY_FACTOR
            a_data['weekend_penalty'] = weekend_penalty

        is_day_before_problem_day = (pd.Timestamp(date.date()) + timedelta(days=1)) in problematic_dates
        if self.is_night_shift(shift_type) and is_day_before_problem_day:
            problem_day_str = (date + timedelta(days=1)).strftime('%Y-m-d')
            candidates_off_tomorrow = [a_data for a_data in available_assistants if
                                       problem_day_str in self.assistants[a_data['name']]['holidays']]
            if candidates_off_tomorrow:
                self.logger(
                    f"\nINFO: จัดเวรดึกวันที่ {date.strftime('%Y-%m-%d')} ให้กับคนที่หยุดในวันถัดไป ({problem_day_str}) ก่อน")
                return min(candidates_off_tomorrow, key=lambda x: (
                x['violates_16h_rule'], x['weekend_penalty'], x['night_count'], x['current_hours'], x['consecutive_days']))

        if self.is_night_shift(shift_type):
            return min(available_assistants, key=lambda x: (
            x['violates_16h_rule'],x['weekend_penalty'], x['night_count'], x['current_hours'], x['consecutive_days']))

        department = self.get_department_from_shift(shift_type)
        if department:
            return min(available_assistants, key=lambda x: (
            x['violates_16h_rule'] ,x['weekend_penalty'], x['specific_shift_counts'].get(shift_type, 0),
            x['department_counts'].get(department, 0), x['current_hours'], x['night_count'], x['consecutive_days']))

        return min(available_assistants, key=lambda x: (
        x['violates_16h_rule'],x['weekend_penalty'], x['specific_shift_counts'].get(shift_type, 0), x['current_hours'], x['night_count'],
        x['consecutive_days']))

    def calculate_schedule_metrics(self, schedule):
        self._recalculate_counts(schedule)
        hours = {p: data['total_hours'] for p, data in self.assistants.items()}
        night_counts_list = [p['night_shift_count'] for p in self.assistants.values()]
        dept_counts = {p: data['department_counts'].values() for p, data in self.assistants.items()}
        total_dept_variance = sum(np.var(list(counts)) for counts in dept_counts.values()) if self.assistants else 0
        return {'hour_diff': stdev(hours.values()) if len(hours) > 1 else 0,
                'night_variance': np.var(night_counts_list) if night_counts_list else 0,
                'department_variance': total_dept_variance}

    def is_schedule_better(self, current_metrics, best_metrics):
        # --- Priority 1: Hour Difference (แก้ไขใหม่ตามโจทย์) ---
        current_hour_diff = current_metrics.get('hour_diff', float('inf'))
        best_hour_diff = best_metrics.get('hour_diff', float('inf'))

        if current_hour_diff < best_hour_diff:
            return True
        if current_hour_diff > best_hour_diff:
            return False

        # --- Priority 2: Combined Score of Other Metrics (Fallback) ---
        # ถ้าชั่วโมงต่างกันเท่ากัน ให้ใช้เกณฑ์อื่นตัดสิน
        weights = {
            'night_variance': 75.0, 
            'department_variance': 200.0
        }
        current_score = sum(weights[k] * current_metrics.get(k, 0) for k in weights)
        best_score = sum(weights[k] * best_metrics.get(k, 0) for k in weights)
        
        return current_score < best_score
    def optimize_schedule(self, dates, iterations=10):
        problematic_dates = self._pre_schedule_staffing_check(dates)
        for name in self.assistants:
            self.assistants[name]['initial_department_counts'] = self.assistants[name]['department_counts'].copy()
            self.assistants[name]['initial_night_count'] = self.assistants[name]['night_shift_count']
            self.assistants[name]['initial_hours'] = self.assistants[name]['total_hours']
            self.assistants[name]['initial_specific_shift_counts'] = self.assistants[name][
                'specific_shift_counts'].copy()
        best_schedule, best_metrics = None, {'hour_diff': float('inf'), 'night_variance': float('inf'),
                                             'department_variance': float('inf')}
        self.logger(f"กำลังเริ่มการคำนวณ {iterations} รอบ...")
        for i in range(iterations):
            self._update_progress((i + 1) / iterations, f"กำลังคำนวณรอบที่ {i + 1}/{iterations}...")
            self.logger(f"\n--- รอบที่ {i + 1}/{iterations} ---")
            current_schedule = self.generate_schedule(dates, problematic_dates, iteration_num=i + 1)
            metrics = self.calculate_schedule_metrics(current_schedule)
            self.logger(
                f"\nผลลัพธ์รอบที่ {i + 1} -> Hour Diff: {metrics['hour_diff']:.2f}, Night Var: {metrics['night_variance']:.2f}, Dept Var: {metrics['department_variance']:.2f}")
            if best_schedule is None or self.is_schedule_better(metrics, best_metrics):
                best_schedule, best_metrics = current_schedule.copy(), metrics.copy()
                self.logger("✅ *** พบตารางที่ดีกว่าเดิม! ***")
        self.logger("\n🎉 คำนวณเสร็จสิ้น! พบตารางที่ดีที่สุดแล้ว")
        self.logger(
            f"ผลลัพธ์สุดท้าย -> Hour Difference: {best_metrics['hour_diff']:.2f}, Night Variance: {best_metrics['night_variance']:.2f}, Dept Var: {best_metrics['department_variance']:.2f}")
        return best_schedule

    def generate_schedule(self, dates, problematic_dates, iteration_num=1):
        max_attempts = 250
        for attempt in range(max_attempts):
            schedule = pd.DataFrame(index=dates, columns=list(self.shift_types.keys()), data='NO SHIFT')
            self._recalculate_counts(pd.DataFrame())
            for assistant, assignments in self.pre_assignments.items():
                if assistant not in self.assistants: continue
                for date_str, shifts in assignments.items():
                    date = pd.to_datetime(date_str)
                    if date in schedule.index:
                        for shift_type in shifts:
                            if shift_type in schedule.columns: schedule.loc[date, shift_type] = assistant
            self._recalculate_counts(schedule)
            problematic_mixing_slots, problematic_other_slots, normal_slots = [], [], []
            all_shifts = list(self.shift_types.keys());
            random.shuffle(all_shifts)
            for shift_type in all_shifts:
                for date in dates:
                    if schedule.loc[date, shift_type] == 'NO SHIFT' and self.is_shift_available_on_date(shift_type,
                                                                                                        date):
                        slot = (date, shift_type)
                        is_problematic = pd.Timestamp(date.date()) in problematic_dates
                        is_mixing = shift_type.startswith('C8')
                        if is_problematic and is_mixing:
                            problematic_mixing_slots.append(slot)
                        elif is_problematic and not is_mixing:
                            problematic_other_slots.append(slot)
                        else:
                            normal_slots.append(slot)
            shuffled_assistants = list(self.assistants.keys());
            random.shuffle(shuffled_assistants)

            def fill_slots(slots_to_fill):
                for date, shift_type in slots_to_fill:
                    available = self._get_available_assistants(shuffled_assistants, date, shift_type, schedule)
                    if available:
                        chosen = self._select_best_assistant(available, shift_type, date, problematic_dates)
                        schedule.loc[date, shift_type] = chosen['name']
                        self._recalculate_counts(schedule)
                    else:
                        schedule.loc[date, shift_type] = 'UNFILLED'

            fill_slots(problematic_mixing_slots)
            fill_slots(problematic_other_slots)

            random.shuffle(normal_slots)

            attempt_failed = False
            for date, shift_type in normal_slots:
                available = self._get_available_assistants(shuffled_assistants, date, shift_type, schedule)
                if available:
                    chosen = self._select_best_assistant(available, shift_type, date, problematic_dates)
                    schedule.loc[date, shift_type] = chosen['name']
                    self._recalculate_counts(schedule)
                else:
                    schedule.loc[date, shift_type] = 'UNFILLED'
                    attempt_failed = True;
                    break

            if not attempt_failed:
                self.logger(f"รอบที่ {iteration_num}: สร้างตารางสำเร็จในความพยายามครั้งที่ {attempt + 1}")
                return schedule

        self.logger(
            f"⚠️ คำเตือน: ไม่สามารถสร้างตารางที่สมบูรณ์ได้ใน {max_attempts} ครั้ง จะสร้างตารางที่ดีที่สุดเท่าที่เป็นไปได้")
        return self._final_fallback_generation(dates)

    def _final_fallback_generation(self, dates):
        schedule = pd.DataFrame(index=dates, columns=list(self.shift_types.keys()), data='NO SHIFT')
        self._recalculate_counts(schedule)
        for assistant, assignments in self.pre_assignments.items():
            if assistant not in self.assistants: continue
            for date_str, shifts in assignments.items():
                date = pd.to_datetime(date_str)
                if date in schedule.index:
                    for shift_type in shifts:
                        if shift_type in schedule.columns: schedule.loc[date, shift_type] = assistant
        self._recalculate_counts(schedule)
        all_slots = []
        all_shifts = list(self.shift_types.keys());
        random.shuffle(all_shifts)
        for shift_type in all_shifts:
            for date in dates:
                if schedule.loc[date, shift_type] == 'NO SHIFT' and self.is_shift_available_on_date(shift_type, date):
                    all_slots.append((date, shift_type))
        shuffled_assistants = list(self.assistants.keys());
        random.shuffle(shuffled_assistants)
        for date, shift_type in all_slots:
            available = self._get_available_assistants(shuffled_assistants, date, shift_type, schedule)
            if available:
                chosen = self._select_best_assistant(available, shift_type, date, [])
                schedule.loc[date, shift_type] = chosen['name']
                self._recalculate_counts(schedule)
            else:
                schedule.loc[date, shift_type] = 'UNFILLED'
        return schedule

    def suggest_negotiations_for_unfilled(self, schedule):
        unfilled_slots = [(index, col) for index, row in schedule.iterrows() for col, value in row.items() if
                          value == 'UNFILLED']
        if not unfilled_slots:
            self.logger("\n✅ ไม่พบเวรว่างในตาราง!")
            return []

        suggestions = []
        self.logger("\n--- ข้อเสนอแนะสำหรับเวรที่ยังว่าง ---")
        for date, shift_type in unfilled_slots:
            date_str = date.strftime('%Y-%m-%d')
            req_skills = [s.strip() for s in self.shift_types[shift_type].get('required_skills', []) if
                          s and pd.notna(s) and s.strip().lower() != 'nan']
            eligible = [name for name, data in self.assistants.items() if
                        all(skill in data['skills'] for skill in req_skills) and date_str not in data['holidays']]
            suggestion_text = random.choice(eligible) if eligible else "ไม่พบผู้ช่วยที่มีคุณสมบัติ"
            self.logger(f"วันที่: {date_str}, เวร: {shift_type} -> ผู้ช่วยที่แนะนำ: {suggestion_text}")
            suggestions.append({'date': date_str, 'shift': shift_type, 'suggestion': suggestion_text})
        self.logger("--------------------------------------------")
        return suggestions

    def check_assistant_availability(self, assistant, date, shift_type, schedule):
        if date.strftime('%Y-%m-%d') in self.assistants[assistant]['holidays']: return False
        required = [s for s in self.shift_types[shift_type]['required_skills'] if
                    s and s.lower() != 'nan' and s.strip() != '']
        if required and not all(s in self.assistants[assistant]['skills'] for s in required): return False
        if self.has_overlapping_shift(assistant, date, shift_type, schedule): return False
        prev_day = date - timedelta(days=1)
        if prev_day in schedule.index:
            if any(
                self.is_night_shift(s) for s in self.get_assistant_shifts(assistant, prev_day, schedule)): return False
        if self.is_night_shift(shift_type):
            next_day = date + timedelta(days=1)
            if next_day in schedule.index and self.get_assistant_shifts(assistant, next_day, schedule): return False
            if self.has_nearby_night_shift(assistant, date, schedule): return False
        return True

    def get_department_from_shift(self, shift_type):
        return next((dept for dept, shifts in self.departments.items() if shift_type in shifts), None)

    def get_assistant_shifts(self, assistant, date, schedule):
        return [st for st, p in schedule.loc[date].items() if p == assistant] if date in schedule.index else []

    def has_overlapping_shift(self, assistant, date, shift_type, schedule):
        if date not in schedule.index: return False
        start1, end1 = self.shift_types[shift_type]['start_time'], self.shift_types[shift_type]['end_time']
        for eshift, epharm in schedule.loc[date].items():
            if epharm == assistant and eshift != shift_type:
                start2, end2 = self.shift_types[eshift]['start_time'], self.shift_types[eshift]['end_time']
                if self.check_time_overlap(start1, end1, start2, end2): return True
        return False

    def has_nearby_night_shift(self, assistant, date, schedule):
        for d in range(-2, 3):
            if d == 0: continue
            c_date = date + timedelta(days=d)
            if c_date in schedule.index and any(
                self.is_night_shift(s) for s in self.get_assistant_shifts(assistant, c_date, schedule)): return True
        return False

    def is_holiday(self, date):
        return date.strftime('%Y-%m-%d') in self.holidays['specific_dates']

    def is_night_shift(self, shift_type):
        return shift_type in self.night_shifts

    def is_shift_available_on_date(self, shift_type, date):
        # <<< MODIFICATION START >>>
        info = self.shift_types[shift_type]
        h = self.is_holiday(date)
        weekday = date.weekday() # Monday is 0, Sunday is 6
        fr = (weekday == 4)
        sa = (weekday == 5)
        su = (weekday == 6)
        
        stype = info['shift_type']
        
        if stype == 'weekday': 
            # Mon-Fri (0-4), excluding holidays
            return not (h or sa or su)
            
        elif stype == 'วันจันทร์-พฤหัส':
            # Mon-Thu (0-3), excluding holidays, Fri, Sat, Sun
            return not (h or fr or sa or su)
            
        elif stype == 'saturday': 
            # Only Saturday, excluding holidays
            return sa and not h
            
        elif stype == 'holiday': 
            # Public holidays, Saturdays, or Sundays
            return h or sa or su
            
        # Default for other types (like 'night' or unlisted) is True
        return True
        # <<< MODIFICATION END >>>

    def count_consecutive_shifts(self, assistant, date, schedule, max_days=2):
        c = 0
        for d in range(1, max_days + 1):
            p_date = date - timedelta(days=d)
            if p_date in schedule.index and assistant in schedule.loc[p_date].values:
                c += 1
            else:
                break
        return c

    def convert_time_to_minutes(self, t):
        if isinstance(t, str):
            h, m = map(int, t.split(':'))
        else:
            h, m = t.hour, t.minute
        return h * 60 + m

    def check_time_overlap(self, s1, e1, s2, e2):
        s1m, e1m, s2m, e2m = map(self.convert_time_to_minutes, [s1, e1, s2, e2])
        if e1m < s1m: e1m += 1440
        if e2m < s2m: e2m += 1440
        return s1m < e2m and e1m > s2m

    def _prepare_daily_summary_data(self, schedule):
        summary_data = {name: {} for name in self.assistants}
        for date in schedule.index:
            date_str = date.strftime('%Y-%m-%d')
            for name in summary_data.keys():
                shifts = self.get_assistant_shifts(name, date, schedule)
                if shifts: summary_data[name][date_str] = shifts
        return summary_data

    def export_to_excel(self, schedule):
        wb = Workbook()
        self.logger("\nกำลังเตรียมข้อมูลสำหรับส่งออก...")
        daily_summary_data = self._prepare_daily_summary_data(schedule)
        ws = wb.active;
        ws.title = 'Main Schedule'
        ws_daily = wb.create_sheet("Daily Summary")
        ws_daily_codes = wb.create_sheet("Daily Summary (Codes)")
        self.create_main_schedule_sheet(ws, schedule)
        self.create_schedule_summaries(ws, schedule, daily_summary_data)
        self.create_daily_summary(ws_daily, schedule, daily_summary_data)
        self.create_daily_summary_with_codes(ws_daily_codes, schedule, daily_summary_data)
        self.create_long_term_fairness_sheet(wb, schedule, daily_summary_data)

        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        return buffer

    def create_main_schedule_sheet(self, ws, schedule):
        header_fill, weekend_fill, holiday_fill, unfilled_fill, no_shift_fill = \
            PatternFill(start_color='D3D3D3', fill_type='solid'), PatternFill(start_color='FFE4E1', fill_type='solid'), \
                PatternFill(start_color='FFB6C1', fill_type='solid'), PatternFill(start_color='FFFF00',
                                                                                  fill_type='solid'), \
                PatternFill(start_color='CCCCCC', fill_type='solid')
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'),
                        bottom=Side(style='thin'))
        ws.cell(row=1, column=1, value='Date').fill = header_fill
        for col, shift_type in enumerate(self.shift_types, 2):
            cell = ws.cell(row=1, column=col,
                           value=f"{self.shift_types[shift_type]['description']}\n({self.shift_types[shift_type]['hours']} hrs)")
            cell.fill, cell.font, cell.alignment = header_fill, Font(bold=True), Alignment(wrap_text=True)
            ws.column_dimensions[get_column_letter(col)].width = 20
        for r_idx, date in enumerate(schedule.index, 2):
            ws.cell(row=r_idx, column=1, value=date.strftime('%Y-%m-%d'))
            is_holiday, is_weekend = self.is_holiday(date), date.weekday() >= 5
            for c_idx, shift_type in enumerate(self.shift_types, 2):
                cell = ws.cell(row=r_idx, column=c_idx, value=schedule.loc[date, shift_type])
                cell.border = border
                if cell.value == 'NO SHIFT':
                    cell.fill = no_shift_fill
                elif cell.value == 'UNFILLED':
                    cell.fill = unfilled_fill
                elif is_holiday:
                    cell.fill = holiday_fill
                elif is_weekend:
                    cell.fill = weekend_fill
        ws.column_dimensions['A'].width = 22

    def create_schedule_summaries(self, ws, schedule, daily_summary_data):
        start_row = len(schedule) + 4
        ws.cell(row=start_row - 1, column=1, value="Summary").font = Font(bold=True, size=14)
        current_month_hours = {
            name: sum(self.shift_types[st]['hours'] for shifts in daily_shifts.values() for st in shifts) for
            name, daily_shifts in daily_summary_data.items()}
        current_month_nights = {
            name: sum(1 for shifts in daily_shifts.values() for st in shifts if self.is_night_shift(st)) for
            name, daily_shifts in daily_summary_data.items()}

        ws.cell(row=start_row, column=1, value="Current Period Summary").font = Font(bold=True)
        ws.cell(row=start_row + 1, column=1, value="Assistant").font = Font(bold=True)
        ws.cell(row=start_row + 1, column=2, value="Hours").font = Font(bold=True)
        ws.cell(row=start_row + 1, column=3, value="Night Shifts").font = Font(bold=True)
        for i, name in enumerate(sorted(self.assistants.keys()), start=start_row + 2):
            ws.cell(row=i, column=1, value=name)
            ws.cell(row=i, column=2, value=current_month_hours.get(name, 0))
            ws.cell(row=i, column=3, value=current_month_nights.get(name, 0))

    def _setup_daily_summary_styles(self):
        return {'header_fill': PatternFill(fill_type='solid', start_color='D3D3D3'),
                'weekend_fill': PatternFill(fill_type='solid', start_color='FFE4E1'),
                'holiday_fill': PatternFill(fill_type='solid', start_color='FFB6C1'),
                'holiday_empty_fill': PatternFill(fill_type='solid', start_color='FFFF00'),
                'off_fill': PatternFill(fill_type='solid', start_color='D3D3D3'),
                'border': Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'),
                                 bottom=Side(style='thin')),
                'fills': {p: PatternFill(fill_type='solid', start_color=c) for p, c in
                          [('I100', '00B050'), ('O100', '00B0F0'), ('Care', 'D40202'), ('C8', 'E6B8AF'),
                           ('I400', 'FF00FF'), ('O400F1', '0033CC'), ('O400F2', 'C78AF2'), ('O400ER', 'ED7D31'),
                           ('ARI', '7030A0')]},
                'fonts': {'O400F1': Font(bold=True, color="FFFFFFFF"), 'ARI': Font(bold=True, color="FFFFFFFF"),
                          'default': Font(bold=True), 'header': Font(bold=True), 'note': Font(italic=True, size=8)}}

    def create_daily_summary(self, ws, schedule, daily_summary_data):
        styles = self._setup_daily_summary_styles()
        ordered_assistants = ["วิภาดา (โอ)", "พักตร์วลัยพร (ปู)", "นิรินทร (อุ้ย)", "วิภาณี (จิ๋ม)", "นัทชา (เก้า)",
                                "ศิริรัตน์ (บี)", "จิรภา (แต๊ก)", "นิรนุช (ปัท)", "ปภัสรินทร์ (ออย)", "ศิริพร (ปุ๋ย)",
                                "พรเพชร (กิ๊ฟ)", "อรุณรัตน์ (เจี๊ยบ)", "กาญจนา (โอ๋)", "พรมงคล (แม็กซ์)", "ฉลอง (โป้ง)",
                                "พราวรวี (แพรว)"
            , "วสุพร (กิ๊ฟ)", "วรัญชลี (เมย์)", "กิตติยา (แนน)", "ปิยะรัตน์ (เบลล์)", "ปนัดดา (แหม่ม)", "ธารวิมล (นัท)",
                                "เกียรติสุดา (ตาต้า)", "ชัญญา (แชมป์)", "แสงเดือน (อั๋น)", "สุกานดา (ตุํกตา)",
                                "พัชรี (ใหม่)", "รุ่งนภา (แพท)", "เบญจวรรณ (จ๊อย)", "จันทร์ภรรัตน์ (อันโน)",
                                "วีระยุทธ (เพ้นท์)", "ฉัตรกมล (ปุ้ย)"
            , "ณัฐฎนิช (มิน)", "วรัญญา (อ้าย)", "ขวัญเนตร (เจเจ)", "มานะ (ตั๊ก)", "รัฎดาวรรณ (เชอรรี่)"]

        ws.cell(row=1, column=1, value='Assistant').fill = styles['header_fill']
        for col, date in enumerate(schedule.index, 2):
            cell = ws.cell(row=1, column=col, value=date.strftime('%d/%m'))
            cell.fill, cell.font = styles['header_fill'], styles['fonts']['header']
            if date.weekday() >= 5: cell.fill = styles['weekend_fill']
            if self.is_holiday(date): cell.fill = styles['holiday_fill']

        current_row = 2
        for assistant in ordered_assistants:
            if assistant not in self.assistants: continue
            ws.cell(row=current_row, column=1, value=assistant).fill = styles['header_fill']
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row + 1, end_column=1)
            ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="center", vertical="center")
            for col, date in enumerate(schedule.index, 2):
                cell1, cell2 = ws.cell(row=current_row, column=col), ws.cell(row=current_row + 1, column=col)
                for cell in [cell1, cell2]: cell.border, cell.alignment = styles['border'], Alignment(
                    horizontal="center", vertical="center")

                date_str = date.strftime('%Y-%m-%d')
                shifts = daily_summary_data.get(assistant, {}).get(date_str, [])
                note_text = self.special_notes.get(assistant, {}).get(date_str)
                if note_text: cell1.value, cell1.font = note_text, styles['fonts']['note']

                if date_str in self.assistants[assistant]['holidays']:
                    cell2.value = 'X';
                    cell1.fill, cell2.fill = styles['off_fill'], styles['off_fill']
                else:
                    if not note_text and len(shifts) == 1:
                        shift = shifts[0]
                        # <<< MODIFICATION START >>>
                        shift_info = self.shift_types[shift]
                        hours_val = int(shift_info['hours'])
                        if self.is_night_shift(shift):
                            cell2.value = f"{hours_val}N"
                        elif shift_info['description'] == "10.00-18.00น" or shift == "I400-8W/2":
                            cell2.value = f"{hours_val}*"
                        else:
                            cell2.value = hours_val
                        # <<< MODIFICATION END >>>
                        
                        prefix = next((p for p in styles['fills'] if shift.startswith(p)), None)
                        if prefix:
                            fill = styles['fills'][prefix];
                            cell1.fill, cell2.fill, cell2.font = fill, fill, styles['fonts'].get(prefix,
                                                                                                 Font(bold=True))
                    elif (note_text and shifts) or len(shifts) > 1:
                        if shifts:
                            shift2 = shifts[0] if note_text else shifts[1]
                            # <<< MODIFICATION START >>>
                            shift_info_2 = self.shift_types[shift2]
                            hours_val_2 = int(shift_info_2['hours'])
                            if self.is_night_shift(shift2):
                                cell2.value = f"{hours_val_2}N"
                            elif shift_info_2['description'] == "10.00-18.00น" or shift2 == "I400-8W/2":
                                cell2.value = f"{hours_val_2}*"
                            else:
                                cell2.value = hours_val_2
                            # <<< MODIFICATION END >>>

                            prefix2 = next((p for p in styles['fills'] if shift2.startswith(p)), None)
                            if prefix2: cell2.fill, cell2.font = styles['fills'][prefix2], styles['fonts'].get(prefix2,
                                                                                                               Font(
                                                                                                                   bold=True))
                        if not note_text and len(shifts) > 1:
                            shift1 = shifts[0]
                            # <<< MODIFICATION START >>>
                            shift_info_1 = self.shift_types[shift1]
                            hours_val_1 = int(shift_info_1['hours'])
                            if self.is_night_shift(shift1):
                                cell1.value = f"{hours_val_1}N"
                            elif shift_info_1['description'] == "10.00-18.00น" or shift1 == "I400-8W/2":
                                cell1.value = f"{hours_val_1}*"
                            else:
                                cell1.value = hours_val_1
                            # <<< MODIFICATION END >>>

                            prefix1 = next((p for p in styles['fills'] if shift1.startswith(p)), None)
                            if prefix1: cell1.fill, cell1.font = styles['fills'][prefix1], styles['fonts'].get(prefix1,
                                                                                                               Font(
                                                                                                                   bold=True))

                    if (self.is_holiday(date) or date.weekday() >= 5) and not shifts and date_str not in \
                            self.assistants[assistant]['holidays']:
                        for cell in [cell1, cell2]:
                            if not cell.value: cell.fill = styles['holiday_empty_fill']
            current_row += 2

        total_row, unfilled_row = current_row + 1, current_row + 2
        ws.cell(row=total_row, column=1, value="Total Hours").fill = styles['header_fill']
        ws.cell(row=unfilled_row, column=1, value="Unfilled Shifts").fill = styles['header_fill']
        for col, date in enumerate(schedule.index, 2):
            total_hours = sum(self.shift_types[st]['hours'] for st, p in schedule.loc[date].items() if
                              p not in ['NO SHIFT', 'UNFILLED'] and st in self.shift_types)
            unfilled_shifts = [st for st, p in schedule.loc[date].items() if p == 'UNFILLED']
            ws.cell(row=total_row, column=col, value=total_hours).border = styles['border']
            unfilled_cell = ws.cell(row=unfilled_row, column=col,
                                    value="\n".join(unfilled_shifts) if unfilled_shifts else "0")
            unfilled_cell.border = styles['border']
            if unfilled_shifts: unfilled_cell.fill = PatternFill(start_color='FFFFFF00', fill_type='solid')

        ws.column_dimensions['A'].width = 25
        for col in range(2, len(schedule.index) + 3): ws.column_dimensions[get_column_letter(col)].width = 7

    def create_daily_summary_with_codes(self, ws, schedule, daily_summary_data):
        styles = self._setup_daily_summary_styles()
        ordered_assistants = ["วิภาดา (โอ)", "พักตร์วลัยพร (ปู)", "นิรินทร (อุ้ย)", "วิภาณี (จิ๋ม)", "นัทชา (เก้า)",
                              "ศิริรัตน์ (บี)", "จิรภา (แต๊ก)", "นิรนุช (ปัท)", "ปภัสรินทร์ (ออย)", "ศิริพร (ปุ๋ย)",
                              "พรเพชร (กิ๊ฟ)", "อรุณรัตน์ (เจี๊ยบ)", "กาญจนา (โอ๋)", "พรมงคล (แม็กซ์)",
                              "ฉลอง (โป้ง)", "พราวรวี (แพรว)", "วสุพร (กิ๊ฟ)", "วรัญชลี (เมย์)", "กิตติยา (แนน)",
                              "ปิยะรัตน์ (เบลล์)", "ปนัดดา (แหม่ม)", "ธารวิมล (นัท)", "เกียรติสุดา (ตาต้า)",
                              "ชัญญา (แชมป์)", "แสงเดือน (อั๋น)", "สุกานดา (ตุํกตา)", "พัชรี (ใหม่)",
                              "รุ่งนภา (แพท)", "เบญจวรรณ (จ๊อย)", "จันทร์ภรรัตน์ (อันโน)", "วีระยุทธ (เพ้นท์)",
                              "ฉัตรกมล (ปุ้ย)", "ณัฐฎนิช (มิน)", "วรัญญา (อ้าย)", "ขวัญเนตร (เจเจ)", "มานะ (ตั๊ก)",
                              "รัฎดาวรรณ (เชอรรี่)"]

        ws.cell(row=1, column=1, value='Assistant').fill = styles['header_fill']
        for col, date in enumerate(schedule.index, 2):
            cell = ws.cell(row=1, column=col, value=date.strftime('%d/%m'))
            cell.fill, cell.font = styles['header_fill'], styles['fonts']['header']
            if date.weekday() >= 5: cell.fill = styles['weekend_fill']
            if self.is_holiday(date): cell.fill = styles['holiday_fill']

        current_row = 2
        for assistant in ordered_assistants:
            if assistant not in self.assistants: continue
            ws.cell(row=current_row, column=1, value=assistant).fill = styles['header_fill']
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row + 1, end_column=1)
            ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="center", vertical="center")
            for col, date in enumerate(schedule.index, 2):
                cell1, cell2 = ws.cell(row=current_row, column=col), ws.cell(row=current_row + 1, column=col)
                for cell in [cell1, cell2]:
                    cell.border, cell.alignment, cell.font = styles['border'], Alignment(horizontal="center",
                                                                                         vertical="center"), Font(
                        bold=True, size=9)

                date_str = date.strftime('%Y-%m-%d')
                shifts = daily_summary_data.get(assistant, {}).get(date_str, [])
                note_text = self.special_notes.get(assistant, {}).get(date_str)
                if note_text: cell1.value, cell1.font = note_text, styles['fonts'].get('note',
                                                                                       Font(italic=True, size=9))

                if date_str in self.assistants[assistant]['holidays']:
                    cell2.value = 'OFF';
                    cell1.fill, cell2.fill = styles['off_fill'], styles['off_fill']
                else:
                    if not note_text and len(shifts) == 1:
                        shift = shifts[0];
                        cell2.value = shift
                        prefix = next((p for p in styles['fills'] if shift.startswith(p)), None)
                        if prefix:
                            fill = styles['fills'][prefix];
                            cell1.fill, cell2.fill, cell2.font = fill, fill, styles['fonts'].get(prefix,
                                                                                                 styles['fonts'][
                                                                                                     'default'])
                    elif (note_text and shifts) or len(shifts) > 1:
                        if shifts:
                            shift2 = shifts[0] if note_text else shifts[1];
                            cell2.value = shift2
                            prefix2 = next((p for p in styles['fills'] if shift2.startswith(p)), None)
                            if prefix2: cell2.fill, cell2.font = styles['fills'][prefix2], styles['fonts'].get(prefix2,
                                                                                                               styles[
                                                                                                                   'fonts'][
                                                                                                                   'default'])
                        if not note_text and len(shifts) > 1:
                            shift1 = shifts[0];
                            cell1.value = shift1
                            prefix1 = next((p for p in styles['fills'] if shift1.startswith(p)), None)
                            if prefix1: cell1.fill, cell1.font = styles['fills'][prefix1], styles['fonts'].get(prefix1,
                                                                                                               styles[
                                                                                                                   'fonts'][
                                                                                                                   'default'])

                    if (self.is_holiday(date) or date.weekday() >= 5) and not shifts and date_str not in \
                            self.assistants[assistant]['holidays']:
                        for cell in [cell1, cell2]:
                            if not cell.value: cell.fill = styles['holiday_empty_fill']
            current_row += 2

    def create_long_term_fairness_sheet(self, wb, schedule, daily_summary_data):
        ws = wb.create_sheet("LongTermFairnessInput")
        header_fill = PatternFill(start_color='B4C6E7', fill_type='solid')
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'),
                        bottom=Side(style='thin'))
        try:
            initial_df = pd.read_excel(self.excel_file_path, sheet_name='Assistants', index_col='Name')
        except Exception as e:
            ws.cell(row=1, column=1, value=f"FATAL ERROR reading original file: {e}");
            return

        shift_code_headers = sorted(list(self.shift_types.keys()))
        headers = ['Name'] + [f'Prev_{sc}' for sc in shift_code_headers] + ['Prev_Night_Shifts', 'Prev_Hours']
        for col, header_text in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header_text);
            cell.fill, cell.font, cell.border = header_fill, Font(bold=True), border

        ordered_assistants = ["วิภาดา (โอ)", "พักตร์วลัยพร (ปู)", "นิรินทร (อุ้ย)", "วิภาณี (จิ๋ม)", "นัทชา (เก้า)",
                              "ศิริรัตน์ (บี)", "จิรภา (แต๊ก)", "นิรนุช (ปัท)", "ปภัสรินทร์ (ออย)", "ศิริพร (ปุ๋ย)",
                              "พรเพชร (กิ๊ฟ)", "อรุณรัตน์ (เจี๊ยบ)", "กาญจนา (โอ๋)", "พรมงคล (แม็กซ์)", "ฉลอง (โป้ง)",
                              "พราวรวี (แพรว)"
            , "วสุพร (กิ๊ฟ)", "วรัญชลี (เมย์)", "กิตติยา (แนน)", "ปิยะรัตน์ (เบลล์)", "ปนัดดา (แหม่ม)", "ธารวิมล (นัท)",
                              "เกียรติสุดา (ตาต้า)", "ชัญญา (แชมป์)", "แสงเดือน (อั๋น)", "สุกานดา (ตุํกตา)",
                              "พัชรี (ใหม่)", "รุ่งนภา (แพท)", "เบญจวรรณ (จ๊อย)", "จันทร์ภรรัตน์ (อันโน)",
                              "วีระยุทธ (เพ้นท์)", "ฉัตรกมล (ปุ้ย)"
            , "ณัฐฎนิช (มิน)", "วรัญญา (อ้าย)", "ขวัญเนตร (เจเจ)", "มานะ (ตั๊ก)", "รัฎดาวรรณ (เชอรรี่)"]

        row_num = 2
        for name in ordered_assistants:
            if name not in initial_df.index: continue
            original_row = initial_df.loc[name]
            current_stats = {'hours': 0, 'nights': 0, 'shifts': {sc: 0 for sc in self.shift_types.keys()}}
            if name in daily_summary_data:
                for shifts in daily_summary_data[name].values():
                    for st in shifts:
                        current_stats['hours'] += self.shift_types[st]['hours']
                        current_stats['shifts'][st] += 1
                        if self.is_night_shift(st): current_stats['nights'] += 1

            ws.cell(row=row_num, column=1, value=name).border = border
            col_num = 2
            for sc in shift_code_headers:
                ws.cell(row=row_num, column=col_num,
                        value=int(original_row.get(f'Prev_{sc}', 0)) + current_stats['shifts'].get(sc,
                                                                                                   0)).border = border
                col_num += 1
            ws.cell(row=row_num, column=col_num,
                    value=int(original_row.get('Prev_Night_Shifts', 0)) + current_stats['nights']).border = border
            ws.cell(row=row_num, column=col_num + 1,
                    value=float(original_row.get('Prev_Hours', 0.0)) + current_stats['hours']).border = border
            row_num += 1

        for col_cells in ws.columns:
            max_length = max(len(str(cell.value)) for cell in col_cells if cell.value is not None)
            ws.column_dimensions[col_cells[0].column_letter].width = max_length + 2


# --- HTML Generation Functions ---

def generate_pharmacist_html_summary(schedule, scheduler):
    styles = {
        'header_fill': '#D3D3D3', 'weekend_fill': '#FFE4E1', 'holiday_fill': '#FFB6C1',
        'holiday_empty_fill': '#FFFF00', 'off_fill': '#D3D3D3',
        'fills': {
            'I100': '#00B050', 'O100': '#00B0F0', 'Care': '#D40202', 'C8': '#E6B8AF',
            'I400': '#FF00FF', 'O400F1': '#0033CC', 'O400F2': '#C78AF2', 'O400ER': '#ED7D31', 'ARI': '#7030A0'
        },
        'font_colors': {'O400F1': '#FFFFFF', 'ARI': '#FFFFFF', 'default': '#000000'}
    }
    ordered_pharmacists = ["ภญ.ประภัสสรา (มิ้น)", "ภญ.ฐิฏิการ (เอ้)", "ภก.บัณฑิตวงศ์ (แพท)", "ภก.ชานนท์ (บุ้ง)",
                           "ภญ.กมลพรรณ (ใบเตย)", "ภญ.กนกพร (นุ้ย)", "ภก.เอกวรรณ (โม)", "ภญ.อาภาภัทร (มะปราง)",
                           "ภก.ชวนันท์ (เท่ห์)", "ภญ.ธนพร (ฟ้า ธนพร)", "ภญ.วิลินดา (เชอร์รี่)", "ภญ.ชลนิชา (เฟื่อง)",
                           "ภญ.ปริญญ์ (ขมิ้น)", "ภก.ธนภรณ์ (กิ๊ฟ)", "ภญ.ปุณยวีร์ (มิ้นท์)", "ภญ.อมลกานต์ (บอม)",
                           "ภญ.อรรชนา (อ้อม)", "ภญ.ศศิวิมล (ฟิลด์)", "ภญ.วรรณิดา (ม่าน)", "ภญ.ปาณิศา (แบม)",
                           "ภญ.จิรัชญา (ศิกานต์)", "ภญ.อภิชญา (น้ำตาล)", "ภญ.วรางคณา (ณา)", "ภญ.ดวงดาว (ปลา)",
                           "ภญ.พรนภา (ผึ้ง)", "ภญ.ธนาภรณ์ (ลูกตาล)", "ภญ.วิลาสินี (เจ้นท์)", "ภญ.ภาวิตา (จูน)",
                           "ภญ.ศิรดา (พลอย)", "ภญ.ศุภิสรา (แพร)", "ภญ.กันต์หทัย (ซีน)", "ภญ.พัทธ์ธีรา (วิว)",
                           "ภญ.จุฑามาศ (กวาง)", 'ภญ. ณัฐพร (แอม)']
    sorted_dates = sorted(schedule.index)

    html = """
    <style>
        table, th, td {
            border: 1px solid #CCCCCC;
            border-collapse: collapse;
            text-align: center;
            vertical-align: middle;
            padding: 4px;
            font-family: sans-serif;
            font-size: 11px;
            white-space: nowrap;
        }
        th { font-weight: bold; }
        .name-col { width: 180px; }
        .date-col { width: 45px; }
        .note { 
            font-style: italic; 
            font-size: 10px; 
            word-wrap: break-word; 
            white-space: normal;
            line-height: 1.1;
        }
    </style>
    """
    html += "<table><thead><tr><th class='name-col' style='background-color: {bg};'>Pharmacist</th>".format(
        bg=styles['header_fill'])
    for date in sorted_dates:
        bg_color = styles['header_fill']
        if scheduler.is_holiday(date):
            bg_color = styles['holiday_fill']
        elif date.weekday() >= 5:
            bg_color = styles['weekend_fill']
        html += f"<th class='date-col' style='background-color: {bg_color};'>{date.strftime('%d/%m')}</th>"
    html += "</tr></thead><tbody>"

    for pharmacist in ordered_pharmacists:
        if pharmacist not in scheduler.pharmacists: continue

        # Build rows first to handle coloring of note cell
        row1_cells, row2_cells, row3_cells = "", "", ""

        for date in sorted_dates:
            date_str = date.strftime('%Y-%m-%d')
            shifts = scheduler.get_pharmacist_shifts(pharmacist, date, schedule)
            is_personal_holiday = date_str in scheduler.pharmacists[pharmacist]['holidays']
            is_off_day = (scheduler.is_holiday(date) or date.weekday() >= 5) and not shifts
            note_text = scheduler.special_notes.get(pharmacist, {}).get(date_str, '&nbsp;')

            # Default colors
            note_bg, cell1_bg, cell2_bg = '#FFFFFF', '#FFFFFF', '#FFFFFF'
            cell1_content, cell2_content = '&nbsp;', '&nbsp;'
            cell1_font_color, cell2_font_color = styles['font_colors']['default'], styles['font_colors']['default']

            if is_personal_holiday:
                note_bg = cell1_bg = cell2_bg = styles['off_fill']
                cell2_content = 'X'
            else:
                if len(shifts) == 1:
                    shift = shifts[0]
                    cell2_content = f"{int(scheduler.shift_types[shift]['hours'])}N" if scheduler.is_night_shift(
                        shift) else str(int(scheduler.shift_types[shift]['hours']))
                    prefix = next((p for p in styles['fills'] if shift.startswith(p)), None)
                    if prefix:
                        note_bg = cell1_bg = cell2_bg = styles['fills'][prefix]
                        cell2_font_color = styles['font_colors'].get(prefix, styles['font_colors']['default'])
                elif len(shifts) > 1:
                    # Cell 1 (bottom in excel)
                    shift1 = shifts[1]
                    cell1_content = f"{int(scheduler.shift_types[shift1]['hours'])}N" if scheduler.is_night_shift(
                        shift1) else str(int(scheduler.shift_types[shift1]['hours']))
                    prefix1 = next((p for p in styles['fills'] if shift1.startswith(p)), None)
                    if prefix1:
                        cell1_bg = styles['fills'][prefix1]
                        cell1_font_color = styles['font_colors'].get(prefix1, styles['font_colors']['default'])
                    # Cell 2 (top in excel)
                    shift2 = shifts[0]
                    cell2_content = f"{int(scheduler.shift_types[shift2]['hours'])}N" if scheduler.is_night_shift(
                        shift2) else str(int(scheduler.shift_types[shift2]['hours']))
                    prefix2 = next((p for p in styles['fills'] if shift2.startswith(p)), None)
                    if prefix2:
                        cell2_bg = styles['fills'][prefix2]
                        cell2_font_color = styles['font_colors'].get(prefix2, styles['font_colors']['default'])

                if is_off_day:
                    note_bg = cell1_bg = cell2_bg = styles['holiday_empty_fill']

            row1_cells += f"<td style='height: 25px; background-color:{note_bg};'><div class='note'>{note_text}</div></td>"
            row2_cells += f"<td style='background-color:{cell1_bg}; color:{cell1_font_color}; font-weight: bold;'>{cell1_content}</td>"
            row3_cells += f"<td style='background-color:{cell2_bg}; color:{cell2_font_color}; font-weight: bold;'>{cell2_content}</td>"

        html += f"<tr><td rowspan='3' class='name-col' style='background-color: {styles['header_fill']}; font-weight: bold;'>{pharmacist}</td>{row1_cells}</tr>"
        html += f"<tr>{row2_cells}</tr>"
        html += f"<tr>{row3_cells}</tr>"

    html += "</tbody></table>"
    return html


def generate_assistant_html_summary(schedule, scheduler):
    styles = {
        'header_fill': '#D3D3D3', 'weekend_fill': '#FFE4E1', 'holiday_fill': '#FFB6C1',
        'holiday_empty_fill': '#FFFF00', 'off_fill': '#D3D3D3',
        'fills': {
            'I100': '#00B050', 'O100': '#00B0F0', 'Care': '#D40202', 'C8': '#E6B8AF',
            'I400': '#FF00FF', 'O400F1': '#0033CC', 'O400F2': '#C78AF2', 'O400ER': '#ED7D31', 'ARI': '#7030A0'
        },
        'font_colors': {'O400F1': '#FFFFFF', 'ARI': '#FFFFFF', 'default': '#000000'}
    }
    ordered_assistants = ["วิภาดา (โอ)", "พักตร์วลัยพร (ปู)", "นิรินทร (อุ้ย)", "วิภาณี (จิ๋ม)", "นัทชา (เก้า)",
                          "ศิริรัตน์ (บี)", "จิรภา (แต๊ก)", "นิรนุช (ปัท)", "ปภัสรินทร์ (ออย)", "ศิริพร (ปุ๋ย)",
                          "พรเพชร (กิ๊ฟ)", "อรุณรัตน์ (เจี๊ยบ)", "กาญจนา (โอ๋)", "พรมงคล (แม็กซ์)", "ฉลอง (โป้ง)",
                          "พราวรวี (แพรว)"
        , "วสุพร (กิ๊ฟ)", "วรัญชลี (เมย์)", "กิตติยา (แนน)", "ปิยะรัตน์ (เบลล์)", "ปนัดดา (แหม่ม)", "ธารวิมล (นัท)",
                          "เกียรติสุดา (ตาต้า)", "ชัญญา (แชมป์)", "แสงเดือน (อั๋น)", "สุกานดา (ตุํกตา)", "พัชรี (ใหม่)",
                          "รุ่งนภา (แพท)", "เบญจวรรณ (จ๊อย)", "จันทร์ภรรัตน์ (อันโน)", "วีระยุทธ (เพ้นท์)",
                          "ฉัตรกมล (ปุ้ย)"
        , "ณัฐฎนิช (มิน)", "วรัญญา (อ้าย)", "ขวัญเนตร (เจเจ)", "มานะ (ตั๊ก)", "รัฎดาวรรณ (เชอรรี่)"]
    sorted_dates = sorted(schedule.index)

    html = """
    <style>
        table, th, td {
            border: 1px solid #CCCCCC;
            border-collapse: collapse;
            text-align: center;
            vertical-align: middle;
            padding: 4px;
            font-family: sans-serif;
            font-size: 11px;
            white-space: nowrap;
        }
        th { font-weight: bold; }
        .name-col { width: 180px; }
        .date-col { width: 45px; }
        .note { 
            font-style: italic; 
            font-size: 10px; 
            word-wrap: break-word; 
            white-space: normal;
            line-height: 1.1;
        }
    </style>
    """
    html += "<table><thead><tr><th class='name-col' style='background-color: {bg};'>Assistant</th>".format(
        bg=styles['header_fill'])
    for date in sorted_dates:
        bg_color = styles['header_fill']
        if scheduler.is_holiday(date):
            bg_color = styles['holiday_fill']
        elif date.weekday() >= 5:
            bg_color = styles['weekend_fill']
        html += f"<th class='date-col' style='background-color: {bg_color};'>{date.strftime('%d/%m')}</th>"
    html += "</tr></thead><tbody>"

    for assistant in ordered_assistants:
        if assistant not in scheduler.assistants: continue
        row1_cells, row2_cells = "", ""

        for date in sorted_dates:
            date_str = date.strftime('%Y-%m-%d')
            shifts = scheduler.get_assistant_shifts(assistant, date, schedule)
            note_text = scheduler.special_notes.get(assistant, {}).get(date_str, '')
            is_personal_holiday = date_str in scheduler.assistants[assistant]['holidays']
            is_off_day = (scheduler.is_holiday(date) or date.weekday() >= 5) and not shifts

            # Defaults
            cell1_content, cell2_content = '&nbsp;', '&nbsp;'
            cell1_bg, cell2_bg = '#FFFFFF', '#FFFFFF'
            cell1_font_color, cell2_font_color = styles['font_colors']['default'], styles['font_colors']['default']

            if is_personal_holiday:
                cell1_bg = cell2_bg = styles['off_fill']
                cell2_content = 'OFF'
            else:
                if note_text:
                    cell1_content = f"<div class='note'>{note_text}</div>"
                    if shifts:  # Note with one shift
                        shift = shifts[0]
                        cell2_content = f"{int(scheduler.shift_types[shift]['hours'])}N" if scheduler.is_night_shift(
                            shift) else str(int(scheduler.shift_types[shift]['hours']))
                        prefix = next((p for p in styles['fills'] if shift.startswith(p)), None)
                        if prefix:
                            cell2_bg = styles['fills'][prefix]
                            cell2_font_color = styles['font_colors'].get(prefix, styles['font_colors']['default'])
                elif len(shifts) == 1:
                    shift = shifts[0]
                    cell2_content = f"{int(scheduler.shift_types[shift]['hours'])}N" if scheduler.is_night_shift(
                        shift) else str(int(scheduler.shift_types[shift]['hours']))
                    prefix = next((p for p in styles['fills'] if shift.startswith(p)), None)
                    if prefix:
                        cell1_bg = cell2_bg = styles['fills'][prefix]
                        cell2_font_color = styles['font_colors'].get(prefix, styles['font_colors']['default'])
                elif len(shifts) > 1:
                    # Cell 1
                    shift1 = shifts[0]
                    cell1_content = f"{int(scheduler.shift_types[shift1]['hours'])}N" if scheduler.is_night_shift(
                        shift1) else str(int(scheduler.shift_types[shift1]['hours']))
                    prefix1 = next((p for p in styles['fills'] if shift1.startswith(p)), None)
                    if prefix1:
                        cell1_bg = styles['fills'][prefix1]
                        cell1_font_color = styles['font_colors'].get(prefix1, styles['font_colors']['default'])
                    # Cell 2
                    shift2 = shifts[1]
                    cell2_content = f"{int(scheduler.shift_types[shift2]['hours'])}N" if scheduler.is_night_shift(
                        shift2) else str(int(scheduler.shift_types[shift2]['hours']))
                    prefix2 = next((p for p in styles['fills'] if shift2.startswith(p)), None)
                    if prefix2:
                        cell2_bg = styles['fills'][prefix2]
                        cell2_font_color = styles['font_colors'].get(prefix2, styles['font_colors']['default'])

                if is_off_day and not is_personal_holiday:
                    cell1_bg = cell2_bg = styles['holiday_empty_fill']

            row1_cells += f"<td style='height: 25px; background-color:{cell1_bg}; color:{cell1_font_color}; font-weight: bold;'>{cell1_content}</td>"
            row2_cells += f"<td style='height: 25px; background-color:{cell2_bg}; color:{cell2_font_color}; font-weight: bold;'>{cell2_content}</td>"

        html += f"<tr><td rowspan='2' class='name-col' style='background-color: {styles['header_fill']}; font-weight: bold;'>{assistant}</td>{row1_cells}</tr>"
        html += f"<tr>{row2_cells}</tr>"

    html += "</tbody></table>"
    return html


# --- Streamlit UI and Main Execution Logic ---

st.set_page_config(layout="wide", page_title="IPSSS", page_icon="⚕️")
st.title("Intelligent Pharmacy Scheduling Support System")

# --- Sidebar for Inputs ---
with st.sidebar:
    st.image("https://github.com/Drugpurchasing/Shift/blob/main/IPSS.png?raw=true", width=100)
    st.title("⚙️ ตั้งค่า")

    scheduler_type = st.selectbox(
        "เลือกประเภทการจัดเวร",
        ("จัดเวรเภสัชกร", "จัดเวรผู้ช่วยเภสัชกร")
    )

    st.divider()

    if scheduler_type == "จัดเวรเภสัชกร":
        excel_url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRJonz3GVKwdpcEqXoZSvGGCWrFVBH12yklC9vE3cnMCqtE-MOTGE-mwsE7pJBBYA/pub?output=xlsx"
        st.info("ใช้ข้อมูล **เภสัชกร** จาก Google Sheet")
    else:  # จัดเวรผู้ช่วยเภสัชกร
        excel_url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSdfJ47lznKm89OatoeciaWoXTXoTtakCaLIDXWYRCZ1hqEy91YBoK80Ih7EosfDQ/pub?output=xlsx"
        st.info("ใช้ข้อมูล **ผู้ช่วยเภสัชกร** จาก Google Sheet")

    mode = st.radio(
        "เลือกรูปแบบ",
        ("จัดทั้งเดือน", "จัดเฉพาะวันที่เลือก"),
        horizontal=True
    )

    # --- Calculate default month and year ---
    current_date = datetime.now()
    default_month = current_date.month + 1
    default_year = current_date.year
    if default_month > 12:
        default_month = 1
        default_year += 1

    dates_to_schedule, year, month = [], 0, 0
    if mode == "จัดทั้งเดือน":
        col1, col2 = st.columns(2)
        with col1:
            year = st.number_input("ปี (ค.ศ.)", min_value=2020, max_value=2050, value=default_year)
        with col2:
            month = st.number_input("เดือน", min_value=1, max_value=12, value=default_month)
    else:  # Specific Dates
        date_range = st.date_input(
            "เลือกช่วงวันที่",
            value=(datetime.now().date(), datetime.now().date() + timedelta(days=2)),
            min_value=datetime(2020, 1, 1)
        )
        if len(date_range) == 2:
            dates_to_schedule = pd.date_range(start=date_range[0], end=date_range[1]).to_pydatetime().tolist()
        elif len(date_range) == 1:
            dates_to_schedule = [date_range[0]]

    st.divider()

    if scheduler_type == "จัดเวรเภสัชกร":
        iterations = st.slider(
            "จำนวนรอบในการหาผลลัพธ์ (เภสัชกร)",
            min_value=5, max_value=300, value=50, step=5,
            help="ยิ่งเยอะ ยิ่งมีโอกาสได้ตารางที่ดีขึ้น แต่จะใช้เวลานานขึ้น"
        )
    else:
        iterations = st.slider(
            "จำนวนรอบในการหาผลลัพธ์ (ผู้ช่วย)",
            min_value=5, max_value=500, value=10,
            help="ยิ่งเยอะ ยิ่งมีโอกาสได้ตารางที่ดีขึ้น แต่จะใช้เวลานานขึ้น"
        )

    st.divider()

    run_button = st.button("🚀 เริ่มจัดตารางเวร", type="primary", use_container_width=True)

# --- Main Area for Output ---
if run_button:
    log_container = st.empty()


    def user_friendly_logger(message):
        log_container.info(message)


    try:
        data_load_progress = st.progress(0, text="กำลังเตรียมโหลดข้อมูล...")

        best_schedule, best_unfilled_info = None, None
        excel_buffer = None
        scheduler_instance = None  # To hold the final scheduler instance

        # --- Pharmacist Scheduler Logic ---
        if scheduler_type == "จัดเวรเภสัชกร":
            scheduler = PharmacistScheduler(excel_url, logger=user_friendly_logger, progress_bar=data_load_progress)
            scheduler_instance = scheduler
            data_load_progress.progress(100, text="โหลดข้อมูลเภสัชกรสำเร็จ!")

            optimization_progress_placeholder = st.empty()
            with optimization_progress_placeholder.container():
                opt_progress_bar = st.progress(0, text="กำลังเตรียมการคำนวณ...")
                if mode == "จัดทั้งเดือน":
                    best_schedule, best_unfilled_info = scheduler.optimize_schedule(year, month, iterations,
                                                                                    opt_progress_bar)
                else:
                    if not dates_to_schedule:
                        st.error("กรุณาเลือกช่วงวันที่ที่ถูกต้อง")
                    else:
                        best_schedule, best_unfilled_info = scheduler.optimize_schedule_for_dates(dates_to_schedule,
                                                                                                  iterations,
                                                                                                  opt_progress_bar)

            optimization_progress_placeholder.empty()
            if best_schedule is not None:
                excel_buffer = scheduler.export_to_excel(best_schedule, best_unfilled_info)

        # --- Assistant Scheduler Logic ---
        else:
            scheduler = AssistantScheduler(excel_url, logger=user_friendly_logger, progress_bar=data_load_progress)
            scheduler_instance = scheduler
            data_load_progress.progress(100, text="โหลดข้อมูลผู้ช่วยเภสัชกรสำเร็จ!")

            final_dates = []
            if mode == "จัดทั้งเดือน":
                start_date = datetime(year, month, 1)
                end_date = datetime(year, month + 1, 1) - timedelta(days=1) if month < 12 else datetime(year, 12, 31)
                final_dates = pd.date_range(start_date, end_date)
            else:
                if not dates_to_schedule:
                    st.error("กรุณาเลือกช่วงวันที่ที่ถูกต้อง")
                else:
                    final_dates = pd.to_datetime(dates_to_schedule).sort_values()

            if len(final_dates) > 0:
                optimization_progress_placeholder = st.empty()
                with optimization_progress_placeholder.container():
                    opt_progress_bar = st.progress(0, text="กำลังเตรียมการคำนวณ...")
                    scheduler.progress_bar = opt_progress_bar
                    best_schedule = scheduler.optimize_schedule(final_dates, iterations)

                optimization_progress_placeholder.empty()
                if best_schedule is not None:
                    scheduler.suggest_negotiations_for_unfilled(best_schedule)
                    excel_buffer = scheduler.export_to_excel(best_schedule)

        # --- Clear logs and display results ---
        log_container.empty()
        data_load_progress.empty()

        if best_schedule is not None and excel_buffer is not None:
            st.success("✅ จัดตารางเวรสำเร็จ!")

            type_prefix = "Pharmacist" if scheduler_type == "จัดเวรเภสัชกร" else "Assistant"
            if mode == "จัดทั้งเดือน":
                output_filename = f"{type_prefix}_Schedule_{year}_{month:02d}.xlsx"
            else:
                date_str = dates_to_schedule[0].strftime('%Y%m%d')
                output_filename = f"{type_prefix}_Schedule_Custom_{date_str}.xlsx"

            st.download_button(
                label="📥 ดาวน์โหลดตารางเวรทั้งหมด (ไฟล์ Excel)",
                data=excel_buffer,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            st.divider()

            # --- Display Styled HTML Summary directly on the page ---
            st.header("🗓️ ตารางสรุปรายวัน (Daily Summary)")
            if scheduler_type == "จัดเวรเภสัชกร":
                html_summary = generate_pharmacist_html_summary(best_schedule, scheduler_instance)
            else:
                html_summary = generate_assistant_html_summary(best_schedule, scheduler_instance)

            st.markdown(html_summary, unsafe_allow_html=True)  # <-- This renders the HTML table

            st.divider()

            # --- Display Other Sheets in Expanders ---
            st.header("📄 ตารางรูปแบบอื่นๆ (ที่มีในไฟล์ Excel)")
            xls = pd.ExcelFile(excel_buffer)
            sheet_names = xls.sheet_names
            for sheet_name in sheet_names:
                if sheet_name != 'Daily Summary':
                    with st.expander(f"ดูตาราง: {sheet_name}"):
                        df = pd.read_excel(xls, sheet_name=sheet_name)
                        st.dataframe(df)
        elif run_button:
            st.error("❌ ไม่สามารถสร้างตารางเวรได้ กรุณาตรวจสอบข้อจำกัดต่างๆ หรือลองเพิ่มจำนวนรอบการคำนวณ")

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดที่ไม่คาดคิด: {e}")
        st.error(
            "อาจเกิดจากปัญหาการเชื่อมต่ออินเทอร์เน็ต, รูปแบบไฟล์ Google Sheet เปลี่ยนไป, หรือลิงก์ไม่ถูกต้อง กรุณาตรวจสอบและลองอีกครั้ง")



