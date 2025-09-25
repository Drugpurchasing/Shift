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
import time

# --- The PharmacistScheduler Class (with modifications for progress bar) ---
# เพิ่มการรับออบเจ็กต์ progress bar เข้ามาเพื่ออัปเดตสถานะการโหลดข้อมูล

class PharmacistScheduler:
    W_CONSECUTIVE = 8
    W_HOURS = 4
    W_PREFERENCE = 4

    def __init__(self, excel_file_path, logger=print, progress_bar=None):
        self.logger = logger
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

        # --- Progress Bar Logic ---
        total_steps = 7
        
        self.read_data_from_excel(self.excel_file_path, progress_bar, 0, total_steps)
        self.load_historical_scores(self.excel_file_path, progress_bar, 6, total_steps)
        
        self._calculate_preference_multipliers()
        self.night_shifts = {'I100-10', 'I100-12N', 'I400-12N', 'I400-10', 'O400ER-12N', 'O400ER-10'}
        self.holidays = {'specific_dates': ['2025-10-13','2025-10-23']}
        for pharmacist in self.pharmacists:
            self.pharmacists[pharmacist]['shift_counts'] = {st: 0 for st in self.shift_types}

    def read_data_from_excel(self, file_path_or_url, progress_bar, start_step, total_steps):
        step = start_step
        
        def update_progress(sheet_name):
            nonlocal step
            step += 1
            if progress_bar:
                text = f"กำลังโหลดข้อมูล: {sheet_name}... ({step}/{total_steps})"
                progress_bar.progress(step / total_steps, text=text)
            time.sleep(0.1) #หน่วงเวลาเล็กน้อยเพื่อให้เห็น progress

        # --- Read Sheets ---
        pharmacists_df = pd.read_excel(file_path_or_url, sheet_name='Pharmacists', engine='openpyxl')
        update_progress("รายชื่อเภสัชกร")
        # ... processing logic ...
        self.pharmacists = {}
        for _, row in pharmacists_df.iterrows():
            name = row['Name']
            max_hours = row.get('Max Hours', 250)
            if pd.isna(max_hours) or max_hours == '' or max_hours is None: max_hours = 250
            else: max_hours = float(max_hours)
            self.pharmacists[name] = {
                'night_shift_count': 0, 'skills': str(row['Skills']).split(','),
                'holidays': [d for d in str(row['Holidays']).split(',') if d != '1900-01-00' and d.strip() and d != 'nan'],
                'shift_counts': {}, 'preferences': {f'rank{i}': row[f'Rank{i}'] for i in range(1, 9)}, 'max_hours': max_hours
            }
            
        shifts_df = pd.read_excel(file_path_or_url, sheet_name='Shifts', engine='openpyxl')
        update_progress("ประเภทกะ")
        # ... processing logic ...
        self.shift_types = {}
        for _, row in shifts_df.iterrows():
            shift_code = row['Shift Code']
            self.shift_types[shift_code] = {
                'description': row['Description'], 'shift_type': row['Shift Type'], 'start_time': row['Start Time'],
                'end_time': row['End Time'], 'hours': row['Hours'], 'required_skills': str(row['Required Skills']).split(','),
                'restricted_next_shifts': str(row['Restricted Next Shifts']).split(',') if pd.notna(row['Restricted Next Shifts']) else []
            }

        departments_df = pd.read_excel(file_path_or_url, sheet_name='Departments', engine='openpyxl')
        update_progress("แผนก")
        # ... processing logic ...
        self.departments = {}
        for _, row in departments_df.iterrows():
            self.departments[row['Department']] = str(row['Shift Codes']).split(',')
        
        pre_assign_df = pd.read_excel(file_path_or_url, sheet_name='PreAssignments', engine='openpyxl')
        update_progress("กะที่กำหนดล่วงหน้า")
        # ... processing logic ...
        pre_assign_df['Date'] = pd.to_datetime(pre_assign_df['Date']).dt.strftime('%Y-%m-%d')
        self.pre_assignments = {p: g.set_index('Date')['Shift'].apply(lambda x: [s.strip() for s in str(x).split(',') if s.strip()]).to_dict() for p, g in pre_assign_df.groupby('Pharmacist')}
        
        try:
            notes_df = pd.read_excel(file_path_or_url, sheet_name='SpecialNotes', index_col=0, engine='openpyxl')
            # ... processing logic ...
            for pharmacist, row_data in notes_df.iterrows():
                if pharmacist in self.pharmacists:
                    for date_col, note in row_data.items():
                        if pd.notna(note) and str(note).strip():
                            date_str = pd.to_datetime(date_col).strftime('%Y-%m-%d')
                            if pharmacist not in self.special_notes: self.special_notes[pharmacist] = {}
                            self.special_notes[pharmacist][date_str] = str(note).strip()
        except Exception: pass
        update_progress("หมายเหตุพิเศษ")
        
        try:
            limits_df = pd.read_excel(file_path_or_url, sheet_name='ShiftLimits', engine='openpyxl')
            # ... processing logic ...
            for _, row in limits_df.iterrows():
                p = row['Pharmacist']
                if p in self.pharmacists:
                    if p not in self.shift_limits: self.shift_limits[p] = {}
                    self.shift_limits[p][row['ShiftCategory']] = int(row['MaxCount'])
        except Exception: pass
        update_progress("โควต้ากะ")

    def load_historical_scores(self, file_path_or_url, progress_bar, start_step, total_steps):
        try:
            df = pd.read_excel(file_path_or_url, sheet_name='HistoricalScores', engine='openpyxl')
            if 'Pharmacist' in df.columns and 'Total Preference Score' in df.columns:
                for _, row in df.iterrows():
                    if row['Pharmacist'] in self.pharmacists:
                        self.historical_scores[row['Pharmacist']] = row['Total Preference Score']
        except Exception: pass
        if progress_bar:
            text = f"กำลังโหลดข้อมูล: คะแนนย้อนหลัง... ({start_step+1}/{total_steps})"
            progress_bar.progress((start_step+1) / total_steps, text=text)
        time.sleep(0.1)
    
    # ... The rest of the class methods are unchanged ...
    # ... Paste all other methods from the previous answer here ...
    def _calculate_preference_multipliers(self):
        if not self.historical_scores:
            self.logger("No historical scores found. All preference multipliers will be 1.0.")
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
                self.logger(f"Pharmacist '{pharmacist}' not in historical data. Assigning a favorable multiplier of {min_multiplier}.")
    def _pre_check_staffing_levels(self, year, month):
        self.logger("\nRunning pre-check for staffing levels (including all shifts + 3 buffer)...")
        start_date = datetime(year, month, 1)
        end_date = datetime(year, month + 1, 1) - timedelta(days=1) if month < 12 else datetime(year, 12, 31)
        dates = pd.date_range(start_date, end_date)
        all_ok = True
        for date in dates:
            available_pharmacists_count = sum(1 for p_name, p_info in self.pharmacists.items() if date.strftime('%Y-%m-%d') not in p_info['holidays'])
            required_shifts_base = sum(1 for st in self.shift_types if self.is_shift_available_on_date(st, date))
            total_required_shifts_with_buffer = required_shifts_base + 3
            if available_pharmacists_count < total_required_shifts_with_buffer:
                all_ok = False
                self.problem_days.add(date)
                self.logger(f"WARNING: Potential shortage on {date.strftime('%Y-%m-%d')}. Available Pharmacists: {available_pharmacists_count}, Required Shifts (with +3 buffer): {total_required_shifts_with_buffer}")
        if all_ok: self.logger("Pre-check complete. All days have sufficient staffing levels for the total workload.")
        else: self.logger("Pre-check complete. Identified days with potential staff shortages. These will be prioritized.")
        return not all_ok
    def convert_time_to_minutes(self, time_input):
        if isinstance(time_input, str): hours, minutes = map(int, time_input.split(':'))
        elif isinstance(time_input, time): hours, minutes = time_input.hour, time_input.minute
        else: raise ValueError("Invalid input type. Expected string (HH:MM) or datetime.time object.")
        return hours * 60 + minutes
    def check_time_overlap(self, start1, end1, start2, end2):
        start1_mins, end1_mins, start2_mins, end2_mins = self.convert_time_to_minutes(start1), self.convert_time_to_minutes(end1), self.convert_time_to_minutes(start2), self.convert_time_to_minutes(end2)
        if end1_mins < start1_mins: end1_mins += 1440
        if end2_mins < start2_mins: end2_mins += 1440
        return start1_mins < end2_mins and end1_mins > start2_mins
    def check_mixing_expert_ratio_optimized(self, schedule_dict, date, current_shift=None, current_pharm=None):
        mixing_shifts = [p for s, p in schedule_dict[date].items() if s.startswith('C8') and p not in ['NO SHIFT', 'UNASSIGNED', 'UNFILLED']]
        if current_shift and current_shift.startswith('C8') and current_pharm: mixing_shifts.append(current_pharm)
        if not mixing_shifts: return True
        return sum(1 for pharm in mixing_shifts if pharm in self.pharmacists and 'mixing_expert' in self.pharmacists[pharm]['skills']) >= (2 * len(mixing_shifts) / 3)
    def count_consecutive_shifts(self, pharmacist, date, schedule, max_days=6):
        count, current_date = 0, date - timedelta(days=1)
        for _ in range(max_days):
            if current_date in schedule.index and pharmacist in schedule.loc[current_date].values:
                count += 1
                current_date -= timedelta(days=1)
            else: break
        return count
    def is_holiday(self, date): return date.strftime('%Y-%m-%d') in self.holidays['specific_dates']
    def calculate_weekend_off_variance(self, schedule, year, month):
        weekend_off_counts = {p: 0 for p in self.pharmacists}
        start_date, end_date = datetime(year, month, 1), datetime(year, month + 1, 1) - timedelta(days=1) if month < 12 else datetime(year, 12, 31)
        for date in pd.date_range(start_date, end_date):
            if date.weekday() >= 5:
                working = {schedule.loc[date, s] for s in schedule.columns if schedule.loc[date, s] not in ['NO SHIFT', 'UNFILLED', 'UNASSIGNED']}
                for p_name in self.pharmacists:
                    if p_name not in working: weekend_off_counts[p_name] += 1
        return np.var(list(weekend_off_counts.values())) if len(weekend_off_counts) > 1 else 0
    def is_night_shift(self, shift_type): return shift_type in self.night_shifts
    def is_shift_available_on_date(self, shift_type, date):
        info, is_holiday, is_sat, is_sun = self.shift_types[shift_type], self.is_holiday(date), date.weekday() == 5, date.weekday() == 6
        if info['shift_type'] == 'weekday': return not (is_holiday or is_sat or is_sun)
        elif info['shift_type'] == 'saturday': return is_sat and not is_holiday
        elif info['shift_type'] == 'holiday': return is_holiday or is_sat or is_sun
        elif info['shift_type'] == 'night': return True
        return False
    def get_department_from_shift(self, shift_type):
        if shift_type.startswith('I100'): return 'IPD100'
        if shift_type.startswith('O100'): return 'OPD100'
        if shift_type.startswith('Care'): return 'Care'
        if shift_type.startswith('C8'): return 'Mixing'
        if shift_type.startswith('I400'): return 'IPD400'
        if shift_type.startswith('O400F1'): return 'OPD400F1'
        if shift_type.startswith('O400F2'): return 'OPD400F2'
        if shift_type.startswith('O400ER'): return 'ER'
        if shift_type.startswith('ARI'): return 'ARI'
        return None
    def _get_shift_category(self, shift_type):
        if self.is_night_shift(shift_type): return 'Night'
        if shift_type.startswith('C8'): return 'Mixing'
        return None
    def get_night_shift_count(self, pharmacist): return self.pharmacists[pharmacist]['night_shift_count']
    def get_preference_score(self, pharmacist, shift_type):
        dept = self.get_department_from_shift(shift_type)
        for rank in range(1, 9):
            if self.pharmacists[pharmacist]['preferences'][f'rank{rank}'] == dept: return rank
        return 9
    def has_restricted_sequence_optimized(self, pharmacist, date, shift_type, schedule_dict):
        prev_date = date - timedelta(days=1)
        if prev_date in schedule_dict:
            for prev_shift, p in schedule_dict[prev_date].items():
                if p == pharmacist and shift_type in self.shift_types[prev_shift].get('restricted_next_shifts', []): return True
        return False
    def has_overlapping_shift_optimized(self, pharmacist, date, new_shift_type, schedule_dict):
        if date not in schedule_dict: return False
        new_start, new_end = self.shift_types[new_shift_type]['start_time'], self.shift_types[new_shift_type]['end_time']
        for es, p in schedule_dict[date].items():
            if p == pharmacist and es != new_shift_type:
                if self.check_time_overlap(new_start, new_end, self.shift_types[es]['start_time'], self.shift_types[es]['end_time']): return True
        return False
    def has_nearby_night_shift_optimized(self, pharmacist, date, schedule_dict):
        for delta in [-2, -1, 1, 2]:
            check_date = date + timedelta(days=delta)
            if check_date in schedule_dict:
                for shift, p in schedule_dict[check_date].items():
                    if p == pharmacist and self.is_night_shift(shift): return True
        return False
    def get_pharmacist_shifts(self, pharmacist, date, current_schedule):
        return [st for st in current_schedule.columns if date in current_schedule.index and current_schedule.loc[date, st] == pharmacist]
    def calculate_total_hours(self, pharmacist, schedule):
        return sum(self.shift_types[st]['hours'] for date in schedule.index for st, p in schedule.loc[date].items() if p == pharmacist and st in self.shift_types)
    def _get_hour_imbalance_penalty(self, hours_dict):
        if not hours_dict or len(hours_dict) < 2: return 0
        vals = list(hours_dict.values())
        hour_range = max(vals) - min(vals)
        return stdev(vals) ** 2 + ((hour_range - 10) ** 2 if hour_range > 10 else 0)
    def calculate_schedule_metrics(self, schedule, year, month):
        hours = {p: self.calculate_total_hours(p, schedule) for p in self.pharmacists}
        night_counts = {p: self.pharmacists[p]['night_shift_count'] for p in self.pharmacists}
        metrics = {'hour_imbalance_penalty': self._get_hour_imbalance_penalty(hours),
                   'night_variance': np.var(list(night_counts.values())) if night_counts else 0,
                   'preference_score': sum(self.calculate_preference_penalty(p, schedule) for p in self.pharmacists),
                   'weekend_off_variance': self.calculate_weekend_off_variance(schedule, year, month)}
        metrics['hour_diff_for_logging'] = stdev(hours.values()) if len(hours) > 1 else 0
        return metrics
    def generate_monthly_schedule_shuffled(self, year, month, progress_bar, shuffled_shifts=None, shuffled_pharmacists=None, iteration_num=1):
        start_date, end_date = datetime(year, month, 1), datetime(year, month + 1, 1) - timedelta(days=1) if month < 12 else datetime(year, 12, 31)
        dates = pd.date_range(start_date, end_date)
        schedule_dict = {date: {shift: 'NO SHIFT' for shift in self.shift_types} for date in dates}
        pharmacist_hours, pharmacist_consecutive_days = {p: 0 for p in self.pharmacists}, {p: 0 for p in self.pharmacists}
        if shuffled_shifts is None: shuffled_shifts = random.sample(list(self.shift_types.keys()), len(self.shift_types))
        if shuffled_pharmacists is None: shuffled_pharmacists = random.sample(list(self.pharmacists.keys()), len(self.pharmacists))
        for p in self.pharmacists: self.pharmacists[p]['night_shift_count'] = 0; self.pharmacists[p]['mixing_shift_count'] = 0; self.pharmacists[p]['category_counts'] = {'Mixing': 0, 'Night': 0}
        for p, assignments in self.pre_assignments.items():
            if p not in self.pharmacists: continue
            for date_str, shift_types in assignments.items():
                date = pd.to_datetime(date_str)
                if date not in schedule_dict: continue
                for st in shift_types:
                    if st in self.shift_types: schedule_dict[date][st] = p; self._update_shift_counts(p, st); pharmacist_hours[p] += self.shift_types[st]['hours']
        all_dates = list(dates); problem_dates_sorted = sorted([d for d in all_dates if d in self.problem_days]); other_dates_sorted = sorted([d for d in all_dates if d not in self.problem_days]); processing_order_dates = problem_dates_sorted + other_dates_sorted
        unfilled_info = {'problem_days': [], 'other_days': []}
        night_shifts = [s for s in shuffled_shifts if self.is_night_shift(s)]; mixing_shifts = [s for s in shuffled_shifts if s.startswith('C8') and not self.is_night_shift(s)]; care_shifts = [s for s in shuffled_shifts if s.startswith('Care') and not self.is_night_shift(s) and not s.startswith('C8')]; other_shifts = [s for s in shuffled_shifts if not self.is_night_shift(s) and not s.startswith('C8') and not s.startswith('Care')]
        standard_order, problem_order = night_shifts + mixing_shifts + care_shifts + other_shifts, mixing_shifts + care_shifts + night_shifts + other_shifts
        for i, date in enumerate(processing_order_dates):
            if progress_bar: progress_bar.progress((i + 1) / len(processing_order_dates), text=f"Iteration {iteration_num}: กำลังจัดตารางวันที่ {date.strftime('%d/%m')}...")
            yesterday_workers = {p for p in schedule_dict[date - timedelta(days=1)].values() if p in self.pharmacists} if date - timedelta(days=1) in schedule_dict else set()
            for p in self.pharmacists: pharmacist_consecutive_days[p] = pharmacist_consecutive_days[p] + 1 if p in yesterday_workers else 0
            shifts_to_process = problem_order if date in self.problem_days else standard_order
            for st in shifts_to_process:
                if schedule_dict[date][st] != 'NO SHIFT' or not self.is_shift_available_on_date(st, date): continue
                available = self._get_available_pharmacists_optimized(shuffled_pharmacists, date, st, schedule_dict, pharmacist_hours, pharmacist_consecutive_days)
                if available:
                    chosen = self._select_best_pharmacist(available, st, date, (date + timedelta(days=1)) in self.problem_days)
                    p_assign = chosen['name']; schedule_dict[date][st] = p_assign; self._update_shift_counts(p_assign, st); pharmacist_hours[p_assign] += self.shift_types[st]['hours']
                else:
                    schedule_dict[date][st] = 'UNFILLED'
                    if date in self.problem_days: unfilled_info['problem_days'].append((date, st))
                    else: unfilled_info['other_days'].append((date, st)); return pd.DataFrame.from_dict(schedule_dict, orient='index').fillna('NO SHIFT'), unfilled_info
        final_schedule = pd.DataFrame.from_dict(schedule_dict, orient='index'); return final_schedule.reindex(columns=list(self.shift_types.keys()), fill_value='NO SHIFT').fillna('NO SHIFT'), unfilled_info
    def _update_shift_counts(self, pharmacist, shift_type):
        if self.is_night_shift(shift_type): self.pharmacists[pharmacist]['night_shift_count'] += 1
        if shift_type.startswith('C8'): self.pharmacists[pharmacist]['mixing_shift_count'] += 1
        category = self._get_shift_category(shift_type)
        if category and category in self.pharmacists[pharmacist]['category_counts']: self.pharmacists[pharmacist]['category_counts'][category] += 1
    def _get_available_pharmacists_optimized(self, pharmacists, date, shift_type, schedule_dict, current_hours_dict, consecutive_days_dict):
        available = []
        night_yesterday = {p for s, p in schedule_dict[date - timedelta(days=1)].items() if p in self.pharmacists and self.is_night_shift(s)} if date - timedelta(days=1) in schedule_dict else set()
        for p in pharmacists:
            if date.strftime('%Y-%m-%d') in self.pharmacists[p]['holidays'] or self.has_overlapping_shift_optimized(p, date, shift_type, schedule_dict) or p in night_yesterday: continue
            if not all(skill.strip() in self.pharmacists[p]['skills'] for skill in self.shift_types[shift_type]['required_skills'] if skill.strip()): continue
            if current_hours_dict[p] + self.shift_types[shift_type]['hours'] > self.pharmacists[p].get('max_hours', 250): continue
            if self.has_restricted_sequence_optimized(p, date, shift_type, schedule_dict): continue
            category = self._get_shift_category(shift_type)
            if category:
                limit = self.shift_limits.get(p, {}).get(category)
                if limit is not None and self.pharmacists[p]['category_counts'][category] >= limit: continue
            if self.is_night_shift(shift_type):
                if self.has_nearby_night_shift_optimized(p, date, schedule_dict): continue
                if p in self.pre_assignments and (date + timedelta(days=1)).strftime('%Y-%m-%d') in self.pre_assignments[p]: continue
            if shift_type.startswith('C8') and not self.check_mixing_expert_ratio_optimized(schedule_dict, date, shift_type, p): continue
            pref = self.get_preference_score(p, shift_type) * self.preference_multipliers.get(p, 1.0)
            available.append({'name': p, 'preference_score': pref, 'consecutive_days': consecutive_days_dict[p], 'night_count': self.pharmacists[p]['night_shift_count'], 'mixing_count': self.pharmacists[p]['mixing_shift_count'], 'current_hours': current_hours_dict[p]})
        return available
    def _calculate_suitability_score(self, d): return self.W_CONSECUTIVE * (d['consecutive_days'] ** 2) + self.W_HOURS * d['current_hours'] + self.W_PREFERENCE * d['preference_score']
    def _select_best_pharmacist(self, available, shift_type, date, is_day_before_problem_day):
        if self.is_night_shift(shift_type) and is_day_before_problem_day:
            problem_day_str = (date + timedelta(days=1)).strftime('%Y-%m-%d')
            off_tomorrow = [p for p in available if problem_day_str in self.pharmacists[p['name']]['holidays']]
            if off_tomorrow: self.logger(f"INFO: Prioritizing night shift on {date.strftime('%Y-%m-%d')} for pharmacists off on problem day {problem_day_str}."); return min(off_tomorrow, key=lambda x: (x['night_count'], self._calculate_suitability_score(x)))
        if self.is_night_shift(shift_type): return min(available, key=lambda x: (x['night_count'], self._calculate_suitability_score(x)))
        elif shift_type.startswith('C8'): return min(available, key=lambda x: (x['mixing_count'], self._calculate_suitability_score(x)))
        else: return min(available, key=lambda x: self._calculate_suitability_score(x))
    def calculate_preference_penalty(self, pharmacist, schedule): return sum(self.get_preference_score(pharmacist, st) for date in schedule.index for st, p in schedule.loc[date].items() if p == pharmacist)
    def is_schedule_better(self, current_metrics, best_metrics):
        current_unfilled, best_unfilled = current_metrics.get('unfilled_problem_shifts', float('inf')), best_metrics.get('unfilled_problem_shifts', float('inf'))
        if current_unfilled != best_unfilled: return current_unfilled < best_unfilled
        weights = {'preference_score': 1.0, 'hour_imbalance_penalty': 25.0, 'night_variance': 800.0, 'weekend_off_variance': 1000.0}
        current_score = sum(weights[k] * current_metrics.get(k, 0) for k in weights)
        best_score = sum(weights[k] * best_metrics.get(k, 0) for k in weights)
        return current_score < best_score
    def optimize_schedule(self, year, month, iterations, progress_bar):
        best_schedule, best_metrics, best_unfilled_info = None, {'unfilled_problem_shifts': float('inf'), 'preference_score': float('inf')}, {}
        self._pre_check_staffing_levels(year, month)
        self.logger(f"\nStarting optimization with {iterations} iterations...")
        for i in range(iterations):
            self.logger(f"\n--- Iteration {i+1}/{iterations} ---")
            current_schedule, unfilled_info = self.generate_monthly_schedule_shuffled(year, month, progress_bar, iteration_num=i+1)
            if unfilled_info['other_days']: continue
            metrics = self.calculate_schedule_metrics(current_schedule, year, month)
            metrics['unfilled_problem_shifts'] = len(unfilled_info['problem_days'])
            self.logger(f"Iteration Results -> Unfilled: {metrics['unfilled_problem_shifts']} | Hour SD: {metrics.get('hour_diff_for_logging', 0):.2f} | Night Var: {metrics.get('night_variance', 0):.2f} | Pref Pen: {metrics.get('preference_score', 0):.1f}")
            if best_schedule is None or self.is_schedule_better(metrics, best_metrics):
                best_schedule, best_metrics, best_unfilled_info = current_schedule.copy(), metrics.copy(), unfilled_info.copy()
                self.logger("*** Found a more balanced schedule! ***")
        if best_schedule is not None: self.logger(f"\nOptimization complete! Final metrics: Unfilled: {best_metrics.get('unfilled_problem_shifts', 0)} | Hour SD: {best_metrics.get('hour_diff_for_logging', 0):.2f} | Night Var: {best_metrics.get('night_variance', 0):.2f} | Pref Pen: {best_metrics.get('preference_score', 0):.1f}")
        else: self.logger("\nOptimization failed to find any valid schedule.")
        return best_schedule, best_unfilled_info
    def export_to_excel(self, schedule, unfilled_info):
        wb = Workbook()
        # ... (rest of export unchanged)
        ws, ws_daily, ws_daily_codes, ws_pref, ws_negotiate = wb.active, wb.create_sheet("Daily Summary"), wb.create_sheet("Daily Summary (Codes)"), wb.create_sheet("Preference Scores"), wb.create_sheet("Negotiation Suggestions")
        ws.title = 'Monthly Schedule'
        # ... (rest of export unchanged)
        buffer = io.BytesIO(); wb.save(buffer); buffer.seek(0); return buffer
    # All `create_...` and specific date methods are unchanged and should be pasted here

# --- Streamlit UI and Main Execution Logic ---

st.set_page_config(layout="wide", page_title="Pharmacist Scheduler")
st.title("⚕️ Pharmacist Shift Scheduler")

# --- Sidebar for Inputs ---
with st.sidebar:
    st.header("⚙️ ตั้งค่าการจัดตาราง")
    
    excel_url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRJonz3GVKwdpcEqXoZSvGGCWrFVBH12yklC9vE3cnMCqtE-MOTGE-mwsE7pJBBYA/pub?output=xlsx"
    st.info("ใช้ข้อมูลจาก Google Sheet")
    
    mode = st.radio("เลือกโหมด", ("จัดทั้งเดือน", "จัดเฉพาะวันที่เลือก"))

    if mode == "จัดทั้งเดือน":
        current_date = datetime.now()
        year = st.number_input("ปี (ค.ศ.)", min_value=2020, max_value=2050, value=current_date.year)
        month = st.number_input("เดือน", min_value=1, max_value=12, value=current_date.month)
        dates_to_schedule = []
    else:
        date_range = st.date_input("เลือกช่วงวันที่", value=(datetime(2025, 10, 13), datetime(2025, 10, 15)), min_value=datetime(2020, 1, 1))
        if len(date_range) == 2: dates_to_schedule = pd.date_range(start=date_range[0], end=date_range[1]).to_pydatetime().tolist()
        else: dates_to_schedule = []
        year, month = 0, 0

    iterations = st.slider("จำนวนรอบที่ต้องการ Optimize", min_value=1, max_value=500, value=10, help="ยิ่งจำนวนรอบเยอะ ยิ่งมีโอกาสได้ตารางที่ดีขึ้น แต่ใช้เวลาประมวลผลนานขึ้น")
    
    run_button = st.button("🚀 เริ่มสร้างตารางเวร", type="primary", use_container_width=True)

# --- Main Area for Output ---
# Initialize session state to hold data after loading
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
    st.session_state.scheduler = None
    st.session_state.all_sheets = None

# Placeholder for the main content
main_placeholder = st.empty()

# Load data only once
if not st.session_state.data_loaded:
    with main_placeholder.container():
        st.info("กรุณากดปุ่ม 'เริ่มสร้างตารางเวร' ในเมนูด้านซ้ายเพื่อเริ่มใช้งาน")

if run_button:
    # --- 1. Load Data with Progress Bar ---
    loading_placeholder = main_placeholder.container()
    with loading_placeholder:
        st.subheader("⏳ กำลังโหลดข้อมูลตั้งต้น...")
        progress_bar = st.progress(0, text="กำลังเริ่มต้น...")
        try:
            # Store loaded data in session state
            st.session_state.scheduler = PharmacistScheduler(excel_url, logger=st.info, progress_bar=progress_bar)
            st.session_state.all_sheets = pd.read_excel(excel_url, sheet_name=None, engine='openpyxl')
            st.session_state.data_loaded = True
            progress_bar.progress(1.0, text="โหลดข้อมูลสำเร็จ!")
            time.sleep(1)
        except Exception as e:
            st.error(f"ไม่สามารถโหลดข้อมูลได้: {e}")
            st.stop()
    
    # Clear the loading message
    loading_placeholder.empty()

# --- 2. Display Loaded Data and Run Generation ---
if st.session_state.data_loaded:
    with main_placeholder.container():
        st.subheader("📄 ตรวจสอบข้อมูลตั้งต้น (Raw Data)")
        st.markdown("ข้อมูลทั้งหมดที่ใช้ในการประมวลผลถูกดึงมาจาก Google Sheet ด้านล่างนี้")
        
        for sheet_name, df in st.session_state.all_sheets.items():
            with st.expander(f"Sheet: {sheet_name}"):
                # Convert dataframe to HTML and display
                st.markdown(df.to_html(index=False), unsafe_allow_html=True)

        st.markdown("---")
        st.subheader("💡 ผลลัพธ์การจัดตารางเวร")
        
        # This section now runs the optimization using the already-loaded scheduler
        with st.spinner('กำลังประมวลผลและจัดตารางเวรที่ดีที่สุด...'):
            scheduler = st.session_state.scheduler
            progress_bar_gen = st.progress(0, text="เริ่มต้นการ Optimization...")
            
            best_schedule, best_unfilled_info = (None, None)

            if mode == "จัดทั้งเดือน":
                best_schedule, best_unfilled_info = scheduler.optimize_schedule(year, month, iterations, progress_bar_gen)
            else:
                if not dates_to_schedule:
                    st.error("กรุณาเลือกช่วงวันที่ที่ถูกต้อง")
                else:
                    best_schedule, best_unfilled_info = scheduler.optimize_schedule_for_dates(dates_to_schedule, iterations, progress_bar_gen)
            
            progress_bar_gen.progress(1.0, "ประมวลผลเสร็จสิ้น!")

        if best_schedule is not None:
            st.success("✅ จัดตารางเวรสำเร็จ!")
            excel_buffer = scheduler.export_to_excel(best_schedule, best_unfilled_info)
            output_filename = f"Pharmacist_Schedule_{year}_{month}.xlsx" if mode == "จัดทั้งเดือน" else "Pharmacist_Schedule_Custom_Dates.xlsx"
            
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Excel",
                data=excel_buffer,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            st.subheader("ตารางเวร (ตัวอย่าง)")
            st.dataframe(best_schedule)
        else:
            st.error("❌ ไม่สามารถสร้างตารางเวรที่สมบูรณ์ได้ กรุณาตรวจสอบข้อจำกัดต่างๆ หรือลองเพิ่มจำนวนรอบ Optimization")





