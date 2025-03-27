import pandas as pd
import re

def run(file_path):
    attendance_df = pd.read_excel("C:\Users\Admin\Desktop\Raashi\data_sample.xlsx", sheet_name="Attendance_data")
    students_df = pd.read_excel("C:\Users\Admin\Desktop\Raashi\data_sample.xlsx", sheet_name="Student_data")

    attendance_df['attendance_date'] = pd.to_datetime(attendance_df['attendance_date'])

    attendance_df = attendance_df.sort_values(by=['student_id', 'attendance_date'])

    attendance_df['absence_group'] = (attendance_df['attendance'] != 'Absent').cumsum()
    
    absent_streaks = attendance_df[attendance_df['attendance'] == 'Absent'].groupby(
        ['student_id', 'absence_group']
    )['attendance_date'].agg(['min', 'max', 'count']).reset_index()

    absent_streaks = absent_streaks[absent_streaks['count'] > 3]

    latest_streaks = absent_streaks.loc[absent_streaks.groupby('student_id')['min'].idxmax()]

    latest_streaks = latest_streaks.rename(
        columns={'min': 'absence_start_date', 'max': 'absence_end_date', 'count': 'total_absent_days'}
    )
    
    final_df = latest_streaks.merge(students_df, on='student_id', how='left')

    def is_valid_email(email):
        pattern = r'^[a-zA-Z_][a-zA-Z0-9_]*@[a-zA-Z]+\.(com)$'
        return bool(re.match(pattern, str(email)))

    final_df['email'] = final_df['parent_email'].apply(lambda x: x if is_valid_email(x) else None)

    final_df['msg'] = final_df.apply(
        lambda row: f"Dear Parent, your child {row['student_name']} was absent from {row['absence_start_date'].date()} to {row['absence_end_date'].date()} for {row['total_absent_days']} days. Please ensure their attendance improves."
        if pd.notna(row['email']) else None, axis=1
    )

    final_df = final_df[['student_id', 'absence_start_date', 'absence_end_date', 'total_absent_days', 'email', 'msg']]

    return final_df

