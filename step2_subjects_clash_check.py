import pandas as pd
from step1_load_data import read_input_files  # your Step 1 loader
import logging

def get_subjects_for_date_session(timetable_df, date_obj, session):
    row = timetable_df[timetable_df['Date'] == date_obj]
    if row.empty:
        print(f"No timetable entry found for date: {date_obj}")
        return []
    subjects = row.iloc[0][session]
    return subjects

def check_clashes(subjects, course_roll_df):
    roll_sets = []
    for subj in subjects:
        rolls = course_roll_df[course_roll_df['course_code'] == subj]['rollno'].unique()
        roll_sets.append(set(rolls))
    n = len(roll_sets)
    for i in range(n):
        for j in range(i + 1, n):
            intersection = roll_sets[i].intersection(roll_sets[j])
            if intersection:
                print(f"Clash detected between {subjects[i]} and {subjects[j]} for roll numbers: {sorted(intersection)}")
                return True
    print("No clashes detected.")
    return False

if __name__ == "__main__":
    logging.basicConfig(
        filename='errors.txt',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    try:
        data = read_input_files()
        timetable_df = data['timetable']
        course_roll_df = data['course_roll']

        date_input = '2016-04-30'  # YYYY-MM-DD
        date_obj = pd.to_datetime(date_input)
        session_input = 'Morning'  # or 'Evening'

        subjects = get_subjects_for_date_session(timetable_df, date_obj, session_input)
        print(f"Subjects scheduled on {date_input} {session_input}: {subjects}")

        has_clash = check_clashes(subjects, course_roll_df)
        if has_clash:
            print("Please resolve clashes before proceeding.")
        else:
            print("No clashes found, you can proceed.")

    except Exception as e:
        logging.error(f"Error in step 2 execution: {e}", exc_info=True)
        print("An error occurred. Check errors.txt for details.")
