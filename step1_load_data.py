import pandas as pd
import logging

def read_input_files():
    """Read and validate all input sheets with proper error handling."""
    try:
        xls = pd.ExcelFile("input_data_tt.xlsx")
        required_sheets = ['in_timetable', 'in_course_roll_mapping', 'in_roll_name_mapping', 'in_room_capacity']
        if not all(sheet in xls.sheet_names for sheet in required_sheets):
            raise ValueError("Missing required sheets in input file")

        # Read and clean timetable
        timetable_df = pd.read_excel(xls, 'in_timetable')
        timetable_df['Date'] = pd.to_datetime(timetable_df['Date'], format='%Y-%m-%d')
        for col in ['Morning', 'Evening']:
            timetable_df[col] = timetable_df[col].apply(
                lambda x: [] if pd.isna(x) else [s.strip() for s in str(x).split(';') if s.strip()]
            ).apply(lambda lst: [s for s in lst if s.upper() != 'NO EXAM'])

        # Read and clean course-roll mapping
        course_roll_df = pd.read_excel(xls, 'in_course_roll_mapping')
        course_roll_df['course_code'] = course_roll_df['course_code'].astype(str).str.strip()
        course_roll_df['rollno'] = course_roll_df['rollno'].astype(str).str.strip()

        # Read and clean roll-name mapping (fix: column is "Roll" in Excel)
        roll_name_df = pd.read_excel(xls, 'in_roll_name_mapping')
        roll_name_df.columns = [c.strip() for c in roll_name_df.columns]
        if 'Roll' in roll_name_df.columns:
            roll_name_df.rename(columns={'Roll': 'rollno', 'Name': 'name'}, inplace=True)
        roll_name_df['rollno'] = roll_name_df['rollno'].astype(str).str.strip()
        roll_name_df['name'] = roll_name_df['name'].astype(str).str.strip()

        # Read and clean room capacity
        room_capacity_df = pd.read_excel(xls, 'in_room_capacity')
        room_capacity_df.columns = [c.strip() for c in room_capacity_df.columns]
        room_capacity_df = room_capacity_df[['Room No.', 'Exam Capacity', 'Block']]
        room_capacity_df['Room No.'] = room_capacity_df['Room No.'].astype(str).str.strip()
        room_capacity_df['Block'] = room_capacity_df['Block'].astype(str).str.strip()

        return {
            'timetable': timetable_df,
            'course_roll': course_roll_df,
            'roll_name': roll_name_df,
            'rooms': room_capacity_df
        }

    except Exception as e:
        logging.error(f"Input file error: {str(e)}", exc_info=True)
        raise

if __name__ == "__main__":
    logging.basicConfig(
        filename='errors.txt',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    try:
        data = read_input_files()
        print("Data loaded successfully:")
        print(f"- Timetable entries: {len(data['timetable'])}")
        print(f"- Unique courses: {data['course_roll']['course_code'].nunique()}")
        print(f"- Rooms: {len(data['rooms'])}")
        print(f"- Sample room: {data['rooms'].iloc[0].to_dict()}")
    except Exception as e:
        print("Critical error during data loading. Check errors.txt.")
