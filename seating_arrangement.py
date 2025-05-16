import os
import pandas as pd
import logging
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font

# --- 1. Setup Logging for Error Tracking ---
logging.basicConfig(
    filename='errors.txt',      # Log file for errors
    level=logging.INFO,         # Log both info and error messages
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# --- 2. Input Data Reading and Cleaning ---
def read_input_files():
    """
    Reads and validates all required input sheets from the Excel file.
    Cleans up whitespace and ensures all necessary columns exist.
    """
    try:
        # Load Excel file
        xls = pd.ExcelFile("input_data_tt.xlsx")
        # Check if all required sheets are present
        required_sheets = ['in_timetable', 'in_course_roll_mapping', 'in_roll_name_mapping', 'in_room_capacity']
        if not all(sheet in xls.sheet_names for sheet in required_sheets):
            raise ValueError("Missing required sheets in input file")

        # Read and clean timetable (subjects split and stripped)
        timetable_df = pd.read_excel(xls, 'in_timetable')
        timetable_df['Date'] = pd.to_datetime(timetable_df['Date'], format='%Y-%m-%d')
        for col in ['Morning', 'Evening']:
            timetable_df[col] = timetable_df[col].apply(
                lambda x: [] if pd.isna(x) else [s.strip() for s in str(x).split(';') if s.strip()]
            ).apply(lambda lst: [s for s in lst if s.upper() != 'NO EXAM'])

        # Course to roll number mapping
        course_roll_df = pd.read_excel(xls, 'in_course_roll_mapping')
        course_roll_df['course_code'] = course_roll_df['course_code'].astype(str).str.strip()
        course_roll_df['rollno'] = course_roll_df['rollno'].astype(str).str.strip()

        # Roll number to student name mapping
        roll_name_df = pd.read_excel(xls, 'in_roll_name_mapping')
        roll_name_df.columns = [c.strip() for c in roll_name_df.columns]
        if 'Roll' in roll_name_df.columns:  # Handle possible column header differences
            roll_name_df.rename(columns={'Roll': 'rollno', 'Name': 'name'}, inplace=True)
        roll_name_df['rollno'] = roll_name_df['rollno'].astype(str).str.strip()
        roll_name_df['name'] = roll_name_df['name'].astype(str).str.strip()

        # Room capacity and details
        room_df = pd.read_excel(xls, 'in_room_capacity').iloc[:, :3]  # Only first 3 columns needed
        room_df.columns = ['Room No.', 'Exam Capacity', 'Block']
        room_df['Room No.'] = room_df['Room No.'].astype(str).str.strip()
        room_df['Block'] = room_df['Block'].astype(str).str.strip()

        return timetable_df, course_roll_df, roll_name_df, room_df

    except Exception as e:
        logging.error(f"Input file error: {str(e)}", exc_info=True)
        raise

# --- 3. Helper: Get Subjects for a Specific Date/Session ---
def get_subjects_for_date_session(timetable_df, date_obj, session):
    """
    Returns list of subjects scheduled for a specific date and session ('Morning' or 'Evening').
    """
    row = timetable_df[timetable_df['Date'] == date_obj]
    if row.empty:
        print(f"No timetable entry found for date: {date_obj}")
        return []
    return row.iloc[0][session]

# --- 4. Clash Checking Among Subjects ---
def check_clashes(subjects, course_roll_df):
    """
    Checks if any student is enrolled in more than one subject (i.e., roll number appears in multiple subject lists).
    Returns True if there is a clash, else False.
    """
    roll_sets = []
    for subj in subjects:
        rolls = set(course_roll_df[course_roll_df['course_code'] == subj]['rollno'])
        roll_sets.append(rolls)
    # Compare each subject's roll numbers with every other subject's roll numbers
    for i in range(len(roll_sets)):
        for j in range(i+1, len(roll_sets)):
            common = roll_sets[i] & roll_sets[j]
            if common:
                logging.error(f"Clash between {subjects[i]} and {subjects[j]}: {common}")
                print(f"Clash detected between {subjects[i]} and {subjects[j]} for rolls: {sorted(list(common))[:5]}...")
                return True
    print("No clashes detected.")
    return False

# --- 5. Main Allocation Logic ---
def allocate_subjects(subjects, course_roll_df, room_df, buffer, arrangement_type):
    """
    Allocates students to rooms while respecting all constraints:
    - Large courses get priority and bigger rooms
    - Sparse/dense arrangement
    - Avoid splitting same subject across buildings/floors if possible
    - Honors buffer and capacity
    Returns allocation dictionary, unallocated students, and updated room capacities
    """
    # Copy room data to avoid altering original dataframe
    room_df = room_df.copy()
    # Extract floor for proximity optimization (if present in room number)
    room_df['floor'] = room_df['Room No.'].str.extract(r'(\d{2,})').astype(float)
    # Calculate the effective remaining seats after applying buffer
    room_df['remaining'] = room_df['Exam Capacity'] - buffer

    # Adjust for sparse arrangement: only half the effective capacity per subject
    if arrangement_type.lower() == 'sparse':
        room_df['remaining'] = room_df['remaining'] // 2
    # Ensure no negative capacities
    room_df['remaining'] = room_df['remaining'].apply(lambda x: max(x, 0))

    # Sort rooms for allocation: Block, Floor, then by capacity (largest first)
    room_df = room_df.sort_values(['Block', 'floor', 'remaining'], ascending=[True, True, False])

    allocation = defaultdict(list)  # {(subject, room): [rolls]}
    unallocated = {}               # subject: [rolls not allocated]
    room_capacity = room_df.set_index('Room No.')['remaining'].to_dict()
    block_rooms = room_df.groupby('Block')['Room No.'].apply(list).to_dict()

    # Sort subjects by number of students (largest first)
    subject_sizes = {subj: len(course_roll_df[course_roll_df['course_code'] == subj]) for subj in subjects}
    sorted_subjects = sorted(subjects, key=lambda x: -subject_sizes[x])

    for subj in sorted_subjects:
        students = sorted(course_roll_df[course_roll_df['course_code'] == subj]['rollno'].tolist())
        remaining_students = students.copy()
        allocated = False

        # Try to allocate all rooms for this subject within the same block first (to minimize movement)
        for block in block_rooms:
            if not remaining_students:
                break
            for room in block_rooms[block]:
                if not remaining_students:
                    break
                available = room_capacity.get(room, 0)
                if available <= 0:
                    continue
                alloc = min(len(remaining_students), available)
                allocation[(subj, room)].extend(remaining_students[:alloc])
                room_capacity[room] -= alloc
                remaining_students = remaining_students[alloc:]
            if not remaining_students:
                allocated = True
                break

        # If still students left, allocate in any available room
        if not allocated and remaining_students:
            for room in room_df['Room No.']:
                if not remaining_students:
                    break
                available = room_capacity.get(room, 0)
                if available <= 0:
                    continue
                alloc = min(len(remaining_students), available)
                allocation[(subj, room)].extend(remaining_students[:alloc])
                room_capacity[room] -= alloc
                remaining_students = remaining_students[alloc:]

        # If students couldn't be allocated, note them as unallocated
        if remaining_students:
            unallocated[subj] = remaining_students
            logging.warning(f"Couldn't allocate {len(remaining_students)} students for {subj}")
            print(f"Cannot allocate {len(remaining_students)} students for {subj} (excess students).")

    return allocation, unallocated, room_capacity

# --- 6. Output Formatting for Room Excel Files ---
def format_room_excel(filepath, date_str, session, room, course, df):
    """
    Formats and writes a room-wise Excel file:
    - Merged header row with exam/session/room/course details
    - Student list (roll number, name)
    - Static placeholder rows for TAs and Invigilators
    """
    wb = Workbook()
    ws = wb.active

    # Add merged and centered header row
    header_text = f"Exam Date: {date_str} | Session: {session} | Room: {room} | Course: {course}"
    ncols = len(df.columns)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    cell = ws.cell(row=1, column=1, value=header_text)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.font = Font(bold=True, size=12)

    # Add student list (header and rows)
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=2):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    # Add 5 TA and 5 Invigilator placeholders (static text)
    start_row = ws.max_row + 2
    for i in range(1, 6):
        ws.cell(row=start_row + i - 1, column=1, value=f"TA {i}")
    for i in range(1, 6):
        ws.cell(row=start_row + 5 + i - 1, column=1, value=f"Invigilator {i}")

    wb.save(filepath)

# --- 7. Output Generation (Folders & Master Excel) ---
def generate_outputs(date_str, day_str, session, allocation, roll_name_dict, room_df, room_capacity):
    """
    Generates all required output files and folders:
    - Per-room Excel files with seating plan
    - Master file with overall arrangement
    - Seats left summary file
    """
    # Create output directory structure: <date>_<day>/<session>/
    base_dir = f"{date_str.replace('-', '')}_{day_str.replace(' ', '_')}"
    session_dir = os.path.join(base_dir, session.lower())
    os.makedirs(session_dir, exist_ok=True)

    # Prepare rows for master Excel file
    master_rows = []
    for (subj, room), rolls in allocation.items():
        # Prepare student list for the room, fill "Unknown Name" if name is missing
        df = pd.DataFrame({
            'Roll Number': rolls,
            'Name': [roll_name_dict.get(r, "Unknown Name") for r in rolls]
        })
        filename = f"{subj}_{room}.xlsx"
        filepath = os.path.join(session_dir, filename)
        format_room_excel(filepath, date_str, session, room, subj, df)

        # Add summary row to master
        master_rows.append({
            'Date': date_str,
            'Day': day_str,
            'course_code': subj,
            'Room': room,
            'Allocated_students_count': len(rolls),
            'Roll_list (semicolon separated_)': ';'.join(rolls)
        })

    # Write/update master file (appends if exists)
    master_file = 'op_overall_seating_arrangement.xlsx'
    if os.path.exists(master_file):
        existing = pd.read_excel(master_file)
        master_df = pd.concat([existing, pd.DataFrame(master_rows)], ignore_index=True)
    else:
        master_df = pd.DataFrame(master_rows)
    master_df.to_excel(master_file, index=False)

    # Generate/overwrite seats left summary
    seats_left = []
    for _, row in room_df.iterrows():
        room = row['Room No.']
        seats_left.append({
            'Room No.': room,
            'Exam Capacity': row['Exam Capacity'],
            'Block': row['Block'],
            'Alloted': int(row['Exam Capacity'] - room_capacity.get(room, row['Exam Capacity'])),
            'Vacant (B-C)': int(room_capacity.get(room, row['Exam Capacity']))
        })
    pd.DataFrame(seats_left).to_excel('op_seats_left.xlsx', index=False)

# --- 8. Main Program Flow ---
def main():
    """
    Main function that coordinates the entire seating arrangement workflow.
    Interacts with the user, processes all days/sessions, and generates outputs.
    """
    try:
        # Load and clean all data from input file
        timetable_df, course_roll_df, roll_name_df, room_df = read_input_files()
        roll_name_dict = roll_name_df.set_index('rollno')['name'].to_dict()  # Quick lookup for names

        # Display quick summary to the user
        print("Data loaded successfully:")
        print(f"- Timetable entries: {len(timetable_df)}")
        print(f"- Unique courses: {course_roll_df['course_code'].nunique()}")
        print(f"- Rooms: {len(room_df)}")

        # Get user input for buffer and arrangement type
        buffer = int(input("Enter buffer (seats to leave empty in each room): "))
        arrangement_type = input("Enter allocation type (sparse/dense): ").lower()
        while arrangement_type not in ['sparse', 'dense']:
            arrangement_type = input("Invalid input! Enter 'sparse' or 'dense': ").lower()

        # Process each date and both sessions
        for _, row in timetable_df.iterrows():
            date_str = row['Date'].strftime('%Y-%m-%d')
            day_str = row['Day']
            for session in ['Morning', 'Evening']:
                subjects = row[session]
                if not subjects:
                    continue  # Skip if no exam in this session
                print(f"\nProcessing {date_str} {session}...")

                # Check for student clashes
                if check_clashes(subjects, course_roll_df):
                    print("Clash detected. Skipping allocation for this slot.")
                    continue

                # Allocate rooms for all subjects in this slot
                allocation, unallocated, remaining_cap = allocate_subjects(
                    subjects, course_roll_df, room_df, buffer, arrangement_type
                )

                # Generate output files for this session
                generate_outputs(date_str, day_str, session, allocation, roll_name_dict, room_df, remaining_cap)
                print(f"Processed {len(subjects)} subjects for {session} session.")
                if unallocated:
                    print(f"Warning: Couldn't allocate {sum(len(v) for v in unallocated.values())} students")

        print("\nAllocation completed successfully! Check the generated Excel files and folders.")

    except Exception as e:
        logging.critical(f"Fatal error: {str(e)}", exc_info=True)
        print("Critical error occurred! Check errors.txt for details.")

# --- 9. Script Entry Point ---
if __name__ == "__main__":
    main()