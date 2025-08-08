import os
import random
import time
from datetime import date, timedelta

# Import the main application class from your script
try:
    from main import AttendanceApp, USER_DATA_PATH
except ImportError:
    print("Error: Make sure this script is in the same folder as 'main.py'")
    exit()

def run_advanced_test():
    """Main function to orchestrate the comprehensive automated test."""
    print("--- Starting Advanced Automated Test ---")

    # --- Test Parameters ---
    TEST_FILENAME = "advanced_automated_test.xlsx"
    SUBJECTS_TO_CREATE = ["PYTHON PROGRAMMING", "DATA STRUCTURES", "OPERATING SYSTEMS"]
    NUM_STUDENTS = 100
    NUM_SESSIONS_PER_SUBJECT = 5

    # --- 1. Clean up previous test file ---
    test_file_path = os.path.join(USER_DATA_PATH, TEST_FILENAME)
    if os.path.exists(test_file_path):
        os.remove(test_file_path)
        print(f"[CLEANUP] Deleted old test file: {TEST_FILENAME}")

    # --- 2. Generate Test Data ---
    print(f"[SETUP] Generating {NUM_STUDENTS} students with random roll numbers...")
    student_names = []
    student_rolls = []
    departments = ["CS", "IT", "ECE", "MECH"]
    for i in range(NUM_STUDENTS):
        student_names.append(f"STUDENT_{i+1:03d}")
        # Generate a complex roll number, e.g., 2025-CS-042
        roll_number = f"2025-{random.choice(departments)}-{i+1:03d}"
        student_rolls.append(roll_number)

    # --- 3. Create an instance of your application ---
    app = AttendanceApp()
    app.update()
    time.sleep(1)

    # --- 4. Setup File, Subjects, and Students ---
    print(f"\n[SETUP] Creating test file: {TEST_FILENAME}")
    app.file_combo.set(TEST_FILENAME)
    app.load_file()
    app.update()

    print("[SETUP] Opening Management Window...")
    app.open_manage_window()
    manage_win = app.manage_win
    manage_win.update()

    # Create all subjects
    for subject in SUBJECTS_TO_CREATE:
        print(f"[SETUP] Creating subject: {subject}")
        manage_win.new_subject_entry.delete(0, "end")
        manage_win.new_subject_entry.insert(0, subject)
        manage_win.add_subject()
        manage_win.update()
        time.sleep(0.5)
    
    # Populate the first subject with the generated student list
    print(f"[SETUP] Populating '{SUBJECTS_TO_CREATE[0]}' with {NUM_STUDENTS} students...")
    manage_win.subject_select_combo.set(SUBJECTS_TO_CREATE[0])
    manage_win.max_students_entry.delete(0, "end")
    manage_win.max_students_entry.insert(0, str(NUM_STUDENTS))
    manage_win.names_textbox.insert("1.0", "\n".join(student_names))
    manage_win.rolls_textbox.insert("1.0", "\n".join(student_rolls))
    manage_win.update_students()
    manage_win.update()
    time.sleep(1)

    # Use the "Copy Data" feature to populate the other subjects
    for i in range(1, len(SUBJECTS_TO_CREATE)):
        target_subject = SUBJECTS_TO_CREATE[i]
        source_subject = SUBJECTS_TO_CREATE[0]
        print(f"[SETUP] Copying student data to '{target_subject}'...")
        manage_win.subject_select_combo.set(target_subject)
        manage_win.copy_source_combo.set(source_subject)
        manage_win.copy_student_data()
        manage_win.update()
        time.sleep(0.5)
        manage_win.update_students()
        manage_win.update()
        time.sleep(1)

    print("[SETUP] Closing Management Window.")
    manage_win.destroy()
    app.update()

    # --- 5. Mark Attendance for all subjects ---
    for subject in SUBJECTS_TO_CREATE:
        print(f"\n[TEST] Now marking attendance for subject: {subject}")
        app.subject_combo.set(subject)
        for i in range(NUM_SESSIONS_PER_SUBJECT):
            session_date = (date.today() + timedelta(days=i)).strftime("%d-%m-%Y")
            session_hours = random.randint(1, 4)
            num_absent = random.randint(0, 20)
            absent_rolls = random.sample(range(1, NUM_STUDENTS + 1), num_absent)
            absent_rolls_str = ", ".join(map(str, absent_rolls))

            print(f"  -> Session {i+1}/{NUM_SESSIONS_PER_SUBJECT}: Marking {num_absent} absentees...")
            
            app.date_entry.delete(0, "end")
            app.hours_entry.delete(0, "end")
            app.rolls_entry.delete(0, "end")
            app.date_entry.insert(0, session_date)
            app.hours_entry.insert(0, str(session_hours))
            app.rolls_entry.insert(0, absent_rolls_str)
            app.update()
            
            # Bypass GUI and call the backend function directly
            sheet = app.wb[subject]
            app.mark_attendance(sheet, NUM_STUDENTS, absent_rolls, session_hours, session_date)
            app.update()
            time.sleep(0.3)

    print("\n--- Advanced Test Finished Successfully! ---")
    app.show_status("Automated test finished!", is_error=False)
    time.sleep(5)
    app.destroy()

if __name__ == "__main__":
    run_advanced_test()