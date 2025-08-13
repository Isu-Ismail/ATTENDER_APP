import os
import random
import time
from datetime import date, timedelta

# Import the main application class and config variables from your modules
try:
    from main import AttendanceApp
    from config import USER_DATA_PATH
    from ui_windows import ManageWindow, MarkEntryWindow
except ImportError as e:
    print(f"Error: Make sure this script is in the same folder as your application files.\nDetails: {e}")
    exit()

def run_advanced_test():
    """Main function to orchestrate the comprehensive automated test."""
    print("--- Starting Advanced Automated Test ---")

    # --- Test Parameters ---
    TEST_FILENAME = "full_system_test.xlsx"
    SUBJECTS = ["ADVANCED PYTHON", "ALGORITHMS"]
    ASSESSMENTS = {"MIDTERM EXAM": 50, "FINAL PROJECT": 100} # Name: Max Marks
    NUM_STUDENTS = 100
    NUM_SESSIONS_PER_SUBJECT = 3

    # --- 1. Clean up previous test file ---
    test_file_path = os.path.join(USER_DATA_PATH, TEST_FILENAME)
    if os.path.exists(test_file_path):
        os.remove(test_file_path)
        print(f"[CLEANUP] Deleted old test file: {TEST_FILENAME}")

    # --- 2. Generate Test Data ---
    print(f"[SETUP] Generating {NUM_STUDENTS} students with random roll numbers...")
    student_names = [f"STUDENT_{i+1:03d}" for i in range(NUM_STUDENTS)]
    student_rolls = [f"2025-{random.choice(['CS', 'IT', 'ECE'])}-{i+1:03d}" for i in range(NUM_STUDENTS)]

    # --- 3. Initialize the Application ---
    app = AttendanceApp()
    app.update()
    time.sleep(1)

    # --- 4. Setup File, Subjects, and Students ---
    print(f"\n[PHASE 1] Testing Subject and Student Management...")
    app.file_combo.set(TEST_FILENAME)
    app.load_file()
    app.update()

    app.open_manage_window()
    manage_win = app.manage_win
    manage_win.update()

    # Create Subjects
    for subject in SUBJECTS:
        print(f"  -> Creating subject: {subject}")
        manage_win.new_subject_entry.insert(0, subject)
        manage_win.add_subject()
        manage_win.update()
        time.sleep(0.5)
    
    # Populate the first subject
    print(f"  -> Populating '{SUBJECTS[0]}' with {NUM_STUDENTS} students...")
    manage_win.subject_select_combo.set(SUBJECTS[0])
    manage_win.max_students_entry.insert(0, str(NUM_STUDENTS))
    manage_win.names_textbox.insert("1.0", "\n".join(student_names))
    manage_win.rolls_textbox.insert("1.0", "\n".join(student_rolls))
    manage_win.update_students()
    manage_win.update()
    time.sleep(1)

    # Test the "Copy Data" feature for the second subject
    if len(SUBJECTS) > 1:
        print(f"  -> Testing 'Copy Data' to '{SUBJECTS[1]}'")
        manage_win.subject_select_combo.set(SUBJECTS[1])
        manage_win.copy_source_combo.set(SUBJECTS[0])
        manage_win.copy_student_data()
        manage_win.update()
        time.sleep(0.5)
        manage_win.update_students()
        manage_win.update()
        time.sleep(1)
    
    manage_win.destroy()
    app.update()

    # --- 5. Mark Attendance for all subjects ---
    print(f"\n[PHASE 2] Testing Attendance Marking...")
    for subject in SUBJECTS:
        print(f"  -> Marking attendance for: {subject}")
        app.subject_combo.set(subject)
        for i in range(NUM_SESSIONS_PER_SUBJECT):
            # Generate random data for the session
            session_date = (date.today() + timedelta(days=i)).strftime("%d-%m-%Y")
            session_hours = random.randint(1, 2)
            absent_rolls = random.sample(range(1, NUM_STUDENTS + 1), 15) # 15 absentees
            
            # Call backend function directly to bypass GUI and confirmation dialogs
            sheet = app.wb[subject]
            app.mark_attendance(sheet, NUM_STUDENTS, absent_rolls, session_hours, session_date)
            print(f"    - Session {i+1}/{NUM_SESSIONS_PER_SUBJECT} marked for {session_date}.")
            app.update()
            time.sleep(0.3)

    # --- 6. Test Mark Entry ---
    print(f"\n[PHASE 3] Testing Mark Entry...")
    app.subject_combo.set(SUBJECTS[0]) # Test on the first subject
    app.open_mark_entry_window()
    mark_win = app.mark_win
    mark_win.update()

    # Add new assessment columns
    for name, max_marks in ASSESSMENTS.items():
        print(f"  -> Adding assessment: {name} (out of {max_marks})")
        app.add_new_assessment_column(app.wb[SUBJECTS[0]], name, str(max_marks))
        mark_win.refresh_assessments()
        mark_win.update()
        time.sleep(0.5)

    # Test bulk mark entry
    assessment_to_test = list(ASSESSMENTS.keys())[0]
    print(f"  -> Testing Bulk Mark Entry for: {assessment_to_test}")
    mark_win.assessment_combo.set(assessment_to_test)
    max_marks_for_test = ASSESSMENTS[assessment_to_test]
    marks_to_paste = [str(random.randint(int(max_marks_for_test*0.6), max_marks_for_test)) for _ in range(NUM_STUDENTS)]
    mark_win.bulk_textbox.insert("1.0", "\n".join(marks_to_paste))
    mark_win.apply_bulk_marks()
    mark_win.save_marks() # Bypassing confirmation for automation
    mark_win.update()
    time.sleep(1)

    mark_win.destroy()

    print("\n--- Advanced Test Finished Successfully! ---")
    app.show_status("Automated test finished!", is_error=False)
    time.sleep(5)
    app.destroy()

if __name__ == "__main__":
    run_advanced_test()