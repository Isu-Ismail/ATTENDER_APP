import customtkinter as ctk
from tkinter import messagebox
import openpyxl as xl
import os
from datetime import date, datetime

# Import from our custom modules
from config import ICON_PATH, resource_path, USER_DATA_PATH
from excel_helpers import count_student_rows

class LowAttendanceWindow(ctk.CTkToplevel):
    """Interactive window to generate low attendance reports."""
    def __init__(self, master, subject_name, sheet):
        super().__init__(master)
        self.title("Low Attendance Report")
        self.geometry("450x500")
        self.transient(master)
        self.focus()
        try:
            self.iconbitmap(resource_path(ICON_PATH))
        except: pass

        self.app = master
        self.sheet = sheet
        self.subject_name = subject_name
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        
        controls_frame = ctk.CTkFrame(self)
        controls_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        controls_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(controls_frame, text="Show students below:").grid(row=0, column=0, padx=(10, 5), pady=10)
        self.percent_entry = ctk.CTkEntry(controls_frame, placeholder_text="e.g., 75", width=60)
        self.percent_entry.grid(row=0, column=1, pady=10, sticky="w")
        self.percent_entry.insert(0, "75")
        ctk.CTkLabel(controls_frame, text="%").grid(row=0, column=2, padx=(2, 10), pady=10)
        self.report_button = ctk.CTkButton(controls_frame, text="Generate Report", command=self.generate_report)
        self.report_button.grid(row=0, column=3, padx=10, pady=10)
        
        self.error_label = ctk.CTkLabel(self, text="", text_color=("#C00000", "#FF8282"))
        self.error_label.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="w")
        
        self.textbox = ctk.CTkTextbox(self, corner_radius=8, font=("", 14))
        self.textbox.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="nsew")
        self.textbox.insert("1.0", "Enter a percentage above and click 'Generate Report'.")
        self.textbox.configure(state="disabled")
        self.generate_report()

    def generate_report(self):
        self.error_label.configure(text="")
        try:
            threshold = float(self.percent_entry.get())
            if not 0 <= threshold <= 100:
                self.error_label.configure(text="Error: Percentage must be between 0 and 100.")
                return
        except (ValueError, TypeError):
            self.error_label.configure(text="Error: Please enter a valid number.")
            return

        student_list = self.app.get_low_attendance_students(self.sheet, threshold)
        self.textbox.configure(state="normal")
        self.textbox.delete("1.0", "end")
        
        if student_list is None:
            report_text = f"Could not find a 'PERCENTAGE' column in {self.subject_name}."
        elif not student_list:
            report_text = f"Congratulations!\n\nNo students in {self.subject_name} below {threshold}%."
        else:
            header = f"Students in {self.subject_name} below {threshold}% attendance:\n{'-'*50}\n"
            report_text = header + "\n".join(student_list)
        
        self.textbox.insert("1.0", report_text)
        self.textbox.configure(state="disabled")

class ManageWindow(ctk.CTkToplevel):
    """Window for creating subjects and managing student lists with separate name/roll number fields."""
    def __init__(self, master):
        super().__init__(master)
        self.title("Subject & Student Management")
        self.geometry("500x700") # Increased height slightly
        self.transient(master)
        self.focus()
        try:
            self.iconbitmap(resource_path(ICON_PATH))
        except: pass

        self.app = master
        self.current_filename = self.app.current_filename 
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1) # Configure main content row to expand
        
        # --- Subject Management (No change) ---
        subject_frame = ctk.CTkFrame(self)
        subject_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        subject_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(subject_frame, text="Add New Subject", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=2, padx=10, pady=(10,0), sticky="w")
        self.new_subject_entry = ctk.CTkEntry(subject_frame, placeholder_text="Enter new subject name")
        self.new_subject_entry.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        self.add_subject_button = ctk.CTkButton(subject_frame, text="Add Subject", width=120, command=self.add_subject)
        self.add_subject_button.grid(row=1, column=1, padx=(0,10), pady=10)

        # --- NEW: Frame to copy data from another subject ---
        copy_frame = ctk.CTkFrame(self)
        copy_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        copy_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(copy_frame, text="Copy student list from:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.copy_source_combo = ctk.CTkComboBox(copy_frame, state="readonly", values=[])
        self.copy_source_combo.grid(row=0, column=1, padx=5, pady=10, sticky="ew")
        self.copy_data_button = ctk.CTkButton(copy_frame, text="Copy Data", width=120, command=self.copy_student_data)
        self.copy_data_button.grid(row=0, column=2, padx=10, pady=10)
        
        # --- Student Management (Redesigned) ---
        student_frame = ctk.CTkFrame(self)
        student_frame.grid(row=2, column=0, padx=20, pady=0, sticky="nsew")
        student_frame.grid_columnconfigure(0, weight=1)
        student_frame.grid_rowconfigure(4, weight=1)
        
        controls_frame = ctk.CTkFrame(student_frame, fg_color="transparent")
        controls_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=(10,0), sticky="ew")
        ctk.CTkLabel(controls_frame, text="Update Student List", font=ctk.CTkFont(weight="bold")).pack(side="left")
        generator_button = ctk.CTkButton(controls_frame, text="Generate Roll Numbers...", width=160, command=self.open_generator_dialog)
        generator_button.pack(side="right")
        
        ctk.CTkLabel(student_frame, text="Max Students in Class:").grid(row=1, column=0, padx=10, pady=(10,0), sticky="w")
        self.max_students_entry = ctk.CTkEntry(student_frame, placeholder_text="e.g., 65")
        self.max_students_entry.grid(row=1, column=1, padx=10, pady=(10,0), sticky="ew")
        
        ctk.CTkLabel(student_frame, text="Select Subject to Manage:").grid(row=2, column=0, padx=10, pady=(10,0), sticky="w")
        self.subject_select_combo = ctk.CTkComboBox(student_frame, state="readonly", command=self.load_student_data)
        self.subject_select_combo.grid(row=2, column=1, padx=10, pady=(10,0), sticky="ew")
        
        textbox_frame = ctk.CTkFrame(student_frame, fg_color="transparent")
        textbox_frame.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        textbox_frame.grid_columnconfigure(0, weight=2)
        textbox_frame.grid_columnconfigure(1, weight=1)
        textbox_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(textbox_frame, text="Student Names (Column B)").grid(row=3, column=0)
        self.names_textbox = ctk.CTkTextbox(textbox_frame, font=("", 14))
        self.names_textbox.grid(row=4, column=0, padx=(0, 5), sticky="nsew")

        ctk.CTkLabel(textbox_frame, text="Roll Numbers (Column C)").grid(row=3, column=1)
        self.rolls_textbox = ctk.CTkTextbox(textbox_frame, font=("", 14))
        self.rolls_textbox.grid(row=4, column=1, padx=(5, 0), sticky="nsew")

        self.update_students_button = ctk.CTkButton(self, text="Save Student List for Selected Subject", command=self.update_students)
        self.update_students_button.grid(row=3, column=0, padx=20, pady=20, sticky="ew")
        
        self.refresh_subject_list()
    
    def copy_student_data(self):
        """Copies student data from a source subject to the current textboxes."""
        source_subject = self.copy_source_combo.get()
        target_subject = self.subject_select_combo.get()

        if not source_subject or "No subjects" in source_subject:
            messagebox.showerror("Error", "Please select a source subject to copy from.", parent=self)
            return
        if not target_subject or "No subjects" in target_subject:
            messagebox.showerror("Error", "Please select a destination subject to copy to.", parent=self)
            return
        if source_subject == target_subject:
            messagebox.showwarning("Warning", "Source and destination subjects are the same.", parent=self)
            return
            
        try:
            source_sheet = self.app.wb[source_subject]
            num_students = count_student_rows(source_sheet)
            
            # Read names and rolls from the source sheet
            names = [str(source_sheet.cell(row=row, column=2).value or '') for row in range(5, num_students + 5)]
            rolls = [str(source_sheet.cell(row=row, column=3).value or '') for row in range(5, num_students + 5)]
            
            # Populate the textboxes
            self.names_textbox.delete("1.0", "end")
            self.names_textbox.insert("1.0", "\n".join(names))
            self.rolls_textbox.delete("1.0", "end")
            self.rolls_textbox.insert("1.0", "\n".join(rolls))
            
            # Update the max students count
            self.max_students_entry.delete(0, "end")
            self.max_students_entry.insert(0, str(num_students))

            messagebox.showinfo("Success", f"Copied {num_students} students from '{source_subject}'.\nClick 'Save Student List' to confirm.", parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"Could not copy data: {e}", parent=self)

    def refresh_subject_list(self):
        """Reloads the list of subjects from the workbook into BOTH dropdowns."""
        if self.app.wb:
            subjects = self.app.wb.sheetnames
            # Configure both dropdowns with the same list of subjects
            self.subject_select_combo.configure(values=subjects)
            self.copy_source_combo.configure(values=subjects)
            
            if subjects:
                self.subject_select_combo.set(subjects[0])
                self.copy_source_combo.set(subjects[0])
                self.load_student_data(subjects[0])
            else:
                self.subject_select_combo.set("No subjects created yet")
                self.copy_source_combo.set("No subjects available")
        else:
            self.subject_select_combo.set("No file loaded")
            self.copy_source_combo.set("No file loaded")

    def load_student_data(self, selected_subject):
        """Loads existing student names and roll numbers into the separate textboxes."""
        # Clear both textboxes
        self.names_textbox.delete("1.0", "end")
        self.rolls_textbox.delete("1.0", "end")
        if not self.app.wb: return
        
        try:
            sheet = self.app.wb[selected_subject]
            # Get data from Column B (Names) and C (Roll Numbers)
            names = [str(sheet.cell(row=row, column=2).value or '') for row in range(5, count_student_rows(sheet) + 5)]
            rolls = [str(sheet.cell(row=row, column=3).value or '') for row in range(5, count_student_rows(sheet) + 5)]
            
            # Insert data into the correct textboxes
            self.names_textbox.insert("1.0", "\n".join(names))
            self.rolls_textbox.insert("1.0", "\n".join(rolls))
        except Exception as e:
            print(f"Error loading student data: {e}")

    # This function is new
    def populate_from_generator(self, rolls_list):
        """Clears the ROLLS textbox and fills it with the generated roll numbers."""
        self.rolls_textbox.delete("1.0", "end")
        self.rolls_textbox.insert("1.0", "\n".join(rolls_list))

    # This function is new
    def open_generator_dialog(self):
        """Opens the Roll Number Generator dialog window."""
        RollGeneratorDialog(self)

    # This function is updated
    def update_students(self):
        """Saves student data from the two separate textboxes."""
        selected_subject = self.subject_select_combo.get()
        if not selected_subject or "No subjects" in selected_subject: return messagebox.showerror("Error", "Please select a valid subject.", parent=self)
        try:
            max_students = int(self.max_students_entry.get())
        except (ValueError, TypeError): return messagebox.showerror("Error", "Please enter a valid number for 'Max Students'.", parent=self)
        
        student_names = [name.strip().upper() for name in self.names_textbox.get("1.0", "end").strip().splitlines() if name.strip()]
        student_rolls = [roll.strip() for roll in self.rolls_textbox.get("1.0", "end").strip().splitlines() if roll.strip()]
        
        if len(student_names) != len(student_rolls):
            return messagebox.showerror("Error", f"Mismatch: There are {len(student_names)} names but {len(student_rolls)} roll numbers. The lists must match.", parent=self)
        if len(student_names) > max_students: return messagebox.showerror("Error", f"Student count ({len(student_names)}) exceeds maximum ({max_students}).", parent=self)
        if not messagebox.askyesno("Confirm", f"Overwrite student list for '{selected_subject}' with {len(student_names)} students?", parent=self): return
        
        try:
            sheet = self.app.wb[selected_subject]
            # Clear old student data from columns A, B, and C
            for row in range(5, sheet.max_row + 5):
                for col in range(1, 4): sheet.cell(row=row, column=col).value = None
            
            # Write new student data from the two textboxes
            for i in range(len(student_names)):
                sheet.cell(row=i+5, column=1).value = i + 1              # Simple Roll No.
                sheet.cell(row=i+5, column=2).value = student_names[i]   # Name
                sheet.cell(row=i+5, column=3).value = student_rolls[i]   # Complex Roll No.
            
            self.app.apply_standard_styles(sheet, len(student_names))
            self.app.wb.save(os.path.join(USER_DATA_PATH, self.current_filename))
            messagebox.showinfo("Success", f"Student list for '{selected_subject}' updated.", parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update students: {e}", parent=self)
            
    # The other functions (add_subject, etc.) remain the same
    def add_subject(self):
        new_name = self.new_subject_entry.get().strip()
        if not new_name: return messagebox.showerror("Error", "Subject name cannot be empty.", parent=self)
        if self.app.wb is None:
            self.app.wb = xl.Workbook()
            self.app.wb.remove(self.app.wb.active)
        if new_name in self.app.wb.sheetnames: return messagebox.showerror("Error", f"A subject named '{new_name}' already exists.", parent=self)
        
        new_sheet = self.app.wb.create_sheet(title=new_name)
        self.app.format_new_sheet(new_sheet)
        try:
            self.app.wb.save(os.path.join(USER_DATA_PATH, self.current_filename))
            messagebox.showinfo("Success", f"Subject '{new_name}' was created.", parent=self)
            self.new_subject_entry.delete(0, "end")
            self.refresh_subject_list()
            self.app.update_main_subject_list()
        except Exception as e:
            messagebox.showerror("Error", f"Could not save file: {e}", parent=self)

class DetailedReportWindow(ctk.CTkToplevel):
    """A new window with tabs for generating reports by date or by name."""
    def __init__(self, master, sheet):
        super().__init__(master)
        self.title("Detailed Report Generator")
        self.geometry("500x600")
        self.transient(master)
        self.focus()
        try:
            self.iconbitmap(resource_path(ICON_PATH))
        except: pass

        self.app = master
        self.sheet = sheet

        # --- Create a Tab View to switch between report types ---
        self.tab_view = ctk.CTkTabview(self, width=480)
        self.tab_view.pack(padx=20, pady=40, fill="both", expand=True)
        self.tab_view.add("By Date")
        self.tab_view.add("By Name")
        
        # --- Populate the "By Date" Tab ---
        self.setup_date_tab()
        
        # --- Populate the "By Name" Tab ---
        self.setup_name_tab()

    def setup_date_tab(self):
        """Creates the widgets for the 'By Date' report tab with a collapsible list."""
        date_tab = self.tab_view.tab("By Date")
        
        top_frame = ctk.CTkFrame(date_tab)
        top_frame.pack(padx=10, pady=10, fill="x")

        ctk.CTkLabel(top_frame, text="Select dates from the list:").pack(side="left", anchor="w", padx=10, pady=5)

        self.date_toggle_button = ctk.CTkButton(top_frame, text="Hide Date List", width=140, command=self.toggle_date_list)
        self.date_toggle_button.pack(side="right", anchor="e", padx=10, pady=5)
        
        # This frame can now be hidden/shown
        self.date_checklist_frame = ctk.CTkScrollableFrame(date_tab, label_text="Select Dates")
        self.date_checklist_frame.pack(padx=10, pady=10, fill="both", expand=True)

        self.date_checkboxes = {}
        all_dates = self.app.get_all_dates_from_sheet(self.sheet)
        for date_str in all_dates:
            var = ctk.StringVar(value="off")
            cb = ctk.CTkCheckBox(self.date_checklist_frame, text=date_str, variable=var, onvalue="on", offvalue="off")
            cb.pack(anchor="w", padx=10, pady=2)
            self.date_checkboxes[date_str] = var

        generate_btn = ctk.CTkButton(date_tab, text="Generate Date Report", command=self.generate_date_report)
        generate_btn.pack(padx=10, pady=10, fill="x")
        
        self.date_results_textbox = ctk.CTkTextbox(date_tab, corner_radius=8, font=("", 14))
        self.date_results_textbox.pack(padx=10, pady=10, fill="both", expand=True)
        self.date_results_textbox.insert("1.0", "Select one or more dates and click 'Generate Report'.")
        self.date_results_textbox.configure(state="disabled")

    # Add this new function inside the DetailedReportWindow class
    def toggle_student_list(self):
        """Shows or hides the student checklist frame."""
        if self.checklist_frame.winfo_viewable():
            self.checklist_frame.pack_forget()
            self.toggle_button.configure(text="Show Student List")
        else:
            self.checklist_frame.pack(padx=10, pady=5, fill="both", expand=True)
            self.toggle_button.configure(text="Hide Student List")
    
    def toggle_date_list(self):
        """Shows or hides the date checklist frame."""
        if self.date_checklist_frame.winfo_viewable():
            self.date_checklist_frame.pack_forget()
            self.date_toggle_button.configure(text="Show Date List")
        else:
            self.date_checklist_frame.pack(padx=10, pady=5, fill="both", expand=True)
            self.date_toggle_button.configure(text="Hide Date List")

    # Replace your old setup_name_tab function with this one
    def setup_name_tab(self):
        """Creates the widgets for the 'By Name' report tab with a collapsible list."""
        name_tab = self.tab_view.tab("By Name")
        
        top_frame = ctk.CTkFrame(name_tab)
        top_frame.pack(padx=10, pady=10, fill="x")
        
        ctk.CTkLabel(top_frame, text="Select students from the list:").pack(side="left", anchor="w", padx=10, pady=5)

        # This button will control the visibility of the checklist
        self.toggle_button = ctk.CTkButton(top_frame, text="Hide Student List", width=140, command=self.toggle_student_list)
        self.toggle_button.pack(side="right", anchor="e", padx=10, pady=5)
        
        # Create a scrollable frame for the student checklist
        self.checklist_frame = ctk.CTkScrollableFrame(name_tab, label_text="Student List")
        self.checklist_frame.pack(padx=10, pady=5, fill="both", expand=True)

        self.student_checkboxes = {}
        student_names = self.app.get_student_list(self.sheet)
        for name in student_names:
            var = ctk.StringVar(value="off")
            cb = ctk.CTkCheckBox(self.checklist_frame, text=name, variable=var, onvalue="on", offvalue="off")
            cb.pack(anchor="w", padx=10, pady=2)
            self.student_checkboxes[name] = var

        generate_btn = ctk.CTkButton(name_tab, text="Generate Name Report", command=self.generate_name_report)
        generate_btn.pack(padx=10, pady=10, fill="x")
        
        self.name_results_textbox = ctk.CTkTextbox(name_tab, corner_radius=8, font=("", 14))
        self.name_results_textbox.pack(padx=10, pady=10, fill="both", expand=True)
        self.name_results_textbox.insert("1.0", "Select students and click 'Generate Report'.")
        self.name_results_textbox.configure(state="disabled")

    def generate_date_report(self):
        """Gathers selected dates and generates the report."""
        selected_dates = [date_str for date_str, var in self.date_checkboxes.items() if var.get() == "on"]
        
        report_lines = self.app.get_report_by_date(self.sheet, selected_dates)
        report_text = "\n\n".join(report_lines)
            
        self.date_results_textbox.configure(state="normal")
        self.date_results_textbox.delete("1.0", "end")
        self.date_results_textbox.insert("1.0", report_text)
        self.date_results_textbox.configure(state="disabled")

    def generate_name_report(self):
        selected_names = [name for name, var in self.student_checkboxes.items() if var.get() == "on"]
        
        if not selected_names:
            report_text = "Please select at least one student from the checklist."
        else:
            report_lines = self.app.get_report_by_name(self.sheet, selected_names)
            report_text = "\n\n".join(report_lines)
            
        self.name_results_textbox.configure(state="normal")
        self.name_results_textbox.delete("1.0", "end")
        self.name_results_textbox.insert("1.0", report_text)
        self.name_results_textbox.configure(state="disabled")

class RollGeneratorDialog(ctk.CTkToplevel):
    """A dialog for generating complex roll numbers based on rules."""
    def __init__(self, master):
        super().__init__(master)
        self.title("Roll Number Generator")
        self.geometry("400x300")
        self.transient(master)
        self.focus()

        self.manage_window = master # Reference to the ManageWindow that opened it

        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self, text="Enter Generation Rules", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=20, pady=(20,10))
        
        main_frame = ctk.CTkFrame(self)
        main_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        main_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(main_frame, text="Prefixes (comma-sep):").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.prefixes_entry = ctk.CTkEntry(main_frame, placeholder_text="e.g., 20235070, 20235073")
        self.prefixes_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        
        ctk.CTkLabel(main_frame, text="Ranges (comma-sep):").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.ranges_entry = ctk.CTkEntry(main_frame, placeholder_text="e.g., 1-53, 1-12")
        self.ranges_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        
        ctk.CTkLabel(main_frame, text="Exclusions (comma-sep):").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.exclusions_entry = ctk.CTkEntry(main_frame, placeholder_text="e.g., 2023507034")
        self.exclusions_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        generate_button = ctk.CTkButton(self, text="Generate and Paste", command=self.generate_and_paste)
        generate_button.grid(row=2, column=0, padx=20, pady=20)

    def generate_and_paste(self):
        """Parses inputs, generates roll numbers, and pastes them into the ManageWindow."""
        try:
            prefixes = [p.strip() for p in self.prefixes_entry.get().split(',')]
            ranges = [r.strip() for r in self.ranges_entry.get().split(',')]
            exclusions = {e.strip() for e in self.exclusions_entry.get().split(',') if e.strip()}

            if len(prefixes) != len(ranges):
                messagebox.showerror("Error", "The number of prefixes must match the number of ranges.", parent=self)
                return

            generated_rolls = []
            for prefix, r_str in zip(prefixes, ranges):
                start, end = map(int, r_str.split('-'))
                for i in range(start, end + 1):
                    # Format with leading zero if needed (e.g., 1 -> 01, 10 -> 10)
                    roll_suffix = f"{i:02d}"
                    full_roll = f"{prefix}{roll_suffix}"
                    if full_roll not in exclusions:
                        generated_rolls.append(full_roll)
            
            # Pass the list back to the ManageWindow
            self.manage_window.populate_from_generator(generated_rolls)
            self.destroy() # Close the generator window

        except Exception as e:
            messagebox.showerror("Error", f"Invalid input format. Please check your entries.\nDetails: {e}", parent=self)

class BulkEntryWindow(ctk.CTkToplevel):
    """A window for entering multiple attendance records at once."""
    def __init__(self, master, sheet):
        super().__init__(master)
        self.title("Bulk Attendance Entry")
        self.geometry("600x650")
        self.transient(master)
        self.focus()
        try:
            self.iconbitmap(resource_path(ICON_PATH))
        except: pass

        self.app = master
        self.sheet = sheet

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1) # Configure log area to expand

        # --- Instructions and Input Area ---
        ctk.CTkLabel(self, text="Paste or type attendance data below.", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=20, pady=(20, 5), sticky="w")
        ctk.CTkLabel(self, text="Format: DATE:HOURS:ABSENTEE_ROLLS (e.g., 08-08-2025:2:1,3,5)", text_color="gray").grid(row=1, column=0, padx=20, pady=(0, 10), sticky="w")
        
        self.input_textbox = ctk.CTkTextbox(self, font=("", 14), height=150)
        self.input_textbox.grid(row=2, column=0, padx=20, pady=5, sticky="nsew")

        self.process_button = ctk.CTkButton(self, text="Process Bulk Entries", command=self.process_entries)
        self.process_button.grid(row=3, column=0, padx=20, pady=10, sticky="ew")

        # --- Results / Log Area ---
        self.results_textbox = ctk.CTkTextbox(self, corner_radius=8, font=("", 12), state="disabled")
        self.results_textbox.grid(row=4, column=0, padx=20, pady=(5, 20), sticky="nsew")

    def log_message(self, message):
        """Adds a message to the results log textbox."""
        self.results_textbox.configure(state="normal")
        self.results_textbox.insert("end", message + "\n")
        self.results_textbox.configure(state="disabled")
        self.results_textbox.see("end") # Auto-scroll to the bottom
        self.update_idletasks() # Force UI to update

    def process_entries(self):
        """Validates and processes each line from the input textbox."""
        self.process_button.configure(state="disabled")
        self.results_textbox.configure(state="normal")
        self.results_textbox.delete("1.0", "end")
        self.results_textbox.configure(state="disabled")

        self.log_message("--- Starting bulk processing ---")
        
        all_lines = self.input_textbox.get("1.0", "end").strip().splitlines()
        total_students = count_student_rows(self.sheet)

        for i, line in enumerate(all_lines):
            line = line.strip()
            if not line: continue

            self.log_message(f"\nProcessing line {i+1}: '{line}'")
            try:
                parts = line.split(':')
                if len(parts) != 3:
                    self.log_message("  -> ERROR: Invalid format. Must be DATE:HOURS:ROLLS.")
                    continue
                
                date_str, hours_str, rolls_str = [p.strip() for p in parts]
                
                # Validate Date
                datetime.strptime(date_str, "%d-%m-%Y")
                
                # Validate Hours
                num_hours = int(hours_str)
                if not 1 <= num_hours <= 8:
                    self.log_message("  -> ERROR: Hours must be between 1 and 8.")
                    continue

                # Validate Roll Numbers
                parsed_rolls = [int(r.strip()) for r in rolls_str.split(',') if r.strip()] if rolls_str else []
                invalid_rolls = [r for r in parsed_rolls if r > total_students or r < 1]
                if invalid_rolls:
                    self.log_message(f"  -> ERROR: Invalid Rolls: {invalid_rolls} out of range (1-{total_students}).")
                    continue
                
                # Check for existing date and ask to overwrite
                existing_date_col = None
                for col in range(4, self.sheet.max_column + 2):
                    if self.sheet.cell(row=2, column=col).value == date_str:
                        existing_date_col = col
                        break
                
                if existing_date_col:
                    if not messagebox.askyesno("Confirm Overwrite", f"An entry for {date_str} already exists.\n\nDo you want to overwrite it?", parent=self):
                        self.log_message(f"  -> SKIPPED: User chose not to overwrite date {date_str}.")
                        continue

                # If all validations pass, call the main mark_attendance function
                success, message = self.app.mark_attendance(self.sheet, total_students, parsed_rolls, num_hours, date_str, overwrite_col=existing_date_col)
                self.log_message(f"  -> STATUS: {message}")

            except Exception as e:
                self.log_message(f"  -> FAILED: An unexpected error occurred. ({e})")
        
        self.log_message("\n--- Bulk processing complete! ---")
        self.process_button.configure(state="normal")
# #class LoadingWindow(ctk.CTkToplevel):
#     """A simple splash screen that shows while the main app is loading."""
#     # In ui_windows.py, inside the LoadingWindow class
#     def __init__(self, master):
#         super().__init__(master)
#         self.title("Loading")
        
#         width, height = 300, 150
#         screen_width = self.winfo_screenwidth()
#         screen_height = self.winfo_screenheight()
#         x = (screen_width / 2) - (width / 2)
#         y = (screen_height / 2) - (height / 2)
#         self.geometry(f'{width}x{height}+{int(x)}+{int(y)}')
#         self.overrideredirect(True)

#         self.main_frame = ctk.CTkFrame(self, corner_radius=10)
#         self.main_frame.pack(fill="both", expand=True, padx=5, pady=5)

#         self.label = ctk.CTkLabel(self.main_frame, text="Attendance Marker", font=ctk.CTkFont(size=16, weight="bold"))
#         self.label.pack(padx=20, pady=(20,10))
        
#         self.progress_bar = ctk.CTkProgressBar(self.main_frame, mode='indeterminate')
        
#         self.start_button = ctk.CTkButton(self.main_frame, text="Start Application", width=200, height=40, command=self.start_build)
#         self.start_button.pack(padx=20, pady=10, expand=True)
    
#     def start_build(self):
#         """Hides the start button, shows the progress bar, and tells the main app to build."""
#         self.start_button.pack_forget()
#         self.label.configure(text="Loading Application...")
#         self.progress_bar.pack(padx=20, pady=10, fill="x", expand=True)
#         self.progress_bar.start()
        
#         # Call the main app's setup function after a short delay
#         self.master.after(100, self.master.setup_main_application)