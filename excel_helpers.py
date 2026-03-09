def count_student_rows(sheet):
    """Counts the number of students by checking for roll numbers in Column A."""
    # Determine the start row based on where the headers are.
    # New format has "PERIODS :" in Row 3, Column B.
    if sheet.cell(row=3, column=2).value and "PERIODS" in str(sheet.cell(row=3, column=2).value):
        start_row = 6
    else:
        start_row = 5
    
    count = 0
    # Check one row past the max_row to be safe in case of empty rows
    for row in range(start_row, sheet.max_row + 2):
        if sheet.cell(row=row, column=1).value is None:
            break
        count += 1
    return count