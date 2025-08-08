def count_student_rows(sheet):
    """Counts the number of students by checking for roll numbers in Column A."""
    count = 0
    # Check one row past the max_row to be safe in case of empty rows
    for row in range(5, sheet.max_row + 2):
        if sheet.cell(row=row, column=1).value is None:
            break
        count += 1
    return count