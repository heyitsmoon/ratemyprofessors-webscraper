import pandas as pd

def create_pivot(file_path): 
    # Load file and data
    df = pd.read_excel(file_path, sheet_name='Data')

    # Create count summaries
    difficulty_count = df['Difficulty'].value_counts().rename_axis('Difficulty').reset_index(name='Count')
    attendance_count = df['Attendance'].value_counts().rename_axis('Attendance').reset_index(name='Count')
    textbook_count = df['Textbook'].value_counts().rename_axis('Textbook').reset_index(name='Count')
    grade_count = df['Grade'].value_counts().rename_axis('Grade').reset_index(name='Count')

    # Write to same workbook, same sheet, different columns
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        difficulty_count.to_excel(writer, sheet_name='Pivot Tables', index=False, startrow=0, startcol=0)
        attendance_count.to_excel(writer, sheet_name='Pivot Tables', index=False, startrow=0, startcol=2)
        textbook_count.to_excel(writer, sheet_name='Pivot Tables', index=False, startrow=0, startcol=4)
        grade_count.to_excel(writer, sheet_name='Pivot Tables', index=False, startrow=0, startcol=6)
