import win32com.client as win32
import os

def clear_pts(ws):
    for pt in ws.PivotTables():
        pt.TableRange2.Clear()

def pt_settings(pt):
    # toggle grand totals
    pt.ColumnGrand = True
    pt.RowGrand = False

    # change subtotal location
    # https://docs.microsoft.com/en-us/office/vba/api/excel.xlsubtotallocationtype
    pt.SubtotalLocation(2) # bottom

    # use Tabular as the report layout
    # https://docs.microsoft.com/en-us/office/vba/api/excel.xllayoutrowtype
    pt.RowAxisLayout(1)

    # update table style
    pt.TableStyle2 = "PivotStyleMedium9"

def create_difficulty_pt(pt):
    field_rows = {}
    field_rows['Difficulty'] = pt.PivotFields("Difficulty")

    field_values = {}
    field_values['Difficulty'] = pt.PivotFields("Difficulty")

    # insert row fields
    # https://docs.microsoft.com/en-us/office/vba/api/excel.xlpivotfieldorientation
    field_rows['Difficulty'].Orientation = 1
    field_rows['Difficulty'].Position = 1

    # insert values field
    # https://docs.microsoft.com/en-us/office/vba/api/excel.xlconsolidationfunction
    field_values['Difficulty'].Orientation = 4
    field_values['Difficulty'].Function = -4112 # value reference
    field_values['Difficulty'].NumberFormat = "#,##0"

def create_grade_pt(pt):
    field_rows = {}
    field_rows['Grade'] = pt.PivotFields("Grade")
    field_rows['Grade'].Orientation = 1
    field_rows['Grade'].Position = 1

    field_values = {}
    field_values['Grade'] = pt.PivotFields("Grade")
    field_values['Grade'].Orientation = 4
    field_values['Grade'].Function = -4112 # value reference
    field_values['Grade'].NumberFormat = "#,##0"

def create_attendance_pt(pt):
    field_rows = {}
    field_rows['Attendance'] = pt.PivotFields("Attendance")
    field_rows['Attendance'].Orientation = 1
    field_rows['Attendance'].Position = 1

    field_values = {}
    field_values['Attendance'] = pt.PivotFields("Attendance")
    field_values['Attendance'].Orientation = 4
    field_values['Attendance'].Function = -4112 # value reference
    field_values['Attendance'].NumberFormat = "#,##0"

def create_overtime_pt(pt):
    field_rows = {}
    field_rows['Year'] = pt.PivotFields('Review Year')
    field_rows['Year'].Orientation = 1
    field_rows['Year'].Position = 1

    field_values = {}
    field_values['Quality'] = pt.PivotFields("Quality")
    field_values['Quality'].Orientation = 4
    field_values['Quality'].Function = -4106 # value reference
    field_values['Quality'].NumberFormat = "#.#0"

    field_values['Difficulty'] = pt.PivotFields("Difficulty")
    field_values['Difficulty'].Orientation = 4
    field_values['Difficulty'].Function = -4106 # value reference
    field_values['Difficulty'].NumberFormat = "#.#0"

def create_textbook_pt(pt):

    field_rows = {}
    field_rows['Textbook'] = pt.PivotFields("Textbook")
    field_rows['Textbook'].Orientation = 1
    field_rows['Textbook'].Position = 1

    field_values = {}
    field_values['Textbook'] = pt.PivotFields("Textbook")
    field_values['Textbook'].Orientation = 4
    field_values['Textbook'].Function = -4112 # value reference
    field_values['Textbook'].NumberFormat = "#,##0"

    

def create_pivots(relative_path):

    # construct the Excel application object
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = True
    #current working directory
    cwd = os.getcwd()
    file_path = os.path.join(cwd, relative_path)


    wb = excel.Workbooks.Open(file_path)
    ws_data = wb.Worksheets("Data")
    ws_report = wb.Worksheets("Pivot Tables")
    
    # clear pivot tables on Report tab
    clear_pts(ws_report)

    # create pt cache connection
    # https://docs.microsoft.com/en-us/office/vba/api/excel.xlpivottablesourcetype
    pt_cache = wb.PivotCaches().Create(1, ws_data.Range("A1").CurrentRegion)

    # insert pivot table designer/editor
    difficulty_pt = pt_cache.CreatePivotTable(ws_report.Range("A1"), "Rating")
    create_difficulty_pt(difficulty_pt)
    pt_settings(difficulty_pt)

    for item in  difficulty_pt.PivotFields("Difficulty").PivotItems():
        if item.Name == "(blank)" or item.Name == "Helpful" or item.Name == "Difficulty":
            item.Visible = False

    attendance_pt = pt_cache.CreatePivotTable(ws_report.Range("C1"), "Attendance")
    pt_settings(attendance_pt)
    create_attendance_pt(attendance_pt)

    for item in  attendance_pt.PivotFields("Attendance").PivotItems():
        if item.Name == "(blank)" or item.Name == "Helpful" or item.Name == "Attendance":
            item.Visible = False

    qd_overtime_pt = pt_cache.CreatePivotTable(ws_report.Range("E1"), "Quality/Difficulty Over Time")
    pt_settings(qd_overtime_pt)
    create_overtime_pt(qd_overtime_pt)
    qd_overtime_pt.ColumnGrand = False
    for item in  attendance_pt.PivotFields("Attendance").PivotItems():
        if item.Name == "(blank)" or item.Name == "Helpful":
            item.Visible = False

    textbook_pt = pt_cache.CreatePivotTable(ws_report.Range("H1"), "Textbook")
    pt_settings(textbook_pt)
    create_textbook_pt(textbook_pt)
    textbook_pt.ColumnGrand = False
    for item in  textbook_pt.PivotFields("Textbook").PivotItems():
        if item.Name == "(blank)" or item.Name == "Helpful":
            item.Visible = False

    grade_pt = pt_cache.CreatePivotTable(ws_report.Range("J1"), "Grade")
    pt_settings(grade_pt)
    create_grade_pt(grade_pt)
    
    grade_pt.ColumnGrand = False
    for item in  grade_pt.PivotFields("Grade").PivotItems():
        if item.Name == "(blank)" or item.Name == "Helpful" or "drop" in item.Name.lower() or "not" in item.Name.lower() or "incomplete" in item.Name.lower() or "textbook" in item.Name.lower():
            item.Visible = False
