from src.attendance_report.processor import process_attendance_report, ProcessingOptions

process_attendance_report(
    '/Users/pimlapat.doa/Documents/testapril/Report finger scan 1-30 April 2026.xlsx',
    '/Users/pimlapat.doa/Documents/testapril/(Draft)AttendanceReport_Exampleinput_APR.xlsx',
    '/Users/pimlapat.doa/Documents/Attendance_Report/Attendance_Report_April2026_validation.xlsx',
    ProcessingOptions(month=4, year=2026, working_days=18),
    '/Users/pimlapat.doa/Documents/testapril/PTVN_Report_Finger_Scan_input_APR.xlsx',
    '/Users/pimlapat.doa/Documents/Attendance_Report/PTVN_Report_Finger_Scan_April2026_validation.xlsx',
    log=print,
)
print("Done")
