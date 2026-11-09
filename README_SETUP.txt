# ExamDesk – Examination Management System (Enhanced Edition)

## Included workflows
- One examination form with General College, Combined multi-subject, and Department modes.
- Excel-only student roster import for each examination; validates required columns, duplicate USNs and selected department.
- Configurable department mixing limit per classroom for Combined examinations.
- Classroom capacity and availability checking by examination date + shift. A classroom booked for another examination in that slot is excluded.
- Automatic room selection using only the number of free rooms needed for the roster.
- Automatic invigilator allocation: one distinct faculty member per used room, with cross-examination conflict checking across General, Combined and Department workflows.
- One additional faculty member is reserved as Overall Supervisor / Chief Superintendent. This person is not reused as a room invigilator or in another examination at the same date and shift.
- Student seating stores classroom and seat number against each student's examination roster record.
- Attendance sheets use the exact examination roster, allow Present/Absent and remarks, and save the records in the database for later updates.
- Individual Form B DOCX files in a ZIP. Form B includes the PDA College heading, examination/date/shift, student's own subject, department, semester, USN, student name, room/seat after seating generation and signature fields.
- Classroom register supports add, edit and delete. Reports and database backup remain available.

## Excel import files
Open `sample_data/import_templates/`:
1. `01_Student_Master_Import.xlsx` – optional student master import.
2. `02_General_Exam_Students.xlsx` – General exam roster.
3. `03_Combined_Exam_Students.xlsx` – Combined roster with each student's own subject.
4. `04_Department_<DEPARTMENT>_Students.xlsx` – use the matching department workbook.
5. `05_Invigilators_Faculty.xlsx` – faculty list.
6. `06_Classrooms.xlsx` – classroom numbers and usable seat capacities.

All sample records are fictional and are for testing only. For live use, replace them with official college data. Keep real student data and `database/exam.db` out of public GitHub repositories.

## First-time setup (Windows)
1. Extract the ZIP.
2. Open the `SmartExam` folder.
3. Install Python 3.10 or newer.
4. Double-click `run_windows.bat`, or open Command Prompt in this folder and run:
   `pip install -r requirements.txt`
   `python app.py`
5. Open `http://127.0.0.1:5000` in your browser.

A virtual environment is optional; it is not required. The app creates the SQLite database and new tables automatically on first start. Before replacing an existing installation, close Flask and back up your `database/exam.db`.

## Recommended import and operation order
1. Import faculty and classrooms first.
2. (Optional) import the Student Master.
3. Create an examination and upload that examination's roster Excel file.
4. Generate seating. ExamDesk checks room availability and faculty availability for the exact date and shift, then stores the room bookings, invigilator duties and overall supervisor.
5. Download individual Form B files after seating has been generated.
6. Open Attendance, select the same examination, mark Present/Absent and remarks, and save.
7. Use Reports to export generated reports and Backup to create a database backup.

## Conflict rules
A classroom or faculty member cannot be allocated to two ExamDesk examinations with the same date and session. This applies regardless of whether the exams are General, Combined or Department type. If free rooms or faculty are insufficient, the system stops and displays the available capacity/count instead of silently reusing resources.

## Important
- Use the same session labels consistently (Morning, Afternoon, Evening).
- A separate overall supervisor is required in addition to one invigilator per occupied room.
- Department legacy records are considered when checking room/faculty clashes.
- This package includes syntax validation; complete live testing should be performed on the college computer with its actual database and official Form B requirements before operational use.


## Combined examination: upload separate department files
The Combined option accepts multiple Excel files in one submission. Select the CSE, Civil, ECE, EEE and/or Mechanical workbooks together using the file picker (Ctrl-click to select more than one). ExamDesk reads each workbook, merges the student rows and keeps each student's department, semester, subject code and subject name. It rejects duplicate USNs across files. Required columns in each Combined file: `usn`, `student_name`, `department`, `semester`, `subject_name`; `subject_code` is optional. The included department sample workbooks now contain these subject columns. You can still upload one pre-combined workbook as before.

SEATING CAPACITY + DEPARTMENT MIXING NOTE
The seating planner now calculates the number of rooms required after applying the selected maximum departments-per-classroom setting. For example, with 80 students across five departments and a mixing limit of 1, it may need five 20-seat rooms (one department per room), even though the raw total capacity of four rooms is 80 seats. Increase the mixing limit to allow more departments to share each classroom, or register more free rooms. Room bookings and faculty conflicts are still checked for the selected date and session.

COMBINED EXAMINATION: MULTIPLE DEPARTMENT FILES
Choose Combined Examination and select multiple department workbooks in the Department Excel files field (Ctrl-click to select several). ExamDesk merges the files, checks duplicate USNs across the upload set, and preserves each student's department, semester, subject name and optional subject code for seating and Form B.

UPLOAD HANDLING IMPROVEMENT
---------------------------
The examination roster reader now searches each worksheet for the actual
header row (including when a title row appears above it), normalises common
column-heading variations, and reports the exact workbook and worksheet when
a required column is missing. The selected examination type is preserved
when a validation error is shown. The Combined upload also accepts multiple
files from the `files` field and compatible older/singular field names.
Clean single-sheet sample upload files are in `sample_data/upload_ready/`.
