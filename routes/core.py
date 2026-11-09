import os, io, json, shutil, zipfile
from datetime import datetime
from pathlib import Path
import pandas as pd
from flask import Blueprint, render_template, request, redirect, url_for, flash, current_app, send_from_directory, session
from sqlalchemy import or_
from database import db
from database.models import Student, Examination, ExaminationStudent, Classroom, Invigilator, DepartmentExamination, DepartmentExamStudent, DepartmentExamRoom, RoomBooking, InvigilatorAssignment, ExaminationSupervisor, AttendanceRecord
from services.documents import make_report
bp=Blueprint('core',__name__)

def _normalise_excel_column(value):
    import re
    text = str(value or '').strip().lower()
    text = re.sub(r'[^a-z0-9]+', '_', text).strip('_')
    aliases = {
        'usn_no': 'usn', 'student_usn': 'usn', 'university_seat_number': 'usn',
        'studentname': 'student_name', 'name_of_student': 'student_name', 'name': 'student_name',
        'dept': 'department', 'department_name': 'department',
        'sem': 'semester', 'semester_no': 'semester', 'semester_number': 'semester',
        'subject': 'subject_name', 'subjectname': 'subject_name', 'subject_title': 'subject_name',
        'course_name': 'subject_name', 'course_title': 'subject_name',
        'course_code': 'subject_code', 'subjectcode': 'subject_code',
        'classroom': 'classroom_number', 'room_number': 'classroom_number',
        'capacity': 'total_seats', 'seats': 'total_seats',
        'faculty_code': 'faculty_id', 'staff_id': 'faculty_id', 'faculty': 'faculty_name',
    }
    return aliases.get(text, text)


def read_excel(file, required, source_name=None):
    # Find the real header row in any worksheet, even when a title or note is
    # placed above it. Normalise common heading variations before validation.
    if not file or not getattr(file, 'filename', '') or not file.filename.lower().endswith(('.xlsx', '.xls')):
        raise ValueError('Please upload an Excel file (.xlsx / .xls).')
    label = source_name or file.filename
    try:
        file.stream.seek(0)
        book = pd.ExcelFile(file.stream)
        best_missing = None
        best_sheet = book.sheet_names[0] if book.sheet_names else 'Sheet1'
        best_headers = []
        for sheet_name in book.sheet_names:
            raw = pd.read_excel(book, sheet_name=sheet_name, header=None, dtype=object).fillna('')
            for header_idx in range(min(len(raw), 15)):
                headers = [_normalise_excel_column(v) for v in raw.iloc[header_idx].tolist()]
                if not any(headers):
                    continue
                missing = [col for col in required if col not in headers]
                if best_missing is None or len(missing) < len(best_missing):
                    best_missing, best_sheet, best_headers = missing, sheet_name, headers
                if not missing:
                    data = raw.iloc[header_idx + 1:].copy()
                    data.columns = headers
                    data = data.loc[:, [c for c in data.columns if c]]
                    data = data.dropna(how='all').fillna('')
                    if 'usn' in data.columns:
                        data = data[data['usn'].astype(str).str.strip() != '']
                    file.stream.seek(0)
                    return data.reset_index(drop=True)
        missing = best_missing if best_missing is not None else list(required)
        found = ', '.join(dict.fromkeys(c for c in best_headers if c)) or 'no usable header row found'
        raise ValueError(
            f"{label} (worksheet '{best_sheet}') is missing required column(s): {', '.join(missing)}. "
            f"Columns found: {found}. Use the matching ExamDesk upload template."
        )
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError(f'Could not read {label}: {exc}. Please open it in Excel and save it as .xlsx.') from exc


def import_rows(model, df, mapping, unique_field):
    added=0; skipped=[]
    for i,(_,r) in enumerate(df.iterrows(),start=2):
        vals={k:(str(r[v]).strip() if v in df.columns else '') for k,v in mapping.items()}
        required_keys = (['usn','student_name','department'] if model is Student else ['faculty_id','faculty_name'])
        if any(not vals.get(k) for k in required_keys): skipped.append(f'Row {i}: blank required value'); continue
        if model.query.filter(getattr(model,unique_field)==vals[unique_field]).first(): skipped.append(f'Row {i}: duplicate {unique_field} {vals[unique_field]}'); continue
        db.session.add(model(**vals)); added+=1
    db.session.commit(); return added,skipped


def _norm_session(value):
    return (value or '').strip().casefold()

def _busy_faculty_ids(exam):
    busy=set()
    rows=(db.session.query(InvigilatorAssignment.invigilator_id)
        .join(Examination,Examination.id==InvigilatorAssignment.examination_id)
        .filter(Examination.exam_date==exam.exam_date,
                db.func.lower(db.func.trim(db.func.coalesce(Examination.session,'')))==_norm_session(exam.session),
                Examination.id!=exam.id).all())
    busy.update(x[0] for x in rows)
    supervisor_rows=(db.session.query(ExaminationSupervisor.invigilator_id)
        .join(Examination,Examination.id==ExaminationSupervisor.examination_id)
        .filter(Examination.exam_date==exam.exam_date,
                db.func.lower(db.func.trim(db.func.coalesce(Examination.session,'')))==_norm_session(exam.session),
                Examination.id!=exam.id).all())
    busy.update(x[0] for x in supervisor_rows)
    legacy=(db.session.query(DepartmentExamRoom.invigilator_id)
        .join(DepartmentExamination,DepartmentExamination.id==DepartmentExamRoom.examination_id)
        .filter(DepartmentExamination.exam_date==exam.exam_date,
                db.func.lower(db.func.trim(db.func.coalesce(DepartmentExamination.session,'')))==_norm_session(exam.session),
                DepartmentExamRoom.invigilator_id.isnot(None)).all())
    busy.update(x[0] for x in legacy)
    return busy

def _busy_room_ids(exam):
    busy=set()
    rows=(db.session.query(RoomBooking.classroom_id)
        .join(Examination,Examination.id==RoomBooking.examination_id)
        .filter(Examination.exam_date==exam.exam_date,
                db.func.lower(db.func.trim(db.func.coalesce(Examination.session,'')))==_norm_session(exam.session),
                Examination.id!=exam.id).all())
    busy.update(x[0] for x in rows)
    legacy=(db.session.query(DepartmentExamRoom.classroom_id)
        .join(DepartmentExamination,DepartmentExamination.id==DepartmentExamRoom.examination_id)
        .filter(DepartmentExamination.exam_date==exam.exam_date,
                db.func.lower(db.func.trim(db.func.coalesce(DepartmentExamination.session,'')))==_norm_session(exam.session)).all())
    busy.update(x[0] for x in legacy)
    return busy

def _plan_seating(roster, free_rooms, mixing_limit):
    """Plan placements first, so department-mixing limits are included in room counts."""
    remaining={}
    for student in roster:
        remaining.setdefault((student.department or 'Unassigned').strip(), []).append(student)
    for dept in remaining:
        remaining[dept].sort(key=lambda x: (x.semester or '', x.usn))
    plan=[]
    limit=max(1, int(mixing_limit or 1))
    for room in free_rooms:
        if not any(remaining.values()):
            break
        chosen=sorted((d for d, items in remaining.items() if items),
                      key=lambda d: (-len(remaining[d]), d))[:limit]
        placements=[]
        seat=1
        # Round-robin across the chosen departments to avoid long same-department blocks.
        while seat <= room.total_seats and any(remaining[d] for d in chosen):
            for dept in chosen:
                if seat > room.total_seats:
                    break
                if remaining[dept]:
                    placements.append((room, seat, remaining[dept].pop(0)))
                    seat += 1
        plan.extend(placements)
    if any(remaining.values()):
        left=sum(len(items) for items in remaining.values())
        raise ValueError(f'Not enough usable classroom capacity under the selected department-mixing limit. {left} students remain unseated. Increase the mixing limit or register more free classrooms.')
    return plan

def _allocate_resources(exam, roster, planned_rooms=None):
    # Choose only free rooms; respect capacity, the seating plan and other exams' bookings.
    busy_rooms=_busy_room_ids(exam)
    all_rooms=Classroom.query.order_by(Classroom.classroom_number).all()
    free_rooms=[r for r in all_rooms if r.id not in busy_rooms]
    if planned_rooms is None:
        chosen=[]; seats=0
        for room in free_rooms:
            if seats>=len(roster): break
            chosen.append(room); seats+=room.total_seats
    else:
        chosen=planned_rooms
        seats=sum(r.total_seats for r in chosen)
    if seats < len(roster):
        occupied_names=[r.classroom_number for r in all_rooms if r.id in busy_rooms]
        free_capacity=sum(r.total_seats for r in free_rooms)
        raise ValueError(f'Not enough free classroom capacity for {len(roster)} students. Available capacity is {free_capacity} seats. Rooms already booked for {exam.exam_date} / {exam.session}: {", ".join(occupied_names) or "none"}.')
    busy_faculty=_busy_faculty_ids(exam)
    staff=Invigilator.query.order_by(Invigilator.faculty_name).all()
    available=[x for x in staff if x.id not in busy_faculty]
    required=len(chosen)+1
    if len(available)<required:
        raise ValueError(f'Not enough free faculty for this shift. {len(available)} faculty are available; {required} are required ({len(chosen)} room invigilators plus one overall supervisor). Faculty already assigned to another exam at this date and shift cannot be reused.')
    RoomBooking.query.filter_by(examination_id=exam.id).delete(synchronize_session=False)
    InvigilatorAssignment.query.filter_by(examination_id=exam.id).delete(synchronize_session=False)
    ExaminationSupervisor.query.filter_by(examination_id=exam.id).delete(synchronize_session=False)
    for room in chosen: db.session.add(RoomBooking(examination_id=exam.id,classroom_id=room.id))
    duties=[]
    for room,faculty in zip(chosen,available):
        db.session.add(InvigilatorAssignment(examination_id=exam.id,classroom_id=room.id,invigilator_id=faculty.id))
        duties.append((room,faculty))
    supervisor=available[len(chosen)]
    db.session.add(ExaminationSupervisor(examination_id=exam.id,invigilator_id=supervisor.id))
    return chosen,duties,supervisor

@bp.app_context_processor
def inject_globals(): return {'current_year':datetime.now().year}

@bp.route('/')
def dashboard():
    return render_template('dashboard.html', students=Student.query.count(), exams=Examination.query.count(), classrooms=Classroom.query.count(), invigilators=Invigilator.query.count(), upcoming=Examination.query.order_by(Examination.exam_date).limit(5).all())

@bp.route('/students')
def students():
    q=request.args.get('q','').strip(); query=Student.query
    if q: query=query.filter(or_(Student.usn.ilike(f'%{q}%'),Student.student_name.ilike(f'%{q}%'),Student.department.ilike(f'%{q}%')))
    return render_template('students/list.html',items=query.order_by(Student.department,Student.usn).all(),q=q)
@bp.route('/students/import',methods=['GET','POST'])
def students_import():
    if request.method=='POST':
        try:
            df=read_excel(request.files.get('file'),['usn','student_name','department'])
            added,skipped=import_rows(Student,df,{'usn':'usn','student_name':'student_name','department':'department','semester':'semester'},'usn')
            flash(f'Imported {added} students. {len(skipped)} rows skipped.','success' if added else 'warning')
            for x in skipped[:8]: flash(x,'warning')
        except Exception as e: db.session.rollback(); flash(str(e),'danger')
        return redirect(url_for('core.students'))
    return render_template('students/import.html')
@bp.route('/students/<int:item_id>/delete',methods=['POST'])
def student_delete(item_id):
    item=Student.query.get_or_404(item_id); db.session.delete(item); db.session.commit(); flash('Student deleted.','success'); return redirect(url_for('core.students'))

@bp.route('/examinations')
def examinations(): return render_template('examinations/list.html',items=Examination.query.order_by(Examination.exam_date).all())
@bp.route('/examinations/add',methods=['GET','POST'])
def exam_add():
    departments = sorted({x.department for x in Student.query.with_entities(Student.department).all() if x.department})
    selected_exam_type = request.form.get('exam_type', 'general').strip() if request.method == 'POST' else 'general'
    if request.method == 'POST':
        try:
            exam_type=selected_exam_type
            exam_name=request.form.get('exam_name','').strip()
            exam_date=request.form.get('exam_date','').strip()
            session_name=request.form.get('session','Morning').strip()
            common_subject=request.form.get('subject_name','').strip()
            common_code=request.form.get('subject_code','').strip()
            semester=request.form.get('semester','').strip()
            selected_dept=request.form.get('department','').strip()
            try: mixing_limit=max(1,int(request.form.get('mixing_limit','1')))
            except ValueError: mixing_limit=1
            if exam_type != 'combined': mixing_limit=1
            if exam_type not in ('general','combined','department'):
                raise ValueError('Select a valid examination type.')
            if not exam_name or not exam_date:
                raise ValueError('Examination name and date are required.')
            if exam_type in ('general','department') and not common_subject:
                raise ValueError('Subject name is required for this examination type.')
            if exam_type=='department' and not selected_dept:
                raise ValueError('Select the department.')
            required=['usn','student_name']
            uploaded_files=[]
            if exam_type=='combined':
                required += ['department','semester','subject_name']
                uploaded_files=[]
                # Current form uses `files`; also accept older/custom field names.
                for field_name in ('files', 'file', 'department_files', 'student_files'):
                    uploaded_files.extend(f for f in request.files.getlist(field_name) if f and f.filename)
                if not uploaded_files:
                    for _, file_list in request.files.lists():
                        uploaded_files.extend(f for f in file_list if f and f.filename)
                if not uploaded_files:
                    raise ValueError('Select at least one department Excel file for the Combined examination.')
            else:
                candidates=[]
                for field_name in ('file', 'files', 'student_files'):
                    candidates.extend(f for f in request.files.getlist(field_name) if f and f.filename)
                if not candidates:
                    for _, file_list in request.files.lists():
                        candidates.extend(f for f in file_list if f and f.filename)
                if not candidates:
                    raise ValueError('Please select the student Excel file.')
                uploaded_files=[candidates[0]]
            records=[]; seen=set()
            for upload in uploaded_files:
                df=read_excel(upload,required, source_name=upload.filename)
                source_name=upload.filename
                for row_no,(_,row) in enumerate(df.iterrows(),start=2):
                    usn=str(row.get('usn','')).strip()
                    name=str(row.get('student_name','')).strip()
                    if not usn and not name: continue
                    if not usn or not name: raise ValueError(f'{source_name}, row {row_no}: USN and student name are both required.')
                    if usn.casefold() in seen: raise ValueError(f'Duplicate USN across uploaded files: {usn}. Check {source_name}.')
                    seen.add(usn.casefold())
                    row_dept=str(row.get('department',selected_dept or 'General')).strip() or (selected_dept or 'General')
                    row_sem=str(row.get('semester',semester)).strip().replace('.0','')
                    if exam_type=='department' and row_dept.casefold()!=selected_dept.casefold():
                        raise ValueError(f'{source_name}, row {row_no}: {usn} does not belong to {selected_dept}.')
                    if exam_type=='general':
                        subject=common_subject; code=common_code
                        row_sem=row_sem or semester
                    elif exam_type=='department':
                        subject=common_subject; code=common_code
                        row_sem=row_sem or semester
                    else:
                        subject=str(row.get('subject_name','')).strip()
                        code=str(row.get('subject_code','')).strip()
                        if not row_dept or not row_sem or not subject:
                            raise ValueError(f'{source_name}, row {row_no}: department, semester and subject name are required.')
                    records.append({'usn':usn,'student_name':name,'department':row_dept,
                        'semester':row_sem,'subject_name':subject,'subject_code':code})
            if not records: raise ValueError('The uploaded Excel file contains no valid student rows.')
            rooms=Classroom.query.order_by(Classroom.classroom_number).all()
            staff=Invigilator.query.order_by(Invigilator.faculty_name).all()
            if not rooms: raise ValueError('Add classrooms before creating an examination.')
            if not staff: raise ValueError('Add/import invigilators before creating an examination.')
            total=sum(r.total_seats for r in rooms)
            if len(records)>total: raise ValueError(f'Not enough seats: {len(records)} students but only {total} seats are registered.')
            exam=Examination(exam_name=exam_name,exam_date=exam_date,session=session_name,
                subject_name=common_subject if exam_type!='combined' else 'Multiple subjects',
                subject_code=common_code if exam_type!='combined' else '',
                semester=semester,exam_type=exam_type,mixing_limit=mixing_limit)
            db.session.add(exam); db.session.flush()
            for rec in records: db.session.add(ExaminationStudent(examination_id=exam.id,**rec))
            db.session.commit()
            flash(f'Examination created successfully for {len(records)} students. Generate seating to allocate rooms and invigilators.','success')
            return redirect(url_for('core.seating'))
        except Exception as e:
            db.session.rollback(); flash(str(e),'danger')
    return render_template('examinations/form.html',item=None,departments=departments,
                           selected_exam_type=selected_exam_type)

@bp.route('/examinations/<int:item_id>/edit',methods=['GET','POST'])
def exam_edit(item_id):
    item=Examination.query.get_or_404(item_id)
    if request.method=='POST':
        for f in ['exam_name','exam_date','session','subject_name','subject_code','semester']: setattr(item,f,request.form.get(f,'').strip())
        if not item.exam_name or not item.exam_date: flash('Exam name and date are required.','danger')
        else:
            RoomBooking.query.filter_by(examination_id=item.id).delete(synchronize_session=False)
            InvigilatorAssignment.query.filter_by(examination_id=item.id).delete(synchronize_session=False)
            ExaminationSupervisor.query.filter_by(examination_id=item.id).delete(synchronize_session=False)
            for st in ExaminationStudent.query.filter_by(examination_id=item.id).all():
                st.classroom_id=None; st.seat_number=None
            db.session.commit(); flash('Examination updated. Generate seating again to reserve rooms and faculty for the revised date/shift.','success'); return redirect(url_for('core.examinations'))
    return render_template('examinations/edit.html',item=item)
@bp.route('/examinations/<int:item_id>/delete',methods=['POST'])
def exam_delete(item_id):
    item=Examination.query.get_or_404(item_id)
    RoomBooking.query.filter_by(examination_id=item.id).delete(synchronize_session=False)
    InvigilatorAssignment.query.filter_by(examination_id=item.id).delete(synchronize_session=False)
    ExaminationSupervisor.query.filter_by(examination_id=item.id).delete(synchronize_session=False)
    AttendanceRecord.query.filter_by(examination_id=item.id).delete(synchronize_session=False)
    ExaminationStudent.query.filter_by(examination_id=item.id).delete(synchronize_session=False)
    db.session.delete(item); db.session.commit(); flash('Examination deleted.','success'); return redirect(url_for('core.examinations'))

@bp.route('/classrooms')
def classrooms(): return render_template('classrooms/list.html',items=Classroom.query.order_by(Classroom.classroom_number).all())
@bp.route('/classrooms/add',methods=['GET','POST'])
def classroom_add():
    if request.method=='POST':
        no=request.form.get('classroom_number','').strip(); cap=request.form.get('total_seats','').strip()
        try:
            if not no or int(cap)<1: raise ValueError()
            if Classroom.query.filter_by(classroom_number=no).first(): flash('That classroom already exists.','danger')
            else: db.session.add(Classroom(classroom_number=no,total_seats=int(cap))); db.session.commit(); flash('Classroom saved.','success'); return redirect(url_for('core.classrooms'))
        except ValueError: flash('Enter a classroom number and a positive seat capacity.','danger')
    return render_template('classrooms/form.html',item=None)
@bp.route('/classrooms/import',methods=['POST'])
def classroom_import():
    try:
        df=read_excel(request.files.get('file'),['classroom_number','total_seats']); added=0; skipped=[]
        for i,(_,r) in enumerate(df.iterrows(),2):
            no=str(r['classroom_number']).strip()
            try: cap=int(r['total_seats'])
            except: cap=0
            if not no or cap<1 or Classroom.query.filter_by(classroom_number=no).first(): skipped.append(f'Row {i}: invalid or duplicate'); continue
            db.session.add(Classroom(classroom_number=no,total_seats=cap)); added+=1
        db.session.commit(); flash(f'Imported {added} classrooms; skipped {len(skipped)}.','success')
    except Exception as e: db.session.rollback(); flash(str(e),'danger')
    return redirect(url_for('core.classrooms'))
@bp.route('/classrooms/<int:item_id>/edit',methods=['GET','POST'])
def classroom_edit(item_id):
    item=Classroom.query.get_or_404(item_id)
    if request.method=='POST':
        number=request.form.get('classroom_number','').strip()
        try: capacity=int(request.form.get('total_seats','0'))
        except ValueError: capacity=0
        duplicate=Classroom.query.filter(Classroom.classroom_number==number,Classroom.id!=item.id).first()
        if not number or capacity<1: flash('Enter a classroom number and a positive seat capacity.','danger')
        elif duplicate: flash('That classroom number is already registered.','danger')
        else:
            item.classroom_number=number; item.total_seats=capacity; db.session.commit()
            flash('Classroom updated.','success'); return redirect(url_for('core.classrooms'))
    return render_template('classrooms/form.html',item=item)

@bp.route('/classrooms/<int:item_id>/delete',methods=['POST'])
def classroom_delete(item_id): db.session.delete(Classroom.query.get_or_404(item_id)); db.session.commit(); flash('Classroom deleted.','success'); return redirect(url_for('core.classrooms'))

@bp.route('/invigilators')
def invigilators(): return render_template('invigilators/list.html',items=Invigilator.query.order_by(Invigilator.faculty_name).all())
@bp.route('/invigilators/import',methods=['GET','POST'])
def invigilator_import():
    if request.method=='POST':
        try:
            df=read_excel(request.files.get('file'),['faculty_id','faculty_name']); added,skipped=import_rows(Invigilator,df,{'faculty_id':'faculty_id','faculty_name':'faculty_name','department':'department','designation':'designation'},'faculty_id')
            flash(f'Imported {added} invigilators; skipped {len(skipped)}.','success')
            for x in skipped[:8]: flash(x,'warning')
        except Exception as e: db.session.rollback(); flash(str(e),'danger')
        return redirect(url_for('core.invigilators'))
    return render_template('invigilators/import.html')
@bp.route('/invigilators/add',methods=['GET','POST'])
def invigilator_add():
    if request.method=='POST':
        vals={f:request.form.get(f,'').strip() for f in ['faculty_id','faculty_name','department','designation']}
        if not vals['faculty_id'] or not vals['faculty_name']: flash('Faculty ID and name are required.','danger')
        elif Invigilator.query.filter_by(faculty_id=vals['faculty_id']).first(): flash('Faculty ID already exists.','danger')
        else: db.session.add(Invigilator(**vals)); db.session.commit(); flash('Invigilator saved.','success'); return redirect(url_for('core.invigilators'))
    return render_template('invigilators/form.html')
@bp.route('/invigilators/<int:item_id>/delete',methods=['POST'])
def invigilator_delete(item_id): db.session.delete(Invigilator.query.get_or_404(item_id)); db.session.commit(); flash('Invigilator deleted.','success'); return redirect(url_for('core.invigilators'))

@bp.route('/seating',methods=['GET','POST'])
def seating():
    if request.method=='POST':
        exam=Examination.query.get_or_404(int(request.form.get('exam_id')))
        roster=ExaminationStudent.query.filter_by(examination_id=exam.id).order_by(ExaminationStudent.department,ExaminationStudent.semester,ExaminationStudent.usn).all()
        if not roster:
            # Legacy exams use the master student list, filtered by semester if selected.
            q=Student.query.filter_by(semester=exam.semester) if exam.semester else Student.query
            legacy=q.order_by(Student.department,Student.usn).all()
            if not legacy:
                flash('No students are available for this examination. Upload the examination student Excel roster first.','warning')
                return redirect(url_for('core.seating'))
            roster=[ExaminationStudent(usn=s.usn,student_name=s.student_name,department=s.department,
                semester=s.semester or '',subject_name=exam.subject_name or '',subject_code=exam.subject_code or '') for s in legacy]
            # Legacy-only records are rendered but cannot be persisted on the transient roster objects.
        try:
            busy_rooms=_busy_room_ids(exam)
            free_rooms=[r for r in Classroom.query.order_by(Classroom.classroom_number).all() if r.id not in busy_rooms]
            placement_plan=_plan_seating(roster,free_rooms,exam.mixing_limit or 1)
            rooms=list(dict.fromkeys(item[0] for item in placement_plan))
            rooms,duties,supervisor=_allocate_resources(exam,roster,planned_rooms=rooms)
            rows=[]
            for room,seat,st in placement_plan:
                if isinstance(st,ExaminationStudent):
                    st.classroom_id=room.id; st.seat_number=seat
                rows.append((room.classroom_number,seat,st.usn,st.student_name,st.department,st.semester or '',st.subject_name or exam.subject_name or ''))
            db.session.commit()
            room_links=[(room.classroom_number,fac.faculty_name,fac.faculty_id) for room,fac in duties]
            session['last_seating']={'exam':exam.exam_name,'rows':rows}
            session['last_duties']={'exam':exam.exam_name,'rows':[(r.classroom_number,f.faculty_id,f.faculty_name,f.department or '') for r,f in duties]+[('Overall Supervisor',supervisor.faculty_id,supervisor.faculty_name,supervisor.department or '')]}
            return render_template('seating/result.html',exam=exam,rows=rows,room_links=room_links,supervisor=supervisor)
        except ValueError as e:
            db.session.rollback(); flash(str(e),'danger')
            return redirect(url_for('core.seating'))
    exams=Examination.query.order_by(Examination.exam_date,Examination.session).all()
    return render_template('seating/form.html',exams=exams,students=Student.query.count(),rooms=Classroom.query.count())

@bp.route('/invigilators/allocate',methods=['GET','POST'])
def allocate_invigilators():
    if request.method=='POST':
        exam=Examination.query.get_or_404(int(request.form.get('exam_id')))
        roster=ExaminationStudent.query.filter_by(examination_id=exam.id).all()
        if not roster:
            q=Student.query.filter_by(semester=exam.semester) if exam.semester else Student.query
            roster=q.all()
        if not roster:
            flash('This examination has no student roster. Upload its Excel student list first.','warning')
            return redirect(url_for('core.allocate_invigilators'))
        try:
            rooms,duties,supervisor=_allocate_resources(exam,roster)
            db.session.commit()
            rows=[(r.classroom_number,f.faculty_id,f.faculty_name,f.department or '') for r,f in duties]
            rows.append(('Overall Supervisor',supervisor.faculty_id,supervisor.faculty_name,supervisor.department or ''))
            session['last_duties']={'exam':exam.exam_name,'rows':rows}
            return render_template('invigilators/duties.html',exam=exam,rows=rows,supervisor=supervisor)
        except ValueError as e:
            db.session.rollback(); flash(str(e),'danger')
            return redirect(url_for('core.allocate_invigilators'))
    return render_template('invigilators/allocate.html',exams=Examination.query.order_by(Examination.exam_date).all(),rooms=Classroom.query.count(),staff=Invigilator.query.count())

@bp.route('/attendance',methods=['GET','POST'])
def attendance():
    if request.method=='POST':
        exam=Examination.query.get_or_404(int(request.form.get('exam_id')))
        roster=ExaminationStudent.query.filter_by(examination_id=exam.id).order_by(ExaminationStudent.department,ExaminationStudent.usn).all()
        if roster:
            groups={}
            for st in roster: groups.setdefault(st.department,[]).append(st)
        else:
            q=Student.query.filter_by(semester=exam.semester) if exam.semester else Student.query
            students=q.order_by(Student.department,Student.usn).all()
            groups={}
            for st in students: groups.setdefault(st.department,[]).append(st)
        saved={x.usn:x for x in AttendanceRecord.query.filter_by(examination_id=exam.id).all()}
        session['attendance_exam_id']=exam.id
        session['last_attendance']={'exam':exam.exam_name,'subject':exam.subject_name or '','rows':[(d,s.usn,s.student_name,(saved[s.usn].status if s.usn in saved else 'Absent'),(saved[s.usn].remarks if s.usn in saved else '')) for d,items in groups.items() for s in items]}
        return render_template('attendance/result.html',exam=exam,groups=groups,saved=saved)
    return render_template('attendance/form.html',exams=Examination.query.order_by(Examination.exam_date).all())

@bp.route('/attendance/<int:exam_id>/save',methods=['POST'])
def attendance_save(exam_id):
    exam=Examination.query.get_or_404(exam_id)
    roster=ExaminationStudent.query.filter_by(examination_id=exam.id).all()
    if not roster:
        q=Student.query.filter_by(semester=exam.semester) if exam.semester else Student.query
        roster=q.all()
    for st in roster:
        status=request.form.get('status_'+st.usn,'Absent')
        if status not in ('Present','Absent'): status='Absent'
        remarks=request.form.get('remarks_'+st.usn,'').strip()[:250]
        rec=AttendanceRecord.query.filter_by(examination_id=exam.id,usn=st.usn).first()
        if rec:
            rec.status=status; rec.remarks=remarks; rec.student_name=st.student_name; rec.department=st.department
        else:
            db.session.add(AttendanceRecord(examination_id=exam.id,usn=st.usn,student_name=st.student_name,
                department=st.department,status=status,remarks=remarks))
    db.session.commit()
    session['last_attendance']={'exam':exam.exam_name,'subject':exam.subject_name or '',
        'rows':[(st.department,st.usn,st.student_name,
                 AttendanceRecord.query.filter_by(examination_id=exam.id,usn=st.usn).first().status,
                 AttendanceRecord.query.filter_by(examination_id=exam.id,usn=st.usn).first().remarks or '') for st in roster]}
    flash('Attendance saved successfully. You can reopen the same examination to update it.','success')
    return redirect(url_for('core.attendance'))


@bp.route('/reports/export-current/<kind>',methods=['POST'])
def export_current(kind):
    payload=session.get({'seating':'last_seating','duties':'last_duties','attendance':'last_attendance'}.get(kind,''))
    if not payload: flash('Generate the preview first, then export it.','warning'); return redirect(url_for('core.reports'))
    if kind=='seating': headers=['Classroom','Seat','USN','Student','Department','Semester','Subject']
    elif kind=='duties': headers=['Classroom','Faculty ID','Invigilator','Department']
    elif kind=='attendance': headers=['Department','USN','Student Name','Signature','Remarks']
    else: flash('Unknown export type.','danger'); return redirect(url_for('core.reports'))
    pdf,docx=make_report(payload['exam']+' - '+kind.title(),headers,payload['rows'],current_app.config['GENERATED_FOLDER'],kind)
    flash(f'Created {pdf} and {docx}.','success')
    return redirect(url_for('core.reports'))

@bp.route('/reports')
def reports():
    files=[]; folder=Path(current_app.config['GENERATED_FOLDER'])
    for p in sorted(folder.glob('*'),key=lambda x:x.stat().st_mtime,reverse=True):
        if p.is_file(): files.append(p.name)
    return render_template('reports/list.html',files=files)
@bp.route('/reports/generate/<kind>',methods=['POST'])
def report_generate(kind):
    folder=current_app.config['GENERATED_FOLDER']
    if kind=='students': rows=[(s.usn,s.student_name,s.department,s.semester or '') for s in Student.query.order_by(Student.department,Student.usn).all()]; title='Student Register'; headers=['USN','Student Name','Department','Semester']
    elif kind=='classrooms': rows=[(r.classroom_number,r.total_seats) for r in Classroom.query.all()]; title='Classroom Register'; headers=['Classroom','Capacity']
    elif kind=='invigilators': rows=[(i.faculty_id,i.faculty_name,i.department or '',i.designation or '') for i in Invigilator.query.all()]; title='Invigilator Register'; headers=['Faculty ID','Name','Department','Designation']
    else: flash('Unknown report type.','danger'); return redirect(url_for('core.reports'))
    pdf,docx=make_report(title,headers,rows,folder,kind); flash(f'Reports generated: {pdf} and {docx}','success'); return redirect(url_for('core.reports'))
@bp.route('/reports/download/<path:filename>')
def report_download(filename): return send_from_directory(current_app.config['GENERATED_FOLDER'],filename,as_attachment=True)

@bp.route('/backup',methods=['GET','POST'])
def backup():
    folder=Path(current_app.config['BACKUP_FOLDER']); folder.mkdir(parents=True,exist_ok=True)
    if request.method=='POST':
        src=Path(current_app.root_path)/'database'/'exam.db'; dest=folder/f'exam_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.db'; shutil.copy2(src,dest); flash(f'Backup created: {dest.name}','success'); return redirect(url_for('core.backup'))
    return render_template('backup.html',files=sorted([p.name for p in folder.glob('*.db')],reverse=True))
@bp.route('/backup/download/<path:filename>')
def backup_download(filename): return send_from_directory(current_app.config['BACKUP_FOLDER'],filename,as_attachment=True)


@bp.route('/examinations/<int:item_id>/form-b', methods=['POST'])
def examination_form_b(item_id):
    exam=Examination.query.get_or_404(item_id)
    students=ExaminationStudent.query.filter_by(examination_id=exam.id).order_by(ExaminationStudent.department,ExaminationStudent.usn).all()
    if not students:
        flash('No uploaded student roster exists for this examination.','warning')
        return redirect(url_for('core.examinations'))
    folder=Path(current_app.config['GENERATED_FOLDER']); folder.mkdir(parents=True,exist_ok=True)
    stamp=datetime.now().strftime('%Y%m%d_%H%M%S')
    safe=''.join(c if c.isalnum() or c in '-_' else '_' for c in exam.exam_name)[:40] or 'exam'
    archive_path=folder/f'{safe}_FormB_{stamp}.zip'
    from docx import Document
    from docx.shared import Inches
    with zipfile.ZipFile(archive_path,'w',zipfile.ZIP_DEFLATED) as archive:
        for n,st in enumerate(students,1):
            doc=Document(); section=doc.sections[0]
            section.top_margin=Inches(0.55); section.bottom_margin=Inches(0.55)
            doc.add_heading('POOJYA DODDAPPA APPA COLLEGE OF ENGINEERING',0)
            p=doc.add_paragraph('KALABURAGI, KARNATAKA'); p.alignment=1
            doc.add_heading('FORM B - STUDENT EXAMINATION DETAILS',1)
            doc.add_paragraph('Examination: '+exam.exam_name)
            doc.add_paragraph('Date: '+exam.exam_date+'    Session: '+(exam.session or ''))
            doc.add_paragraph('Department: '+st.department)
            doc.add_paragraph('Semester: '+(st.semester or ''))
            doc.add_paragraph('Subject: '+st.subject_name+(f' ({st.subject_code})' if st.subject_code else ''))
            doc.add_paragraph('USN: '+st.usn); doc.add_paragraph('Student Name: '+st.student_name)
            room=Classroom.query.get(st.classroom_id) if st.classroom_id else None
            if room: doc.add_paragraph(f'Classroom: {room.classroom_number}    Seat: {st.seat_number}')
            doc.add_paragraph('Student Signature: ______________________________')
            doc.add_paragraph('Invigilator Signature: ____________________________')
            doc.add_paragraph('Chief Superintendent / Overall Supervisor: ____________________________')
            doc.add_paragraph('Office Use: Attendance ______   Answer Booklet No. ______')
            stream=io.BytesIO(); doc.save(stream)
            safe_usn=''.join(c if c.isalnum() or c in '-_' else '_' for c in st.usn)[:40] or f'student_{n}'
            archive.writestr(f'Form_B_{safe_usn}.docx',stream.getvalue())
    return send_from_directory(folder,archive_path.name,as_attachment=True)

# ---------------------------------------------------------------------------
# Department-specific examinations. These use separate tables and never alter
# the existing college-wide Examination / Student workflow.
# ---------------------------------------------------------------------------
@bp.route('/department-examinations', methods=['GET'])
def department_examinations():
    return redirect(url_for('core.examinations'))

@bp.route('/department-examinations/add', methods=['GET', 'POST'])
def department_exam_add():
    return redirect(url_for('core.exam_add'))

# Legacy department exam creation implementation retained below for compatibility.
@bp.route('/department-examinations/add-legacy', methods=['GET', 'POST'])
def department_exam_add_legacy():
    departments = sorted({x.department for x in Student.query.with_entities(Student.department).all() if x.department})
    departments = sorted(set(departments) | {x.department for x in DepartmentExamination.query.with_entities(DepartmentExamination.department).all() if x.department})
    # Department examinations automatically use all registered classrooms and
    # distribute the registered invigilators across those rooms.
    rooms = Classroom.query.order_by(Classroom.classroom_number).all()
    staff = Invigilator.query.order_by(Invigilator.faculty_name).all()
    if request.method == 'POST':
        try:
            dept = request.form.get('department', '').strip()
            exam_name = request.form.get('exam_name', '').strip()
            exam_date = request.form.get('exam_date', '').strip()
            subject = request.form.get('subject_name', '').strip()
            subject_code = request.form.get('subject_code', '').strip()
            semester = request.form.get('semester', '').strip()
            session_name = request.form.get('session', '').strip()
            if not all([dept, exam_name, exam_date, subject]):
                raise ValueError('Department, examination name, date and subject are required.')
            df = read_excel(request.files.get('file'), ['usn', 'student_name'])
            if 'department' in df.columns:
                mismatched = df[df['department'].astype(str).str.strip().str.casefold() != dept.casefold()]
                if len(mismatched):
                    raise ValueError(f'{len(mismatched)} uploaded student row(s) do not belong to {dept}.')
            if 'semester' in df.columns and semester:
                df = df[df['semester'].astype(str).str.strip().str.replace(r'\\.0$', '', regex=True) == semester]
            if df.empty:
                raise ValueError('The uploaded Excel file contains no students for this examination.')
            # No manual room/faculty selection in this workflow:
            # use every classroom and invigilator already registered in ExamDesk.
            selected_rooms = Classroom.query.order_by(Classroom.classroom_number).all()
            selected_staff = Invigilator.query.order_by(Invigilator.faculty_name).all()
            if not selected_rooms:
                raise ValueError('No classrooms are registered yet. Add classrooms in the Classrooms section first.')
            if not selected_staff:
                raise ValueError('No invigilators are registered yet. Import/add invigilators in the Invigilators section first.')
            records = []
            seen = set()
            for row_no, (_, row) in enumerate(df.iterrows(), start=2):
                usn = str(row.get('usn', '')).strip()
                name = str(row.get('student_name', '')).strip()
                row_dept = str(row.get('department', dept)).strip() or dept
                row_sem = str(row.get('semester', semester)).strip() or semester
                if not usn or not name:
                    continue
                if row_dept.casefold() != dept.casefold():
                    raise ValueError(f'Row {row_no}: student department does not match {dept}.')
                if usn.casefold() in seen:
                    raise ValueError(f'Duplicate USN in uploaded file: {usn}.')
                seen.add(usn.casefold())
                records.append((usn, name, dept, row_sem))
            if not records:
                raise ValueError('No valid student records found in the Excel file.')
            capacity = sum(x.total_seats for x in selected_rooms)
            if len(records) > capacity:
                raise ValueError(f'Not enough seats: {len(records)} students but selected classrooms have only {capacity} seats.')
            exam = DepartmentExamination(exam_name=exam_name, department=dept, exam_date=exam_date,
                    session=session_name, subject_name=subject, subject_code=subject_code, semester=semester)
            db.session.add(exam)
            db.session.flush()
            for usn, name, student_dept, sem in records:
                db.session.add(DepartmentExamStudent(examination_id=exam.id, usn=usn,
                    student_name=name, department=student_dept, semester=sem))
            for index, room in enumerate(selected_rooms):
                db.session.add(DepartmentExamRoom(examination_id=exam.id, classroom_id=room.id,
                    invigilator_id=selected_staff[index % len(selected_staff)].id))
            db.session.commit()
            flash(f'Department examination created for {dept} with {len(records)} students. All {len(selected_rooms)} registered classrooms were used and invigilators were assigned automatically.', 'success')
            return redirect(url_for('core.department_exam_detail', item_id=exam.id))
        except Exception as e:
            db.session.rollback()
            flash(str(e), 'danger')
    return render_template('department_exams/form.html', departments=departments, rooms=rooms, staff=staff)

@bp.route('/department-examinations/<int:item_id>')
def department_exam_detail(item_id):
    exam = DepartmentExamination.query.get_or_404(item_id)
    students = DepartmentExamStudent.query.filter_by(examination_id=exam.id).order_by(DepartmentExamStudent.usn).all()
    room_links = DepartmentExamRoom.query.filter_by(examination_id=exam.id).all()
    assignments = []
    index = 0
    for link in room_links:
        for seat in range(1, link.classroom.total_seats + 1):
            if index >= len(students):
                break
            st = students[index]
            assignments.append((link.classroom.classroom_number, seat, st.usn, st.student_name, st.department))
            index += 1
    return render_template('department_exams/result.html', exam=exam, students=students,
                           rooms=room_links, rows=assignments)

@bp.route('/department-examinations/<int:item_id>/export', methods=['POST'])
def department_exam_export(item_id):
    exam = DepartmentExamination.query.get_or_404(item_id)
    students = DepartmentExamStudent.query.filter_by(examination_id=exam.id).order_by(DepartmentExamStudent.usn).all()
    room_links = DepartmentExamRoom.query.filter_by(examination_id=exam.id).all()
    assignments = []
    index = 0
    for link in room_links:
        for seat in range(1, link.classroom.total_seats + 1):
            if index >= len(students): break
            st = students[index]
            assignments.append((link.classroom.classroom_number, seat, st.usn, st.student_name, st.department))
            index += 1
    folder = Path(current_app.config['GENERATED_FOLDER'])
    folder.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    safe = ''.join(c if c.isalnum() or c in '-_' else '_' for c in exam.exam_name)[:40] or 'exam'
    safe_dept = ''.join(c if c.isalnum() or c in '-_' else '_' for c in exam.department)[:30] or 'department'
    archive_path = folder / f'{safe}_{safe_dept}_FormB_{stamp}.zip'
    from docx import Document
    from docx.shared import Inches
    with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as archive:
        for n, st in enumerate(students, start=1):
            doc = Document()
            doc.add_heading('FORM B - STUDENT EXAMINATION DETAILS', 0)
            doc.add_paragraph('Examination: ' + exam.exam_name)
            doc.add_paragraph('Department: ' + exam.department)
            doc.add_paragraph('Subject: ' + exam.subject_name + (f' ({exam.subject_code})' if exam.subject_code else ''))
            doc.add_paragraph('Date: ' + exam.exam_date + '    Session: ' + (exam.session or ''))
            doc.add_paragraph('USN: ' + st.usn)
            doc.add_paragraph('Student Name: ' + st.student_name)
            doc.add_paragraph('Semester: ' + (st.semester or exam.semester or ''))
            placement = next((a for a in assignments if a[2] == st.usn), None)
            if placement:
                doc.add_paragraph(f'Classroom: {placement[0]}    Bench/Seat: {placement[1]}')
            doc.add_paragraph('Student Signature: ______________________________')
            doc.add_paragraph('Invigilator Signature: ____________________________')
            doc.add_paragraph('Chief Superintendent / Overall Supervisor: ____________________________')
            doc.add_paragraph('Office Use: Attendance ______   Answer Booklet No. ______')
            stream = io.BytesIO()
            doc.save(stream)
            safe_usn = ''.join(c if c.isalnum() or c in '-_' else '_' for c in st.usn)[:40] or f'student_{n}'
            archive.writestr(f'Form_B_{safe_usn}.docx', stream.getvalue())
        pdf, docx = make_report(f'{exam.exam_name} - {exam.department} Seating', 
            ['Classroom','Seat','USN','Student','Department'], assignments,
            str(folder), f'{safe}_{safe_dept}_seating')
        archive.write(folder / pdf, arcname=pdf)
        archive.write(folder / docx, arcname=docx)
    return send_from_directory(folder, archive_path.name, as_attachment=True)

@bp.route('/department-examinations/<int:item_id>/delete', methods=['POST'])
def department_exam_delete(item_id):
    exam = DepartmentExamination.query.get_or_404(item_id)
    db.session.delete(exam)
    db.session.commit()
    flash('Department examination deleted.', 'success')
    return redirect(url_for('core.department_examinations'))
