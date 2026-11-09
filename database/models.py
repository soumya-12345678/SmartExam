from database import db
class Student(db.Model):
    id=db.Column(db.Integer, primary_key=True)
    usn=db.Column(db.String(30), unique=True, nullable=False)
    student_name=db.Column(db.String(100), nullable=False)
    department=db.Column(db.String(50), nullable=False)
    semester=db.Column(db.String(20))
class Examination(db.Model):
    id=db.Column(db.Integer, primary_key=True)
    exam_name=db.Column(db.String(150), nullable=False)
    exam_date=db.Column(db.String(20), nullable=False)
    session=db.Column(db.String(30))
    subject_name=db.Column(db.String(150))
    subject_code=db.Column(db.String(50))
    semester=db.Column(db.String(20))
    exam_type=db.Column(db.String(30), default='general', nullable=False)
    mixing_limit=db.Column(db.Integer, default=1, nullable=False)
    students=db.relationship('ExaminationStudent', backref='examination', cascade='all, delete-orphan', lazy=True)
class ExaminationStudent(db.Model):
    id=db.Column(db.Integer, primary_key=True)
    examination_id=db.Column(db.Integer, db.ForeignKey('examination.id'), nullable=False, index=True)
    usn=db.Column(db.String(30), nullable=False)
    student_name=db.Column(db.String(100), nullable=False)
    department=db.Column(db.String(50), nullable=False)
    semester=db.Column(db.String(20))
    subject_name=db.Column(db.String(150), nullable=False)
    subject_code=db.Column(db.String(50))
    classroom_id=db.Column(db.Integer, db.ForeignKey('classroom.id'))
    seat_number=db.Column(db.Integer)
    __table_args__=(db.UniqueConstraint('examination_id','usn',name='uq_exam_student_usn'),)
class Classroom(db.Model):
    id=db.Column(db.Integer, primary_key=True)
    classroom_number=db.Column(db.String(50), unique=True, nullable=False)
    total_seats=db.Column(db.Integer, nullable=False)
class Invigilator(db.Model):
    id=db.Column(db.Integer, primary_key=True)
    faculty_id=db.Column(db.String(30), unique=True, nullable=False)
    faculty_name=db.Column(db.String(100), nullable=False)
    department=db.Column(db.String(50))
    designation=db.Column(db.String(100))


class DepartmentExamination(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    exam_name = db.Column(db.String(150), nullable=False)
    department = db.Column(db.String(50), nullable=False, index=True)
    exam_date = db.Column(db.String(20), nullable=False)
    session = db.Column(db.String(30))
    subject_name = db.Column(db.String(150), nullable=False)
    subject_code = db.Column(db.String(50))
    semester = db.Column(db.String(20))
    created_at = db.Column(db.DateTime, default=db.func.now())
    students = db.relationship('DepartmentExamStudent', backref='examination',
                                cascade='all, delete-orphan', lazy=True)
    room_assignments = db.relationship('DepartmentExamRoom', backref='examination',
                                cascade='all, delete-orphan', lazy=True)

class DepartmentExamStudent(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    examination_id = db.Column(db.Integer, db.ForeignKey('department_examination.id'), nullable=False, index=True)
    usn = db.Column(db.String(30), nullable=False)
    student_name = db.Column(db.String(100), nullable=False)
    department = db.Column(db.String(50), nullable=False)
    semester = db.Column(db.String(20))
    __table_args__ = (db.UniqueConstraint('examination_id', 'usn', name='uq_dept_exam_student'),)

class DepartmentExamRoom(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    examination_id = db.Column(db.Integer, db.ForeignKey('department_examination.id'), nullable=False, index=True)
    classroom_id = db.Column(db.Integer, db.ForeignKey('classroom.id'), nullable=False)
    invigilator_id = db.Column(db.Integer, db.ForeignKey('invigilator.id'))
    classroom = db.relationship('Classroom')
    invigilator = db.relationship('Invigilator')
    __table_args__ = (db.UniqueConstraint('examination_id', 'classroom_id', name='uq_dept_exam_room'),)


class RoomBooking(db.Model):
    id=db.Column(db.Integer, primary_key=True)
    examination_id=db.Column(db.Integer, db.ForeignKey('examination.id'), nullable=False, index=True)
    classroom_id=db.Column(db.Integer, db.ForeignKey('classroom.id'), nullable=False)
    examination=db.relationship('Examination')
    classroom=db.relationship('Classroom')
    __table_args__=(db.UniqueConstraint('examination_id','classroom_id',name='uq_room_booking_exam_room'),)

class InvigilatorAssignment(db.Model):
    id=db.Column(db.Integer, primary_key=True)
    examination_id=db.Column(db.Integer, db.ForeignKey('examination.id'), nullable=False, index=True)
    classroom_id=db.Column(db.Integer, db.ForeignKey('classroom.id'), nullable=False)
    invigilator_id=db.Column(db.Integer, db.ForeignKey('invigilator.id'), nullable=False)
    examination=db.relationship('Examination')
    classroom=db.relationship('Classroom')
    invigilator=db.relationship('Invigilator')
    __table_args__=(db.UniqueConstraint('examination_id','classroom_id',name='uq_exam_room_duty'),
                    db.UniqueConstraint('examination_id','invigilator_id',name='uq_exam_faculty_duty'),)

class ExaminationSupervisor(db.Model):
    id=db.Column(db.Integer, primary_key=True)
    examination_id=db.Column(db.Integer, db.ForeignKey('examination.id'), nullable=False, unique=True, index=True)
    invigilator_id=db.Column(db.Integer, db.ForeignKey('invigilator.id'), nullable=False)
    examination=db.relationship('Examination')
    invigilator=db.relationship('Invigilator')

class AttendanceRecord(db.Model):
    id=db.Column(db.Integer, primary_key=True)
    examination_id=db.Column(db.Integer, db.ForeignKey('examination.id'), nullable=False, index=True)
    usn=db.Column(db.String(30), nullable=False)
    student_name=db.Column(db.String(100), nullable=False)
    department=db.Column(db.String(50), nullable=False)
    status=db.Column(db.String(12), nullable=False, default='Absent')
    remarks=db.Column(db.String(250), default='')
    marked_at=db.Column(db.DateTime, default=db.func.now(), onupdate=db.func.now())
    __table_args__=(db.UniqueConstraint('examination_id','usn',name='uq_attendance_exam_usn'),)
