import os
from flask import Flask, render_template
from config import Config
from database import db
from database.models import Student, Examination, ExaminationStudent, Classroom, Invigilator
from routes.core import bp

def create_app():
    app=Flask(__name__); app.config.from_object(Config)
    for key in ['UPLOAD_FOLDER','GENERATED_FOLDER','BACKUP_FOLDER']:
        os.makedirs(app.config[key],exist_ok=True)
    db.init_app(app); app.register_blueprint(bp)
    with app.app_context():
        db.create_all()
        # Lightweight additive migration for databases created by earlier ExamDesk versions.
        from sqlalchemy import inspect, text
        inspector=inspect(db.engine)
        if 'examination' in inspector.get_table_names():
            columns={c['name'] for c in inspector.get_columns('examination')}
            with db.engine.begin() as conn:
                if 'exam_type' not in columns:
                    conn.execute(text("ALTER TABLE examination ADD COLUMN exam_type VARCHAR(30) NOT NULL DEFAULT 'general'"))
                if 'mixing_limit' not in columns:
                    conn.execute(text('ALTER TABLE examination ADD COLUMN mixing_limit INTEGER NOT NULL DEFAULT 1'))
    return app
app=create_app()
if __name__=='__main__': app.run(debug=True)
