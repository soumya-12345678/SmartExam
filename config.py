import os
BASE_DIR=os.path.abspath(os.path.dirname(__file__))
class Config:
    SECRET_KEY=os.environ.get('SMARTEXAM_SECRET_KEY','change-this-local-development-key')
    SQLALCHEMY_DATABASE_URI='sqlite:///'+os.path.join(BASE_DIR,'database','exam.db')
    SQLALCHEMY_TRACK_MODIFICATIONS=False
    UPLOAD_FOLDER=os.path.join(BASE_DIR,'uploads')
    GENERATED_FOLDER=os.path.join(BASE_DIR,'generated_reports')
    BACKUP_FOLDER=os.path.join(BASE_DIR,'backups')
    MAX_CONTENT_LENGTH=20*1024*1024
