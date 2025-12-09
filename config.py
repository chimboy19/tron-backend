# backend/config.py
import os
from dotenv import load_dotenv
from datetime import timedelta
load_dotenv()

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'hard-to-guess-string-for-dev-only'
    SQLALCHEMY_DATABASE_URI = os.environ.get('DATABASE_URL') or 'sqlite:///tokyo_tron.db'
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    JWT_SECRET_KEY = os.getenv('JWT_SECRET_KEY') or 'jwt-secret-string-for-dev-only'
    JWT_ACCESS_TOKEN_EXPIRES = timedelta(hours=24)  

    # API Keys loaded from environment
    DIGIKEY_CLIENT_ID = os.getenv('DIGIKEY_CLIENT_ID')
    DIGIKEY_CLIENT_SECRET = os.getenv('DIGIKEY_CLIENT_SECRET')
    OPENAI_API_KEY = os.getenv('OPENAI_API_KEY')

    # Email settings (if using SMTP)
    MAIL_SERVER = os.environ.get('MAIL_SERVER')
    MAIL_PORT = int(os.environ.get('MAIL_PORT', 587))
    MAIL_USE_TLS = os.environ.get('MAIL_USE_TLS', 'True').lower() in ['true', 'on', '1']
    MAIL_USERNAME = os.environ.get('MAIL_USERNAME')
    MAIL_PASSWORD = os.environ.get('MAIL_PASSWORD')
    GRAPH_CLIENT_ID = os.getenv("GRAPH_CLIENT_ID")
    GRAPH_TENANT_ID = os.getenv("GRAPH_TENANT_ID")
    GRAPH_CLIENT_SECRET = os.getenv("GRAPH_CLIENT_SECRET")
    GRAPH_MAILBOX = os.getenv("GRAPH_MAILBOX")

    # Application settings
    CHECK_INTERVAL = int(os.getenv('CHECK_INTERVAL', '30'))
    PROCESSED_FOLDER_NAME = os.environ.get('PROCESSED_FOLDER_NAME', 'Quotations Processed')
    OUTLOOK_PROFILE_NAME = os.getenv('OUTLOOK_PROFILE_NAME') 