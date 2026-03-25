# Punto de entrada WSGI para hosting con cPanel (Passenger) o gunicorn
# cPanel: apunta "Application startup file" a este archivo y "Application Entry point" a "application"
# Gunicorn: gunicorn --bind 0.0.0.0:8000 wsgi:application

from app import app as application

if __name__ == '__main__':
    application.run()
