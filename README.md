# College Event Report Generation System

A web-based application for managing college events and generating professional event reports with college letterhead in Word format.

## Features

- **User Authentication**: Secure registration and login system with password hashing
- **Event Management**: Add events with photos and attendance sheets
- **Report Generation**: Generate Word documents (.docx) with college letterhead template
- **Report History**: View and download previously generated reports
- **Clean UI**: Professional, responsive design suitable for academic environments

## Technology Stack

- **Backend**: Python Flask
- **Frontend**: HTML, CSS (no JavaScript frameworks)
- **Database**: SQLite
- **Report Generation**: python-docx
- **Authentication**: Flask sessions with Werkzeug password hashing

## Installation

### Prerequisites

- Python 3.7 or higher
- pip (Python package manager)

### Steps

1. **Clone or download this project**

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Create the college letterhead template**:
   - Create a Word document (.docx) with your college letterhead
   - Save it as `college_letterhead.docx` in the `word_templates/` folder
   - The letterhead should be at the top of the document
   - The report content will be added below the letterhead

4. **Run the application**:
   ```bash
   python app.py
   ```

5. **Access the application**:
   - Open your browser and go to: `http://localhost:5000`
   - Register a new account or login

## Project Structure

```
college_event_report/
├── app.py                      # Main Flask application
├── database.db                 # SQLite database (created automatically)
├── requirements.txt            # Python dependencies
├── README.md                   # This file
├── templates/                  # HTML templates
│   ├── base.html              # Base template with navbar
│   ├── login.html             # Login page
│   ├── register.html          # Registration page
│   ├── add_event.html         # Add event form
│   ├── events.html            # Events listing page
│   └── reports.html           # Reports listing page
├── static/                     # Static files
│   ├── css/
│   │   └── style.css          # Main stylesheet
│   ├── event_photos/          # Uploaded event photos
│   └── attendance_photos/     # Uploaded attendance sheets
├── word_templates/             # Word templates
│   └── college_letterhead.docx # College letterhead template (create this)
└── generated_reports/          # Generated report files
```

## Usage Guide

### For Users

1. **Registration/Login**:
   - Click "Register" to create a new account
   - Fill in username, email, and password
   - Click "Login" to access your account

2. **Viewing Events**:
   - Click "Events" in the navbar to see all events
   - Events are displayed in card format with photos

3. **Generating Reports**:
   - While logged in, click "Generate Report" on any event
   - The system will create a Word document with:
     - College letterhead (from template)
     - Event details
     - Event photo (if uploaded)
     - Attendance sheet (if uploaded)

4. **Viewing/Downloading Reports**:
   - Click "Reports" in the navbar
   - View all generated reports
   - Click "Download" to save any report

### For Admins

1. **Adding Events**:
   - Click "Add Event" in the navbar
   - Fill in event details:
     - Event Title (required)
     - Date (required)
     - Venue (required)
     - Department (required)
     - Description (optional)
     - Event Photo (optional)
     - Attendance Sheet Photo (optional)
   - Click "Add Event" to save

## Database Schema

### Users Table
- `id`: Primary key
- `username`: Unique username
- `email`: Unique email
- `password`: Hashed password
- `created_at`: Timestamp

### Events Table
- `id`: Primary key
- `title`: Event title
- `date`: Event date
- `venue`: Event venue
- `department`: Department name
- `description`: Event description
- `event_photo`: Path to event photo
- `attendance_photo`: Path to attendance sheet
- `created_at`: Timestamp
- `created_by`: Foreign key to users table

### Reports Table
- `id`: Primary key
- `event_id`: Foreign key to events table
- `file_path`: Path to generated report file
- `created_at`: Timestamp
- `created_by`: Foreign key to users table

## Important Notes

1. **Letterhead Template**: You must create the `college_letterhead.docx` file in the `word_templates/` folder before generating reports. The template should contain your college letterhead at the top.

2. **File Uploads**: 
   - Supported image formats: JPG, PNG, GIF
   - Maximum file size: 16MB
   - Files are stored in `static/event_photos/` and `static/attendance_photos/`

3. **Security**: 
   - Change the `secret_key` in `app.py` before deploying to production
   - Passwords are hashed using Werkzeug's security utilities
   - File uploads are validated for type and size

4. **Session Management**: 
   - Users remain logged in until they click "Logout"
   - Session data is stored server-side

## Troubleshooting

### Issue: "Letterhead template not found"
**Solution**: Create a Word document named `college_letterhead.docx` in the `word_templates/` folder.

### Issue: Cannot upload images
**Solution**: 
- Check file format (JPG, PNG, GIF only)
- Ensure file size is under 16MB
- Verify folder permissions for `static/event_photos/` and `static/attendance_photos/`

### Issue: Database errors
**Solution**: Delete `database.db` and restart the application. The database will be recreated automatically.

## Development

To modify the application:

1. **Changing Styles**: Edit `static/css/style.css`
2. **Modifying Templates**: Edit files in `templates/`
3. **Adding Features**: Modify `app.py` and add new routes
4. **Database Changes**: Update the `init_db()` function in `app.py`

## Viva/Project Explanation

### Key Concepts Explained

1. **Flask Framework**: Lightweight Python web framework for building web applications
2. **SQLite Database**: File-based database perfect for small to medium applications
3. **Session Management**: Flask sessions store user login state securely
4. **Password Hashing**: Passwords are hashed using Werkzeug (never stored in plain text)
5. **File Upload**: Secure file handling with validation and safe storage
6. **Word Generation**: python-docx library creates Word documents programmatically
7. **Template Inheritance**: Jinja2 templates use inheritance (base.html) for consistent layout

### Security Features

- Password hashing (Werkzeug security)
- Session-based authentication
- File upload validation
- SQL injection prevention (parameterized queries)
- XSS protection (Jinja2 auto-escaping)

### Design Patterns Used

- MVC (Model-View-Controller) pattern
- Template inheritance
- Decorator pattern (login_required)
- Separation of concerns (routes, database, templates)

## License

This project is for educational purposes.

## Support

For issues or questions, refer to the Flask documentation:
- Flask: https://flask.palletsprojects.com/
- python-docx: https://python-docx.readthedocs.io/
