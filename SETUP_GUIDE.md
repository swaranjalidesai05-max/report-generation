# Quick Setup Guide

## Step-by-Step Installation

### 1. Install Python Dependencies

Open terminal/command prompt in the project directory and run:

```bash
pip install -r requirements.txt
```

This will install:
- Flask (web framework)
- Werkzeug (security utilities)
- python-docx (Word document generation)

### 2. Create College Letterhead Template

**Important**: You must create this file before generating reports!

1. Open Microsoft Word
2. Create a new document
3. Add your college letterhead (logo, name, address, etc.) at the top
4. Save the file as `college_letterhead.docx` in the `word_templates/` folder

**Path**: `word_templates/college_letterhead.docx`

### 3. Run the Application

```bash
python app.py
```

You should see:
```
 * Running on http://127.0.0.1:5000
```

### 4. Access the Application

1. Open your web browser
2. Go to: `http://localhost:5000`
3. You will be redirected to the login page

### 5. Create Your First Account

1. Click "Register here" or go to the Register page
2. Fill in:
   - Username
   - Email
   - Password
3. Click "Register"
4. Login with your credentials

### 6. Add Your First Event

1. After logging in, click "Add Event" in the navbar
2. Fill in the event details:
   - Event Title (required)
   - Date (required)
   - Venue (required)
   - Department (required)
   - Description (optional)
   - Upload event photo (optional)
   - Upload attendance sheet (optional)
3. Click "Add Event"

### 7. Generate a Report

1. Go to "Events" page
2. Find your event
3. Click "Generate Report"
4. The report will be created with:
   - College letterhead
   - Event details
   - Event photo (if uploaded)
   - Attendance sheet (if uploaded)

### 8. View and Download Reports

1. Click "Reports" in the navbar
2. View all generated reports
3. Click "Download" to save any report

## Troubleshooting

### Error: "ModuleNotFoundError: No module named 'flask'"

**Solution**: Install dependencies:
```bash
pip install -r requirements.txt
```

### Error: "Letterhead template not found"

**Solution**: Create `college_letterhead.docx` in the `word_templates/` folder

### Error: "Port 5000 already in use"

**Solution**: 
1. Close other applications using port 5000, OR
2. Change the port in `app.py`:
   ```python
   app.run(debug=True, port=5001)
   ```

### Images not displaying

**Solution**: 
- Check file format (JPG, PNG, GIF only)
- Ensure file size is under 16MB
- Check that files were uploaded successfully

### Database errors

**Solution**: Delete `database.db` and restart the application

## Project Structure

```
college_event_report/
├── app.py                      # Main application file
├── database.db                 # SQLite database (auto-created)
├── requirements.txt            # Python dependencies
├── README.md                   # Full documentation
├── SETUP_GUIDE.md             # This file
├── templates/                  # HTML templates
│   ├── base.html
│   ├── login.html
│   ├── register.html
│   ├── add_event.html
│   ├── events.html
│   └── reports.html
├── static/
│   ├── css/
│   │   └── style.css
│   ├── event_photos/          # Uploaded event photos
│   └── attendance_photos/     # Uploaded attendance sheets
├── word_templates/
│   └── college_letterhead.docx # CREATE THIS FILE!
└── generated_reports/          # Generated Word reports
```

## Testing the Application

1. **Test Registration**: Create a new account
2. **Test Login**: Logout and login again
3. **Test Event Creation**: Add an event with and without photos
4. **Test Report Generation**: Generate a report for an event
5. **Test Report Download**: Download a generated report
6. **Test View Events**: View events as logged-out user (should not see "Generate Report" button)

## For Viva/Project Defense

### Key Points to Explain:

1. **Technology Stack**: Flask (Python), SQLite, HTML/CSS, python-docx
2. **Authentication**: Session-based with password hashing
3. **File Upload**: Secure file handling with validation
4. **Report Generation**: Word document generation using templates
5. **Database Design**: Three tables (users, events, reports) with relationships
6. **Security**: Password hashing, file validation, SQL injection prevention

### Demo Flow:

1. Show registration and login
2. Add an event with photos
3. Generate a report
4. Download the report
5. Show the database structure
6. Explain the code structure

## Next Steps

- Customize the college letterhead template
- Add more events
- Generate reports for events
- Explore the code in `app.py` to understand how it works

## Support

Refer to `README.md` for detailed documentation and troubleshooting.
