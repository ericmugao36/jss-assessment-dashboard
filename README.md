# Assessment Dashboard

A web-based school assessment management system built with Python and FastAPI. The project is designed to help teachers manage learner assessment information and allow parents to view their children's academic progress through separate role-based portals.

## Overview

The Assessment Dashboard provides a central place to manage learner records, assessment results, and teacher feedback. It separates teacher and parent access while using a shared application and database.

This is a project under active development.

## Features

- **Separate portals:** Dedicated teacher and parent login pages.
- **Authentication:** Password-based login and session management.
- **Role-based access:** Restricts portal access according to the user's role.
- **Learner records:** View learner profiles and associated information.
- **Assessment management:** Record and view assessment marks.
- **Performance tracking:** Display learner assessment performance in a graph.
- **Teacher remarks:** Allow parents to view recorded teacher feedback.
- **Data import:** Support for importing school assessment data, where configured.
- **Database persistence:** Store application records using SQLAlchemy and SQLite.

## Technology Stack

- Python
- FastAPI
- SQLAlchemy
- SQLite
- HTML and CSS
- Uvicorn
- Git and GitHub

## Project Structure

```text
jss-dash/
├── app.py
├── requirements.txt
├── .gitignore
├── README.md
└── Other application files
```

The exact project structure may vary as development continues.

## Getting Started

### 1. Prerequisites

Install Python and Git. A code editor such as Visual Studio Code is recommended.

### 2. Clone the repository

If you have access to the private GitHub repository:

```bash
git clone https://github.com/ericmugao36/jss-assessment-dashboard.git
cd jss-assessment-dashboard
```

### 3. Create a virtual environment

On Windows:

```powershell
python -m venv .venv
```

Activate it in PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

If PowerShell blocks activation, you can use the environment's Python directly:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### 4. Install dependencies

With the virtual environment active:

```powershell
python -m pip install -r requirements.txt
```

### 5. Run the application

From the project directory:

```powershell
python -m uvicorn app:app --reload
```

Open the local application in your browser:

- Application: http://127.0.0.1:8000/
- Teacher portal: http://127.0.0.1:8000/teacher/login
- Parent portal: http://127.0.0.1:8000/parent/login
- API documentation: http://127.0.0.1:8000/docs

Use existing test accounts configured for your local application. A fresh database may require an initial setup process before login is possible.

## Security and Data Privacy

- Keep passwords, API keys, session secrets, and other credentials out of source control.
- Do not commit real learner records, parent details, assessment spreadsheets, or database files.
- Use synthetic or anonymized data when testing and demonstrating the application.
- Keep dependencies updated and configure secure settings before any production deployment.
- Use HTTPS and appropriate access controls when deploying outside your local development environment.

## Development Status

This project is being developed incrementally, with a focus on role-based access, learner assessment tracking, parent visibility, and reliable school data management.

## Author

**Eric Mugao**  
Python Developer | Automation & AI Tools

GitHub: [@ericmugao36](https://github.com/ericmugao36)

## License

No open-source license has been specified yet. Unless a license is added, all rights are reserved by the copyright holder.
