# CPaaS Connector Plugin

## Project Overview

The CPaaS Connector Plugin is a Python-based web application that simulates messaging workflows through a dashboard. It allows users to manage message records, track delivery statuses, create order notifications, and export message history.

**Note:** This project currently simulates message delivery locally. It is not connected to a real SMS provider.

## Technologies Used

* **Programming Language:** Python
* **Database Language:** SQL
* **Backend Framework:** FastAPI
* **Database:** SQLite
* **Frontend:** HTML, CSS, JavaScript
* **Development Server:** Uvicorn
* **Libraries:** phonenumbers, openpyxl


## Features

* International phone number validation
* Unique message ID generation
* Message history and status management
* Search and filter messages
* Order notification creation
* Order history management
* SQLite database integration
* Export message records to CSV and Excel

## Project Structure

* `app.py` — Main web application and dashboard
* `connector.py` — Message processing and phone number validation
* `database.py` — Database operations and record management
* `README.md` — Project documentation

## How to Run

1. Install Python.

2. Open the project folder in a terminal.

3. Create a virtual environment:

   `python -m venv .venv`

4. Activate it on Windows:

   `.\.venv\Scripts\Activate.ps1`

5. Install the required dependencies from `requirements.txt`, if provided:

   `python -m pip install -r requirements.txt`

6. Start the application:

   `python -m uvicorn app:app --reload`

7. Open `http://127.0.0.1:8000` in your browser.

## Learning Outcomes

This project demonstrates practical experience with Python programming, SQL database operations, backend development, input validation, application workflows, and data export functionality.

## Future Scope

The application could be integrated with a real SMS provider through an API in the future.

Live SMS delivery is not currently implemented.
