# Excel Application: Timekeeping & Payroll

A Python + Tkinter application designed for small businesses (such as coffee shops) to automate attendance and payroll management.  
It processes attendance data from Excel files, calculates working hours, applies penalties, supports flexible shifts (including split shifts), and exports formatted payroll reports.

---

## 🚀 Features

- Import and process attendance data from Excel sheets  
- Calculate working hours with high precision (using Decimal)  
- Apply penalties for late check-ins automatically  
- Support flexible/split shifts (morning and afternoon)  
- Customizable shift rules for specific employees  
- Generate well-formatted Excel payroll reports ready for business use  
- User-friendly Tkinter interface with multiple tabs (for different locations)  
- Configuration saved in JSON files (per business branch)  

---

## 🛠️ Installation

Clone the repository:

```bash
git clone https://github.com/ngkhoa2708joy-github/excel-application-timekeeping-payroll.git
cd excel-application-timekeeping-payroll
```
# Install dependencies (recommend using virtual environment):
```yami
python -m venv .venv
.venv\Scripts\activate   # Windows
pip install -r requirements.txt
```
## 🔑 Google API Setup (for Excel / Google Sheets integration)
This app uses the Google Sheets API for reading/writing online attendance sheets.

# Step 1. Create a Google Cloud project
Visit Google Cloud Console.

Create a new project.

Enable the Google Sheets API and Google Drive API.

# Step 2. Create OAuth Credentials
Go to APIs & Services > Credentials.

Click Create Credentials > OAuth client ID.

Choose Desktop App.

Download the file → rename it to credentials.json.

Place credentials.json in the root of your project.

# Step 3. First-time authentication
When you run the program for the first time:
A browser window will open asking you to log in with your Google account.

After granting access, a file named token.json will be created automatically.

The app will use token.json for future runs.

Note
Do NOT commit credentials.json or token.json to GitHub.

Add them to your .gitignore file.
```
# secrets
credentials.json
token.json
.env
*.xlsx
*.json
```
