# Description

Parses emails received hourly (or every 3 hours after business hours) and enters information from them into a database.

## Requirements

* Additional file named connection.py located in root folder, containing:

```
EMAIL_ACCOUNT = "your email at gmail dot com"
EMAIL_FOLDER = "folder name"
PASSWORD = 'your password here'

connStr = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=yourDatabaseNameHere.accdb;'
```

* Existing Database in Microsoft Access
* Enable IMAP in Gmail settings
* Allow access for less secure apps in Gmail settings

## Technologies

* Python3
* pyodbc
* Regex
* SQL
