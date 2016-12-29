#Description

Parses emails received hourly and enters information from them into a database.

##Requirements

* File named connection.py in root folder, containing:

```
EMAIL_ACCOUNT = "your email at gmail dot com"
EMAIL_FOLDER = "folder name"
PASSWORD = 'your password here'

connStr = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=yourDatabaseNameHere.accdb;'
```

* Existing Database in Microsoft Access

##Technologies

* Python3
* pyodbc
* Regex
* SQL
