
#Description

Email parser for Uneda Member List Emails. Runs every two minutes (approximately) to parse data from emails received, and insert it into existing Microsoft Access Database.

##Requirements to Run

* File named connection.py in root directory, containing:

```
EMAIL_ACCOUNT = "your email at gmail dot com"
EMAIL_FOLDER = "folder name"
PASSWORD = 'your password here'

connStr = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=yourDatabaseNameHere.accdb;'
```

* Existing Database in Microsoft Access

<p align="center">
  <img src="email-parser-erd.png?raw=true" alt="ERD"/>
</p>

##Technologies:

* Python3
* pyodbc 
* imaplib
* regex
* Microsoft Access
