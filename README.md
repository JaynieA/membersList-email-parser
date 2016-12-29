
Description
-------------
Email parser for Member List Emails. Parses emails received and inserts them into existing Microsoft Access Database.

Other Requirements to Run
--------------
* File named connection.py, containing:

```
EMAIL_ACCOUNT = "your email at gmail dot com"
EMAIL_FOLDER = "folder name"
PASSWORD = 'your password here'

connStr = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=yourDatabaseNameHere.accdb;'
```

* Existing Database in Microsoft Access
