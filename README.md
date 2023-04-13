# MailGuard
A script for hmailserver 5 that controls the outbound addresses that mail can be sent to for each account

MailGuard consists of a script, some extra tables that are created in hmaildb and an optional front end GUI to manage the ACL written in .Net


Installation steps.
1. Install the ODBC driver for MariaDB/MySQL and create a system DSN using ODBC sources. I usually create a separate user for accessing the hmail database, so configure this DSN with the same credentials. Name the DSN "MailGuard"
2. Create a folder called "mailguard" to store any log files from the script. The location can be changed in EventHandlers.vbs
3. Execute the SQL file to create the stored procedure "MGCheck" and the 3 supporting tables inside the hmail db
4. Copy  EventHandlers.vbs into hmailserver\events. If you have other custom scripts defined, please merge the scripts carefully.
5. The frontend to manage the entries in the table is also provided. After opening it, specify the server, username and password to access the hmaildb and connect. You can then select the email address and specify which addresses / domains it is allowed to send mails to. You can also specify which IP addresses the email address can connect from (to deter spoofing)
