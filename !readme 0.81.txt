==============================================================
                   R    E    A    D    M    E
==============================================================
2eNetWorX TableEditoR v0.81 Beta
http://www.2enetworx.com/dev/projects/tableeditor.asp
--------------------------------------------------------------

What's New?
- Exporting is a permission assigned to each user in the users table.
- Security checks for adding records/fields/table without having permission.
- High Security Login (*Check below on how it works).
- Remember Me option.
- Level-based User/Database access (*Check below how to use it).
- Compatible with most browsers:
  + Automatic check for javascript, if not found, TE will work without javascript.
  + Fixed Javascripts for more compatabilaty.
  + Export/Delete selected records work with brwosers with no javascript.

- Stored Procedures work better now
- Excel export now runs 2000% faster
- XML Export with schema (XML files 99% compatible with Access XP)
- Database Backup now make a copy of the original database.
- Server Database browser became easier to use.
- Option to enable/disable Popups.
- Option for "Default Per Page" (idea from Peter Stucke).
- Page Titles give and idea of the content of the page.
- OLE Images are showed with better script.
- Records in related table work better (add/edit).
- Show Schema table is like all te tables.
- Show Table contents (IE mode and normal mode) work better with queries.

-----------------------------------------------------------------------------

Q: How does "High Security Login" work?
A: To login, you need to get a fresh copy of the homepage (which means loading it while on the internet), this will let te add a unique string to the login form, then login normaly.
If you try to login with a cached copy, the string will be different than the expected one, so login will fail.

Also when this option is on, Clicking "Home" of the TableEditor will be like Logging out.


Q: What is "Level-based User/Database" access?
A: User with value "0", means users can view all of the connections (except administrator one).
Admin sets the value for each user, the higher the value, the more restrictive (i.e., a user-session value of "2" allows the user to view connections with a value greater than or equal to 2).

It works opposite for the DB_privileges value assigned to each connection in the "Databases" table in teadmin.mdb. The lower the value, the more restrictive the connection (i.e., if arrConn(1)'s DB_privileges value is set to "0", then users with a session("rConnectionViews") of "1" or greater cannot view it.
