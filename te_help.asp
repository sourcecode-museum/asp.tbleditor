<!--#include file="te_header.asp"-->
<pre style="font-size:14">
Considerations

Protection

Table Editor features a protection for your databases.
Protection is enabled by default. If you want to disable
protection, open te_config.asp file and change
bProtected = False

However, it is not recommended to disable protection because
any visitor who knows the exact location of the TableEditor 
files may view or change the information in your databases.

When logged in as an administrator, you will be able to
administer your TableEditoR users and permissions.

Permissions for users:
rAdmin		: Administrator
rRecAdd		: Can add records?
rRecEdit	: Can edit records?
rRecDel		: Can delete records?
rQueryExec	: Can execute stored queries?
		  User with QueryExec permission can view and execute stored
		  queries in the database.
rSQLExec	: Can execute queries?
		  Only the admin should have SQLExec permisson, because a user
		  with this permission can execute any type of sql statements
		  which includes delete and update type queries.
rTableAdd	: Can create tables?
rTableEdit	: Can edit table structures?
		  This permission also grants the compacting of databases.
rTableDel	: Can remove tables?
rFldAdd		: Can add new field definitions to tables?
rFldDel		: Can delete field definitions from tables?
rFldEdit	: Can edit field definitions on tables?

NOTE:
rTableAdd, rTableEdit and rTableDel permissions apply on
queries as well.

Editing, Viewing and Deleting records
TableEditor 0.7 and better versions now autodetects the primary keys of the
tables. When it's not possible to detect the primary key,
or when a primary key doesnt exist; the following case
is still applicable as previous versions:
"When finding the record to be edited, viewed or deleted;
table editor assumes that the first field of the table 
to be the key field. For this reason, you should be using 
such a table structure that the first field is either unique 
or better is a primary key."


Creating New Queries

Queries created with TableEditoR may not be visible
in Access. We are looking for a workaround.

Adding records

If your table has auto increment fields, you should leave
blank these fields to be auto incremented. Auto increment
field values are automatically assigned by data provider.

If you specify a number for a auto increment field, table
editor won't return an error, and field value will be set
to what you have entered. However, this is not recommended.
You should leave blank the auto increment fields when adding
new records.


Browsing Queries

When you are browsing queries, depending on the type of the
query, some features may not be accesible.

Known issues:
* If the query already includes an order by statement, clicking
on the field names will no longer perform sorting.
* If the query is not updatable, you cannot add, edit or
delete records.

HTML Content
<em>In case HTML Encoding is disabled in configuration:</em>
If the records include html content, they will be rendered as html.
It's a feature which may be considered as both good or bad depending
on the record contents. For example if a field value is
&lt;img src="images/logo.gif"&gt; then table editor will show the image.

Known issues:
If a field with html content is the first field in the table,
undesirable content might be rendered. Actually, this is not a common
case, but be sure to have a unique field in your tables and queries 
as the first field.

OLE Fields
OLE fields are ignored by table editor now. When you edit records
existing OLE fields in the tables will not be overwritten.

What's new in version v0.8?

* TableEditoR v0.8 is made possible by the team:
Rami Kattan - Development and implementation
Jeff Wilkinson - Testing, feedback and ideas

1. Most contributions are compiled into this release
2. Client side TDC dataview for Internet Explorer
3. Dynamically configure your settings over the web
4. Dynamically define your connections over the web
5. Dynamic paging support
6. Export to XML
7. Export to Excel
8. Better support for parameterized queries and stored procedures
9. Security fixes
10. Fixes for well known and small bugs


What's new in version 0.7?

I will be leaving the Asp community in April 1, 2001 for the obligatory 
army service. I'll be back 8 months later. I wish I had some more time for 
new features. I've tried to include most wanted features before I leave ;)

1. Support for SQL Server and DSN.
2. Foreign Key support for related tables. (Credits: Danival A. Souza)
3. Support for multiple primary keys (constraints).
4. Support for NT4 systems which sometimes causes problems with Unicode characters
5. Bug fixes regarding editing queries.
6. Executing parameterized queries.
7. Delete multiple records.
8. Detailed information about connections (optional)
9. More information about tables, views and procedures.
10. Many small bugfixes and enhancements.


What's new in version 0.6?

1. Editing table structures (adding, changing and deleting fields and indexes).
2. Compacting Access Mdb files on the server.
3. Creating new tables.
4. Creating, editing and deleting queries.
5. More permissions on table structure modification.
6. Configuration switches for HTML content.
7. Configuration switches for action on Null values. (by Kevin Yochum)
8. Various bug fixes and cosmetic enhancements (as always ;).


What's new in version 0.5? (Included in version 0.6 as well)

1. Multi-user permission-level protection.
2. Multiple connections for more than one database.
3. Listing stored queries and browsing them like tables.
4. Searching tables (by Kevin Yochum).
5. Viewing table and query structure information.
6. Running SQL statements and browsing them like tables.
7. Sorting data by clicking on the field names.
8. Various bug fixes and cosmetic enhancements.
</pre>
<!--#include file="te_header.asp"-->