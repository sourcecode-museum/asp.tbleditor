==============================================================
                   R    E    A    D    M    E
==============================================================
2eNetWorX TableEditoR v0.81 Beta
http://www.2enetworx.com/dev/projects/tableeditor.asp
--------------------------------------------------------------

1. How to Install?
Please see install.txt.

2. What's New in Summary?

* TableEditoR v0.81 Beta is made possible by the team:

Rami Kattan - Development and implementation
Jeff Wilkinson - Testing, feedback and ideas


» Most contributions are compiled into this release
» Client side TDC dataview for Internet Explorer
» Dynamically configure your settings over the web
» Dynamically define your connections over the web
» Dynamic paging support
» Export to XML
» Export to Excel
» Better support for parameterized queries and stored procedures
» Security fixes
» Fixes for well known and small bugs

3. What's New in Detail?

This update was based on official TableEditor v0.7 Beta release.


What's new in v0.81 including features in v0.8:
-----------------------------------------------
* Table contents are viewed using Microsoft Tabular Data Control (TDC),
  that means the table is built client-side, not server side, server
  only sends the base data in CSV (Comma Seperated Values) format.
  This saves a lot of the transfered data (see note 1 below).
  NOTES: * I don't know how to test Stored Procedures, so this TDC may have
           errors.
         * Search also work with this TDC view.
	 * TDC works in Internet Explorer (4.0 and above), I don't have
	   (and don't like) Netscape, so not guaranteed to work with it.
	 * Sometimes, when adding/deleteing records, and pages become more
	   or less, and navigation buttons doesn't show, click once or twice
	   on "Refresh".

* Added May 2, 2002: In addition to the TDC viewing, normal table viewing
  is also available (automatic check for browser).

- Sorting: sorts the according to any column Ascending or Descending.
         
+ Databases are defined in a database, and can be administrated via web.

+ Defining new Access databases is easy, user can select through a list 
  (Explorer-like) of databases in a folder.

+ Configurations are in the administrator database, not te_config.asp,
  can be changed easily with a web browser.

+ Administrator can check Active Users (in Config, enable: CountActiveUsers).

* Autoincrement fields are automatically incremented, no need to fill them.

* Date fields have an option to insert today's date.

* Required fields (don't accept null) are marked with [*] (Under experiment)

* Interface changes, highlight rows on mouse over/click (in Config, enable: HighLight).

+ Dynamic paging functions (in Config, enable: PageSelector).

* Compact/Backup all access databases (administrators only) (in Config, enable: BulkCompact).

+ Switch between tables and databases using combo boxes (in Config, enable: ComboTables).


Fixes and update:
+ fixed export to excel 
+ better compatibility with more browsers
- fixed the security risk in v0.7 (by Jeff Wilkinson)
+ display OLE images (by schwarzk)
+ view all records in a table (by Pete Stucke)
* dynamic recordset paging (by PStucke) + fixed when records per page become more,
  while we are on last page, it will go automatically to the real last page
  (will not display blank page)
+ export to excel (by Pete Stucke) (in Config, enable: ExportExcel)
+ multidelete check/uncheck all (by Pete Stucke)
+ backup database (by unknown)
- ... and some few other fixes, that are not so visible (at least to my testing).

									Legend:
									    + Added
									    * Improved
									    - fixed

Compatibility:
--------------
Browser:
	- Internet Explorer 4.0 and above (for TDC advanced viewing)

	- Other browsers (Normal viewing):
		+ Netscape 6.0 (most features work)
	        + Konqueror for Linux (most features work)

Databases:
	- Access
	- DSN
	- SQL Server 2000 (basic testing)

-------------------------------------------------------------------
(1) Performance experiment
Tested data transfer for a 800 record database, results were:

New format: Page  = 14566 Byte   DB = 492443 Byte
            total =  507009 Byte

Old format: Page  = 1018065 Byte

Saves nearly 50% of the transfered data
-------------------------------------------------------------------