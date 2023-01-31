<!--#include file="te_includes.asp"-->
<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_config.asp
	' Description: Configuration File for TableEditoR
	' Initiated By Hakan Eskici on Nov 01, 2000
	'--------------------------------------------------------------
	' Copyright (c) 2002, 2eNetWorX/dev.
	'
	' TableEditoR is distributed with General Public License.
	' Any derivatives of this software must remain OpenSource and
	' must be distributed at no charge.
	' (See license.txt for additional information)
	'
	' See Credits.txt for the list of contributors.
	'
	' Change Log:
	'--------------------------------------------------------------
	' # Nov 16, 2000 by Kevin Yochum
	' Added switches for converting null values
	' # Mar 31, 2001 by Hakan Eskici
	' Changed defining connections
	' Added support for SQL Server, DSN and DirectConnections
	' # April 22, 2002 by Rami Kattan
	' Configurations & Database defenition in the teadmin.mdb database
	' this file loads all database definitions and configurations
	' # May 11, 2002 by Hakan Eskici
	' Modified recordcount calculation
	' # May 14, 2002 by Rami Kattan
	' Browser check if can execute javascipts
	' # May 30, 2002 by Rami Kattan
	' Added option for default Per Page value.
	' Option for high security login
	' Response.buffer enabled in some pages for better performance
	' User permissions for databases using user/database leveling
	' # Jun 3, 2002 by Rami Kattan
	' Two new config options in database loaded here
	'==============================================================

	sScript = lcase(Request.ServerVariables("SCRIPT_NAME"))
	if instr(sScript, "te_xml") or instr(sScript, "te_readdb") or instr(sScript, "te_excel") then
		response.buffer = true
	else
		response.buffer = false
	end if
		

	Server.ScriptTimeout = 120

	'--[»]--- Define Main Table Editor Connection ---------------------
	' Table Editor Administration db.
	' >>>>>>>>>>>>>>>>>>>>>> CHECK THIS NOTE <<<<<<<<<<<<<<<<<<<<<<<<<<<
	' NOTE: teadmin.mdb Should be placed in a folder with WRITE permission.
	' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	te_arrDBs = "teadmin.mdb;"
	TempPath = "/public"

	' -->> That's all, go to the application via web, login, and then
	'      make the other configuration.
	' -->> Default login is: admin/admin
	'------------------------------------------------------------------

	'Different Connection Types
	const tedbAccess = 1
	const tedbSQLServer = 2
	const tedbDsn = 3
	const tedbConnStr = 4

	aParams = split(te_arrDBs, ";")
	sDBName = Server.MapPath(trim(aParams(0)))
	if ubound(aParams) > 0 then
		sPassword = trim(aParams(1))
	end if
	te_arrConn = "Provider=Microsoft.Jet.OLEDB.4.0;" &_
					   "Persist Security Info=False;" &_
					   "Data Source=" & sDBName & ";" & _
					   "Jet OLEDB:Database Password=" & sPassword & ";"

	OpenRS2 te_arrConn

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Under Experiments, not enabled till v0.8 is finished and 0.8.1 begins
	' Permissions section: I added a new session variable, session("rConnectionViews").
	' Default value is "0", which means users can view all of the connections. 
	' Admin sets the value for each user, the higher the value, the more 
	' restrictive (I.e., a user-session value of "2" only allows the user to view connections 
	' with a value greater than or equal to 2).
	'
	' It works opposite for the DB_view value assigned to each connection in the "Databases" table
	' in teadmin.mdb. The lower the value, the more restrictive the connection 
	' (I.e., if arrConn(1)'s DB_view value is set to "0", then users
	' with a session("rTableViews") of "1" or greater cannot view it.
	
	if session("rConnectionViews")  = "" then
		rConnectionViews  = "0"
	else
		rConnectionViews  = session("rConnectionViews")
	end if
	' One more addition below.
	PrivWhere = " WHERE DB_privileges >=" & rConnectionViews
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ By Pete Stucke ~~~~~~~~~~~~~~~~~~

	dim iTotalConnections

	' Added by Hakan on May 11, 2002
	' Get the count via sql in case the recordset cannot read recordcount
	strSql = "SELECT COUNT(ID) AS Total FROM Databases" & PrivWhere
	rs2.open strSql,,,adCmdTable
	iTotalConnections = CDbl(rs2("Total"))
	rs2.close

	strSql = "SELECT * FROM Databases" & PrivWhere & " ORDER BY DB_Desc"
	rs2.open strSql,,,adCmdTable


	'Using Redim causes a performance degredation
	'But it's OK since array size is small
	redim arrDbs(iTotalConnections)
	redim arrDesc(iTotalConnections)
	redim arrType(iTotalConnections)
	redim arrConn(iTotalConnections)

	arrType(0) = tedbAccess
	arrDBs(0) = te_arrDBs 'reget TableEditoR user administration db
	arrDesc(0) = "Table Editor Administration"
	aParams = split(arrDBs(0), ";")
	sDBName = Server.MapPath(trim(aParams(0)))
	if ubound(aParams) > 0 then
		sPassword = trim(aParams(1))
	end if
	arrConn(0) = "Provider=Microsoft.Jet.OLEDB.4.0;" &_
					   "Persist Security Info=False;" &_
					   "Data Source=" & sDBName & ";" & _
					   "Jet OLEDB:Database Password=" & sPassword & ";"

	'Construct connection strings
	for iConnection = 1 to iTotalConnections

		arrType(iConnection) = rs2.Fields("DB_type")
		arrDBs (iConnection) = rs2.Fields("DB_loc")
		arrDesc(iConnection) = rs2.Fields("DB_Desc")

		sDBName = ""
		sComputerName = ""
		sUserName = ""
		sPassword = ""
		select case arrType(iConnection)
			case tedbAccess
				'Access
				aParams = split(arrDBs(iConnection), ";")
				sDBName = Server.MapPath(trim(aParams(0)))
				if ubound(aParams) > 0 then
					sPassword = trim(aParams(1))
				end if
				arrConn(iConnection) = "Provider=Microsoft.Jet.OLEDB.4.0;" &_
					   		       "Persist Security Info=False;" &_
							       "Data Source=" & sDBName & ";" & _
								   "Jet OLEDB:Database Password=" & sPassword & ";"
			case tedbSQLServer
				'SQL Server
				aParams = split(arrDBs(iConnection), ";")
				if isArray(aParams) then
					sDBName = trim(aParams(0))
					sComputerName = trim(aParams(1))
					sUserName = trim(aParams(2))
					sPassword = trim(aParams(3))
				end if
				arrConn(iConnection)  = "Provider=SqlOLEDB;Network Library=DBMSSOCN;" & _ 
									"Data Source=" & sComputerNAme & ";" &_
									"Initial Catalog=" & sDBName & ";" & _
									"User Id=" & sUserName & ";" &_
									"Password=" & sPassword & ";"
			case tedbDsn
				'Data Source Name
				arrConn(iConnection)  = "dsn=" & arrDBs(iConnection)
			case tedbConnStr
				'Direct connection string
				arrConn(iConnection) = arrDBs(iConnection)
		end select
		rs2.movenext
	next
	
	CloseRS2

	OpenRS2 arrConn(0)
	strSql = "SELECT * FROM config WHERE id = " & 1
	rs2.open strSql,,,adCmdTable

		'Encode HTML tags?
		'Turn this on if you have problems with displaying
		'records with html content.
		dim bEncodeHTML
		bEncodeHTML = rs2.Fields("EncodeHTML")

		'Maximum number of chars to display in table view (0 : no limit)
		'Warning: If you have HTML content in your fields;
		'you should set bEncodeHTML to True if you specify a limit
		dim lMaxShowLen
		lMaxShowLen = rs2.Fields("MaxShowLen")
		
		'Show Related Table Contents as ComboBoxes for foreign key fields?
		dim Relation
		Relation = rs2.Fields("Relation")
		
		'Show connection details? (Number of tables, views and procs)
		'This requires all connections to be opened, so te_admin.asp will run slow
		dim bShowConnDetails
		bShowConnDetails = rs2.Fields("ShowConnDetails")
		
		'Should blank fields be converted to NULL when the field is nullable?
		'Convert '' to null in non-numeric and non-date fields?
		dim bConvertNull
		bConvertNull = rs2.Fields("ConvertNull")

		'Convert '' and 0 to null in numeric fields?
		dim bConvertNumericNull
		bConvertNumericNull = rs2.Fields("ConvertNumericNull")

		'Convert '' and 0 to null in date fields?
		dim bConvertDateNull
		bConvertDateNull = rs2.Fields("ConvertDateNull")

		'enable/disable the table highlight?
		dim bTableHighlight
		bTableHighlight = rs2.Fields("HighLight")

		'enable/disable the MultiDelete buttons?
		dim bBulkDelete
		bBulkDelete = rs2.Fields("BulkDelete")

		'enable/disable the export to Excel button?
		dim bExportExcel
		bExportExcel = rs2.Fields("ExportExcel")

		'enable/disable the export to XML button?
		dim bExportXML
		bExportXML = rs2.Fields("ExportXML")
		XMLExportSchema = true

		'enable/disable the Bulk compact (admin only)?
		dim bBulkCompact
		bBulkCompact = rs2.Fields("BulkCompact")

		'enable/disable Active Users Logging?
		dim bActiveUsers
		bActiveUsers = rs2.Fields("CountActiveUsers")

		'enable/disable the Dynamic Page Selectors?
		dim bPageSelector
		bPageSelector = rs2.Fields("PageSelector")
		iPageSelectorMax = 25

		' values: 5, 10, 15, 20, 30, 40 or 0 for all
		dim iDefaultPerPage
		iDefaultPerPage = rs2.Fields("DefaultPerPage")

		'enable/disable view databases-tables as combo?
		dim bComboTables
		bComboTables = rs2.Fields("ComboTables")

		dim bPopUps
		bPopUps = rs2.Fields("PopUps")

		'enable/disable the new interface for IE4+ (only work with IE)?
		dim bIEAdvancedMode
		bIEAdvancedMode = rs2.Fields("IEAdvancedMode")

		'enable/disable the Higher security login?
		dim RequireSecurityID
		RequireSecurityID = rs2.Fields("HighSecurityLogin")

	CloseRS2

	bJSEnable = session("JavaScriptEnabled")

	' ----------- EXPERIMENTAL FEATURES ----------------
	' No support still for them, not recommended to enable
		bRequiredField = false	' enable/disable required field notice [*] (beta testing, wrong results sometimes)?
		ConvertURL = false		' convert unclickable urls into clickable urls
		bUserLogging = false	' Log logins.
		bSQLBuilder = false		' SQL builder for queries
		bAllowImport = false	' Allow XML import (Under Development)
		te_debug = false		'leave false, true only for debugging
	' --------------------------------------------------

	if bPopUps and not bJSEnable then bPopUps = false

	' BrowserCompat
	sU = ucase(Request.ServerVariables("HTTP_USER_AGENT"))
	isOp = instr(sU, "OPERA") > 0
	isIE = instr(sU, "MSIE") > 0 AND not isOp
	isNS = instr(sU, "NETSCAPE") > 0 AND not isOp
	isKo = instr(sU, "KONQ") > 0

	if isIE AND bIEAdvancedMode AND bJSEnable then
		TableViewerCompat = "te_showtable.asp"
	else
		TableViewerCompat = "te_showtable2.asp"
	end if
%>