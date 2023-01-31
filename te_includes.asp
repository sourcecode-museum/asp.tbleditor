<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_includes.asp
	' Description: Constants and Public Functions
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
	' # Nov 15, 2000 by Hakan Eskici
	' Added permission assignment for Field functions
	' Added constants for fields
	' # Nov 29, 2000 by Hakan Eskici
	' Changed the file name check for protection which assumed that
	' the file name is lower case
	' # May 14, 2002 by Rami Kattan
	' Extra OpenRS and CloseRS for other functions, sometimes caused
	' problems closing and opening the same connection
	'==============================================================

	const bProtected = true

	'If protection is on, make sure that user has logged in
	if instr(lcase(request.servervariables("script_name")), "index.asp") = 0 then
		if bProtected then
			if session("teUserName") = "" then
				response.redirect "index.asp?comebackto=" & request.servervariables("script_name") & "?" & server.urlencode(request.querystring)
			end if
		end if
	end if

	if bProtected then
		'If protection is on, get permissions
		'for the user from the session
		bAdmin = session("rAdmin")
		bRecAdd = session("rRecAdd")
		bRecEdit = session("rRecEdit")
		bRecDel = session("rRecDel")
		bQueryExec = session("rQueryExec")
		bSQLExec = session("rSQLExec")
		bTableAdd = session("rTableAdd")
		bTableEdit = session("rTableEdit")
		bTableDel = session("rTableDel")
		bFldAdd = session("rFldAdd")
		bFldEdit = session("rFldEdit")
		bFldDel = session("rFldDel")
		bAllowExport = session("rAllowExport")
	else
		'Not protected, give Full control
		bAdmin = True
		bRecAdd = True
		bRecEdit = True
		bRecDel = True
		bQueryExec = True
		bSQLExec = True
		bTableAdd = True
		bTableEdit = True
		bTableDel = True
		bFldAdd = True
		bFldEdit = True
		bFldDel = True
		bAllowExport = true
	end if

	'Pre-create connection and recordset objects
	set conn = Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")
	
	'Opens a given connection and initializes rs (For main TE)
	sub OpenRS(sConn)
		conn.open sConn
		set rs.ActiveConnection = conn
		rs.CursorType = adOpenStatic
	end sub
	
	'Closes open connections and releases objects (For main TE)
	sub CloseRS()
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	end sub

	set conn2 = Server.CreateObject("ADODB.Connection")
	set rs2 = Server.CreateObject("ADODB.Recordset")

	'Opens a given connection and initializes rs (For extra TE functions)
	sub OpenRS2(sConn)
		if (conn2 is nothing) then
			set conn2 = Server.CreateObject("ADODB.Connection")
			set rs2 = Server.CreateObject("ADODB.Recordset")
		end if
		conn2.open sConn
		set rs2.ActiveConnection = conn2
		rs2.CursorType = adOpenStatic
	end sub
	
	'Closes open connections and releases objects (For extra TE functions)
	sub CloseRS2()
		rs2.close
		conn2.close
		set rs2 = nothing
		set conn2 = nothing
	end sub

	'---- CursorTypeEnum Values ----
'	Const adOpenForwardOnly = 0
'	Const adOpenKeyset = 1
'	Const adOpenDynamic = 2
'	Const adOpenStatic = 3

	'---- CursorLocationEnum Values ----
'	Const adUseServer = 2
'	Const adUseClient = 3

	'---- CommandTypeEnum Values ----
'	Const adCmdUnknown = &H0008
'	Const adCmdText = &H0001
'	Const adCmdTable = &H0002
'	Const adCmdStoredProc = &H0004
'	Const adCmdFile = &H0100
'	Const adCmdTableDirect = &H0200
	
	'---- SchemaEnum Values ----
'	Const adSchemaTables = 20
'	Const adSchemaPrimaryKeys = 28
'	Const adSchemaIndexes = 12
'const adSchemaViews = 23
'	Const adSchemaForeignKeys = 27
'	Const adSchemaProcedures = 16
	
	'---- DataTypeEnum Values ----
'	Const adEmpty = 0
'	Const adTinyInt = 16
'	Const adSmallInt = 2
'	Const adInteger = 3
'	Const adBigInt = 20
'	Const adUnsignedTinyInt = 17
'	Const adUnsignedSmallInt = 18
'	Const adUnsignedInt = 19
'	Const adUnsignedBigInt = 21
'	Const adSingle = 4
'	Const adDouble = 5
'	Const adCurrency = 6
'	Const adDecimal = 14
'	Const adNumeric = 131
'	Const adBoolean = 11
'	Const adError = 10
'	Const adUserDefined = 132
'	Const adVariant = 12
'	Const adIDispatch = 9
'	Const adIUnknown = 13
'	Const adGUID = 72
'	Const adDate = 7
'	Const adDBDate = 133
'	Const adDBTime = 134
'	Const adDBTimeStamp = 135
'	Const adBSTR = 8
'	Const adChar = 129
'	Const adVarChar = 200
'	Const adLongVarChar = 201
'	Const adWChar = 130
'	Const adVarWChar = 202
'	Const adLongVarWChar = 203
'	Const adBinary = 128
'	Const adVarBinary = 204
'	Const adLongVarBinary = 205
'	Const adChapter = 136
'	Const adFileTime = 64
'	Const adPropVariant = 138
'	Const adVarNumeric = 139
'	Const adArray = &H2000	
	
	'---- FieldAttributeEnum Values ----
'	Const adFldMayDefer = &H00000002
'	Const adFldUpdatable = &H00000004
'	Const adFldUnknownUpdatable = &H00000008
'	Const adFldFixed = &H00000010
'	Const adFldIsNullable = &H00000020
'	Const adFldMayBeNull = &H00000040
'	Const adFldLong = &H00000080
'	Const adFldRowID = &H00000100
'	Const adFldRowVersion = &H00000200
'	Const adFldCacheDeferred = &H00001000
'	Const adFldIsChapter = &H00002000
'	Const adFldNegativeScale = &H00004000
'	Const adFldKeyColumn = &H00008000
'	Const adFldIsRowURL = &H00010000
'	Const adFldIsDefaultStream = &H00020000
'	Const adFldIsCollection = &H00040000	
	
'	Const adColFixed = 1
'	Const adColNullable = 2	
%>
<!--#include file="te_functions.asp"-->