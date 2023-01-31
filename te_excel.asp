<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_excel.asp
	' Description: Displays all records from selected table in a spreadsheet
	' Initiated By Pete Stucke on Apr 11, 2002
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
	' # Mar 26, 2001 by Hakan Eskici
	' Added support for automatic primary key detection
	' Added support for multiple primary keys
	' # Mar 28, 2001 by Hakan Eskici
	' Modified the recordset paging control
	' # Mar 29, 2001 by Hakan Eskici
	' Added support for SQL Server boolean values
	' Modified request's to .form or .querystring
	' Added support for deleting multiple records
	' # May 22, 2002 By Rami Kattan
	' Enabled response buffering, which increased performance by more then 2000%
	' Made Server.ScriptTimeout dynamic, according to number of records to be exported.
	' Check if browser is still connected, so not to use extra server resources
	' Allow the functionality with non-javascript browsers
	' Security check if user can export
	'==============================================================

%><!--#include file="te_config.asp"-->
<%
lConnID = request.querystring("cid")
sTableName = request.querystring("tablename")
sQuery = request.querystring("q")

ExcelTableName = sTableName
if instr(ucase(sTableName), "SELECT") then
	ExcelTableName = "QueryResult"
end if

sNoJscript = request.querystring("nojs")
if sNoJscript = "1" then
	if not ValidSecurityID("Javaless_browser", request.querystring("SecID")) then
		response.write "Error: you must be <a href=""index.asp"">logged</a> on this site."
		response.end
	end if
end if

if not bAllowExport then
%><!--#include file="te_header.asp"-->
<p class="smallerheader">You have no permission to export data.</p>
<!--#include file="te_footer.asp"-->
<%
	response.end
end if

if not te_debug then
	Response.ContentType = "application/vnd.ms-excel"
else
	Response.ContentType = "text/html"
end if
if not te_debug then Response.AddHeader "content-disposition", "attachment; filename=" & ExcelTableName & ".xls"

	if sQuery <> "" then
		bQuery = True
		sTableName = replace(sTableName, """", "'")
	end if

	function isPrimaryKey(sFieldName)
		bPrimaryKey = False
		for iPK = 0 to ubound(aPrimaryKeys)
			if LCase(sFieldName) = LCase(aPrimaryKeys(iPK)) then
				bPrimaryKey = True
				exit for
			end if
		next
		isPrimaryKey = bPrimaryKey
	end function

	OpenRS arrConn(lConnID)
	
	'Added by Hakan
	'Find the primary key of the given table
	dim aPrimaryKeys

	if arrType(lConnID) = tedbAcess then
		set rsX = conn.openSchema(adSchemaPrimaryKeys)
		do while not rsX.eof
			if (rsX("table_name") = sTableName) then
				if sPrimaryKeyFieldName = "" then
					sPrimaryKeyFieldName = rsX("column_name")
				else
					sPrimaryKeyFieldName = sPrimaryKeyFieldName & "," & rsX("column_name")
				end if
			end if
			rsX.movenext
		loop
		rsX.close
	end if

	if (sPrimaryKeyFieldName = "") and (bQuery = False) then
		if arrType(lConnID) = tedbDsn then
			response.write "Automatic primary key detection is not possible for DSN Connections.<br><br>"
		else
			response.write "This table does not have any primary keys.<br><br>"
		end if
	else
		'response.write "Primary key(s): " & sPrimaryKeyFieldName & "<br><br>"
	end if

	'Set the primary key field to first field in the list by default
	if sPrimaryKeyFieldName = "" then sPrimaryKeyFieldName = 0
	
	'Search Support Added by Kevin
	if request.form("cmdSearch") <> "" then
		bQuery = True
		
		'Renamed cmdSubstring to chkSubstring
		if request.form("chkSubstring") <> "" then
			bSubstring = True
		end if
		
		'For different data types, added enuming fields
		'rather than form fields as Kevin did
		
		sSQL = "SELECT * FROM [" & sTableName & "] "

		on error resume next
		rs.Open sSQL,,,adCmdTable
		
		for each fld in rs.fields
			if request.form(fld.name) <> "" then
				sOP = " = "
				select case fld.type
					case adBoolean
						'BUG: What if the user dont want to perform a distinction on the boolean field?
						'Added by Hakan
						select case arrType(lConnID)
							case tedbSqlServer
								bTrue = "1"
								bFalse = "0"
							case else
								bTrue = "True"
								bFalse = "False"
						end select
						if len(request.form(fld.name))>0 then sFieldVal = bTrue else sFieldVal = bFalse
					case adLongVarBinary
						'no search on OLE fields
					case adDate
						if isDate(request.form(fld.name)) then sFieldVal = "#" & request.form(fld.name) & "#"
					case adSmallInt, adInteger, adCurrency, adUnsignedTinyInt
						if isNumeric(request.form(fld.name)) then sFieldVal = request.form(fld.name)
					case else
						sFieldVal = "'" & replace(request.form(fld.name),"'", "") & "'"
		                if bSubstring then
							sOp = " LIKE "
							sFieldVal = "'%" & request.form(fld.name) & "%'"
						else
							sFieldVal = "'" & request.form(fld.name) & "'"
		            	end if

				end select
				
				iSearchFieldCount = iSearchFieldCount + 1
				
				if iSearchFieldCount = 1 then
					sWhere = " WHERE " & fld.name & " " & sOp & sFieldVal
				else
					sWhere = sWhere & " AND " & fld.name & " " & sOp & sFieldVal
				end if
				
			end if
		next

		sTableName = "SELECT * FROM [" & sTableName & "] " & sWhere
		rs.close
	end if

		if request.querystring("nojs") = "1" then
			sFieldValues = request.querystring("chkDel")
			sFieldNames = request.querystring("txtFieldName")
			sFieldTypes = request.querystring("txtFieldType")
		else
			sFieldValues = request.form("chkDel")
			sFieldNames = request.form("txtFieldName")
			sFieldTypes = request.form("txtFieldType")
		end if
		
		aFieldNames = split(sFieldNames, ";")
		aFieldTypes = split(sFieldTypes, ";")
		aFieldValues = split(sFieldValues, ",")
		
		select case arrType(lConnID)
			case tedbSQLServer
				sDateSeperator = "'"
			case else
				sDateSeperator = "#"
		end select
		
		if ubound(aFieldNames) >= 0 then sFieldName = aFieldNames(0)
		if ubound(aFieldTypes) >= 0 then lFieldType = CLng(aFieldTypes(0))

		for iFld=0 to ubound(aFieldValues)
			sFieldValue = trim(aFieldValues(iFld))

			select case lFieldType
				case adDate, adDBDate, adDBTime, adDBTimeStamp
					if isDate(sFieldValue) then 
						sFieldValue = cDate(sFieldValue)
						sFieldValue = month(sFieldValue) & "/" & day(sFieldValue) & "/" & year(sFieldValue)
					end if
					
					if sWhereFields = "" then
						sWhereFields = "([" & sFieldName & "]=" & sDateSeperator & sFieldValue & sDateSeperator & ")"
					else
						sWhereFields = sWhereFields & " OR ([" & sFieldName & "]=" & sDateSeperator & sFieldValue & sDateSeperator & ")"
					end if
				case adTinyInt, adSmallInt, adInteger, adBigInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adUnsignedBigInt, adSingle, adDouble, adCurrency, adDecimal, adNumeric, adBoolean
					'Added by Hakan
					'Convert decimal point to dot if it's a comma
					sFieldValue = replace(sFieldValue, ",", ".")
					if sWhereFields = "" then
						sWhereFields = "([" & sFieldName & "]=" & sFieldValue & ")"
					else
						sWhereFields = sWhereFields & " OR ([" & sFieldName & "]=" & sFieldValue & ")"
					end if
				case else
					'Added by Hakan
					'Prepare SQL value by replacing single quote with two single quotes
					sFieldValue = replace(sFieldValue, "'", "''")
					if sWhereFields = "" then
						sWhereFields = "([" & sFieldName & "]='" & sFieldValue & "')"
					else
						sWhereFields = sWhereFields & " OR ([" & sFieldName & "]='" & sFieldValue & "')"
					end if
			end select
		next
		if sWhereFields <> "" then sWhere = " WHERE " & sWhereFields


	if request.form("excel_ordering") <> "" then
		sOrderBy = " ORDER BY [" & request.form("excel_ordering") & "] "
		select case request.form("excel_ordering_dir") 
			case "DESC"
				sOrderBy = sOrderBy & " DESC"
			case else
				sOrderBy = sOrderBy & " ASC"
		end select
	end if

	
	if instr(lcase(sTableName), "order by") <> 0 then
		sOrderBy = ""
	end if

	'Added by Danival
	'Modified by Hakan
	bProc = request.querystring("proc")
	if instr(1, ucase(sTableName), "SELECT") then
		sSQL =  sTableName & sOrderBy
	else
		if bProc <> "" then
			bRecAdd = False
			bRecEdit = False
			bRecDel = False
			sParamString = request.querystring("paramstring")
			sProcURL = "&proc=1&paramstring=" & sParamString
			sSQL = "EXEC [" & sTableName & "] " & sParamString
		else
			sSQL = "SELECT * FROM [" & sTableName & "]" & sWhere & sOrderBy
		end if
	end if
	on error resume next
	
'	response.write "<BR>" & sSQL & "<BR>"

	rs.CursorLocation = adUseServer
	rs.Open sSQL, conn, adOpenStatic
	
	if err <> 0 then
		response.write "Error: " & err.description & "<br><br>"
		if bQuery then
			response.write "SQL : " & sSQL & "<br><br>"
		end if

		CloseRS

		response.end
	end if
%>
<table border=1 cellspacing=1 cellpadding=1 bgcolor = "#ffe4b5" width=100%>
	<%
	for each fld in rs.fields
		if fld.type <> adLongVarBinary then
			if request("orderby") = fld.name then
				if request("dir") = "asc" then
					sDirection = "desc"
				else
					sDirection = "asc"
				end if
			else
				sDirection = "asc"
			end if
			response.write "<td class=""smallerheader"">"

			response.write fld.name
			response.write "</a>"
			response.write "</td>"
		else
			response.write "<td class=""smallerheader"">"
			response.write fld.name
			response.write "</td>"
		end if

		'Added by Hakan
		'Support for automatic primary key detection
		'Support for multiple primary keys
		aPrimaryKeys = split(sPrimaryKeyFieldName, ",")
		sPKFieldNames = ""
		sPKFieldValues = ""
		sPKFieldTypes = ""
		for iPK = 0 to ubound(aPrimaryKeys) 
			if isNumeric(aPrimaryKeys(iPK)) then aPrimaryKeys(iPK) = 0
			set fld = rs.fields(aPrimaryKeys(iPK))
			if sPKFieldNames = "" then sPKFieldNames = fld.name else sPKFieldNames = sPKFieldNames & ";" & fld.name
			'if sPKFieldValues = "" then sPKFieldValues = fld.value else sPKFieldValues = sPKFieldValues & ";" & fld.value
			if sPKFieldTypes = "" then sPKFieldTypes = fld.type else sPKFieldTypes = sPKFieldTypes & ";" & fld.type
		next
	next

	lRecs = rs.RecordCount
	TimeOutAfter = int(lRecs / 700) + 60
	'on my computer (700 @ 889 MHz, 384 MB ram), it made 750 recs per second
	Server.ScriptTimeout = TimeOutAfter

	DoneLoops = 0
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' This column is necessary for the records to properly align with their respective column names...
' If anyone can figure something better out, please help.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	do while not rs.eof 
		DoneLoops = DoneLoops + 1
		if (DoneLoops MOD 100) = 0 then Response.Flush
		if not Response.IsClientConnected then exit do

		response.write "<tr bgcolor=""#ffffff"">"
		for each fld in rs.fields
			response.write "<td class=""smallertext"">"
				select case fld.type
					case adSmallInt, adInteger
						response.write rs(fld.name)
					case adDate
						if isdate(rs(fld.name)) then
							response.write rs(fld.name)
						end if
					case adBoolean
						if rs(fld.name)=true then
							response.write "True"
						else
							response.write "False"
						end if
					case adLongVarBinary
						Response.write "EXCEL EXPORTER: Currently OLE Data not supported"
					case adVarWChar, adLongVarWChar		'Text, Memo
						sVal = rs(fld.name)
						if (bEncodeHTML) and (len(sVal) > 0)then
							response.write server.htmlencode(sVal)
						else
							response.write sVal
						end if
					case else
						response.write rs(fld.name)
				end select
			response.write "</td>"
		next
		response.write "</tr>"
		rs.movenext
	loop
	CloseRS
%>
</table>