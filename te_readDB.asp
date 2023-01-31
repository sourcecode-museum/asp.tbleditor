<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_readDB.asp
	' Description: Generate CSV file for te_showtable.asp (IE mode)
	' Initiated By Hakan Eskici on Nov 07, 2000
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
	' # April 18, 2002 by Rami Kattan
	' this file generate only CSV text file to be read by the
	' te_showtable.asp and view the records
	'-------------------------------------------------------------
	Response.ContentType = "text/csv"

' Get the requested number of records per page
cPerPage = CLng(request.QueryString("cPerPage"))
If cPerPage = 0 or cPerPage= "" then cPerPage = iDefaultPerPage

%>
<!--#include file="te_config.asp"-->
<%
	lConnID = request("cid")
	sTableName = request("tablename")
	sQuery = request("q")
	
	'------------------------------
	'added 8/10/01 by j.wilkinson, jwilkinson@mail.com
	'added a check for nonAdmin users trying to view the admin table
	'This is just checking that the connection ID = 0, assumes that
	'non-admin users have no legitimate reason to get to that db at all.
	'  note that this may not protect against using queries to view
	'  this db and table
	if lConnID=0 and not bAdmin then
		response.end
	end if
	'------------------------------

    const csvchar = ","
	
	if sQuery <> "" then
		bQuery = True
		sTableName = replace(sTableName, """", "'")
	end if

	function CheckData(inData)
		if not isnull(inData) then
'			inData = replace(inData, vbCrLf, "")
'			inData = replace(inData, vbCr, "")
'			inData = replace(inData, vbLf, "")
			inData = replace(inData, "\", "\\")
			inData = replace(inData, ",", "\,")
			inData = replace(inData, ";", "\;")
			inData = replace(inData, "<form", "&lt;form")
			inData = replace(inData, "</form", "&lt;/form")
			inData = replace(inData, """", "\""")
		end if
		CheckData = inData
	end function

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
	if arrType(lConnID) = tedbDsn then
		'response.write "Automatic primary key detection is not possible for DSN Connections. " & sSoWhat & "<br><br>"
	else
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

	'Set the primary key field to first field in the list by default
	if sPrimaryKeyFieldName = "" then sPrimaryKeyFieldName = 0
	
	if request.querystring("orderby") <> "" then
		sOrderBy = " ORDER BY [" & request.querystring("orderby") & "] "
		sOrderByLink = "&orderby=" & request.querystring("orderby")
		select case request.querystring("dir") 
			case "desc"
				sOrderBy = sOrderBy & " DESC"
				sOrderByLink = sOrderByLink & "&dir=desc"
			case "asc"
				sOrderBy = sOrderBy & " ASC"
				sOrderByLink = sOrderByLink & "&dir=asc"
			case else
				sOrderBy = sOrderBy & " ASC"
				sOrderByLink = sOrderByLink & "&dir=asc"
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
	
	rs.CursorLocation = adUseServer
	rs.Open sSQL, conn, adOpenStatic
	
	if err <> 0 then
		response.write "Error: " & err.description & "<br><br>"
		if bQuery then
			response.write "SQL : " & sSQL & "<br><br>"
		end if
		response.write "Click here to <a href=""javascript:history.back()"">go back</a>.<br><br>"
		CloseRS
		%><!--#include file="te_footer.asp"--><%
		response.end
	end if

	on error goto 0
	
	'Performance Issue:
	'Getting the recordset properties may take long time for tables with many records
	lRecs = rs.RecordCount
	lFields = rs.Fields.Count
	
	if isNumeric(request("ipage")) then iPage = CLng(request("ipage"))
	rs.PageSize = cPerPage
	rs.CacheSize = cPerPage
	iPageCount = rs.PageCount

	if iPage < 1 then iPage = 1
	if lRecs > 0 then rs.AbsolutePage = iPage
	
'	if bQuery then
'		response.write sSQL & ""
'	end if

	response.write "Action"
	for each fld in rs.fields
' Type			Description 
' -------------------------
' String		Text data 
' Date			Calendar date 
' Boolean		Logical data 
' Int			Integer number 
' Float			Floating-point number 

		
		select case fld.Type
			case adSmallInt, adInteger, adCurrency	: tdcType = "Int"
			case adVarWChar, adLongVarWChar			: tdcType = "String"
			case adBoolean		: tdcType = "Boolean"
			case adDate			: tdcType = "date"
			case else			: tdcType = ""
		end select

		response.write csvchar & fld.name
		if tdcType <> "" then response.write ":" & tdcType

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

	response.write ";"

	iRecCount = 0


	'Key Field form elements for Multiple delete
	do while not rs.eof 
		if (iRecCount = cPerPage) or (not Response.IsClientConnected) then exit do
		if (iRecCount MOD 15) = 0 then Response.Flush
		response.write vbCrLf

		if arrType(lConnID) = tedbAccess or arrType(lConnID) = tedbSQLServer then
			'Only Access and SQL can do this
			if rs.AbsolutePage <> iPage then exit do
		end if

		sPKFieldValues = ""
		for iPK = 0 to ubound(aPrimaryKeys) 
			if isNumeric(aPrimaryKeys(iPK)) then aPrimaryKeys(iPK) = 0
			set fld = rs.fields(aPrimaryKeys(iPK))
			'if sPKFieldNames = "" then sPKFieldNames = fld.name else sPKFieldNames = sPKFieldNames & ";" & fld.name
			if sPKFieldValues = "" then sPKFieldValues = fld.value else sPKFieldValues = sPKFieldValues & ";" & fld.value
			'if sPKFieldTypes = "" then sPKFieldTypes = fld.type else sPKFieldTypes = sPKFieldTypes & ";" & fld.type
		next

		sPKURL = "<a href=""javascript:EW('" & server.URLEncode(sPKFieldValues) & "')"">"

		if bRecEdit then response.write "<img src=""images/edit.gif"" alt=""edit"" class=""smimg"" onclick=""EW('" & server.URLEncode(sPKFieldValues) & "')"">&nbsp\;"
		if bRecDel then 
			'One click delete link
			response.write "<img src=""images/del.gif"" alt=""delete"" onclick=""DW('" & server.URLEncode(sPKFieldValues) & "')"" class=""smimg"">"
			'Multi Delete Check box
			response.write "<input type=""checkbox"" name=""chkDel"" value=""" & sPKFieldValues & """>"
		end if
'		response.write csvchar
		iFieldCount = 0
		for each fld in rs.fields
			response.write csvchar
			iFieldCount = iFieldCount + 1
			if isPrimaryKey(fld.name) = True then
				response.write sPKURL & rs(fld.name) & "</a>"' & csvchar
			else
				select case fld.type
					case adSmallInt, adInteger
						response.write rs(fld.name)
					case adDate
						if isdate(rs(fld.name)) then
							response.write """" & rs(fld.name) & """"
						end if
					case adBoolean
						response.write """<input type=\""checkbox\"" name=\""chk\"" disabled"
						if rs(fld.name)=true then
							response.write " checked>"""
						else
							response.write ">"""
						end if
					case adLongVarBinary
							If Not isNull(rs(fld.name)) Then
								response.write "<img align=""absmiddle"" src=""te_imagesdb.asp?cid=" & lConnID & "&tablename=" & server.UrlEncode(sTableName) & "&fld=" & sPKFieldNames & "&val=" & sPKFieldValues & "&fldtype=" & sPKFieldTypes & "&olefield=" & server.UrlEncode(fld.name) & """>"
							End if
					case adVarWChar, adLongVarWChar		'Text, Memo
						if lMaxShowLen > 0 then
							'If max # of chars is specified
							sVal = left(rs(fld.name), lMaxShowLen)
						else
							sVal = rs(fld.name)
						end if
						sVal = CheckData(sVal)
						sVal = MakeURL(sVal)
						if (bEncodeHTML) and (len(sVal) > 0)then
							response.write """" & server.htmlencode(sVal) & """"
						else
							response.write """" & sVal & """"
						end if
					case else
						response.write """" & CheckData(rs(fld.name)) & """"
				end select
'				response.write csvchar
			end if
		next
		response.write ";" '& VbCrLf
		rs.movenext
		iRecCount = iRecCount + 1
	loop
	CloseRS
%>