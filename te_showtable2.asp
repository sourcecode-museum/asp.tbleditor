<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_showtable2.asp
	' Description: Displays the selected table contents (Generic)
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
	' # May 11, 2002 by Hakan Eskici
	' Added page selector max check
	' Added button for exporting the whole table
	' # May 20, 2002 by Rami Kattan
	' More backward browser compatibility
	' Support for non-javascript browsers
	' Fixed export from Stored Procedures
	' Options for working with no-popups
	' # Jun 3, 2002 by Rami Kattan
	' Fixes for Konqueror for linux
	'==============================================================
%>
<!--#include file="te_config.asp"-->
<%
' Get the requested number of records per page
cPerPage = request.querystring("cPerPage")
if cPerPage = "" then cPerPage = iDefaultPerPage

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
		response.redirect "te_admin.asp"
	end if
	'------------------------------

%>
<!--#include file="te_header.asp"-->
<%
	lConnID = request.querystring("cid")
	sTableName = request.querystring("tablename")
	sQuery = request.querystring("q")
	
	if sQuery <> "" and sQuery <> "0" then
		bQuery = True
		sTableName = replace(sTableName, """", "'")
	end if

	if request.querystring("cid") = "0" AND request.querystring("tablename") = "Databases" then
		EditScriptName = "te_addDB"
	else
		EditScriptName = "te_showrecord"
	end if
if bJSEnable then 
%>
<!-- ------------------------------------------------------------- -->
<script language="JavaScript" type="text/javascript" >

//Check all radio/check buttons script- by javascriptkit.com
//Visit JavaScript Kit (http://javascriptkit.com) for script
//Credit must stay intact for use  

function checkall(thestate){
var el_collection=eval("document.forms.frmAddDelete.chkDel")
for (c=0;c<el_collection.length;c++)
el_collection[c].checked=thestate
}

<% if bComboTables then %>

function ChangeTable(){
	if (GetObject("allTables").options[GetObject("allTables").selectedIndex].value != "<<")
		location.href='<% =TableViewerCompat %>?cid=<%=lConnID%>&tablename=' + GetObject("allTables").options[GetObject("allTables").selectedIndex].value;
	else
		location.href='te_listtables.asp?cid=<%=lConnID%>';
}
<% end if %>
</script>
<!-- ------------------------------------------------------------- -->
<% end if %>
<table border=0 cellspacing=1 cellpadding=2 bgcolor = "#ffe4b5" width="100%">
	<tr>
		<td class="smallertext">
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <% allTablesCombo() %> » <% if bQuery then 
					response.write "Query" 
			else 
				response.write "Table: ["
				allTablesCombo2(lConnID)
				response.write "]"
			end if%>
		</td>
		<td class="smallerheader" width=130 align=right>
		<%
		if bProtected then 
			response.write session("teFullName")
			response.write " (<a href=""te_logout.asp"">logout</a>)"
		end if
		%>
		</td>
	</tr>
</table>
<br>
<%
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
	if arrType(lConnID) <> tedbDsn then
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
		sSoWhat = "(<a title=""TableEditoR needs at least one unique key field to distinguish the records. The key field of this table either doesn't exist or cannot be automatically detected. TableEditoR will use the first field as the key. (Click on the page image in the Action column to edit the record anyway)."" style=""cursor: hand"">So What?</a>)"
		if arrType(lConnID) = tedbDsn then
			response.write "Automatic primary key detection is not possible for DSN Connections. " & sSoWhat & "<br><br>"
		else
			response.write "This table does not have any primary keys. " & sSoWhat & "<br><br>"
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
	
	if cPerPage = 0 then
		RSPagesize = lRecs
	else
		RSPagesize = cPerPage
	end if

	if isNumeric(request("ipage")) then iPage = CLng(request("ipage"))
	rs.PageSize = RSPagesize
	rs.CacheSize = RSPagesize
	iPageCount = rs.PageCount

	if iPage < 1 then iPage = 1
	if lRecs > 0 then rs.AbsolutePage = iPage
	
	if bQuery or te_debug then
		response.write sSQL & "<br><br>"
	end if

	if bPopUps then AddRecURL = "javascript:openWindow('"
	AddRecURL = AddRecURL & EditScriptName & ".asp?cid=" & lConnID & "&tablename=" & server.UrlEncode(sTableName) & "&add=1&ipage=" & iPage & "&recs=" & RSPagesize
	if bPopUps then AddRecURL = AddRecURL & "')"

	FormAction = "#"
	if not bJSEnable then FormAction = "te_formaction.asp"

if bJSEnable then
%>
<!-- -------------------------------------------------------------------- -->
<!-- 

If anyone can help with this, I would like to populate the value
	 of the select box with the correct number after the onChange event -->
<script language="JavaScript" type="text/javascript">

<!--

function ChangePerPage() {
	if (GetObject("URLSelect").options[GetObject("URLSelect").selectedIndex].value != 0 )
		location.href = "te_showtable2.asp?cid=<% =lConnID %>&tablename=<% =server.UrlEncode(sTableName) %>&q=<% =bQuery %>&ipage=1&cPerPage=" + GetObject("URLSelect").options[GetObject("URLSelect").selectedIndex].value;
	else
		location.href = "te_showtable2.asp?cid=<% =lConnID %>&tablename=<% =server.UrlEncode(sTableName) %>&q=<% =bQuery %>&ipage=1&cPerPage=<% =lRecs %>";
}

function ChangePageNum() {
	<% if bPageSelector and iPageCount < iPageSelectorMax then %>
		var iGotoPage = GetObject("Pg").options[GetObject("Pg").selectedIndex].value;
	<% else %>
		var iGotoPage = GetObject("Pg").value;
	<% end if %>
	location.href = "te_showtable2.asp?cid=<% =lConnID %>&tablename=<% =server.UrlEncode(sTableName) %>&q=<% =bQuery %>&cPerPage=<% =cPerPage %>&ipage=" + iGotoPage + "<% =sOrderByLink & sProcURL %>"
}
//-->
</script>
<% end if %>
<table border=0 cellspacing=1 cellpadding=2 bgcolor = "#ffdead" width="100%">
<form name="headerForm">
	<tr>
    <td width="10"></td>
		<td class="smallerheader"><%if bQuery then response.write "Query" else response.write sTableName%></td>
		<td class="smallertext" width=100><%=lRecs%> records</td>
		<% if bRecAdd then %>
		<td class="smallertext" width=100>
		<a href="<% =AddRecURL %>">Add Record</a>
		</td>
		<% end if %>
		
		<td class="smallertext">
		<%
			'Build navigation bar
			if iPage <> 1 then response.write "<a href=""te_showtable2.asp?cid=" & lConnID & "&tablename=" & server.UrlEncode(sTableName) & "&q=" & bQuery & "&ipage=1" & "&cPerPage=" & cperPage & sOrderByLink & sProcURL & """>first</a> :: " else response.write "first :: "
			if iPage > 1 then response.write "<a href=""te_showtable2.asp?cid=" & lConnID & "&tablename=" & server.UrlEncode(sTableName) & "&q=" & bQuery & "&ipage=" & iPage - 1  & "&cPerPage=" & cperPage & sOrderByLink & sProcURL & """>previous</a> :: " else response.write "previous :: "
			if iPage < iPageCount then response.write "<a href=""te_showtable2.asp?cid=" & lConnID & "&tablename=" & server.UrlEncode(sTableName) & "&q=" & bQuery & "&ipage=" & iPage + 1  & "&cPerPage=" & cperPage & sOrderByLink & sProcURL & """>next</a> :: " else response.write "next :: "
			if iPage <> iPageCount then response.write "<a href=""te_showtable2.asp?cid=" & lConnID & "&tablename=" & server.UrlEncode(sTableName) & "&q=" & bQuery & "&ipage=" & iPageCount  & "&cPerPage=" & cperPage & sOrderByLink & sProcURL & """>last</a>" else response.write "last"
		%>
		</td>
		<td class="smallertext" align=right>
			Page <%	

			If (bPageSelector and bJSEnable) Then 
				' Added by Hakan on May 11, 2002
				' Don't display the combo if there are too many pages
				if iPageCount < iPageSelectorMax then
				    response.write "<select id=""Pg"" name=""Pg"" onchange=""ChangePageNum()"" class=""smallertext"">"
					For i = 1 to iPageCount
					    response.write "<option value=""" & i & """"
						if iPage = i then response.write " selected"
						response.write ">" & i & "</option>"
					Next
					response.write "</select>" & vbcrlf
				else
					%><input type="text" id="Pg" name="Pg" onchange="ChangePageNum()" class="smallbutton" size="6" value="<%=iPage%>"><input type="button"	onclick="ChangePageNum()" class="smallbutton" value="»"><%
				end if
			Else
			    response.write iPage
			End If
		%> of <%=iPageCount%>
</td>
<% if bJSEnable then %>
<td class="smallertext" align="right"> Show 
		<select id="URLSelect" name="URLSelect" onchange="ChangePerPage()" size="1" class="smallertext">
        	<option value="5"<% isPerPage cPerPage, 5 %>>5</option>
        	<option value="10"<% isPerPage cPerPage, 10 %>>10</option>
        	<option value="15"<% isPerPage cPerPage, 15 %>>15</option>
        	<option value="20"<% isPerPage cPerPage, 20 %>>20</option>
        	<option value="30"<% isPerPage cPerPage, 30 %>>30</option>
        	<option value="40"<% isPerPage cPerPage, 40 %>>40</option>
        	<option value="0"<% isPerPage cPerPage, 0 %>>All</option>
        </select>
		records per page
</td>
<% end if %>
</tr>
</form>
</table>

<form id="frmAddDelete" name="frmAddDelete" action="<% =FormAction %>" method="post">
<input type="hidden" name="excel_ordering" value="<% =request.QueryString("orderby") %>">
<input type="hidden" name="excel_ordering_dir" value="<% =request.QueryString("dir") %>">
<table border=0 cellspacing=1 cellpadding=2 bgcolor=#ffe4b5 width="100%" <% if bTableHighlight then response.write "style=""behavior:url(tablehl.htc);"" slcolor='#ffffcc' hlcolor='#bec5de'"%>>
<thead>
<tr bgcolor="#fffaf0"><td class="smallertext" width=10></td>
	<%
	response.write "<td class=""smallerheader"" width=30>Action</td>"
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
			response.write "<a href=""te_showtable2.asp?cid=" & lConnID & "&tablename=" & server.UrlEncode(sTableName) & "&q=" & bQuery & "&ipage=" & iPage & "&cPerPage=" & cperPage & "&orderby=" & fld.name & "&dir=" & sDirection & sProcURL &  """>"
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

	mainFrmExt_str = "?cid=" & lConnID & "&tablename=" & server.urlencode(sTableName)
	if request.querystring("proc") <> "" then
		mainFrmExt_str = mainFrmExt_str & sProcURL
	end if
	if request.querystring("q") <> "" then
		mainFrmExt_str = mainFrmExt_str & "&q=1"
	end if
	%>
<input type="hidden" id="mainFrmExt" name="mainFrmExt" value="<% =mainFrmExt_str %>">
<% 	if bJSEnable then %>
<script language="JavaScript" type="text/javascript">
<!--
function EW(inData){
	url = '<% =EditScriptName %>.asp?cid=<%=lConnID%>&q=<% =bQuery %>&tablename=<% =server.UrlEncode(sTableName) %>&fld=<% =server.URLEncode(sPKFieldNames) %>&val=' + inData + '&fldtype=<% =server.URLEncode(sPKFieldTypes) %>&ipage=<% =iPage %>'
<%
	if bPopUps then
		response.write  "	openWindow(url);"
	else
		response.write  "	location.href = url;"
	end if
%>
}

function DW(inData){
	url = 'te_deleterecord.asp?cid=<%=lConnID%>&q=<% =bQuery %>&tablename=<% =server.UrlEncode(sTableName) %>&fld=<% =server.URLEncode(sPKFieldNames) %>&val=' + inData + '&fldtype=<% =server.URLEncode(sPKFieldTypes) %>&ipage=<% =iPage %>'
<%
	if bPopUps then
		response.write  "	openWindow(url);"
	else
		response.write  "	location.href = url;"
	end if
%>
}
//-->
</script>
<%
end if

	response.write "<td width=10></td>"
	response.write "</tr></thead>"

	'Key Field form elements for Multiple delete
	response.write "<input type=""hidden"" name=""txtFieldName"" value=""" & sPKFieldNames & """>"
	response.write "<input type=""hidden"" name=""txtFieldType"" value=""" & sPKFieldTypes & """><tbody>"

	do while not rs.eof 

		if iRecCount = RSPagesize then exit do

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
	
		response.write "<tr bgcolor=""#fffaf0"">" & vbCrLf & vbTab
		response.write "<td width=10></td>" & vbCrLf & vbTab
		response.write "<td nowrap>"

		if bJSEnable then
			sPKURL    = "<a href=""javascript:EW('" & server.URLEncode(sPKFieldValues) & "')"">"
			sPKURLDel = "<a href=""javascript:DW('" & server.URLEncode(sPKFieldValues) & "')"">"
		else
			sPKURL    = "<a href=""te_showrecord.asp?cid=" & lConnID & "&q=" & bQuery & "&tablename=" & server.UrlEncode(sTableName) & "&fld=" & sPKFieldNames & "&val=" & sPKFieldValues & "&fldtype=" & sPKFieldTypes &  "&ipage=" & iPage &""">"
			sPKURLDel = "<a href=""te_deleterecord.asp?cid=" & lConnID & "&q="& bQuery & "&tablename=" & server.UrlEncode(sTableName) & "&fld=" & sPKFieldNames & "&val=" & sPKFieldValues & "&fldtype=" & sPKFieldTypes & "&ipage=" & iPage & """>"
		end if

		if bRecEdit then response.write sPKURL & "<img src=""images/edit.gif"" width=9 height=11 alt=""edit"" border=0></a>&nbsp;"
		if bRecDel then 
			'One click delete link
			response.write sPKURLDel & "<img src=""images/del.gif"" width=9 height=11 alt=""delete"" border=0></a>"
			'Multi Delete Check box
			response.write "<input type=""checkbox"" name=""chkDel"" value=""" & sPKFieldValues & """>"
		end if
		response.write "</td>"
		iFieldCount = 0
		for each fld in rs.fields
			iFieldCount = iFieldCount + 1
			response.write "<td class=""smallertext"">"
			if isPrimaryKey(fld.name) = True then
				response.write sPKURL & rs(fld.name) & "</a>"
			else
				select case fld.type
					case adSmallInt, adInteger
						response.write rs(fld.name)
					case adDate
						if isdate(rs(fld.name)) then
							response.write rs(fld.name)
						end if
					case adBoolean
						response.write "<input type=""checkbox"" name=""chk"""
						if isIE or isNS or isOP or isKo then response.write " disabled"
						if rs(fld.name)=true then
							response.write " checked>"
						else
							response.write ">"
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
						sVal = MakeURL(sVal)
						if (bEncodeHTML) and (len(sVal) > 0)then
							response.write server.htmlencode(sVal)
						else
							response.write sVal
						end if
					case else
						response.write rs(fld.name)
				end select
			end if
			response.write "</td>" & vbCrLf & vbTab
		next
		response.write "<td width=10></td>"
		response.write "</tr>" & vbCrLf
		rs.movenext
		iRecCount = iRecCount + 1
	loop
		
	CloseRS

%>
</tbody></table>
<table border="0" cellspacing="0" cellpadding="2" bgcolor="#ffe4b5" width="100%">
<% 
if bJSEnable then
	if bRecDel and bBulkDelete then %>
<tr bgcolor="#fffaf0">
	  <td class="smallertext"> 
		  <input type="button" name="cmdChkAll" value="Check All" class="cmdflat" onclick="checkall(true)">
		  <input type="button" name="cmdUnChkAll" value="Uncheck All" class="cmdflat" onclick="checkall(false)">
          <input type="button" name="cmdMultiDel" value="Delete Selected" class="cmdflat" onclick="MainFormAction('multidelete')">
      </td>
</tr>
<% end if  'del
   if (bExportExcel or bExportXML) and bJSEnable and bAllowExport then %>
<tr bgcolor="#fffaf0">
  <td class="smallertext"> 
	<% if bExportExcel then %><input type="button" name="cmdExportExcel" class="cmdflat" onclick="MainFormAction('excel')" value="Export to Excel">
	<% end if
	   if bExportXML then
	%><input type="button" name="cmdExportXML" class="cmdflat" onclick="MainFormAction('XML')" value="Export to XML">
	<% end if %>
  </td>
</tr>
<% end if  'excel or xml

else   ' for if bJSEnabled
%>
<tr bgcolor="#fffaf0">
<td>
<% if bRecDel and bBulkDelete then %><input type="submit" value="Delete Selected" name="action" class="cmdflat"><% end if%>
<% if bExportExcel then %><input type="submit" value="Export to Excel" name="action" class="cmdflat"><% end if %>
<% if bExportXML then %><input type="submit" value="Export to XML" name="action" class="cmdflat"><% end if %>
</td>
</tr>
<% end if ' bJSEnabled %>
</table>
</form>
<!--#include file="te_footer.asp"-->