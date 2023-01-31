<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_showtable.asp
	' Description: Displays the selected table contents (IE Mode)
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
	' # Apr 20, 2002 by Rami Kattan
	' IE mode, using client-side data viewer
	' More dynamic navigation
	' Combo Selector for databases/tables
	' # May 11, 2002 by Hakan Eskici
	' Added page selector max check
	' Added button for exporting the whole table
	' # May 21, 2002 by Rami Kattan
	' Fixed navigation of Stored Procedures
	' Options for working with no-popups
	'==============================================================
%>
<!--#include file="te_config.asp"-->
<%
' Get the requested number of records per page
cPerPage = request.QueryString("cPerPage")
if cPerPage="" then cPerPage = iDefaultPerPage
'If cPerPage = 0 Then cPerPage = 10

	lConnID = request("cid")
	sTableName = request("tablename")
	sQuery = request("q")

	if (instr(sU, "MSIE")=0 OR instr(sU, "OPERA") > 0) then response.redirect "te_showtable2.asp?cid=" & lConnID & "&tablename=" & sTableName
	
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

	dim bOnPageLoad  ' to tell header to add onload code to body tag or not.
	bOnPageLoad = true
%>
<!--#include file="te_header.asp"-->
<%
	if sQuery <> "" then
		bQuery = True
		sTableName = replace(sTableName, """", "'")
	end if

SendQueryString = ""
htmlSorter = ""
'For each varr in Request.form
'	if ucase(varr) <> "CID" and ucase(varr) <> "TABLENAME" then
'		SendQueryString = SendQueryString & "&" &varr & "=" & Request.form(varr)
'	end if
'next

%>
<span id="DisplayMessageBox" style="display:none; position:absolute; padding: 18 20 20 20; filter: progid:DXImageTransform.Microsoft.Shadow(direction=135,color=#B0B0B0,strength=5) progid:DXImageTransform.Microsoft.Alpha( style=0,opacity=75) progid:DXImageTransform.Microsoft.Gradient(gradientType=1,startColorStr=#A5CBF7,endColorStr=#08246B); font: bold 14pt/1.3 verdana; color: #FFFFFF; height: 60px" ><img src="images/loading.gif" align="absmiddle"> Please wait While loading data...</span>

<table border=0 cellspacing=1 cellpadding=2 bgcolor="#ffe4b5" width="100%">
	<tr>
		<td class="smallertext">
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <% allTablesCombo() %> » <%if bQuery then 
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
	
	if arrType(lConnID) = tedbDsn then
		'response.write "Automatic primary key detection is not possible for DSN Connections. " & sSoWhat & "<br><br>"
	else
		sSoWhat = "(<a title=""TableEditor needs at least one unique key field to distinguish the records. The key field of this table either doesn't exist or cannot be automatically detected. TableEditoR will use the first field as the key. (Click on the page image in the Action column to edit the record anyway)."" style=""cursor: hand"">So What?</a>)"

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

		if (sPrimaryKeyFieldName = "") and (bQuery = False) then
			if arrType(lConnID) = tedbDsn then
				response.write "Automatic primary key detection is not possible for DSN Connections. " & sSoWhat & "<br><br>"
			else
				response.write "This table does not have any primary keys. " & sSoWhat & "<br><br>"
			end if
		else
			'response.write "Primary key(s): " & sPrimaryKeyFieldName & "<br><br>"
		end if
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
	bProc = request("proc")
	if instr(1, ucase(sTableName), "SELECT") then
		sSQL =  sTableName & sOrderBy
	else
		if bProc <> "" then
			bRecAdd = False
			bRecEdit = False
			bRecDel = False
			sParamString = request("paramstring")
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
	if iPageCount = 0 then iPageCount = 1

	if iPage < 1 then iPage = 1
	if lRecs > 0 then rs.AbsolutePage = iPage
	
	if bQuery or te_debug then
		response.write sSQL & "<br><br>"
	end if

	'Added by Rami Kattan
	'for each fld in rs.fields
	'		htmlSorter = htmlSorter & "<option value=""" & fld.name & """>" & fld.name
	'next

	for each fld in rs.fields
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
			if sPKFieldTypes = "" then sPKFieldTypes = fld.type else sPKFieldTypes = sPKFieldTypes & ";" & fld.type
		next
	next

	if request.querystring("cid") = "0" AND request.querystring("tablename") = "Databases" then
		EditScriptName = "te_addDB"
	else
		EditScriptName = "te_showrecord"
	end if

	mainFrmExt_str = "?cid=" & lConnID & "&tablename=" & server.urlencode(sTableName)
	if request.querystring("proc") <> "" then
		mainFrmExt_str = mainFrmExt_str & sProcURL
	end if

	if request.querystring("q") <> "" then
		mainFrmExt_str = mainFrmExt_str & "&q=1"
	end if

	AddRecURL = EditScriptName & ".asp?cid=" & lConnID & "&tablename=" & server.UrlEncode(sTableName) & "&add=1&recs=" & RSPagesize & "&ipage="

%>
<script language="JavaScript" type="text/javascript">
<!--
function EW(inData){
	url = '<% =EditScriptName %>.asp?cid=<%=lConnID%>&q=<% =bQuery %>&tablename=<% =server.UrlEncode(sTableName) %>&fld=<% =server.URLEncode(sPKFieldNames) %>&val=' + inData + '&fldtype=<% =server.URLEncode(sPKFieldTypes) %>&ipage=' + GetCurrentPage()
<%
	if bPopUps then
		response.write  "	openWindow(url);"
	else
		response.write  "	location.href = url;"
	end if
%>
}

function DW(inData){
	url = 'te_deleterecord.asp?cid=<%=lConnID%>&q=<% =bQuery %>&tablename=<% =server.UrlEncode(sTableName) %>&fld=<% =server.URLEncode(sPKFieldNames) %>&val=' + inData + '&fldtype=<% =server.URLEncode(sPKFieldTypes) %>&ipage=' + GetCurrentPage()
<%
	if bPopUps then
		response.write  "	openWindow(url);"
	else
		response.write  "	location.href = url;"
	end if
%>
}
<% if bComboTables then %>
function cboChangeDB(){
	location.href = "te_listtables.asp?cid=" + GetObject("allDBs").value
}
function ChangeTable(){
	if (GetObject("allTables").value != "<<")
		location.href = '<% =TableViewerCompat %>?cid=<%=lConnID%>&tablename=' + GetObject("allTables").value;
	else
		location.href = 'te_listtables.asp?cid=<%=lConnID%>';
}
<% end if %>

function AddRecord(){
<%
	if bPopUps then
		response.write  "openWindow('" & AddRecURL & "'+GetCurrentPage());"
	else
		response.write  "location.href = '" & AddRecURL & "'+GetCurrentPage();"
	end if
%>
}
//-->
</script>
<OBJECT id="DynamicData" CLASSID="clsid:333C7BC4-460F-11D0-BC04-0080C7055A83" ondatasetchanged="ShowTransientMessage();" ondatasetcomplete="HideTransientMessage();">
	<PARAM NAME="DataURL" VALUE="te_readDB.asp<% =mainFrmExt_str %>&cPerPage=<% =RSPagesize %>&iPage=<% =iPage %>">
	<PARAM NAME="UseHeader" VALUE="True">
	<PARAM NAME="FieldDelim" VALUE=",">
	<PARAM NAME="RowDelim" VALUE=";">
	<PARAM NAME="TextQualifier" VALUE='"'>
	<PARAM NAME="EscapeChar" VALUE="\">
</OBJECT>
<OBJECT id="RecordCounter" CLASSID="clsid:333C7BC4-460F-11D0-BC04-0080C7055A83">
	<PARAM NAME="DataURL" VALUE="te_readDB_counter.asp<% =mainFrmExt_str %>">
	<PARAM NAME="UseHeader" VALUE="True">
	<PARAM NAME="FieldDelim" VALUE=",">
	<PARAM NAME="EscapeChar" VALUE="\">
</OBJECT>
<input type="hidden" name="Current_Page" value="<% =iPage %>">
<input type="hidden" name="totalRecords" value="0" DataSrc="#RecordCounter" DataFld="TotalRecords">
<input type="hidden" name="db_lConnID" value="<% =lConnID %>">
<input type="hidden" name="db_sTableName" value="<%=server.UrlEncode(sTableName)  & sProcURL %>">
<input type="hidden" name="db_string" value="<% =sProcURL %>">
<input type="hidden" name="db_Ordering" value="">
<input type="hidden" id="mainFrmExt" name="mainFrmExt" value="<% =mainFrmExt_str %>">
<table border=0 cellspacing=1 cellpadding=2 bgcolor = "#ffdead" width="100%">
	<tr>
    <td width=10>&nbsp;</td>
		<td class="smallerheader"><%if bQuery then response.write "Query" else response.write sTableName%></td>
		<td class="smallertext" width=100><span id="TotRecs" DataSrc="#RecordCounter" DataFld="TotalRecords"><%=lRecs%></span> records</td>
		<% if bRecAdd then %>
		<td class="smallertext" width=100>
		<button id="btnAddRec" class="smallertext" onclick="AddRecord()">Add Record</button>
		</td>
		<% end if %>
		
		<td class="smallertext">
		<button id="btnFirst" class="smallertext" onclick="gotoFirst()" disabled>First</button>
		<button id="btnPrev" class="smallertext" onclick="gotoPrev()" disabled>Previous</button>
		<button id="btnNext" class="smallertext" onclick="gotoNext()" disabled>Next</button>
		<button id="btnLast" class="smallertext" onclick="gotoLast()" disabled>Last</button>
		&nbsp;&nbsp;&nbsp;
		<button id="btnRefresh" class="smallertext" onclick="update_view()">Refresh</button>
		</td>
		<td class="smallertext" align=right>
			Page <% 
	if bPageSelector then
		if iPageCount < iPageSelectorMax then 
			%><select id="pg" name="pg" onchange="update_view()" class="smallertext"><% 
			For i = 1 to iPageCount
			    response.write "<option value=""" & i & """"
				if i = iPage then response.write " selected"
				response.write ">" & i
			Next 
			%></select><% 
		else 
			%><input type="text" id="pg" name="pg" onchange="update_view()" class="smallbutton" size="6" value="1"><button onchange="update_view()" class="smallbutton">»</button><% 
		end if
	else 
		%><span id="pg">1</span><% 
	end if 
	%> of <span id="pagecount"><%=iPageCount%></span>
		</td>
	<td class="smallertext" align="right"> Show 
		<select NAME="URLSelect" onchange="update_view()" class="smallertext">
        	<option value="5"<% isPerPage cPerPage, 5 %>>5</option>
        	<option value="10"<% isPerPage cPerPage, 10 %>>10</option>
        	<option value="15"<% isPerPage cPerPage, 15 %>>15</option>
        	<option value="20"<% isPerPage cPerPage, 20 %>>20</option>
        	<option value="30"<% isPerPage cPerPage, 30 %>>30</option>
        	<option value="40"<% isPerPage cPerPage, 40 %>>40</option>
        	<option value="0"<% isPerPage cPerPage, 0 %>>All</option>
        </select>
 		records per page </td>
	<td>&nbsp;</td>
</tr>
<tr id="filtering" style="display:none">
<td colspan="9" align="center" class="smallertext">Filter records&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
field:
<select name="FilterField" class="smallertext">
	<% =htmlSorter %>
</select>
<input type="text" name="FilterWord" class="smallertext">
<button onclick="FilterData()" class="smallertext">Go</button>
</td>
</tr>
</table>

<form id="frmAddDelete" name="frmAddDelete" action="te_javascript.asp?cid=<%=lConnID%>&tablename=<%=server.urlencode(sTableName)%>" method="post">
<input type="hidden" name="excel_ordering" value="">
<input type="hidden" name="excel_ordering_dir" value="">
<table border=0 cellspacing=1 cellpadding=2 bgcolor=#ffe4b5 width="100%" <% if bTableHighlight then response.write "style=""behavior:url(tablehl.htc);"" slcolor='#ffffcc' hlcolor='#bec5de'"%>DataSrc="#DynamicData">
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
'			response.write "<a href=""" & TableViewerCompat & "?cid=" & lConnID & "&tablename=" & server.UrlEncode(sTableName) & "&q=" & bQuery & "&ipage=" & iPage & "&orderby=" & fld.name & "&dir=" & sDirection & sProcURL &  """>"
			response.write fld.name
'			response.write "</a>"
			response.write " <nobr><img class=""updown_sort"" src=""images/up.gif"" onclick=""changeOrderingAsc('" & fld.name & "')"" title='Ascending'> <img class=""updown_sort"" src=""images/down.gif"" onclick=""changeOrderingDesc('" & fld.name & "')"" title='Descending'></nobr>"
			response.write "</td>"
		else
			response.write "<td class=""smallerheader"">"
			response.write fld.name
			response.write "</td>"
		end if
		response.write vbCrLf

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
	response.write "<td width=10></td>"
	response.write "</tr></thead>"

	'Key Field form elements for Multiple delete
	response.write "<input type=""hidden"" name=""txtFieldName"" value=""" & sPKFieldNames & """>"
	response.write "<input type=""hidden"" name=""txtFieldType"" value=""" & sPKFieldTypes & """><tbody>" & vbCrLf

	response.write "<tr bgcolor=""#fffaf0""><td width=10></td>" & vbCrLf
	response.write "<td class=""smallertext"" nowrap><span DATAFORMATAS=html DATAFLD=""action""></span></td>" & vbCrLf

	for each fld in rs.fields
			response.write "<td class=""smallertext""><span DATAFORMATAS=html DATAFLD=""" & fld.name & """></span></td>" & vbCrLf
	next
	response.write "<td width=10></td>"

	CloseRS
%>
</tbody></table>
<table border="0" cellspacing="0" cellpadding="2" bgcolor="#ffe4b5" width="100%">
<% if bRecDel and bBulkDelete then %>
<tr bgcolor="#fffaf0">
	  <td class="smallertext"> 
		  <input type="button" name="cmdChkAll" value="Check All" class="cmdflat" onClick="checkall(true)">
		  <input type="button" name="cmdUnChkAll" value="Uncheck All" class="cmdflat" onClick="checkall(false)">
          <input type="button" name="cmdMultiDel" value="Delete Selected" class="cmdflat" onclick="MainFormAction('multidelete')">
      </td>
</tr>
<% end if 
   if (bExportExcel or bExportXML) and bAllowExport then %>
<tr bgcolor="#fffaf0">
  <td class="smallertext"> 
	<% if bExportExcel then %><button name="cmdExportExcel" class="cmdflat" onclick="MainFormAction('excel')">Export to Excel</button>&nbsp;
	<% end if
	if bExportXML then
	%>
	<button name="cmdExportXML" class="cmdflat" onclick="MainFormAction('xml')">Export to XML</button>
	<% end if %>
  </td>
</tr>
<% end if 
   check1 = inStr(replace(ucase(sU)," ",""), "MOZILLA/4.0(COMPATIBLE;MSIE") = 0
   check2 = inStr(replace(ucase(sU)," ",""), "OPERA") > 0
   if check1 OR check2 then
%>
<tr bgcolor="#fffaf0"><td align="right"><span class="smallertext">If you don't see the table, or if it is mixed, <a href="te_showtable2.asp?cid=<% =lConnID %>&tablename=<% =sTableName %>">click here</a>.</span></td>
</tr>
<% end if %>
</table>
</form>
<!--#include file="te_footer.asp"-->