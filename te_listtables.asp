<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_listtables.asp
	' Description: Lists the tables for a given connection
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
	' # Nov 14, 2000 by Hakan Eskici
	' Depending on the permissions, users will either see "info"
	' or "edit" for the action with a link to te_tableedit.asp
	' # Nov 17, 2000 by Hakan Eskici
	' Added drop table links for tables and queries
	' # Mar 27, 2001 by Hakan Eskici
	' Added showing stored procedures
	' # Mar 29, 2001 by Hakan Eskici
	' Fixed the cid=0 security bug
	' # May 28, 2001 by Alain Jacquot
	' Fixed the "Object or provider is not capable of performing requested operation" bug
	' # June 03, 2001 by Rakesh Jain
	' Rearranged the code so that system tables and data tables do not show under Queries
	'==============================================================
%>
<!--#include file="te_config.asp"-->
<!--#include file="te_header.asp"-->
<table border=0 cellspacing=1 cellpadding=2 bgcolor="#ffe4b5" width="100%">
	<tr>
		<td class="smallertext">
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <% 
				if bComboTables then
					allTablesCombo() 
				else
					response.write arrDesc(request("cid"))
				end if %></td>
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

<p class="smallerheader">
	Select a table
	<%
	if bSQLExec then
		response.write " or <a href=""te_query.asp?cid=" & request("cid") & """>Run a Query</a>"
	end if
	%>
</p>

<%

'Which connection to use?
lConnID = request("cid")

if (lConnID = 0) and bAdmin = False then
	response.write "Not authorised to view this connection."
	%><!--#include file="te_footer.asp"--><%
	response.end
end if

conn.Open arrConn(lConnID)

set rs = conn.OpenSchema(adSchemaTables)

%>

<table border=0 cellspacing=1 cellpadding=2 bgcolor = "#ffe4b5" width="100%"<% if bTableHighlight then response.write " style=""behavior:url(tablehl.htc);"" slcolor='#ffffcc' hlcolor='#bec5de'"%>>
<thead>
	<tr>
		<td class="smallerheader" bgcolor="#ffe4b5" width=10>»</td>
		<td bgcolor="#ffe4b5" class="smallerheader" colspan=3>
			Tables
			<%
			if bTableAdd then
				response.write " &nbsp;&nbsp;&nbsp;(<a href=""te_tablecreate.asp?cid=" & request("cid") & """>Create a new table</a>)"
			end if
			if bAllowImport then
				response.write " &nbsp;&nbsp;&nbsp;(<a href=""te_xml_import.asp?cid=" & request("cid") & """>Import Data</a>)"
			end if
			%>	
		</td>
	</tr>
	<tr>
		<td class="smallerheader" bgcolor="#fffaf0" width=10></td>
		<td class="smallerheader" bgcolor="#fffaf0">Table Name</td>
		<td class="smallerheader" bgcolor="#fffaf0">Description</td>
		<td class="smallerheader" bgcolor="#fffaf0" align=right width=150>Action</td>
		<td bgcolor="#fffaf0" width=10></td>
	</tr>
</thead>
<tbody>
<%
	'Addition: Search Link
	'By Kevin Yochum on Nov 07, 2000

	if bTableEdit then sAction = "edit" else sAction = "info"

	do while not rs.eof
		if rs("table_type") = "TABLE" then
			%>
			<tr bgcolor="#fffaf0">
				<td></td>
				<td><a href="<% =TableViewerCompat %>?cid=<%=lConnID%>&tablename=<%=rs("table_name")%>"><%=rs("table_name")%></a></td>
				<td class="smallertext"><%=rs("descrIptIon")%></td>
				<td align=right>
				<a href="te_searchtable.asp?cid=<%=lConnID%>&tablename=<%=rs("table_name")%>">search</a>
				<a href="te_tableedit.asp?cid=<%=lConnID%>&tablename=<%=rs("table_name")%>&type=table"><%=sAction%></a>
				<% if bTableDel then %>
				<a href="te_tableremove.asp?cid=<%=lConnID%>&tablename=<%=rs("table_name")%>">remove</a>
				<% end if %>
				</td>
				<td></td>
			</tr>
			<%
		end if
		rs.movenext
	loop
	rs.close
	
%>
</tbody>
</table>

<%	
	if bQueryExec then
%>
<br>
<table border=0 cellspacing=1 cellpadding=2 bgcolor="#ffe4b5" width="100%"<% if bTableHighlight then response.write " style=""behavior:url(tablehl.htc);"" slcolor='#ffffcc' hlcolor='#bec5de'"%>>
<thead>
	<tr>
		<td class="smallerheader" bgcolor="#ffe4b5" width=10>»</td>
		<td bgcolor="#ffe4b5" class="smallerheader" colspan=3>
			Queries
			<%
			if bTableAdd then
				response.write " &nbsp;&nbsp;&nbsp;(<a href=""te_queryedit.asp?cid=" & request("cid") & "&add=1"">Create a new query</a>)"
			end if
			%>	
		</td>
	</tr>
	<tr>
		<td class="smallerheader" bgcolor="#fffaf0" width=10></td>
		<td class="smallerheader" bgcolor="#fffaf0">Query Name</td>
		<td class="smallerheader" bgcolor="#fffaf0">Description</td>
		<td class="smallerheader" bgcolor="#fffaf0" width=150 align=right>Action</td>
		<td bgcolor="#fffaf0" width=10></td>
	</tr>
</thead>
<tbody>
<%
	
	' # May 28, 2001 by Alain Jacquot
	' Fixed the "Object or provider is not capable of performing requested operation" bug
		
	' # June 03, 2001 by Rakesh Jain
	' Rearranged the code so that system tables and data tables do not show under Queries
	
	select case arrType(lConnID)
		Case tedbAccess
			set rs = conn.OpenSchema(adSchemaViews)
		Case tedbSQLServer
			set rs = conn.OpenSchema(adSchemaTables)	
		Case tedbDSN
			set rs = conn.OpenSchema(adSchemaTables)
		Case tedbConnStr
			set rs = conn.OpenSchema(adSchemaViews)
		Case else
			if (lConnID = 0) and bAdmin = True then
				set rs = conn.OpenSchema(adSchemaTables)	
			else
				response.write "Not authorised to view this connection."
				%><!--#include file="te_footer.asp"--><%
				response.end
	end if
	end select
	
	do while not rs.eof
		if arrType(lConnID) = tedbAccess then
			bProceed = True
		else
			if rs("table_type") = "VIEW" then
				bProceed = True
			end if
		end if
		
		if bProceed then
			%>
			<tr bgcolor="#fffaf0">
				<td></td>
				<td><a href="<% =TableViewerCompat %>?cid=<%=lConnID%>&tablename=<%=rs("table_name")%>"><%=rs("table_name")%></a></td>
				<td class="smallertext"><%=rs("descrIptIon")%></td>
				<td align=right>
				<% if arrType(lConnID) = tedbAccess then %>
				<a href="te_queryinfo.asp?cid=<%=lConnID%>&tablename=<%=rs("table_name")%>&type=query">info</a>
				<% end if %>
				<% if bTableEdit and arrType(lConnID) = tedbAccess then %>
				<a href="te_queryedit.asp?cid=<%=lConnID%>&queryname=<%=rs("table_name")%>">edit</a>
				<% end if %>
				<% if bTableDel then %>
				<a href="te_tableremove.asp?cid=<%=lConnID%>&tablename=<%=rs("table_name")%>">remove</a>
				<% end if %>
				</td>
				<td></td>
			</tr>
			<%
		end if
		rs.movenext
	loop
	rs.close
	response.write "</tbody></table>"
end if	

set rs = conn.OpenSchema(adSchemaProcedures)

%>
<br>
<table border=0 cellspacing=1 cellpadding=2 bgcolor = "#ffe4b5" width="100%"<% if bTableHighlight then response.write " style=""behavior:url(tablehl.htc);"" slcolor='#ffffcc' hlcolor='#bec5de'"%>>
<thead>
	<tr>
		<td class="smallerheader" bgcolor="#ffe4b5" width=10>»</td>
		<td bgcolor="#ffe4b5" class="smallerheader" colspan=4>
			Procedures
			<%
			if bTableAdd then
				response.write " &nbsp;&nbsp;&nbsp;(<a href=""te_procedit.asp?add=1&cid=" & lConnID & """>Create a new stored procedure</a>)"
			end if
			%>	
		</td>
	</tr>
	<tr>
		<td class="smallerheader" bgcolor="#fffaf0" width=10></td>
		<td class="smallerheader" bgcolor="#fffaf0">Name</td>
		<td class="smallerheader" bgcolor="#fffaf0">Parameters</td>
		<td class="smallerheader" bgcolor="#fffaf0">Description</td>
		<td class="smallerheader" bgcolor="#fffaf0" width=150 align=right>Action</td>
		<td bgcolor="#fffaf0" width=10></td>
	</tr>
</thead>
<tbody>
<%
	set cat = server.createobject("adox.catalog")
	cat.ActiveConnection = conn
	
	do while not rs.eof
		sProcName = rs("procedure_name")
		sProcDef = rs("descrIptIon")
		iProcType = rs("procedure_type")
		select case arrType(lConnID) 
			case tedbSQLServer
				aProcName = split(sProcName, ";")
				iProcType = 0
				if isArray(aProcName) then
					sProcName = aProcName(0)
					iParamCount = aProcName(1)
				end if
			case tedbAccess
				iParamCount = cat.Procedures(sProcName).command.parameters.count
			case else
				iParamCount = 0
				iProcType = 0
		end select

		'if (iProcType = 3) and arrType(lConnID)<>tedbSQLServer then
		'	sProcURL = "<a href=""te_execproc.asp?cid=" & lConnID & "&procname=" & server.htmlencode(sProcName) & """>" & sProcName & "</a>"
		'else
		'	sProcURL = sProcName & "(" & sProcType & ")"
		'end if
		
		select case iProcType
			case 2	'Non-parameterized query
				if iParamCount > 0 then
					sProcURL = "<a href=""te_execproc.asp?cid=" & lConnID & "&procname=" & server.htmlencode(sProcName) & """>" & sProcName & "</a>"
				else
					sProcURL = "<a href=""" & TableViewerCompat & "?cid=" & lConnID & "&tablename=" & server.htmlencode(sProcName) & """>" & sProcName & "</a>"
				end if
			case 3	'Query with parameters
				sProcURL = "<a href=""te_execproc.asp?cid=" & lConnID & "&procname=" & server.htmlencode(sProcName) & """>" & sProcName & "</a>"
			case else
				sProcURL = "<a href=""te_procmanager.asp?cid=" & lConnID & "&action=alter&procname=" & server.htmlencode(sProcName) & """>" & sProcName & "</a>"
				'sProcURL = sProcName
		end select

		%>
			<tr bgcolor="#fffaf0">
				<td></td>
				<td class="smallertext"><%=sProcURL%></td>
				<td class="smallertext"><%=iParamCount%></td>
				<td class="smallertext"><%=sProcDef%></td>
				<td align=right>
				<% if arrType(lConnID) = tedbAccess then %>
				<a href="te_queryinfo.asp?cid=<%=lConnID%>&tablename=<%=server.htmlencode(sProcName)%>&type=proc">info</a>
				<% end if %>
				<% if bTableEdit and arrType(lConnID) = tedbAccess then %>
				<a href="te_procedit.asp?cid=<%=lConnID%>&queryname=<%=server.htmlencode(sProcName)%>">edit</a>
				<% end if %>
				<% if bTableDel then %>
				<a href="te_tableremove.asp?cid=<%=lConnID%>&tablename=<%=server.htmlencode(sProcName)%>">remove</a>
				<% end if %>
				</td>
				<td></td>
			</tr>
		<%
		rs.movenext
	loop
%>
</tbody>
</table>
<!--#include file="te_footer.asp"-->