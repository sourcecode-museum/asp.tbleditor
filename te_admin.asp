<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_admin.asp
	' Description: Lists defined connections
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
	' # Mar 26, 2001 by Hakan Eskici
	' Added displaying table, view and procedure counts
	' # April 18, 2002 by Rami Kattan
	' Dynamic configurations, highlighting and more.
	'==============================================================
%>
<!--#include file="te_config.asp"-->
<!--#include file="te_header.asp"-->
<%
	sub GetDBDetails (sConnStr, ByRef sTableCount, ByRef sViewCount, ByRef sProcCount)
		on error resume next

		conn.Open sConnStr
		Set adox = Server.CreateObject("ADOX.Catalog") 
		adox.ActiveConnection = conn

		sTableCount = adox.tables.count
		sViewCount = adox.views.count
		if err <> 0 then sViewCount = "n/a"
		sProcCount = adox.procedures.count
		conn.close
	
	end sub
%>
<script language="JavaScript" type="text/javascript">
<!--
function openWindow(url) {
  popupWin = window.open(url,'new_page','width=540,height=600,resizable=yes')
}
//-->
</script>
<table border=0 cellspacing=1 cellpadding=3 bgcolor="#ffe4b5" width="100%">
	<tr>
		<td class="smallertext">
			<a href="index.asp">Home</a> » Connections
		</td>
		<td class="smallerheader" width=130 align=right>
		<%
		if bProtected then 
			if session("teLastAccess") <> "" then
				response.write "<span title=""User last access was on '" & formatdatetime(session("teLastAccess"), 1) & "'" & vbCrLf & "at '"& formatdatetime(session("teLastAccess"), 3) & "'"">"
			else
				response.write "<span title=""First time access, Welcome."">"
			end if
			response.write session("teFullName") & "</span>"
			response.write " (<a href=""te_logout.asp"">logout</a>)"
		end if
		%>
		</td>
	</tr>
</table>
<noscript>
<div align="center"><h2><font color="#FF0000"><i>Highly</i> Recommended having <b><u>JavaScript</u></b> enabled</font></h2></div>
</noscript>

<p class="smallerheader">
Select a connection <nobr>
<% if bAdmin then %>
<script language="JavaScript" type="text/javascript">
<!--
function goto_admin(){
	var URL = GetObject("AdminWhat").options[GetObject("AdminWhat").selectedIndex].value;
	if (URL != "") location.href = URL;
}
//-->
</script>
<% if bJSEnable and bPopUps then %>
or <select id="AdminWhat" name="AdminWhat" onchange="goto_admin()" class="smallbutton">
		<option value="">--- Select ---</option>
		<option value="<% =TableViewerCompat %>?cid=0&amp;tablename=Users">Administer Table Editor Users</option>
		<option value="te_dynamic_config.asp">Change Table Editor Configurations</option>
		<option value="<% =TableViewerCompat %>?cid=0&amp;tablename=Databases">Administer Database Connections</option>
		<% if bUserLogging then %>
		<option value="<% =TableViewerCompat %>?cid=0&amp;tablename=Logging">View Login Logs</option>
		<%	end if
			if bActiveUsers then %>
		<option value="te_view_active_users.asp">View Active Users</option>
		<% end if 
		   if bBulkCompact then %>
			<option value="">--------------------</option>
			<option value="te_compactall.asp?onlybackup=true">Backup All Databases</option>
			<option value="te_compactall.asp">Compact All Databases</option>
		<% end if %>
	</select><input type="button" name="AdminWhatGo" value="»" onclick="goto_admin()" class="smallbutton">
<% else ' if JS enabled %>
or <a href="te_admin_options.asp">Configure Table Editor</a>
<% end if ' if JS enabled %>
<% end if ' if bAdmin %>
</p>

<table border=0 cellspacing=1 cellpadding=2 bgcolor = "#ffe4b5" width="100%"<% if bTableHighlight then response.write " style=""behavior:url(tablehl.htc);"" slcolor='#ffffcc' hlcolor='#bec5de'"%>>
<thead>
	<tr>
		<td class="smallerheader" bgcolor="#ffe4b5" width=10>»</td>
		<td bgcolor="#ffe4b5" class="smallerheader">Connections</td>
	</tr>
	<tr>
		<td class="smallerheader" bgcolor="#fffaf0" width=10></td>
		<td class="smallerheader" bgcolor="#fffaf0">Database Name</td>
		<%
		if bShowConnDetails = True then
		%>
		<td class="smallerheader" bgcolor="#fffaf0" width="60">Tables</td>
		<td class="smallerheader" bgcolor="#fffaf0" width="60">Views</td>
		<td class="smallerheader" bgcolor="#fffaf0" width="60">Procs</td>
		<%
		end if
		%>
		<td class="smallerheader" bgcolor="#fffaf0" width="100">Type</td>
		<td bgcolor="#fffaf0" width=160 align="right" class="smallerheader">Action</td>
		<td bgcolor="#fffaf0" width=10></td>
	</tr>
</thead>
<tbody>
<%
for i = 1 to ubound(arrConn)
	select case arrType(i)
		case tedbAccess
			sConnType = "Access"
		case tedbSQLServer
			sConnType = "SQL Server"
		case tedbDsn
			sConnType = "DSN"
		case tedbConnStr
			sConnType = "ConnStr"
	end select
%>
	<tr bgcolor="#fffaf0">
	<td></td>
	<td>
	<a href="te_listtables.asp?cid=<%=i%>"><%=arrDesc(i)%></a>
	</td>
	<%
	if bShowConnDetails = True then
		GetDBDetails arrConn(i), sTableCount, sViewCount, sProcCount
	%>
	<td class="smallertext"><%=sTableCount%></td>
	<td class="smallertext"><%=sViewCount%></td>
	<td class="smallertext"><%=sProcCount%></td>
	<%
	end if
	%>
	<td class="smallertext"><%=sConnType%></td>
	<td width=160 align="right">
	<%
	if bTableEdit and arrType(i) = tedbAccess then
		%>
		<a href="te_compactdb.asp?cid=<%=i%>">compact</a>
		<a href="te_compactdb.asp?cid=<%=i%>&amp;onlybackup=true">backup</a>
		<%
	end if
	if bAdmin then
	%>
	<a href="te_showschema.asp?cid=<%=i%>">schema</a>
	<%
	end if
	%>
	</td>
	<td></td></tr>
<%
next
%>
</tbody>
</table>
<!--#include file="te_footer.asp"-->