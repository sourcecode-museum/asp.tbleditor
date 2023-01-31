<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_admin_options.asp
	' Description: Administrate TE for non javascript browsers
	' Initiated By Rami Kattan on May 20, 2002
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
	'==============================================================
%>
<% Response.Buffer = true %>
<!--#include file="te_config.asp"-->
<!--#include file="te_header.asp"-->
<%
	if bAdmin = False then	
		response.write "Not authorized to view this page."
		%><!--#include file="te_footer.asp"--><%
		response.end
	end if
%>
<table border=0 cellspacing=1 cellpadding=3 bgcolor="#ffe4b5" width="100%">
	<tr>
		<td class="smallertext">
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » Table Editor Administration
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
<ul>
	<li><a href="<% =TableViewerCompat %>?cid=0&amp;tablename=Users">Administer Table Editor Users</a></li>
	<li><a href="te_dynamic_config.asp?popup=no">Change Table Editor Configurations</a></li>
	<li><a href="<% =TableViewerCompat %>?cid=0&amp;tablename=Databases">Administer Database Connections</a></li>
<% if bUserLogging then %>
	<li><a href="<% =TableViewerCompat %>?cid=0&amp;tablename=Logging">View Login Logs</a></li>
<% end if 
   if bActiveUsers then %>
	<li><a href="te_view_active_users.asp">View Active Users</a></li>
<% end if 
   if bBulkCompact then %>
	<li><a href="te_compactall.asp?onlybackup=true">Backup All Databases</a></li>
	<li><a href="te_compactall.asp">Compact All Databases</a></li>
<% end if %>
</ul><!--#include file="te_footer.asp"-->