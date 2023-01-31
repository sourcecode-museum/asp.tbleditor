<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_xml_import.asp
	' Description: Import XML data into tables
	' Initiated By Rami Kattan on Jun 05, 2002
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
	' # June 5, 2002 by Rami Kattan
	' Under Development.
	'==============================================================
%>
<!--#include file="te_config.asp"-->
<!--#include file="te_header.asp"-->
<%

lConnID = request("cid")

if lConnID=0 and not bAdmin then
	response.redirect "te_admin.asp"
end if
%>
	<br><br>
	<table border=0 cellspacing=1 cellpadding=2 bgcolor="#ffe4b5" width="100%">
		<tr>
			<td class="smallertext">
				<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <% allTablesCombo() %> » Import XML Data</td>
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
	<br /><br /><br />
	<h1 align="center">Under Development !!!</h1>
	<h3 align="center">Well be available as add-on on <a style=" font-size:18" href="http://www.2enetworx.com/dev/projects/uploads.asp?pid=2">User Uploads</a> </h3>
	<br /><br /><br />
<!--#include file="te_footer.asp"-->