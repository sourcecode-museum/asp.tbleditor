<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_showschema.asp
	' Description: Displays the database schema
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
	' # May 30, 2002 by Rami Kattan
	' Table format is like other TE tables :)
	'==============================================================
%>
<!--#include file="te_config.asp"-->
<!--#include file="te_header.asp"-->
<table border=0 cellspacing=1 cellpadding=2 bgcolor = "#ffe4b5" width=100%>
	<tr>
		<td class="smallertext">
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » Schema List
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
<strong>Schema : <% =sTitle %></strong><br><br>
<table border=0 cellspacing=1 cellpadding=2 bgcolor="#ffe4b5" width="100%"<% if bTableHighlight then response.write " style=""behavior:url(tablehl.htc);"" slcolor='#ffffcc' hlcolor='#bec5de'"%>>
<thead>
<%

	lConnID = request.querystring("cid")

	OpenRS arrConn(lConnID)

	sub ShowSchema(sTitle, iSchema)
		set rsX = conn.openSchema(iSchema)
		
'		response.write "<table border=0>"
		response.write "<tr>"
		for each fld in rsX.fields
			response.write "<td class=""smallerheader"">" & fld.name & "</td>"
		next
		response.write "</tr></thead><tbody>"
	
		do while not rsX.eof
			response.write "<tr bgcolor=""#fffaf0"">"
			for each fld in rsX.fields
				response.write "<td class=""smallertext"">" & rsX(fld.name) & "&nbsp;</td>"
			next
			response.write "</tr>"
			rsX.MoveNext
		loop
		response.write "</tbody></table><br><br><br>"
		
		rsX.close
	end sub
	
	ShowSchema "Tables", 20
%>
<!--#include file="te_footer.asp"-->