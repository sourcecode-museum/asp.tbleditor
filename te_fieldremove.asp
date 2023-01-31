<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_fieldremove.asp
	' Description: Removes a field from the table
	' Initiated By Hakan Eskici on Nov 17, 2000
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
<!--#include file="te_config.asp"-->
<%

	lConnID = request("cid")
	sTableName = request("tablename")
	sFieldName = request("fldname")
	
	OpenRS arrConn(lConnID)

	'Check if the user approved delete and has permissions to delete
	if (request("sure") <> "") and bFldDel then

		'Delete the indexes if any
		set rs = conn.openSchema(adSchemaIndexes)
		do while not rs.eof
			if rs("table_name") = sTableName then
				if (rs("column_name") = sFieldName) then
					sSQL = "ALTER TABLE [" & sTableName & "] DROP CONSTRAINT [" & rs("index_name") & "]"
					conn.execute sSQL
				end if
			end if
			rs.movenext
		loop

		sSQL = "ALTER TABLE [" & sTableName & "] DROP COLUMN [" & sFieldName & "]"
		conn.execute sSQL
		
		response.redirect "te_tableedit.asp?cid=" & lConnID & "&tablename=" & sTableName
	end if

	'Check if the field is indexed
	set rs = conn.openSchema(adSchemaIndexes)
	do while not rs.eof
		if rs("table_name") = sTableName then
			if (rs("column_name") = sFieldName) then
				sErr = "Field '" & sFieldName & "' is indexed. If you delete this field, all related indices will also be deleted.<br><br>"
			end if
		end if
		rs.movenext
	loop
	
	CloseRS

%>
<!--#include file="te_header.asp"-->

<table border=0 cellspacing=1 cellpadding=2 bgcolor = "#ffe4b5" width=100%>
	<tr>
		<td class="smallertext">
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <a href="te_listtables.asp?cid=<%=request("cid")%>"><%=arrDesc(request("cid"))%></a> » <a href="te_tableedit.asp?cid=<%=request("cid")%>&tablename=<%=server.urlencode(request("tablename"))%>">Edit Table [<%=request("tablename")%>]</a> » Remove Field
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
<% if bFldDel then %>
			<p class="smallerheader"><%=sErr%>Are you sure that you want to delete the record?</p>
			<a href="te_fieldremove.asp?<%=request.querystring%>&sure=1">Yes</a>&nbsp;
			<a href="<%=request.servervariables("http_referer")%>">No</a>
<% else %>
			<p class="smallerheader">You have no permission to delete fields.</p>
<% end if %>
<!--#include file="te_footer.asp"-->