<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_tableremove.asp
	' Description: Drops the given table
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
	
	OpenRS arrConn(lConnID)

	'Check if the user approved delete and has permissions to delete
	if (request("sure") <> "") and bFldDel then
		sSQL = "DROP TABLE [" & sTableName & "]"
		conn.execute sSQL
		response.redirect "te_listtables.asp?cid=" & lConnID
	end if

	conn.close
	set rs=nothing
	set conn=nothing

%>
<!--#include file="te_header.asp"-->

<table border=0 cellspacing=1 cellpadding=2 bgcolor = "#ffe4b5" width=100%>
	<tr>
		<td class="smallertext">
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <a href="te_listtables.asp?cid=<%=request("cid")%>"><%=arrDesc(request("cid"))%></a> » Remove Table
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

	<p class="smallerheader"><%=sErr%>Are you sure that you want to remove the table "<%=sTableName%>"?</p>
	<a href="te_tableremove.asp?<%=request.querystring%>&sure=1">Yes</a>&nbsp;
	<a href="<%=request.servervariables("http_referer")%>">No</a>

<!--#include file="te_footer.asp"-->