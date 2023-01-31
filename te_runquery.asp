<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_runquery.asp
	' Description: Executes a query
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
	' Mar 31, 2001 by Hakan Eskici
	' Added security check
	'==============================================================
%>
<!--#include file="te_config.asp"-->
<%

	if not bSQLExec then
		response.redirect "index.asp"
	end if

	lConnID = request("cid")
	sSQL = request("txtSQL")

	'Query returns records?
	if request("chkRec") = "" then
		bRecordReturn = False
	else
		bRecordReturn = True
		response.redirect TableViewerCompat & "?cid=" & lConnID & "&q=1&tablename=" & server.urlencode(sSQL)
	end if

%>
<!--#include file="te_header.asp"-->

<table border=0 cellspacing=1 cellpadding=2 bgcolor = "#ffe4b5" width=100%>
	<tr>
		<td class="smallertext">
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <a href="te_listtables.asp?cid=<%=request("cid")%>"><%=arrDesc(request("cid"))%></a> » Run Query
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
	response.write "<p class=""smallerheader"">Executing Query:</p>"
	response.write "<p class=""smallertext"">" & sSQL & "</p>"
	
	
	on error resume next

	conn.Open arrConn(lConnID)
	set rs = conn.Execute (sSQL, lRecordsAffected)
	
	if Err <> 0 then
		response.write "<p class=""smallerheader"">Error:</p>"
		response.write err.description
	end if
	
	if not bRecordReturn then
		'If the query doesnt return records;
		response.write lRecordsAffected & " records affected."
	end if

%>
<!--#include file="te_footer.asp"-->