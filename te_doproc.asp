<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_doproc.asp
	' Description: Makes changes to a stored procedure
	' Initiated By raf on 31 Mar, 2002
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
	Dim query 
	Const adClipString = 2

	lConnID = request("cid")
	sProcName = request("procname")

	action = request("action")
	if action="new" then
		query=trim(request("txtproc"))
	else
		if action="alter" then
			query=replace(trim(request("txtproc")),"CREATE", "ALTER")				
		else
			if action="remove" then
				query="DROP PROCEDURE dbo." & sProcName
			end if
		end if
	end if	
	
	conn.Open arrConn(lConnID)	
	
  	Conn.Execute(query) 

	if action="new" then		
		proctxtArray=split(query," ",-1,1)
		sProcName=trim(proctxtArray(2))
'response.write sProcName	
		query="SELECT CURRENT_USER as currUser"
		set RS=Conn.Execute(query) 

		query="EXEC sp_changeobjectowner '" & trim(RS("currUser")) & "." & mid(sProcName,1,len(sProcName)-2) & "', 'dbo'"
'response.write query		
  		Conn.Execute(query) 
  	end if
  		
  	conn.close

%>
<!--#include file="te_header.asp"-->
<table border=0 cellspacing=1 cellpadding=2 bgcolor = "#ffe4b5" width="100%">
	<tr>
		<td class="smallertext">
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <a href="te_listtables.asp?cid=<%=request("cid")%>"><%=arrDesc(request("cid"))%></a> » execute procedure
		</td>
		<td class="smallerheader" width=130 align=right>
		<%
		if bProtected then 
			response.write session("teFullName")
			response.write "<a href=""te_logout.asp""> (logout)</a>" 
		end if
		%>
		</td>
	</tr>
</table>
<!--#include file="te_footer.asp"-->