<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_procmanager.asp
	' Description: View and modify stored procedures
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
	Const adClipString = 2
	
	Function GetSPText(objConn, strSPName) 
	   'On Error Resume Next 
	   Dim sql 
	   Dim rs 
	   Set rs = Server.CreateObject("ADODB.Recordset") 
	    
	   sql = "sp_helptext " & strSPName 
	   
	   Set rs = objConn.Execute(sql) 
	   GetSPText = rs.GetString(adClipString,," "," ") 
	   Set rs = Nothing 
	End Function 
		
	lConnID = request("cid")
	sProcName = request("procname")

	conn.Open arrConn(lConnID)

	action = request("action")
	if action="new" then
		SPText=""
		labelbutton="New"
		header="Write new stored procedure"
	else
		if action="alter" then
			SPText=GetSPText(Conn, sProcName)
			labelbutton="Alter"
			'button="<input type='button' name='cmdExec2' class='cmdFlat' value=' Remove ' onclick='javascript:goURL(3)'>"
			button="<input type='button' name='cmdExec2' class='cmdFlat' value=' Remove ' onclick='javascript:changeAction()'>"
			newAction="te_doproc.asp?cid=" & lConnID & "&action=remove&procname="&sProcName
			header="Modify stored procedure"
		else
			if action="remove" then
				SPText=GetSPText(Conn, sProcName)
				labelbutton="Remove"
				'button="<input type='button' name='cmdExec2' class='cmdFlat' value=' Alter ' onclick='javascript:goURL(2)'>"
				button="<input type='button' name='cmdExec2' class='cmdFlat' value=' Alter ' onclick='javascript:changeAction()'>"
				newAction="te_doproc.asp?cid=" & lConnID & "&action=alter&procname="&sProcName
				header="Remove stored procedure"
			end if
		end if
	end if	
	
	
	

	
		
					
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
<br>
<script>
function changeAction(){	
	doproc.action="<%=newAction%>"
	doproc.submit()
}

function newSP(){		
	location.href="te_procmanager.asp?cid=<% = lConnID %>&action=new"			
}

</script>
	<p class="smallerheader"><%=header%></p>
	
	<form name=doproc action="te_doproc.asp?cid=<%=lConnID%>&action=<%=action%>&procname=<%=server.htmlencode(sProcName)%>&cmdExec=1" method="post">
		<table border=0>
			<tr>
				<td>
					<textarea cols=60 rows=10 WRAP="physical" name="txtproc"><%=SPText%></textarea>
				</td>
			</tr>	
		</table>
		<% if bTableEdit then %>		
		<input type="submit" name="cmdExec" class="cmdFlat" value=" <%=labelbutton%> ">
		<%if action="alter" or action="remove" then%>			
		<% if bTableAdd then %>		
		<input type="button" name="cmdExec1" class="cmdFlat" value=" New " onclick="javascript:newSP()">					
		<%end if	%>
		<%end if	%>
		<%=button%>
		<%end if	%>
	</form>
<!--#include file="te_footer.asp"-->