<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_view_active_users.asp
	' Description: Displays active TableEditoR users
	' Initiated By Rami Kattan on Apr 10, 2002
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
	if not bAdmin then	
		response.write "Not authorized to view this page."
		%><!--#include file="te_footer.asp"--><%
		response.end
	end if

set my_Conn = Server.CreateObject("ADODB.Connection")
my_Conn.Open arrConn(0)

nRefreshTime = Request.Form("RefreshTime")
%>
<table border=0 cellspacing=1 cellpadding=2 bgcolor = "#ffe4b5" width="100%">
	<tr>
		<td class="smallertext">
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » Active Users
		</td>
		<td class="smallerheader" width="130" align="right">
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
<script language="JavaScript" type="text/javascript">
<!--
function autoReload()
{
	GetObject("ReloadFrm").submit()
}
//-->
</script>

<div align="center">
<table cellpadding="0" cellspacing="0" border="0" width="90%">
<form name="ReloadFrm" action="te_view_active_users.asp" method="post">
<tr height="25">
	<td bgcolor="#ffe4b5" width="4"></td>
	<td bgcolor="#ffe4b5" nowrap class="smallerheader">»&nbsp;&nbsp;Active
      Users &nbsp;</td>
	<td bgcolor="#ffe4b5" align="right" class="smallerheader">
    <select name="RefreshTime" size="1" onchange="autoReload();" class="smallertext">
        <option value="0"  <% if nRefreshTime = "0" then Response.Write(" SELECTED")%>>Don't reload automatically</option>
        <option value="1"  <% if nRefreshTime = "1" then Response.Write(" SELECTED")%>>Reload page every minute</option>
        <option value="5"  <% if nRefreshTime = "5" then Response.Write(" SELECTED")%>>Reload page every 5 minutes</option>
        <option value="10" <% if nRefreshTime = "10" then Response.Write(" SELECTED")%>>Reload page every 10 minutes</option>
        <option value="15" <% if nRefreshTime = "15" then Response.Write(" SELECTED")%>>Reload page every 15 minutes</option>
        <option value="30" <% if nRefreshTime = "30" then Response.Write(" SELECTED")%>>Reload page every 30 minutes</option>
    </select>
	</td>
	<td bgcolor="#ffe4b5" width="4">&nbsp;</td>
</tr>
</form>
<tr>
	<td bgcolor="#FFFFFF" colspan="4">
<script language="JavaScript" type="text/javascript">
<!--
if (GetObject("RefreshTime").options[GetObject("RefreshTime").selectedIndex].value > 0) {
	reloadTime = 60000 * GetObject("RefreshTime").options[GetObject("RefreshTime").selectedIndex].value
	self.setInterval('autoReload()', 60000 * GetObject("RefreshTime").options[GetObject("RefreshTime").selectedIndex].value)
}
//-->
</script>
	</td>
</tr>
<tr>
	<td bgcolor="#fffaf0" colspan="4">
      <div align="left">
        <table border=0 cellspacing=1 cellpadding=5 bgcolor="#ffe4b5" width="100%" <% if bTableHighlight then response.write " style=""behavior:url(tablehl.htc);"" slcolor='#ffffcc' hlcolor='#bec5de'"%>>
		<thead><tr>
			<td class="smallerheader">User</td>
			<td class="smallerheader">Page Viewing</td>
			<td class="smallerheader">Logged On</td>
			<td class="smallerheader">Last Active</td>
			<td class="smallerheader">Online Time</td>
		</tr></thead>
		<tbody>
<%	
	if not bActiveUsers then %>
          <tr>
            <td bgcolor="#fffaf0" class="smallertext" colspan="5" align="center"><b>Active Users not enabled...</b></td>
          </tr>
		  		</tbody>
		</table>
      </div>
    </td>
</tr>
</table>
</div>
<!--#include file="te_footer.asp"-->
<% 
	response.end
end if 


strSQL = "SELECT * FROM Active_Users"
set rs =  my_Conn.Execute (strSQL)
if rs.EOF or rs.BOF then
%>
          <tr>
            <td bgcolor="#fffaf0" class="smallertext" colspan="5">No active users found</td>
          </tr>
<%
else
	strActiveUsersRecordCount = 1
	strActiveUsersI = 0
	do until rs.eof
	if strActiveUsersI = 0 then
		strActiveUsersCColor = "#fffaf0"
	else
		strActiveUsersCColor = "#fffaf0"
	end if

	strActiveUsersUserID = rs("UserID")
	strActiveUsersUserIP = rs("UserIP")
	strActiveUsersPageViewing = rs("PageViewing")
	strActiveUsersCheckedIn = rs("CheckedIn")
	strActiveUsersLastCheckedIn = rs("LastCheckedIn")

	strActiveUsersTotalOnlineTime = datediff("n", strActiveUsersCheckedIn, NOW)

	If strActiveUsersTotalOnlineTime > 60 then
		' they must have been online for like an hour or so.
		strActiveUsersTotalOnlineHours = 0
		do until strActiveUsersTotalOnlineTime < 60
			strActiveUsersTotalOnlineTime = (strActiveUsersTotalOnlineTime - 60)
			strActiveUsersTotalOnlineHours = strActiveUsersTotalOnlineHours + 1
		loop
		strActiveUsersTotalOnlineTime = strActiveUsersTotalOnlineHours & " Hours " & strActiveUsersTotalOnlineTime & " Minutes"
	Else
		strActiveUsersTotalOnlineTime = strActiveUsersTotalOnlineTime & " Minutes"
	End If

	strActiveUsersPageViewingOriginal = strActiveUsersPageViewing

	strActiveUsersHTTPHost = Request.ServerVariables("HTTP_HOST")
	strActiveUsersHTTPHost = "http://" & strActiveUsersHTTPHost
	strActiveUsersPathLength = Len(strActiveUsersPageViewing) - (Len(strActiveUsersHTTPHost))
	If instr(strActiveUsersPageViewing, strActiveUsersHTTPHost) Then
		strActiveUsersPageViewing = Right(strActiveUsersPageViewing, strActiveUsersPathLength)
		If Left(strActiveUsersPageViewing, 1) = "/" Then
			strActiveUsersPageViewing = Right(strActiveUsersPageViewing, (Len(strActiveUsersPageViewing) - 1))
		End If
	End If
%>
          <tr>
            <td bgcolor="<%=strActiveUsersCColor%>" class="smallertext"><span title="IP: <% =strActiveUsersUserIP %>"><%=strActiveUsersUserID%></span></td>
            <td bgcolor="<%=strActiveUsersCColor%>" class="smallertext"><a title="<%=strActiveUsersPageViewingOriginal%>" href="<%=strActiveUsersPageViewingOriginal%>" target="_blank"><%=strActiveUsersPageViewing%></a></td>
            <td bgcolor="<%=strActiveUsersCColor%>" class="smallertext"><%=strActiveUsersCheckedIn%></td>
            <td bgcolor="<%=strActiveUsersCColor%>" class="smallertext"><%=strActiveUsersLastCheckedIn%></td>
            <td bgcolor="<%=strActiveUsersCColor%>" class="smallertext"><%=strActiveUsersTotalOnlineTime%></td>
          </tr>
<%
	strActiveUsersRecordCount = strActiveUsersRecordCount + 1
	strActiveUsersI  = strActiveUsersI + 1
	if strActiveUsersI = 2 then
		strActiveUsersI = 0
	end if
	rs.MoveNext
	loop
end if

rs.close
set rs = nothing

'// FIND CURRENT DATE AND TIME
strActiveUsersCurrentDate = Now

'// SET CHECK OUT TIME FOR 11 MINUTES OF INACTIVITY
strActiveUsersCheckOutTime = DATEADD("s", -660, strActiveUsersCurrentDate)

'strSQL = "DELETE FROM Active_Users WHERE LastCheckedIn < '" & strActiveUsersCheckOutTime & "'"
'my_Conn.Execute (strSQL)

my_conn.close
set my_conn = nothing

%>
		</tbody>
		</table>
      </div>
    </td>
</tr>
</table>
</div>
<!--#include file="te_footer.asp"-->