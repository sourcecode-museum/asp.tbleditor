<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_count_active_users.asp
	' Description: Counts active TableEditoR users
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
<%
if not instr(request.ServerVariables("SCRIPT_NAME"), "index.asp") > 0 AND bActiveUsers then
'	ON ERROR RESUME NEXT
	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Open arrConn(0)

	'// FIND OUT WHAT PAGE THEY ARE VIEWING
	strActiveUsersPathHost = Request.ServerVariables("HTTP_HOST")
	strActiveUsersPathInfo = Request.ServerVariables("PATH_INFO")
	strUserID = session("teUserName")
	strTableName = request("tablename")

	if strUserID = "" then strUserID = "Guest"

	strActiveUsersQueryString = Request.QueryString
	
	If strActiveUsersQueryString <> "" Then
		strActiveUsersQueryString = "?" & strActiveUsersQueryString
	End If
	strActiveUsersPageViewing = "http://" & strActiveUsersPathHost & strActiveUsersPathInfo & strActiveUsersQueryString

	'// GET THE USERS IP ADDRESS
	strActiveUsersIPAddress = Request.ServerVariables("REMOTE_ADDR")

	'// ENCODE THIS INFO FOR SQL SERVER
'	strActiveUsersPageViewing = ActiveUsersSQLencode(strActiveUsersPageViewing)

	'// FIND CURRENT DATE AND TIME
	strActiveUsersCurrentDate = Now

	'// SET CHECK OUT TIME FOR 11 MINUTES OF INACTIVITY
	strActiveUsersCheckOutTime = DATEADD("s", -660, strActiveUsersCurrentDate)

	'// LETS DELETE ALL INACTIVE USERS
	strSQL = "DELETE FROM Active_Users WHERE LastCheckedIn < #" & strActiveUsersCheckOutTime & "#"
	my_Conn.Execute (strSQL)

	strSQL = "SELECT * FROM Active_Users WHERE Active_Users.UserID='" & strUserID & "'"
	set rs1 =  my_Conn.Execute (strSQL)

	if rs1.eof or rs1.bof then
		'// THEY ARE NOT IN THE DATABASE SO LETS ADD THEM
		strSQL =  "INSERT INTO Active_Users (UserID,UserIP,CheckedIn,LastCheckedIn,PageViewing) VALUES ('" & strUserID & "','"
		strSQL = strSQL & strActiveUsersIPAddress & "','" & strActiveUsersCurrentDate & "','" & strActiveUsersCurrentDate & "','" & strActiveUsersPageViewing & "')"
		my_Conn.Execute (strSQL)
	else
		'// THEY ARE IN THE DATABASE SO LETS UPDATE THERE STUFF

		'// FIRST LETS MAKE SURE THEY ARE NOT SUPOSED TO BE TIMED OUT
		strSQL = "SELECT Active_Users.LastCheckedIn "
		strSQL = strSQL & "FROM Active_Users "
		strSQL = strSQL & "WHERE Active_Users.UserID = '" & strUserID & "' "
		strSQL = strSQL & "AND Active_Users.LastCheckedIn < #" & strActiveUsersCheckOutTime & "# "
		set rs2 =  my_Conn.Execute (strSQL)

		if rs2.eof or rs2.bof then
			'// NOW LETS UPDATE THEM SINCE THEY ARE NOT SUPOSED TO BE TIMED OUT
			strSQL = "UPDATE Active_Users SET Active_Users.PageViewing='" & strActiveUsersPageViewing & "' , Active_Users.LastCheckedIn='" & strActiveUsersCurrentDate & "' WHERE Active_Users.UserID='" & strUserID & "'"
			my_Conn.Execute (strSQL)
		else
			'// DELETE THEM
			strSQL = "DELETE FROM Active_Users "
			strSQL = strSQL & "WHERE Active_Users.UserID = '" & strUserID & "' "
			strSQL = strSQL & "AND Active_Users.LastCheckedIn < #" & strActiveUsersCurrentDate & "# "
			my_Conn.Execute (strSQL)

			strSQL =  "INSERT INTO Active_Users (UserID,UserIP,CheckedIn,LastCheckedIn,PageViewing) VALUES ('" & strUserID & "','"
			strSQL = strSQL & strActiveUsersIPAddress & "','" & strActiveUsersCurrentDate & "','" & strActiveUsersCurrentDate & "','" & strActiveUsersPageViewing & "')"
			my_Conn.Execute (strSQL)
		end if
		rs2.close
		set rs2 = nothing
	end if

	rs1.close
	my_Conn.close
	set rs1 = nothing
	set my_Conn = nothing
end if
%>