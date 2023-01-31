<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_addon_logging.asp
	' Description: Log user logon information
	' Initiated By Rami Kattan on May 10, 2002
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

sub LogUserLogin(UserName)
		OpenRS2 arrConn(0)
		
		on error resume next
		sSQL = "SELECT * FROM logging"
		rs2.Open sSQL, , , adCmdTable
	
		if err <> 0 then
			response.write "Error: " & err.description & "<br><br>"
			CloseRS2
			response.end
		end if

		rs2.AddNew
		rs2("UserName") = UserName
		rs2("DateTime") = now
		rs2("IP") = Request.ServerVariables("REMOTE_HOST")
		rs2.update

		CloseRS2
end sub

%>