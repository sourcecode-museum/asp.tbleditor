<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: index.asp
	' Description: Default page for TableEditoR
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
	' # Nov 15, 2000 by Hakan Eskici
	' Added permission assignment for Table and Field functions	
	' # May 05, 2002 by Rami Kattan
	' Record in users table user's last access
	' # May 14, 2002 by Rami Kattan
	' Check if browser can execute javascripts
	' # May 30, 2002 by Rami Kattan
	' Remember Me option to remember the last user name
	' High Security login method
	'==============================================================

	dim bOnPageLoad  ' to tell header to add onload code to body tag or not.
	bOnPageLoad = true
	bOnPageLoadPage = "index"

%>
<!--#include file="te_config.asp"-->
<!--#include file="te_addon_logging.asp"-->
<%
'We may come here as a result of a session timeout
'or direct access without opening a session
'redirection will occur after login
if request("comebackto") <> "" then
	sReferer = request("comebackto")
	sGoBackTo = "?" & request.querystring
end if

if RequireSecurityID then session("teUserName") = ""

sub AskForLogin(sText)
%>
<p class="smallheader">
	<%
		sMsg = sText
		if request("comebackto") <> "" then
			sMsg = sMsg & "<br>Please re-login."
			if sText <> "" then
				%><!--#include file="te_header.asp"--><%
			end if
		else
			if sText = "" then 
				sMsg = sMsg & "<br>Please login."
			else
				%><!--#include file="te_header.asp"--><%
			end if
		end if
		response.write "<p class=""smallheader""><br>" & sMsg

		LastUserName = Request.Cookies("TableEditor8")("UserName")
		LastRemember = (Request.Cookies("TableEditor8")("RememberMe") = "true")

		if LastRemember then RememberCheck = " checked"

	%>
</p>

<script language="JavaScript" type="text/javascript">
<!--
function HelpRemember(){
	alert("Checking this option will save the last User Name\nlogged on this machine, only User Name.");
}
//-->
</script>

<form action="index.asp<%=sGoBackTo%>" method="post">
<table border=0>
	<tr>
		<td class="smallerheader"><u>U</u>ser Name</td>
		<td><input type="text" name="txtUserName" class="tbflat" value="<% =LastUserName %>" tabindex="1" accesskey="u"></td>
	</tr>
	<tr>
		<td class="smallerheader"><u>P</u>assword</td>
		<td><input type="password" name="txtPassword" class="tbflat" tabindex="2" accesskey="p"></td>
	</tr>
	<tr>
		<td></td>
		<td><input type="submit" name="cmdLogin" class="cmdflat" value=" Login " accesskey="l" tabindex="3"></td>
	</tr>
	<tr>
		<td class="smallerheader"><span title="Checking this option will save the last User Name logged on this machine, only User Name."><u>R</u>emember Me<sup><a href="#" onclick="HelpRemember()">?</a></sup></span></td>
		<td><input type="checkbox" name="chkRemember"<% =RememberCheck %> accesskey="r" tabindex="4"></td>
	</tr>
</table>
<input type="hidden" id="JavaScriptEnabled" name="JavaScriptEnabled" value="false">
<input type="hidden" name="SecurityID" value="<% =GetSecurityID("Login") %>">
</form>
<script language="JavaScript" type="text/javascript">
<!--
	function enableJS(){
		// If java is not enabled, it will not understand this function,
		// and then i'll get the parameter as false  ;)
		GetObject("JavaScriptEnabled").value = "true";
	}
//-->
</script>
<%
	if sText = "" then
	else
		%><!--#include file="te_footer.asp"--><%
	end if

end sub

	if request("cmdLogin") <> "" then
		'User provided the user name and password
		
		'Open the connections and create the recordset object
		OpenRS arrConn(0)
		
		sUserName = trim(request("txtUserName"))
		sPassword = trim(request("txtPassword"))
		sSecurityID = trim(request("SecurityID"))
		GoodLogin = not RequireSecurityID
		if ValidSecurityID("Login", sSecurityID) and RequireSecurityID then GoodLogin = true

		session("JavaScriptEnabled") = trim(request.form("JavaScriptEnabled"))

		sSQL = "SELECT * FROM Users WHERE UserName = " & SQLEncode(sUserName)
		rs.Open sSQL, , , adCmdTable
		if not (rs.bof or rs.eof) then
			if rs("Password") = sPassword and GoodLogin then
				'Login succeeded
				'Store info into session
				session("teUserName") = sUserName
				session("teFullName") = rs("FullName")
				session("rAdmin") = rs("rAdmin")
				session("rRecAdd") = rs("rRecAdd")
				session("rRecEdit") = rs("rRecEdit")
				session("rRecDel") = rs("rRecDel")
				session("rQueryExec") = rs("rQueryExec")
				session("rSQLExec") = rs("rSQLExec")
				session("rTableAdd") = rs("rTableAdd")
				session("rTableEdit") = rs("rTableEdit")
				session("rTableDel") = rs("rTableDel")
				session("rFldAdd") = rs("rFieldAdd")
				session("rFldEdit") = rs("rFieldEdit")
				session("rFldDel") = rs("rFieldDel")
				session("rAllowExport") = rs("rAllowExport")
				session("rConnectionViews") = rs("rTablePrivileges")
				session("teLastAccess") = rs("LastAccess")
				
				' Added by Hakan on May 11, 2002
				' If te install folder is not write enabled, ignore
				on error resume next
				rs("LastAccess") = Now
				rs.update
				on error goto 0
			
'				rs.close
'				conn.close

				if bUserLogging then call LogUserLogin(sUserName)

				chkRemember = (Request.form("chkRemember") = "on")
				if chkRemember then
					Response.Cookies("TableEditor8")("UserName") = sUserName
					Response.Cookies("TableEditor8")("RememberMe") = "true"
					Response.Cookies("TableEditor8").Expires = now + 30

				else
					Response.Cookies("TableEditor8")("UserName") = ""
					Response.Cookies("TableEditor8")("RememberMe") = "false"
				end if


				if sReferer = "" then
					response.redirect "te_admin.asp"
				else
					response.redirect sReferer
				end if
			else
				if not GoodLogin then
					'Security ID Failed, using a login method outsite the site
					AskForLogin "Login Error: you must use the login form on this site.<br>If you were already, please try refreshing the login form page before logging in. Cached form will not work."
				else
					'Login failed - Wrong password
					AskForLogin "Incorrect password."
				end if
			end if
		else
			'User not found
			AskForLogin "Incorrect credentials."
		end if
		CloseRS
	else
%>
<!--#include file="te_header.asp"-->
<%
		if bProtected then
		'If protection is ON, ask for login
			AskForLogin ""
		else
		'if protection is OFF, display a warning
%>
	<p class="smallheader">
		You are not using protection!
	</p>
	Any visitor who knows the exact location of the TableEditor files may view or change the information in your databases.<br>
	To enable protection, open <strong>te_config.asp</strong> file and set <strong>bProtected = True</strong>.<br><br>
	You may go to <a href="te_admin.asp" style="color: #0000FF;">Admin Page</a> now.

<%
		end if
%><!--#include file="te_footer.asp"--><%
	end if
%>