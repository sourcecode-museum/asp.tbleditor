<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_dbpassword.asp
	' Description: Ask for database password, for IE browsers
	' Initiated By Rami Kattan on May 25, 2002
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<title>Database Password</title>
<link rel="stylesheet" type="text/css" href="te.css">
</head>
<body bgcolor="threedlightshadow" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" scroll="no" style="border: 0px; overflow: hidden; margin: 0pt;">
<script language="JavaScript" type="text/javascript">
<!--
function DoOK(){
	window.returnValue = window.document.all.Password.value;
    window.close();
}

function DoCancel(){
	window.returnValue = "";
    window.close();
}

//-->
</script>
<form method="post">
<table align="center">
<tr>
	<td class="OpenDialogueFont" style="font-size: 16px">Password:</td>
	<td><input type="password" name="Password"></td>
</tr>
<tr>
	<td colspan="2" align="center">
	<button onclick="DoOK()" style="width: 80;">OK</button>&nbsp;&nbsp;
	<button onclick="DoCancel()" style="width: 80;">Cancel</button>
	</td>
</tr>
</table>
</form>

</body>
</html>