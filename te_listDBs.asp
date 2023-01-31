<!--#include file="te_config.asp"-->
<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_listDBs.asp
	' Description: List mdb database files on the server
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
%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>Table Editor 0.8 &raquo; Server Browser</title>
	<link rel="stylesheet" type="text/css" href="te.css">
	<script language="JavaScript" type="text/javascript" src="functions.js"></script>
</head>
<%
dbFolder = request.form("Base_Loc")
LastCurrentFolder = request.form("CurrentFolder")
if LastCurrentFolder = "/" then LastCurrentFolder = ""
if instr(dbFolder, "..") then dbFolder = ""
if dbFolder = "" then dbFolder = request.form("folder")
num_files = 0
strDBs = ""

' Added by Hakan on May 11, 2002
dbFolder = replace(dbFolder, "\", "/")

if dbFolder = "" then dbFolder = "/"    'if first view, get virtual root path
if dbFolder <> "/" AND right(dbFolder, 1) = "/" then dbFolder = left(dbFolder, len(dbFolder)-1)
if dbFolder <> "/" AND left(dbFolder, 1) <> "/" then dbFolder = LastCurrentFolder & "/" & dbFolder

function formatSize(filesize)
	formatSize = FormatNumber(filesize/1000, 0)
end function

	folder_found = true
	Set fso = CreateObject("Scripting.FileSystemObject")

	If fso.FolderExists(Server.MapPath(dbFolder)) = false Then
		response.write "Folder '" & dbFolder & "' not found...<br>"
		dbFolder = LastCurrentFolder
		folder_found = false
	end if
		Set theCurrentFolder = fso.GetFolder(Server.MapPath(dbFolder))
		Set inFolders = theCurrentFolder.SubFolders
		if dbFolder <> "/" then
			upFolder = left(dbFolder, inStrRev(dbFolder, "/")-1)
			if upFolder = "" then upFolder = "/"
			strDBs = "<option class=""folder"" value=""" & upFolder & """>../" & "</option>" & vbCrLf
			StateImageTruth = true
		end if
		if dbFolder = "/" then dbFolder = ""
		For Each folder in inFolders
			strDBs = strDBs & "<option class='folder' value=""" & dbFolder & "/" & folder.name & """>" & ucase(folder.name) & vbCrLf
		Next

		Set curFiles = theCurrentFolder.Files 

		strDesc = ""				
		For Each fileItem in curFiles
			fname = fileItem.Name
			fext = InStrRev( fname, "." )
			If fext < 1 Then fext = "" Else fext = Mid(fname,fext+1)

			if fext = "mdb" and fname <> "te_admin.mdb" and fname <> "teadmin.mdb" then
					fileClass = "mdbNew"
					for each FClass in arrDBs
						if instr(ucase(FClass), ucase(fname)) > 0 then
							fileClass = "mdbExist"
							exit for
						end if
					next
					fname = lcase(fname) & "&nbsp;&nbsp;&nbsp;&nbsp; (" & formatSize(fileItem.size) & " KB)"
					strDBs = strDBs & "<option class=""" & fileClass & """ value=""" & dbFolder & "/" & fname & """>" & fname & "</option>" & vbCrLf
				num_files = num_files + 1
			end if

		Next

		set  curFiles = nothing
		set  theCurrentFolder = nothing
    set  fso = nothing

	if dbFolder = "" then dbFolder = "/"

	btnDisabled = " disabled"

	if StateImageTruth then
		StateImage = "s1"
	else
		StateImage = "s3"
	end if

	if isIE then
		PageBGColor = "buttonface"
	else
		PageBGColor = "#D6D3CE"
	end if
%>
<body bgcolor="<% =PageBGColor %>" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" scroll="no" style="border: 0px; overflow: hidden; margin: 0pt;">
<%
	if bAdmin = False then	
		response.write "Not authorized to view this connection."
		%><!--#include file="te_footer.asp"--><%
		response.end
	end if
%>
<script language="JavaScript">
document.oncontextmenu = function() {return false};
</script>
<table border=0 cellspacing=1 cellpadding=2 width="100%">
	<tr>
		<td class="OpenDialogueFont">
			<font size="3"><b>Table Editor Administration » Select Database</b></font>
		</td>
	</tr>
</table>
<script language="JavaScript" type="text/javascript">
<!--

function DoOpen(){
	var loca = GetObject("folder").options[GetObject("folder").selectedIndex].value;
	if (loca.indexOf(".mdb") == -1)	document.frm.submit();
	else AddThis();
}

function AddThis(){
	isIE5 = (document.all && document.getElementById) ? true : false;
	var password = "";
	if (GetObject("PasswordProtected").checked) {
		var password
		
		if (isIE5)
			password = window.showModalDialog("te_dbpassword.asp", null,"font-size:10px; dialogWidth:265px;dialogHeight:105px");
		else
	        password = prompt("Please enter the database password", "");
	}
	if (password == undefined) password = ""
	FileName = GetObject("folder").options[GetObject("folder").selectedIndex].value;
	FileName = FileName.substr(0, FileName.indexOf(".mdb")+4) + ";" + password;
	window.opener.document.frm.DB_loc.value = FileName;
	window.close();
}

function CheckAdd(){
	GetObject("btnOpen").disabled = false;
	var loca = GetObject("folder").options[GetObject("folder").selectedIndex].value;
	if (loca.indexOf(".mdb") == -1){
		GetObject("btnOpen").value = "Open Folder";
	} else {
		GetObject("btnOpen").value = "Select File";
		FileName = GetObject("folder").options[GetObject("folder").selectedIndex].text
		FileName = FileName.substr(0, FileName.indexOf(".mdb")+4);
		GetObject("FileName").value = FileName;
	}
}

function onDoubleClick(){
	if ((GetObject("folder").value).indexOf(".mdb") == -1)
		GetObject("frm").submit();
	else 
		AddThis();
}

function UpFolder(){
	var locPath = GetObject("folder").options[0].value;
	var locName = GetObject("folder").options[0].text;
	if (locName == "../"){
		GetObject("Base_Loc").value = locPath;
		GetObject("frm").submit();
	}
}

function changeFolderImage(state){
	var img;
	if (<% =lcase(StateImageTruth=true) %>) {
		if (state == "up") img = "s2.gif";
		else img = "s1.gif"
		GetObject("FolderImage").src = "images/" + img;
	}
}
//-->
</script>
<table border=0 cellspacing=1 cellpadding=2 width="100%">
<form method="post" action="te_listDBs.asp" name="frm">
<tr valign="center">
<td width="120" class="OpenDialogueFont">Looking in:</td>
<td width="250"><input type="text" value="<%=dbFolder%>" readonly class="smallertext" style="width:100%;"></td>
<td class="OpenDialogueFont"><img id="FolderImage" src="images/<% =StateImage %>.gif" onclick="UpFolder();" onmouseover="changeFolderImage('up');" onmouseout="changeFolderImage('out');"></td>
</tr>
<tr valign="center"><td width="120" class="OpenDialogueFont">Enter <u>v</u>irtual folder:</td>
<td width="250"><input accesskey="v" type="text" name="Base_Loc" id="Base_Loc" value="" class="smallertext" style="width:100%;"></td>
<td width="140"><input type="submit" value="Check folder" style="width:145;"></td>
</tr>
<tr  valign="center">
<td colspan="3" class="OpenDialogueFont">or <u>S</u>elect from list:<select name="folder" id="folder" style="width:100%;" class="smallertext" size=8 accesskey="s" onchange="CheckAdd();" ondblclick="onDoubleClick()">
<% =strDBs %>
</select></td></tr>
<tr><td class="OpenDialogueFont">
<input type="hidden" name="CurrentFolder" value="<% =dbFolder %>">
File<u>n</u>ame:</td><td><input accesskey="n" type="text" name="FileName" id="FileName" style="width:100%;"></td>
<td>
<input type="button" accesskey="o" onclick="DoOpen();" id="btnOpen" style="width:145;" value="Open Folder"<% =btnDisabled%>></td>
</tr>
<tr>
<td class="OpenDialogueFont">Files of <u>t</u>ype:</td>
<td><select accesskey="t" style="width:100%;"><option>Microsoft Access (*.mdb)</option></select></td>
<td><input type="button" onclick="window.close()" style="width:145;" accesskey="esc" value="Cancel">
</td>
</tr>
<tr>
<td>&nbsp;</td><td colspan="2" class="OpenDialogueFont"><input type="checkbox" name="PasswordProtected" id="PasswordProtected" value="true"><label for="PasswordProtected">Database is password protected.</label></td>
</tr>
<tr><td colspan=3 class="OpenDialogueFont">
<%
	response.write "<b>Found " & num_files & " database file"
	if num_files > 1 then response.write "s"
	response.write ".</b>"
%></td></tr>
<tr><td  class="OpenDialogueFont">Color Codes:</td><td class="OpenDialogueFont" colspan="2">
<span class="folder">Folder</span> | <span class="mdbExist">Similar filename already exist</span> | <span class="mdbNew">New filename</span></td></tr></form></table>
<hr size="2">
<table border=0 width="100%">
	<tr>
	<td class="OpenDialogueFont">
		Visit <a href="http://www.2enetworx.com/dev" style="color: #0000FF; text-decoration: underline;">2eNetWorX</a> for more OpenSource VB and ASP Projects.
	</td>
	<td align="right">
		<a href="http://www.2enetworx.com/dev/projects/tableeditor.asp"><img src="images/te.gif" width=90 height=30 alt="Table Editor" border="0" align="middle"></a>
	</td>
	</tr>
</table>
</body>
</html>