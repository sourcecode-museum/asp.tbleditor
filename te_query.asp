<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_query.asp
	' Description: Asks for an SQL statement to be executed
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
	' # Mar 31, 2001 by Hakan Eskici
	' Added security check
	' # May 30, 2002 by Rami Kattan
	' SQL Builder, for making SQL statements easily, under development
	'==============================================================
%>
<!--#include file="te_config.asp"-->
<!--#include file="te_header.asp"-->
<table border=0 cellspacing=1 cellpadding=2 bgcolor="#ffe4b5" width="100%">
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
	if bSQLExec then
%>
<form action="te_runquery.asp?cid=<%=request("cid")%>" method="post">
	<table border=0 cellspacing=2 cellpadding=3 bgcolor = "#ffe4c4" width=600>
<% 	if bJSEnable and bSQLBuilder then
		bAdd = True
		sAdd = "&add=1"
		sAction = "Create new Query"
%>
<script language="JavaScript" type="text/javascript">
<!--

text = "";

function SQLBuilder(method){
	var Operands;
	switch (method){
		case 'Select':	Operands = prompt("Enter field names to select", "*");
						if (Operands == null) Operands = "";
						insertMethod = "SELECT " + Operands + " ";
						break;
		case 'From'  :	Operands = prompt("Enter table name to select from", "");
						if (Operands == null) Operands = "";
						insertMethod = "FROM " + Operands + " ";
						break;
		case 'Where'  :	Operands = prompt("Enter Conditions", "");
						if (Operands == null) Operands = "";
						insertMethod = "WHERE " + Operands + " ";
						break;
		case 'GroupBy':	Operands = prompt("Enter Grouping column", "");
						if (Operands == null) Operands = "";
						insertMethod = "GROUP BY " + Operands + " ";
						break;
		case 'Having':	Operands = prompt("Enter Having condition", "");
						if (Operands == null) Operands = "";
						insertMethod = "HAVING " + Operands + " ";
						break;
		case 'OrderBy':	Operands = prompt("Enter ordering column", "");
						if (Operands == null) Operands = "";
						insertMethod = "ORDER BY " + Operands + " ";
						break;
	}
	
	GetObject("txtSQL").value = GetObject("txtSQL").value + insertMethod;
}

function getActiveText(selectedtext) { 
	text = (document.all) ? document.selection.createRange().text : document.getSelection();
		if (selectedtext.createTextRange) {	
   			selectedtext.caretPos = document.selection.createRange().duplicate();	
  		}
		return true;
}
function setfocus() {
  GetObject("txtSQL").focus();
}

function AddText(NewCode) {
	if (text != "" && GetObject("txtSQL").createTextRange && GetObject("txtSQL").caretPos) {
		var caretPos = GetObject("txtSQL").caretPos;
		caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ? NewCode + ' ' : NewCode;
	}
	setfocus();
}

function InsertFX(Fx) {
		AddTxt = Fx + "("+text+")";
		AddText(AddTxt);
}

function insertFX(){
	FX = GetObject("SQLFunction").value;
	if (FX != "") InsertFX(FX);
//	alert(GetObject("txtSQL").caretPos.text);
}

function OpenSQLBuilder(){
	popupWin = window.open('te_SQLBuilder.asp?cid=<% =lConnID %>','sql_builder','width=540,height=395,top=120,left=80,scrollbars=no,resizable=no,status=yes');
}
//-->
</script>
<tr><td class="smallerheader" colspan="2">Query Builder &nbsp;&nbsp;&nbsp;&nbsp; &lt;<span onclick="SQLBuilder('Select')" class="SQLBuilder">SELECT</span>&gt; &lt;<span onclick="SQLBuilder('From')" class="SQLBuilder">FROM</span>&gt; &lt;<span onclick="SQLBuilder('Where')" class="SQLBuilder">WHERE</span>&gt; &lt;<span onclick="SQLBuilder('GroupBy')" class="SQLBuilder">GROUP BY</span>&gt; &lt;<span onclick="SQLBuilder('Having')" class="SQLBuilder">HAVING</span>&gt; &lt;<span onclick="SQLBuilder('OrderBy')" class="SQLBuilder">ORDER BY</span>&gt;
<select name="SQLFunction" id="SQLFunction" class="smallertext">
<option value=""></option>
<option value="COUNT">Count</option>
<option value="MAX">Max</option>
<option value="MIN">Min</option>
<option value="AVG">Avg</option>
</select><input type="button" value="Insert" onclick="insertFX()">
</td></tr>
<tr>
<td width=350>
	<textarea cols=60 rows=10 name="txtSQL" id="txtSQL" class="tbflat" onblur="getActiveText(this)" onfocus="getActiveText(this)" onclick="getActiveText(this)" onchange="getActiveText(this)">SELECT * FROM r</textarea>
</td>
<% else %>
<tr>
	<td width=350>
		<textarea cols=60 rows=10 name="txtSQL" id="txtSQL" class="tbflat"></textarea>
	</td>
<% end if %>
			<td class="smallertext" valign=top align=left>
				<input type="checkbox" name="chkRec" value="1" checked>Query returns records<br><br>
				<input type="submit" name="cmdExecute" value=" Execute " class="cmdFlat">
			</td>
		</tr>
	</table>
</form>
<%
	else
%>
	<p class="smallheader">
	You don't have permission to execute sql statements.
	</p>
<%
	end if
%>
<!--#include file="te_footer.asp"-->