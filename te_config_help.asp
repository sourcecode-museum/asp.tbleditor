<%	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_config_help.asp
	' Description: Displays Help for config options
	' Initiated By Rami Kattan on May 14, 2002
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
	' Jun 7, 2002 by Hakan
	' Provided simple help contents
	'--------------------------------------------------------------
	'==============================================================
%>
<!--#include file="te_config.asp"-->
<!--#include file="te_header.asp"-->
<table border=0 cellspacing=1 cellpadding=2 bgcolor="#ffe4b5" width="100%">
	<tr>
		<td class="smallertext">
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » Help
		</td>
	</tr>
</table>
<br>
<%
FieldNames_long  = array("Encode HTML" ,"Relation","Show Connection Details","Convert Nulls","Convert Numeric Null","Convert Date Null","Max Length to Show","Bulk Delete","Bulk Compact","High Light","Export To Excel","Count Active Users","Page Selector","Combo Tables","IE Advanced Mode","Export XML","Use Popups","Default Record Per Page","High Security Login")
iHelpID = Request.QueryString("ID")
'response.write "<div class=""HelpTitle"">" & FieldNames_long(iHelpID) & "</div>"
response.write "<div class=""HelpText"">"
response.write "<div class=""smallheader"">" & FieldNames_long(iHelpID) & "</div><br>"
select case iHelpID
	case 0
		%>If your data includes HTML content, by enabling this option you can tell TableEditoR to display the content as html source instead of rendered html.<br><br><em><strong>Warning: </strong>If HTML Encode is turned on, you might encounter problems if you set <a href="te_config_help.asp?ID=6">Max Length to Show</a>.</em><%
	case 1
		%>Enable this option if you would like to make use of inter table relations. TableEditoR will display drop down boxes for related foreign fields if a table has a foreign key. <br><br><em><strong>Warning: </strong>If the foreign table data record count is large, the performance might suffer.</em><%
	case 2
		%>Enable this option if you would like to view the number of tables, procedures etc. on Connections screen.<%
	case 3
		%>If enabled, empty fields will be converted to null.<%
	case 4
		%>If enabled, empty and zero numeric fields will be converted to null.<%
	case 5 
		%>If enabled, empty and zero date fields will be converted to null.<%
	case 6
		%>Sets the maximum number of characters to display in grid view. Set to 0 (zero) for no limit.<%
	case 7
		%>Enables deleting more than one record at once.<%
	case 8
		%>Enables compacting more than one database at once.<%
	case 9
		%>If enabled, TableEditoR will highlight table rows when mouse is over.<%
	case 10
		%>If enabled, a "Export to Excel" button will be displayed under the data grid.<%
	case 11
		%>If enabled, TableEditoR will maintain a list of online users. <%
	case 12
		%>If enabled, TableEditoR will display a page selector in grid view. For tables with less data a drop down will be displayed, for large data sets, a textbox is displayed.<%
	case 13
		%>If enabled, navigation link will be displayed as a drop down with all tables.<%
	case 14
		%>Enable this option if you would like to make use of Internet Explorer's special TDC view. In Advanced IE mode, the amount of data transferred is almost 50% less compared to classic mode.<%
	case 15
		%>If enabled, a "Export to XML" button will be displayed under the data grid.<%
	case 16
		%>If enabled, TableEditoR will launch popup windows for editing tasks.<%
	case 17
		%>Sets the default value for number of records per page drop down.<%
	case 18
		%>Enable this option for more secure login. <br><br><em><strong>Warning: </strong>You must always request a fresh copy of the login page for this feature to work. TableEditoR will assign a unique hidden value to the login form and will expect the same upon submit. Therefore if your login page is retrieved from cache, you will not be able to login.</em><%
end select

%>
</div><br>
<!--#include file="te_footer.asp"-->