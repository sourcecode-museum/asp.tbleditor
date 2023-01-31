<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_queryinfo.asp
	' Description: Displays query information
	' Initiated By Hakan Eskici on Nov 07, 2000
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
	' # Nov 22, 2000 by Hakan Eskici
	' Renamed the file from te_tableinfo.asp (which is replaced by
	' te_tableedit.asp) to te_queryinfo.asp
	' Removed listing of indexes
	'==============================================================
%>
<!--#include file="te_config.asp"-->
<!--#include file="te_header.asp"-->
<table border=0 cellspacing=1 cellpadding=2 bgcolor="#ffe4b5" width="100%">
	<tr>
		<td class="smallertext">
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <a href="te_listtables.asp?cid=<%=request("cid")%>"><%=arrDesc(request("cid"))%></a> » Table Info [<%=request("tablename")%>]
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

<%
	lConnID = request("cid")
	sTableName = request("tablename")
	sType = request("type")
	
	OpenRS arrConn(lConnID)
	
	set adox = server.createobject("adox.catalog")
	set cmd = server.createobject("adodb.command")
	adox.ActiveConnection = conn

	select case sType
		case "query"
			response.write "<br>Query : <br>" 
			response.write "<textarea cols=120 rows=5 class=""tbflat"" readonly>" & adox.views(sTableName).command.commandtext & "</textarea><br>"

%>
<br>
<table border=0 cellspacing=1 cellpadding=2 bgcolor = "#ffdead" width=100%>
	<tr>
		<td width=10></td>
		<td class="smallerheader" colspan=4>Field Definitions</td>
	</tr>
	<tr bgcolor="#fffaf0">
		<td width=10></td>
		<td class="smallerheader">Field Name</td>
		<td class="smallerheader">Field Type</td>
		<td class="smallerheader">Size</td>
		<td class="smallerheader">Attributes</td>
		<td width=10></td>
	</tr>
<%

	rs.Open "SELECT * FROM [" & sTableName & "]", , ,adCmdTable
	
	for each fld in rs.fields
		sAttributes = ""
		sFieldType = ""
		select case fld.Type
			case adSmallInt			: sFieldType = "integer"
			case adInteger			: sFieldType = "long"
			case adBoolean			: sFieldType = "boolean"
			case adDate				: sFieldType = "date"
			case adCurrency			: sFieldType = "currency"
			case adVarWChar			: sFieldType = "text"
			case adLongVarWChar		: sFieldType = "memo"
			case adLongVarBinary	: sFieldType = "ole"
			case adGUID				: sFieldType = "guid"
			case adUnsignedTinyInt	: sFieldType = "byte"
			case else				: sFieldType = fld.type
		end select
		
		if fld.properties("IsAutoIncrement") = true then
			sAttributes = sAttributes & "(auto increment)"
		end if
		
		'Actually this won't work, any recommendations are welcome
		if (fld.attributes and adFldKeyColumn) = adFldKeyColumn then
			sAttributes = sAttributes & " (primary key)"
		end if

		if (fld.attributes and adFldUpdatable) = adFldUpdatable then
			sAttributes = sAttributes & " (updatable)"
		end if
		
		if (fld.attributes and adFldIsNullable) = adFldIsNullable then
			sAttributes = sAttributes & " (nullable)"
		end if

		if (fld.attributes and adFldFixed) = adFldFixed then
			sAttributes = sAttributes & " (fixed)"
		end if

		'if (fld.attributes and adFldMayBeNull) = adFldMayBeNull then
		'	sAttributes = sAttributes & " (may be null)"
		'end if

		if (fld.attributes and adFldLong) = adFldLong then
			sAttributes = sAttributes & " (long)"
		end if

		if (fld.attributes and adFldRowID) = adFldRowID then
			sAttributes = sAttributes & " (row id)"
		end if

		if (fld.attributes and adFldIsRowURL) = adFldIsRowURL then
			sAttributes = sAttributes & " (url)"
		end if
		
		
		
		%>
		<tr bgcolor="#fffaf0">
			<td width=10></td>
			<td class="smallertext"><%=fld.name%></td>	
			<td class="smallertext"><%=sFieldType%></td>
			<td class="smallertext"><%=fld.definedsize%></td>	
			<td class="smallertext"><%=sAttributes%></td>
			<td width=10></td>
		</tr>
		<%
	next
	
	response.write "</table><br>"
	
	CloseRS
	
		case "proc"
			response.write "<br>Procedure : <br>" 
			response.write "<textarea cols=120 rows=5 class=""tbflat"">" & adox.procedures(sTableName).command.commandtext & "</textarea><br>"
	end select
%>
<!--#include file="te_footer.asp"-->