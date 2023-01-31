<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_tableedit.asp
	' Description: Displays the table structure for modification
	' Initiated By Hakan Eskici on Nov 15, 2000
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
<!--#include file="te_header.asp"-->
<table border=0 cellspacing=1 cellpadding=2 bgcolor = "#ffe4b5" width="100%">
	<tr>
		<td class="smallertext">
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <a href="te_listtables.asp?cid=<%=request("cid")%>"><%=arrDesc(request("cid"))%></a> » Edit Table [<%=request("tablename")%>]
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
	lConnID = request("cid")
	sTableName = request("tablename")
	
	OpenRS arrConn(lConnID)

	if arrType(lConnID) = tedbDsn then
		response.write "Table Structure modification might not work as expected with Dsn connections.<br><br>"
	end if

%>
<table border=0 cellspacing=1 cellpadding=2 bgcolor="#ffdead" width="100%"<% if bTableHighlight then response.write " style=""behavior:url(tablehl.htc);"" slcolor=#ffffcc hlcolor=#bec5de"%>>
<thead>
	<tr>
		<td width=10></td>
		<td class="smallerheader" colspan=4>Field Definitions
		<%
		if bFldAdd and bTableEdit then	
		%>
		&nbsp;&nbsp;&nbsp; (<a href="te_fieldedit.asp?cid=<%=lConnID%>&tablename=<%=server.urlencode(sTableName)%>">Add a new field</a>)
		<%	end if 	%>
		</td>
	</tr>
	<tr>
		<td bgcolor="#fffaf0" width=10></td>
		<td bgcolor="#fffaf0" class="smallerheader">Field Name</td>
		<td bgcolor="#fffaf0" class="smallerheader">Field Type</td>
		<td bgcolor="#fffaf0" class="smallerheader">Size</td>
		<td bgcolor="#fffaf0" class="smallerheader">Attributes</td>
		<td bgcolor="#fffaf0" class="smallerheader" align=right width=100>Action</td>
		<td bgcolor="#fffaf0" width=10></td>
	</tr>
</thead>
<tbody>
<%
	on error resume next

	rs.Open "[" & sTableName & "]",,,adCmdTable
	if rs.fields.count = 0 then
		response.write "No fields defined for table."
		CloseRS
		%><!--#include file="te_footer.asp"--><%
		response.end
	end if
	
	if err <> 0 then
		response.write err.number & ":" & err.description & "<br><br>"
	end if
	
	for each fld in rs.fields
		sFieldType = ""
		sAttributes = ""
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
      
      ' begin added for SQL2000 2001-04-20 15:46 <ivan@konsep.net>
			case adBigInt           : sFieldType = "bigint"
			case adBinary           : sFieldType = "binary"
			case adNumeric          : sFieldType = "decimal"
		case adVarChar		      : sFieldType = "varchar"
		case adDouble			      : sFieldType = "double"
		case adWChar			      : sFieldType = "nchar"
		case adVariant		      : sFieldType = "variant"
		case adVarBinary	      : sFieldType = "varbinary"
			case adChar			        : sFieldType = "char"
			case adDBTimeStamp      : sFieldType = "datetime"
			case adLongVarChar      : sFieldType = "text (long)"
			case adSingle           : sFieldType = "single"
      ' end added for SQL2000 2001-04-20 15:46 <ivan@konsep.net>

			case else				: sFieldType = fld.type
		end select
		isAutoIncrement4DSN = ""
		if fld.properties("IsAutoIncrement") = true then
			sAttributes = sAttributes & "(auto increment)"
			if arrType(lConnID) <> tedbAccess then isAutoIncrement4DSN = "&autoinc=1"
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

'			sAttributes = sAttributes & " (" & fld.attributes & ")"
		
		%>
		<tr bgcolor="#fffaf0">
			<td width=10></td>
			<td class="smallertext"><%=fld.name%></td>	
			<td class="smallertext"><%=sFieldType%></td>
			<td class="smallertext"><%=fld.definedsize%></td>	
			<td class="smallertext"><%=sAttributes%></td>
			<td class="smallertext" align=right>
				<% 	if bFldEdit and bTableEdit then %>
				<a href="te_fieldedit.asp?cid=<%=lConnID%>&tablename=<%=server.urlencode(sTableName)%>&fldname=<%=server.urlencode(fld.name)%><% =isAutoIncrement4DSN %>">edit</a>&nbsp;
				<% 
					end if
					if bFldDel and bTableEdit then
				%>
				<a href="te_fieldremove.asp?cid=<%=lConnID%>&tablename=<%=server.urlencode(sTableName)%>&fldname=<%=server.urlencode(fld.name)%>">remove</a>
				<%	end if %>
			</td>
			<td width=10></td>
		</tr>
		<%
	next

	response.write "</tbody></table><br>"

	rs.close
	
	'List Indexes of this table
	set rs = conn.openSchema(adSchemaIndexes)
	%>

<table border=0 cellspacing=1 cellpadding=2 bgcolor="#ffdead" width="100%"<% if bTableHighlight then response.write " style=""behavior:url(tablehl.htc);"" slcolor=#ffffcc hlcolor=#bec5de "%>>
<thead>
	<tr>
		<td width=10></td>
		<td class="smallerheader" colspan=4>Indexes</td>
	</tr>
	<tr>
		<td bgcolor="#fffaf0" width=10></td>
		<td bgcolor="#fffaf0" class="smallerheader">Index Name</td>
		<td bgcolor="#fffaf0" class="smallerheader">Indexed Column</td>
		<td bgcolor="#fffaf0" class="smallerheader">Primary Key</td>
		<td bgcolor="#fffaf0" class="smallerheader">Unique</td>
		<td bgcolor="#fffaf0" class="smallerheader" align=right width=100>Action</td>
		<td bgcolor="#fffaf0" width=10></td>
	</tr>
</thead>
<tbody>
	<%
	if arrType(lConnID) = tedbDsn then
		response.write "<tr bgcolor=""#fffaf0""><td></td><td class=""smallertext"" colspan=""5"">Unable to get indexes for Dsn connections.</td><td></td></tr>"
	end if
	
	do while not rs.eof
		if rs("table_name") = sTableName then
			response.write "<tr bgcolor=""#fffaf0""><td width=10></td>"
			response.write "<td class=""smallertext"">" & rs("Index_name") & "</td>"
			response.write "<td class=""smallertext"">" & rs("column_name") & "</td>"
			response.write "<td class=""smallertext"">" & rs("prImary_key") & "</td>"
			response.write "<td class=""smallertext"">" & rs("unIque") & "</td>"
			response.write "<td class=""smallertext"" align=right>" 
			if bFldDel then
			%>
				<a href="te_indexremove.asp?cid=<%=lConnID%>&tablename=<%=server.urlencode(sTableName)%>&idxname=<%=rs("Index_name")%>">remove</a>
			<%
			end if
			response.write "</td>"
			response.write "<td width=10></td></tr>"
		end if
		rs.movenext
	loop
	response.write "</tbody></table>" & vbCrLf
	
	CloseRS
%>
<!--#include file="te_footer.asp"-->