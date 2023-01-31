<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_fieldedit.asp
	' Description: Adds or edits a field in a table
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
	'-------------------------------------------------------------
	' # Nov 21, 2000 by Hakan Eskici
	' Fixed a bug regarding the index to be in the specified table
	' Added the check to see if a field is indexed (not being a PK)
	' Added sSQL2 to create an index not being a constraint
	' # Nov 27, 2000 by Hakan Eskici
	' Changed sql data type mappings for memo and boolean
	' # May 11, 2002 by Hakan Eskici
	' Small fixes for autoincrement, unique and primary key issues
	' # May 30, 2002 by Rami Kattan
	' Security check if user can add fields
	' Security fix, fieldedit require that user have the TableEdit permission too
	' Fix bug: DSN and SQL cannot show autoincrement with ADOX
	'==============================================================
%>
<!--#include file="te_config.asp"-->
<%

	sub ShowFieldForm
		if sName = "" then sName = sNewName
%>
	<!--#include file="te_header.asp"-->
	
	<table border=0 cellspacing=1 cellpadding=2 bgcolor = "#ffe4b5" width=100%>
		<tr>
			<td class="smallertext">
				<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <a href="te_listtables.asp?cid=<%=request("cid")%>"><%=arrDesc(request("cid"))%></a> » <a href="te_tableedit.asp?cid=<%=request("cid")%>&tablename=<%=server.urlencode(request("tablename"))%>">Edit Table [<%=request("tablename")%>]</a> » Edit Field
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
	<% if (bFldEdit or bFldAdd) and bTableEdit then	%>
	<p class="smallertext"><%=sErr%></p>
	<form action="te_fieldedit.asp?cid=<%=lConnID%>&tablename=<%=sTableName%>&fldname=<%=sName%>" method="post">
		<table border=0 width=400>
			<tr>
				<td class="smallerheader">Name</td>
				<td><input type="text" name="txtName" class="tbflat" value="<%=sName%>"></td>
			</tr>
			<tr>
				<td class="smallerheader">Type</td>
				<td>
					<select name="cboType" class="tbflat">
					<%
					for iTemp = 1 to 10
						response.write "<option value=""" & aT(iTemp,1) & """"
						if lType = aT(iTemp,1) then
							response.write " selected>"
						else
							response.write ">" 
						end if
						response.write aT(iTemp,2)
					next
					%>
					</select>
				</td>
			</tr>
			<tr>
				<td class="smallerheader">Size (if text)</td>
				<td><input type="text" name="txtSize" value="<%=lSize%>" size="5" maxlength="10" class="tbflat"></td>
			</tr>
			<tr>
				<td class="smallerheader" valign=top>Attributes</td>
				<td class="smallertext">
					<input type="checkbox" name="chkAcceptNulls" value="1"<%=sAcceptNulls%>>Accept Nulls?<br>
					<input type="checkbox" name="chkIndexed" value="1"<%=sIndexed%>>Indexed? <br>
					<input type="checkbox" name="chkIndexUnique" value="1"<%=sIndexUnique%>>Unique? <br>
					<input type="checkbox" name="chkPrimaryKey" value="1"<%=sPrimaryKey%>>Primary Key?<br>
					<input type="checkbox" name="chkAutoIncrement" value="1"<%=sAutoIncrement%>>Auto Increment? (if numeric)<% if sNote <> "" then response.write " *"%><br>
				</td>
			</tr>
			<tr>
				<td></td>
				<td><input type="submit" name="cmdSave" value=" Save " class="cmdflat"></td>
			</tr><%
		if sNote <> "" then response.write "<tr><td colspan=2><p class=""smallertext"">" & sNote & "</p></td></tr>"
	%>
			</table>
		</form>
<%
			else
				Response.Write "<p class=""smallerheader"">You have no permission to add/edit fields</p>"
			end if
	end sub

	dim aT(10,2)

	aT(1,1) = adVarWChar		:aT(1,2) = "text"
	aT(2,1) = adInteger			:aT(2,2) = "long"
	aT(3,1) = adSmallInt		:aT(3,2) = "integer"
	aT(4,1) = adBoolean			:aT(4,2) = "boolean"
	aT(5,1) = adDate			:aT(5,2) = "date"
	aT(6,1) = adCurrency		:aT(6,2) = "currency"
	aT(7,1) = adLongVarWChar	:aT(7,2) = "memo"
	aT(8,1) = adLongVarBinary	:aT(8,2) = "ole"
	aT(9,1) = adGUID			:aT(9,2) = "guid"
	aT(10,1)= adUnsignedTinyInt	:aT(10,2)= "byte"
	
	lConnID = request("cid")
	sTableName = request("tablename")
	sName = request("fldname")
	if sName = "" then bNewField = True
	
	'on error resume next	

	if request("cmdSave") <> "" and (bFldEdit or bFldAdd) then
		'Submitted new field definition
		
		OpenRS arrConn(lConnID)
		
		sNewName = request("txtName")
		sBefore = request("txtBefore")
		lType = CLng(request("cboType"))
		if request("txtSize") <> "" then lSize = CLng(request("txtSize"))
		if request("chkAutoIncrement") <> "" then sAutoIncrement = " checked" 
		if request("chkAcceptNulls") <> "" then sAcceptNulls = " checked"
		if request("chkPrimaryKey") <> "" then sPrimaryKey = " checked"
		if request("chkIndexed") <> "" then sIndexed = " checked"
		if request("chkIndexUnique") <> "" then sIndexUnique = " checked"
		
		sSQL = "ALTER TABLE [" & sTableName & "] "
		
		if bNewField then
			'We are adding a new field
			sSQL = sSQL & " ADD COLUMN [" & sNewName & "] "
		else
			'We are editing a field
			sSQL = sSQL & " ALTER COLUMN [" & sName & "] "
		end if
		
		select case lType
			case adSmallInt			: sFieldType = "SMALLINT"
			case adInteger			: sFieldType = "INTEGER"
			case adBoolean			: sFieldType = "BIT"
			case adDate				: sFieldType = "DATE"
			case adCurrency			: sFieldType = "CURRENCY"
			case adVarWChar			: sFieldType = "VARCHAR"
			case adLongVarWChar		: sFieldType = "TEXT"
			case adLongVarBinary	: sFieldType = "LONGVARBINARY"
			case adGUID				: sFieldType = "GUID"
			case adUnsignedTinyInt	: sFieldType = "TINYINT"
		end select
		
		sSQL = sSQL & sFieldType
		
		if lType = adVarWChar then
			'Set field size
			if lSize = "" then lSize = 50
			sSQL = sSQL & "(" & lSize & ") "
		end if
		
		if sAcceptNulls <> "" then
			sSQL = sSQL & " NULL "
		else
			sSQL = sSQL & " NOT NULL "
		end if
		
		if sAutoIncrement <> "" then
			sSQL = sSQL & " IDENTITY "
		end if

		if sPrimaryKey <> "" then
			'If the field has a primary key index
			sSQL = sSQL & " CONSTRAINT [pk" & sNewName & "] PRIMARY KEY "
		else
			'If the field is indexed (not PK)
			if sIndexed <> "" then
				if sIndexUnique <> "" then
					sSQL = sSQL & " CONSTRAINT [idx" & sNewName & "] UNIQUE"
				else
					'Index is no constraint, so another sql stat required
					sSQL2 = "CREATE INDEX [idx" & sNewName & "] ON [" & sTableName & "] ([" & sNewName & "]) "
				end if
			end if
		end if
		
		if (sName <> "") and (sNewName <> sName) then
			'Rename field
			'sSQL = "ALTER TABLE [" & sTableName & "] ALTER COLUMN RENAME [" & sName & "] TO [" & sNewName & "]"
			sErr = "Cannot rename field."
			bErr = True
		end if
		
		if bErr = True then
			conn.close
			set rs = nothing
			set conn = nothing
			ShowFieldForm
		else
'			response.write sSQL
			conn.execute sSQL 
			if sSQL2 <> "" then
				conn.execute sSQL2
			end if
	
			conn.close
			set rs = nothing
			set conn = nothing
			if err <> 0 then
				%><!--#include file="te_header.asp"--><%
				response.write err.description & "<br>" & sSQL
			else
				response.redirect "te_tableedit.asp?cid=" & lConnID & "&tablename=" & server.urlencode(request("tablename"))
			end if
		end if

	else
		if sName <> "" then
			'We are editing a field
			'Get the attributes for the field
		
			OpenRS arrConn(lConnID)
			set AdoX = server.createobject("adox.catalog")
'			set fld = server.createobject("adox.column")
	
			AdoX.ActiveConnection = conn
			set fld = AdoX.Tables(sTableName).Columns(sName)
			
			lType = fld.Type
			lSize = fld.DefinedSize
			
			if (fld.attributes AND adFldIsNullable) = adFldIsNullable then
				sAcceptNulls = " checked"
			end if
			
			'Enumerate the indices to check that if this field has a primary key'd index
			set rs = conn.openSchema(adSchemaIndexes)
			do while not rs.eof
				if rs("table_name") = sTableName then
					if (rs("column_name") = fld.Name) then
						sIndexed = " checked"
						if (rs("unIque") = True) then
							sIndexUnique = " checked"
						end if
						if (rs("prImary_key") = True) then 
							sPrimaryKey = " checked"
						end if
					end if
				end if
				rs.movenext
			loop
			
			if arrType(lConnID) <> tedbAccess then
					if request.querystring("autoinc") = "1" then
						sAutoIncrement = " checked disabled"
					else
						sAutoIncrement = " disabled"
					end if

					sNote = "* Cannot change <i>Auto Increment</i> fields."
			else
				if fld.properties("Autoincrement") = True then
					sAutoIncrement = " checked"
				end if
			end if

		end if
		
		'Show form for a new field definition
		if bNewField then 
			sErr = "Add new field."
		else
			sErr = "Edit field definition."
		end if
		ShowFieldForm
	end if
%>
<!--#include file="te_footer.asp"-->