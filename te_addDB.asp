<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_addDB.asp
	' Description: Adds new database definitions
	' Initiated By Rami Kattan on April 22, 2002
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
	' # Jun 3, 2002 by Rami Kattan
	' Note about Brwose button if Popups is disabled
	' Hide browse button for linux Konqueror
	'==============================================================
%>
<!--#include file="te_config.asp"-->
<!--#include file="te_header.asp"-->
<%
	if len(request.querystring("q"))>0 then bQuery = True else bQuery = False

%>
<table border=0 cellspacing=1 cellpadding=2 bgcolor="#ffe4b5" width="100%">
	<tr>
		<td class="smallertext">
		<% if bPopUps then %>
			<% =arrDesc(Cint(request.querystring("cid"))) %> » Table [<% =request.querystring("tablename") %>] » <% if request.querystring("add") then response.write "Add" else response.write "Edit" %> Database
		<% else %>
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <a href="te_listtables.asp?cid=<%=request.querystring("cid")%>"><%=arrDesc(Cint(request.querystring("cid")))%></a> » <a href="te_showtable.asp?cid=<%=request.querystring("cid")%>&tablename=<%=server.urlencode(request.querystring("tablename"))%>&ipage=<%response.write(request.querystring("ipage"))%><% if bQuery then response.write("&q=1") end if %>">Table [<%=request.querystring("tablename")%>]</a> » <%if request.querystring("add") then response.write "Add" else response.write "Edit"%> Database
		</td>
		<td class="smallerheader" width=130 align=right>
			<%
			if bProtected then 
				response.write session("teFullName")
				response.write " (<a href=""te_logout.asp"">logout</a>)" 
			end if
			%>
		<% end if %>
		</td>
	</tr>
</table>
<br>
<script language="JavaScript" type="text/javascript">
<!--
function ShowSample(){
	db_type = GetObject("DB_type").options[GetObject("DB_type").selectedIndex].value;
	if (db_type == 1) {
		GetObject("DB_loc").value = "";
		if (GetObject("db_lister") != null) GetObject("db_lister").disabled = false;
	}
	if (db_type == 2) {
		GetObject("DB_loc").value = "[DatabaseName];[ComputerName];[UserName];[Password]";
		if (GetObject("db_lister") != null) GetObject("db_lister").disabled = true;
	}
	if (db_type == 3) {
		GetObject("DB_loc").value = "[DataSourceName]";
		if (GetObject("db_lister") != null) GetObject("db_lister").disabled = true;
	}
	if (db_type == 4) {
		GetObject("DB_loc").value = "";
		if (GetObject("db_lister") != null) GetObject("db_lister").disabled = true;
	}
}

function opendbwin(){
	popupWin = window.open('te_listDBs.asp','select_dbs','width=540,height=375,top=120,left=80,scrollbars=yes,resizable=no,status=no');
}
//-->
</script>
<%
	'disallow connections to admin table for nonadmin users 11/14/01
	lConnID = request("cid")
	if (lConnID = 0) and bAdmin = False then	
		response.write "Not authorized to view this connection."
		%><!--#include file="te_footer.asp"--><%
		response.end
	end if

	sTableName = request.querystring("tablename")
	sFieldNames = request.querystring("fld")
	sFieldValues = request.querystring("val")
	sFieldTypes = request.querystring("fldtype")
	iPage = request.querystring("ipage")
	sQuery = request.querystring("q")
	if request.querystring("add") then bAdd = True else bAdd = False

	if request.form("cmdSave") <> "" AND instr(request.form("DB_loc"), "teadmin.mdb") then
		response.write "This connection cannot be added."
		%><!--#include file="te_footer.asp"--><%
		response.end
	end if

    sParentName = server.urlencode(sTableName)

	OpenRS arrConn(lConnID)

	if not bAdd then
	
		'Added by Hakan
		'Support for multiple primary keys
		aFieldNames = split(sFieldNames, ";")
		aFieldTypes = split(sFieldTypes, ";")
		aFieldValues = split(sFieldValues, ";")
		
		select case arrType(lConnID)
			case tedbSQLServer
				sDateSeperator = "'"
			case else
				sDateSeperator = "#"
		end select
		
		for iFld = 0 to ubound(aFieldNames)
			sFieldName = aFieldNames(iFld)
			lFieldType = CLng(aFieldTypes(iFld))
			sFieldValue = aFieldValues(iFld)
			
			select case lFieldType
				case adDate, adDBDate, adDBTime, adDBTimeStamp
					if isDate(sFieldValue) then 
						sFieldValue = cDate(sFieldValue)
						sFieldValue = month(sFieldValue) & "/" & day(sFieldValue) & "/" & year(sFieldValue)
					end if
					
					if sWhereFields = "" then
						sWhereFields = "([" & sFieldName & "]=" & sDateSeperator & sFieldValue & sDateSeperator & ")"
					else
						sWhereFields = sWhereFields & " AND ([" & sFieldName & "]=" & sDateSeperator & sFieldValue & sDateSeperator & ")"
					end if
				case adTinyInt, adSmallInt, adInteger, adBigInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adUnsignedBigInt, adSingle, adDouble, adCurrency, adDecimal, adNumeric, adBoolean
					'Added by Hakan
					'Convert decimal point to dot if it's a comma
					sFieldValue = replace(sFieldValue, ",", ".")
					if sWhereFields = "" then
						sWhereFields = "([" & sFieldName & "]=" & sFieldValue & ")"
					else
						sWhereFields = sWhereFields & " AND ([" & sFieldName & "]=" & sFieldValue & ")"
					end if
				case else
					'Added by Hakan
					'Prepare SQL value by replacing single quote with two single quotes
					sFieldValue = replace(sFieldValue, "'", "''")
					if sWhereFields = "" then
						sWhereFields = "([" & sFieldName & "]='" & sFieldValue & "')"
					else
						sWhereFields = sWhereFields & " AND ([" & sFieldName & "]='" & sFieldValue & "')"
					end if
			end select
		next
		sWhere = " WHERE " & sWhereFields
	else
		sWhere = ""
	end if
	
	iPlace = instr(1, sTableName, " wHeRe ", 1)
	if iPlace then
		sTableName = left(sTableName, iPlace)
	end if

	'Added by Danival
	if instr(1, ucase(sTableName),"SELECT") then
		sSQL =  sTableName & sWhere
	else
		sSQL = "SELECT * FROM [" & sTableName & "]" & sWhere
	end if

	'response.write sSQL

        rs.Open sSQL, , , adCmdTable
        
   'Added By Brad Orgill
   'Reads FK's and PK's into an array
   
     
	if arrType(lConnID) <> tedbDsn then
		Set rs3 = conn.OpenSchema(adSchemaForeignKeys)
       
		Dim fkeyary()
		Dim numRows
		numRows = 0
		Do While NOT rs3.EOF
		numRows = numRows + 1
		ReDim Preserve fkeyary(1, numRows)
		fkeyary(0, numRows - 1) = rs3("FK_COLUMN_NAME") 'reads FK names
		fkeyary(1, numRows - 1) = rs3("PK_Column_Name") 'reads PK names
		rs3.MoveNext
		Loop

		rs3.Close
	end if

	Dim h  
     
		if request.form("cmdSave") <> "" and bRecEdit then
			if bAdd then
				rs.AddNew
			end if
			
			for each fld in rs.fields
				ifield = ifield + 1
				'If field is AutoIncrement, just skip it.
				
				
			   'Added By Brad Orgill
			   donada = false
			   fldName = fld.name
			   
			   'on error resume next
				
				'modified By Brad Orgill
				if (fld.properties("IsAutoIncrement") = false) AND (not donada) then
					For h = 0 to numrows - 1
       	    			if (fldName = fkeyary(1, h)) then 'OR (fldName = fkeyary(0, h)) then  remove the coment to disallow FK updates 
							donada = true
				   		end if
					Next	
				 if true then    ' By Rami, line was   "if (not donada) then"
					select case fld.type
						case adBoolean
							if len(request.form(fld.name)) then rs(fld.name) = True else rs(fld.name) = False
						case adLongVarBinary
							'noop
						case adDate
							if ( request.form(fld.name) = "" or request.form(fld.name) = "0" ) and (fld.attributes and adFldIsNullable) = adFldIsNullable and bConvertDateNull then
								rs(fld.name) = NULL
							else
								if isDate(request.form(fld.name)) then 
									rs(fld.name) = request.form(fld.name)
								end if
							end if
						case adSmallInt, adInteger, adCurrency, adUnsignedTinyInt
							if ( request.form(fld.name) = "" or request.form(fld.name) = "0" ) and (fld.attributes and adFldIsNullable) = adFldIsNullable and bConvertNumericNull then
								rs(fld.name) = NULL
							else
								if isNumeric(request.form(fld.name)) then 
									rs(fld.name) = request.form(fld.name)
								end if						
							end if
						case else
							if request.form(fld.name) = "" and (fld.attributes and adFldIsNullable) = adFldIsNullable and bConvertNull then
							    rs(fld.name) = NULL
							else
							    rs(fld.name) = request.form(fld.name)
							end if						
					end select
					if Err <> 0 then
						response.write "<strong>Error</strong> while updating field '" & fld.name & "'<br>"
						response.write "Description: " & err.description & "<br>"
						select case err
							case -2147352571	'Type mismatch
								response.write "If this is an <strong>auto increment</strong> field, you may <strong>ignore</strong> this error.<br><br>"
							case -2147217887	'Field cannot be updated
								response.write "If this is an <strong>auto increment</strong> field, you may <strong>ignore</strong> this error.<br><br>"
							case else
								response.write "<br>"
								bError = True
						end select
						err = 0
					else
						'Added to enable editing of first field.
						if iField = 1 then
							sFieldValue = request.form(fld.name)
						end if
					end if
					end if
				end if
			
			next
			rs.update
			if err <> 0 then
				response.write "<strong>Error</strong> while updating.<br>"
				response.write "Description: " & err.description & "<br>"
				bError = True
				err = 0
			end if
			if not bError then
				if request.querystring("cmdsave") <> "" and not request.querystring("add") then
					response.write "Database updated"
				else
					response.write "Database added"
				end if
			end if
		end if
		on error goto 0

		if not bAdd then
			DB_Desc = rs("DB_Desc")
			DB_type = rs("DB_type")
			DB_loc = rs("DB_loc")
			DB_priv = rs("DB_privileges")
		end if

		if not bPopUps then PopUpAsterisk = "*"

		bBrowseButton = bJSEnable and not isKo
%>

<form action="te_addDB.asp?cid=0&tablename=Databases&fld=<%=sFieldName%>&val=<%=sFieldValue%>&fldtype=<%=lFieldType%>&ipage=<%=iPage%>&add=<%=bAdd%>&cmdsave=1<% if bQuery then response.Write("&q=1") %>" method="post" name="frm">
<input type="hidden" name="DBAdder" value=1>
<input type="hidden" name="add" value=1>

<table border=0 cellspacing=2 cellpadding=3 bgcolor="#ffe4b5" width=500>
		
<tr><td class="smallerheader">Description</td>
<td><input class="tbflat" size=64 type="text" name="DB_Desc" id="DB_Desc" value="<%=DB_Desc%>" maxlength="100"></td></tr>

<tr><td class="smallerheader">Type</td>
<td><select class="tbflat" name="DB_type" id="DB_type" onchange="ShowSample()">
	<option value="1"<% if DB_type = 1 then response.write " Selected"%>>Access</option>
	<option value="2"<% if DB_type = 2 then response.write " Selected"%>>SQL Server</option>
	<option value="3"<% if DB_type = 3 then response.write " Selected"%>>DSN</option>
	<option value="4"<% if DB_type = 4 then response.write " Selected"%>>Connection String</option>
</select>
</td></tr>

<tr><td class="smallerheader">Location</td>
<td><input class="tbflat" size=64 type="text" name="DB_loc" id="DB_loc" value="<% =DB_loc %>" maxlength="100"><% if bBrowseButton then %>&nbsp;<input type="button" onclick="opendbwin();" class="cmdflat" name="db_lister" id="db_lister" value="Browse<% =PopUpAsterisk %>"><% end if %></td></tr>

<tr><td class="smallerheader">Required Privileges</td>
<td>
<select class="tbflat" name="DB_privileges" id="DB_privileges">
<% 

OpenRS2 arrConn(0)
rs2.open "SELECT * FROM tePrivileges",,,adCmdTable
		do while not rs2.eof
			if len(DB_priv) > 0 and len(rs2("PrivValue")) > 0 then
				if cint(DB_priv) = cint(rs2(0)) then
				   flagnbh = " selected"
				else
					flagnbh = ""
				end if
			end if
			Response.Write "<option value=""" & rs2("PrivValue") & """" & flagnbh & ">" & rs2("PrivName") & "</option>" & vbcrlf
			rs2.movenext
		loop
CloseRS2
%>
</select>
</td></tr>
<tr><td></td>
	<td><input type="submit" name="cmdSave" id="cmdSave" value="  Save  " class="cmdflat" accesskey="s">
<% if bPopUps then %>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" onclick="window.close()" class="cmdflat" value="  Close  ">
<% end if %>
</td></tr><%
if bBrowseButton and not bPopUps then response.write "<tr><td colspan=""2"" class=""smallertext"">* Your browser must allow popup windows for the Browse function to work.</td></tr>"

%></table>
</form>
<!--#include file="te_footer.asp"-->