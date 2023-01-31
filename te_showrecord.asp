<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_showrecord.asp
	' Description: Displays the selected record for editing
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
	' # Nov 16, 2000 by Kevin Yochum
	' Added null value checks depending on the config switches
	' # Nov 23, 2000 by Hakan Eskici
	' Added replacing HTML Content
	' # Nov 29, 2000 by Hakan Eskici
	' Changed displaying boolean field which may cause incorrect
	' checkbox view when HTMLEncoding is True
	' # Feb 17, 2001 by Danival A. Souza
	' Added Foreign Key Support
	' # Mar 26, 2001 by Hakan Eskici
	' Added support for multiple primary keys
	' Changed request()'s to use form or querystring for fixing the
	' unicode form submission bug in NT4 systems
	' # Aug 16, 2001 by Brad Orgill
	' Added check from non-auto incrament primary keys and forigne keys
	' # Nov 14, 2001 by Jeff Wilkinson (jwilkinson@mail.com)
	' security fix entered per Dilyias suggested fix 10/29/01
	' prevents nonadmin users from accessing the admin db (conn=0)
	' # Nov 27, 2001 by Jeff Wilkinson (jwilkinson@mail.com
	' added in some vbcrlf's and </tr>'s to make the html cleaner
	' # April 20, 2002 by Rami Kattan
	' Autoincrement can be ignored
	' Option to autofill date fields
	' # May 30, 2002 by Rami Kattan
	' Security check updates, to allow adding/editing records
	' Fixed Query record updates
	'==============================================================
%>
<!--#include file="te_config.asp"-->
<!--#include file="te_header.asp"-->
<%
	if len(request.querystring("q"))>0 then bQuery = True else bQuery = False

	iPage = request.querystring("ipage")
	sQuery = request.querystring("q")

%>
<table border=0 cellspacing=1 cellpadding=2 bgcolor="#ffe4b5" width="100%">
	<tr>
		<td class="smallertext">
		<% if bPopUps then %>
			<% =arrDesc(Cint(request.querystring("cid"))) %> » Table [<% =request.querystring("tablename") %>] » <% if request.querystring("add") then response.write "Add" else response.write "Edit" %> Record
		<% else %>
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <a href="te_listtables.asp?cid=<% =request.querystring("cid") %>"><% =arrDesc(Cint(request.querystring("cid"))) %></a> » <a href="<% =TableViewerCompat %>?cid=<% =request.querystring("cid") %>&tablename=<% =server.urlencode(request.querystring("tablename")) %>&ipage=<% =request.querystring("ipage") %><% if bQuery then response.write("&q=" & request.querystring("q")) %>">Table [<%=request.querystring("tablename")%>]</a> » <% if request.querystring("add") then response.write "Add" else response.write "Edit"%> Record
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
function Insert_Today(form_element){
	var d = new Date();
	document.frm[form_element].value = "<% =formatDateTime(now, 0) %>";
	//d.getDate() + "/" +  parseInt(d.getMonth() + 1) + "/" +  d.getFullYear();
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

	sTableName = trim(request.querystring("tablename"))
	sFieldNames = request.querystring("fld")
	sFieldValues = request.querystring("val")
	sFieldTypes = request.querystring("fldtype")
	if request.querystring("add") <> "" and request.querystring("cmdsave") <> "1" and not bRecAdd then
		response.write "<p class=""smallerheader"">You have no pemissions to add records.</p>" %>
		<!--#include file="te_footer.asp"-->
		<%
		response.end
	end if
	if request.querystring("add") then bAdd = True else bAdd = False

    sParentName = server.urlencode(sTableName)
	TableNameIsQuery = instr(1, ucase(sTableName),"SELECT") > 0

	OpenRS arrConn(lConnID)
	
	' the following if check if adding or editing, if editing it will build the where sql statement
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
	
	iPlace = instr(1, ucase(sTableName), " WHERE ", 1)
	if iPlace and TableNameIsQuery then
		sTableName = left(sTableName, iPlace)
	end if

	'Added by Danival
	if TableNameIsQuery then
		if right(sTableName, 1) = ";" then sTableName = left(sTableName, len(sTableName)-1)
		sSQL =  sTableName & sWhere
	else
		sSQL = "SELECT * FROM [" & sTableName & "]" & sWhere
	end if

	if te_debug then response.write sSQL & "<br>"

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

	set adox = server.createobject("adox.catalog")
	adox.ActiveConnection = conn

		if request.form("cmdSave") <> "" and (bRecEdit or bRecAdd) then
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
				if arrType(lConnID) <> tedbAccess or TableNameIsQuery then
					isAutoIncrement = fld.properties("IsAutoIncrement")
				else
					set fldx = AdoX.Tables(sTableName).Columns(fld.name)
					isAutoIncrement = fldx.properties("Autoincrement")
					set fldx = nothing
				end if

				if (not isAutoIncrement) AND (not donada) then
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
				response.write "Record saved"
			end if
		end if
		'on error goto 0
	if bRecEdit or bRecAdd then
%>
	<form action="te_showrecord.asp?cid=<%=lConnID%>&tablename=<%=sParentName%>&fld=<%=sFieldName%>&val=<%=sFieldValue%>&fldtype=<%=lFieldType%>&ipage=<%=iPage%>&add=<%=bAdd%>&cmdsave=1<% if bQuery then response.Write("&q=" & request.querystring("q")) %>" method="post" name="frm">

	<table border=0 cellspacing=2 cellpadding=3 bgcolor="#ffe4b5" width="500">
		<%
			lFields = rs.Fields.count
			for each fld in rs.fields
				if not bAdd then
					sValue = rs(fld.name)
				end if

				if bEncodeHTML then
					if sValue <> "" then sValue = server.htmlencode(sValue)
				end if

				response.write vbcrlf & "<tr>"
				response.write "<td class=""smallerheader"">" & fld.name & "</td>" & vbcrlf
				select case fld.type
					case adLongVarWChar	'memo
						response.write "<td><textarea class=""tbflat"" cols=60 rows=6 name=""" & fld.name & """>" & sValue & "</textarea>"
						if (fld.attributes and 102) AND bRequiredField then response.write "<span class='smallertext'>*</span>"
						response.write "</td>"
					case adLongVarBinary	'ole
						'response.write "<td><img src=""te_getole.asp?cid=" & lConnID & "&tablename=" & sTableName & "&fld=" & fld.name & """></td>" 
						response.write "<td class='smallertext'>Cannot be accessed.</td>" 
					case adBoolean
						response.write "<td><input type=""checkbox"" name=""" & fld.name & """"
						'Changed sValue to rs(fld.name) which may cause incorrect display when
						'HTML Encoding is True
						if rs(fld.name) = true then
							response.write " checked></td>"
						else
							response.write "></td>"
						end if
					case adDate
						response.write "<td><input class=""tbflat"" size=64 type=""text"" name=""" & fld.name &  """ value=""" & sValue & """ maxlength=""25""> <img src=""images/edit.gif"" title=""Insert today"" onclick=""Insert_Today('" & fld.name & "')"">"
						if (fld.attributes and 102) AND bRequiredField then response.write "<span class='smallertext'>*</span>"
						response.write "</td>"
					case else
						'Replace quotation marks with html &quot;
                        if sValue <> "" then sValue = replace(sValue, """", "&quot;")
						select case fld.type
							case adSmallInt, adInteger, adCurrency, adUnsignedTinyInt, adDate, adDBDate, adDBTime, adDBTimeStamp
							
							if arrType(lConnID) <> tedbDsn then
								'-[2]- Added by Danival
								flag = false
								set rs2 = conn.OpenSchema(adSchemaForeignKeys)
								do while not rs2.eof	
									if rs2("FK_TABLE_NAME") = sTableName and rs2("FK_COLUMN_NAME") = fld.name  and relation then
										Set Consulta = Server.CreateObject("ADODB.Command")
										Set Consulta.ActiveConnection = conn
										consulta.CommandText = "SELECT * FROM " & rs2("PK_TABLE_NAME")
										set RSConsulta = consulta.execute
										response.write "<td><select class=""tbflat""  name=""" & fld.name &  """>" & vbcrlf
										Response.Write "<option value="""">        </option>"
										do while not RSConsulta.eof
											if len(svalue) > 0 and len (RsConsulta(0)) > 0 then
												if CLng(sValue) = CLng(RSConsulta(0)) then
												   flagnbh = " selected"
												else
													flagnbh = ""
												end if
											end if
											Response.Write "<option value=""" & RSConsulta(trim(rs2("PK_COLUMN_NAME"))) & """" & flagnbh & ">" & RSConsulta(trim(rs2("PK_COLUMN_NAME"))) & " - " & RSConsulta(1) & "</option>" & vbcrlf
											RSConsulta.movenext
										loop
										set rsconsulta = nothing
										set consulta = nothing
										Response.Write "</select></td>"
										flag = true
									end if
									rs2.movenext
								loop
								set rs2 = nothing
							end if

							if arrType(lConnID) <> tedbAccess or TableNameIsQuery then
								isAutoIncrement = fld.properties("IsAutoIncrement")
							else
								set fldx = AdoX.Tables(sTableName).Columns(fld.name)
								isAutoIncrement = fldx.properties("Autoincrement")
								set fldx = nothing
							end if
							
							if flag = false then
								if not isAutoIncrement then
									response.write "<td><input class=""tbflat"" size=64 type=""text"" name=""" & fld.name &  """ value=""" & sValue & """>"
									if (fld.attributes = 102) AND bRequiredField then response.write "<span class='smallertext'>*</span>"
									response.write "</td>"
								else
									response.write "<td class=""smallertext"">Auto Increment"
									if sValue <> "" then response.write " (" & sValue & ")"
									response.write "</td>"
								end if
							end if	
							'-[2]- End Addition

							case else
								response.write "<td><input class=""tbflat"" size=64 type=""text"" name=""" & fld.name &  """ value=""" & sValue & """ maxlength=""" & fld.definedsize & """>"
								if (fld.attributes = 102) AND bRequiredField then response.write "<span class='smallertext'>*</span>"
								response.write "</td>"
						end select
				end select
				response.write "</tr>" & vbcrlf
			next
			if bRecEdit or bRecAdd then
		%>
			<tr><td></td>
				<td><input type="submit" name="cmdSave" value="  Save  " class="cmdflat">
				<% if bPopUps then %>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" onclick="window.close()" class="cmdflat" value="  Close  ">
				<% end if %>
		<%
			end if
		if bRequiredField then response.write "<tr><td colspan=2 class='smallertext'>* means the field is required.</td></tr>"
		%>
			</table>
		</form>
<%	else
		response.write "<p class=""smallerheader"">You do not have permissions to edit records.</p>"
	end if
%>
<!--#include file="te_footer.asp"-->