<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_dynamic_config.asp
	' Description: Displays the configurations of TE
	' Initiated By Rami Kattan on May 20, 2002
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
	' Drop Down selector for DefaultPerPage option
	'==============================================================
bPopUps2 = (request.querystring("popup") <> "")
bPopUps2 = true
%>
<!--#include file="te_config.asp"-->
<!--#include file="te_header.asp"-->
<table border=0 cellspacing=1 cellpadding=2 bgcolor="#ffe4b5" width="100%">
	<tr>
		<td class="smallertext">
		<% if bPopUps and not bPopUps2 then %>
			Table Editor Administration » Configurations » Change Configurations
		<% else %>
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <a href="te_listtables.asp?cid=0">Table Editor Administration</a> » Configurations » Change Configurations
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
<%
	'disallow connections to admin table for nonadmin users 11/14/01
	lConnID = 0
	if (lConnID = 0) and bAdmin = False then	
		response.write "Not authorized to view this connection."
		%><!--#include file="te_footer.asp"--><%
		response.end
	end if

	sTableName = "config"
	sFieldNames = "ID"
	sFieldValues = 1
	sFieldTypes = 3
	sQuery = request.querystring("q")
	if request.querystring("add") then bAdd = True else bAdd = False

    sParentName = server.urlencode(sTableName)

	OpenRS arrConn(lConnID)

	set adox = server.createobject("adox.catalog")
	adox.ActiveConnection = conn

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
     
		if request.form("cmdSave") <> "" and bRecEdit = True then
			if bAdd then
				rs.AddNew
			end if
			
			for each fld in rs.fields
				ifield = ifield + 1
				'If field is AutoIncrement, just skip it.
				
				
				'Added By Brad Orgill
				donada = false
				fldName = fld.name
			   
				set fldx = AdoX.Tables(sTableName).Columns(fld.name)

				'modified By Brad Orgill
				if (not fldx.properties("Autoincrement")) AND (not donada) then
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
%>
			<form action="te_dynamic_config.asp?cid=0&tablename=config&fld=ID&val=<%=sFieldValue%>&fldtype=3&ipage=&add=<%=bAdd%>&cmdsave=1<% if bQuery then response.Write("&q=1") %><% if bPopUps2 then response.Write("&popup=no") %>" method="post" name="frm">

			<table border="0" cellspacing="2" cellpadding="3" bgcolor="#ffe4b5" width="400">
		<%

		FieldNames_short = array("EncodeHTML"  ,"Relation","ShowConnDetails"        ,"ConvertNull"  ,"ConvertNumericNull"  ,"ConvertDateNull"  ,"MaxShowLen"        ,"BulkDelete" ,"BulkCompact" ,"HighLight" ,"ExportExcel"    ,"CountActiveUsers"  ,"PageSelector" ,"ComboTables" ,"IEAdvancedMode"  ,"ExportXML" ,"PopUps"    ,"DefaultPerPage"         ,"HighSecurityLogin")
		FieldNames_long  = array("Encode HTML" ,"Relation","Show Connection Details","Convert Nulls","Convert Numeric Null","Convert Date Null","Max Length to Show","Bulk Delete","Bulk Compact","High Light","Export To Excel","Count Active Users","Page Selector","Combo Tables","IE Advanced Mode","Export XML","Use Popups","Default Record Per Page","High Security Login")

		function HelpURL(info)
			if bPopups then URL = "javascript:openWindow('"
			URL = URL & "te_config_help.asp?ID=" & info & ""
			if bPopups then URL = URL & "')"
			HelpURL = URL
		end function

			lFields = rs.Fields.count
			for each fld in rs.fields
				if not bAdd then
					sValue = rs(fld.name)
				end if

				if bEncodeHTML then
					if sValue <> "" then sValue = server.htmlencode(sValue)
				end if

				FieldName = fld.name
				NameID = 0
				do while NameID <= ubound(FieldNames_short)
					if ucase(FieldNames_short(NameID)) = ucase(FieldName) then
						FieldName = FieldNames_long(NameID)
						Exit do
					end if
					NameID = NameID + 1
				loop

				if FieldName <> "ID" then 
					response.write vbcrlf & "<tr>"
					response.write "<td class=""smallerheader"">" & FieldName & " <sup><a href=""" & HelpURL(NameID) & """>?</a></sup>"& "</td>" & vbcrlf
				end if
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
												if cint(sValue) = cint(RSConsulta(0)) then
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
							end if
							
							if fld.name = "DefaultPerPage" then 
								response.write "<td><select name=""DefaultPerPage"" class=""smallertext"">"
								response.write "<option value=""5""" & isSelected(5, sValue) & ">5</option>"
								response.write "<option value=""10""" & isSelected(10, sValue) & ">10</option>"
								response.write "<option value=""15""" & isSelected(15, sValue) & ">15</option>"
								response.write "<option value=""20""" & isSelected(20, sValue) & ">20</option>"
								response.write "<option value=""30""" & isSelected(30, sValue) & ">30</option>"
								response.write "<option value=""40""" & isSelected(40, sValue) & ">40</option>"
								response.write "<option value=""0""" & isSelected(0, sValue) & ">All</option>"
								response.write "</select></td>" & vbcrlf
								flag = true
							end if

							set fldx = AdoX.Tables(sTableName).Columns(fld.name)
							if not flag then
								if not fldx.properties("autoincrement") then
									response.write "<td><input class=""tbflat"" size=6 type=""text"" name=""" & fld.name &  """ value=""" & sValue & """>"
									if (fld.attributes = 102) AND bRequiredField then response.write "<span class='smallertext'>*</span>"
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
				if FieldName <> "ID" then response.write "</tr>" & vbcrlf
			next
			if bRecEdit then
		%>
			<tr><td>&nbsp;</td>
				<td><input type="submit" name="cmdSave" value="  Save  " class="cmdflat">
				<% if bPopUps and not bPopUps2 then %>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" onclick="window.close()" class="cmdflat" value="  Close  ">
				<% end if %>
		<%
			end if
		if bRequiredField then response.write "<tr><td colspan=2 class='smallertext'>* means the field is required.</td></tr>"
		%>
			</table>
		</form>
<!--#include file="te_footer.asp"-->