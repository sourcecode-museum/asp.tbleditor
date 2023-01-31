<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_XML.asp
	' Description: Export data and schema to XML
	' Initiated By Rami Kattan on May 11, 2002
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
	' # May 12, 2002 By Rami Kattan
	' Exported XML is more Microsoft Access XP Valid
	' # May 14, 2002 By Rami Kattan & Peter Stucke
	' Fixed more XML special characters in data and field/table names
	' # May 17, 2002 By Rami Kattan
	' Fixed XML encoding from UTF-8 to ISO-8859-1, no need now for special character encoding
	' # May 22, 2002 By Rami Kattan
	' Enabled response buffering, which increased performance by more then 2200%
	' Made Server.ScriptTimeout dynamic, according to number of records to be exported.
	' Check if browser is still connected, so not to use extra server resources
	' # May 29, 2002 By Rami Kattan
	' Export to XML work also with queries
	' Security check if user can export
	'==============================================================

%><!--#include file="te_config.asp"--><%
lConnID = request.querystring("cid")
sTableName = request.querystring("tablename")
sQuery = request.querystring("q")

XMLTableName = sTableName

if instr(ucase(sTableName), "SELECT") then
	XMLTableName = "QueryResult"
	sTableName = replace(sTableName, ";", "")
end if

sNoJscript = request.querystring("nojs")
if sNoJscript = "1" then
	if not ValidSecurityID("Javaless_browser", request.querystring("SecID")) then
		response.write "Error: you must be <a href=""index.asp"">logged</a> on this site."
		response.end
	end if
end if

if not bAllowExport then
%><!--#include file="te_header.asp"-->
<p class="smallerheader">You have no permission to export data.</p>
<!--#include file="te_footer.asp"-->
<%
	response.end
end if

if (isNS or isIE) and not te_debug then
	Response.ContentType = "application/octet-stream"
else
	Response.ContentType = "text/xml"
end if

if not te_debug then Response.AddHeader "content-disposition", "attachment; filename=" & XMLTableName & ".xml"

%><?xml version="1.0" encoding="ISO-8859-1" ?>
<%
	if sQuery <> "" then
		bQuery = True
		sTableName = replace(sTableName, """", "'")
		XMLExportSchema = false
	end if

	OpenRS arrConn(lConnID)
	
	'Added by Hakan
	'Find the primary key of the given table
	dim aPrimaryKeys
	if arrType(lConnID) = tedbDsn then
		'response.write "Automatic primary key detection is not possible for DSN Connections. " & sSoWhat & "<br><br>"
	else
		set rsX = conn.openSchema(adSchemaPrimaryKeys)

		do while not rsX.eof
			if (rsX("table_name") = sTableName) then
				if sPrimaryKeyFieldName = "" then
					sPrimaryKeyFieldName = rsX("column_name")
				else
					sPrimaryKeyFieldName = sPrimaryKeyFieldName & "," & rsX("column_name")
				end if
			end if
			rsX.movenext
		loop
		rsX.close
		sPrimaryKeyFieldExist = true
	end if

	'Set the primary key field to first field in the list by default
	if sPrimaryKeyFieldName = "" then
		sPrimaryKeyFieldName = 0
		sPrimaryKeyFieldExist = false
	end if

	PrimaryKeys = replace(sPrimaryKeyFieldName, ",", " ")
	PrimaryKeysArr = Split(sPrimaryKeyFieldName, ",")
	XMLTableName = FormatXML(XMLTableName)

if te_debug then XMLExportSchema = false

if XMLExportSchema then
%><root xmlns:xsd="http://www.w3.org/2000/10/XMLSchema" xmlns:od="urn:schemas-microsoft-com:officedata">
<xsd:schema>
<xsd:element name="dataroot">
<xsd:complexType>
<xsd:choice maxOccurs="unbounded">
<xsd:element ref="<% =XMLTableName %>"/>
</xsd:choice>
</xsd:complexType>
</xsd:element>
<xsd:element name="<% =XMLTableName %>">
<xsd:annotation>
<xsd:appinfo>
<% 
	if sPrimaryKeyFieldExist then
%><od:index index-name="PrimaryKey" index-key="<% =PrimaryKeys %> " primary="yes" unique="yes" clustered="no"/>
<% 
end if

	if arrType(lConnID) <> tedbDsn then
		set rs = conn.openSchema(adSchemaIndexes)
		do while not rs.eof
		if rs("table_name") = sTableName AND rs("Index_name") <> "PrimaryKey" then
			if rs("prImary_key") then
				isPrimary = "yes"
			else
				isPrimary = "no"
			end if
			if rs("unIque") then
				isUnique = "yes"
			else
				isUnique = "no"
			end if

			response.write "<od:index index-name=""" & FormatXML(rs("Index_name")) & """ index-key=""" & FormatXML(rs("column_name")) & """ primary=""" & isPrimary & """ unique=""" & isUnique & """ clustered=""no""/>" & vbCrLf
		end if
		rs.movenext
	loop
	rs.close
	end if
%></xsd:appinfo>
</xsd:annotation>
<xsd:complexType>
<xsd:sequence>
<%
		sSQL = "SELECT * FROM [" & sTableName & "] " & sWhere

		on error resume next
		rs.Open sSQL,,,adCmdTable
		for each fld in rs.fields
				jetType = ""
				sqlSType= ""
				ExtraString = ""
				minOccurs = " minOccurs=""0"""
			select case fld.type
				case adBoolean
					jetType = "yesno"
					sqlSType= "bit"
					minOccurs = ""
					ExtraString = " type=""xsd:byte"""
				case adDate
					jetType = "datetime"
					sqlSType= "datetime"
					ExtraString = " type=""xsd:timeInstant"""
				case adInteger
					jetType = "longinteger"
					sqlSType= "int"
				case adUnsignedTinyInt
					jetType = "byte"
					sqlSType= "tinyint"
					ExtraString = " type=""xsd:unsignedByte"""
				case adSmallInt
					jetType = "integer"
					sqlSType= "smallint"
					ExtraString = " type=""xsd:short"""
				case adCurrency
					jetType = "currency"
					sqlSType= "money"
					ExtraString = " type=""xsd:double"""
				case adVarWChar
					jetType = "text"
					sqlSType= "nvarchar"
				case adLongVarWChar
					jetType = "memo"
					sqlSType= "ntext"
				case adLongVarBinary
					jetType = "oleobject"
					sqlSType= "image"
				case adGUID
					jetType = "replicationid"
					sqlSType= "uniqueidentifier"
'				case adHyperLink
'					jetType = "hyperlink"
'					sqlSType= "ntext"
'					ExtraString = " od:hyperlink=""yes"""
				case else
					jetType = "text"
					sqlSType= "nvarchar"
			end select
			if not (fld.attributes and adFldIsNullable) = adFldIsNullable then
				ExtraString = " od:nonNullable=""yes""" & ExtraString
			end if

			if fld.properties("IsAutoIncrement") = true then
				ExtraString = " od:autoUnique=""yes""" & ExtraString
				jetType = "autonumber"
				minOccurs = ""
			end if
%><xsd:element name="<% =FormatXML(fld.name) %>"<% =minOccurs %> od:jetType="<% =jetType %>" od:sqlSType="<% =sqlSType %>"<% =ExtraString %>>
<%
if fld.type = adLongVarWChar or fld.type = adVarWChar then 
%><xsd:simpleType>
<xsd:restriction base="xsd:string">
<xsd:maxLength value="<% =fld.definedsize %>"/>
</xsd:restriction>
</xsd:simpleType>
<%
end if 

if sqlSType = "int" then 

%><xsd:simpleType>
<xsd:restriction base="xsd:integer" /> 
</xsd:simpleType>
<%
   end if 

   if jetType = "replicationid" then 

%><xsd:simpleType>
<xsd:restriction base="xsd:string">
<xsd:maxLength value="38" /> 
</xsd:restriction>
</xsd:simpleType>
<%

   end if 
   
   if sqlSType = "image" then
%><xsd:simpleType>
<xsd:restriction base="xsd:binary">
<xsd:encoding value="base64"/>
<xsd:maxLength value="1476395008"/>
</xsd:restriction>
</xsd:simpleType>
<% end if
%></xsd:element>
<%
		next
		rs.close
%></xsd:sequence>
</xsd:complexType>
</xsd:element>
</xsd:schema>
<%
end if ' For if Export Schema = true

		if request.querystring("nojs") = "1" then
			sFieldValues = request.querystring("chkDel")
			sFieldNames = request.querystring("txtFieldName")
			sFieldTypes = request.querystring("txtFieldType")
		else
			sFieldValues = request.form("chkDel")
			sFieldNames = request.form("txtFieldName")
			sFieldTypes = request.form("txtFieldType")
		end if

		aFieldNames = split(sFieldNames, ";")
		aFieldTypes = split(sFieldTypes, ";")
		aFieldValues = split(sFieldValues, ",")
		
		select case arrType(lConnID)
			case tedbSQLServer
				sDateSeperator = "'"
			case else
				sDateSeperator = "#"
		end select
		for iFld=0 to ubound(aFieldValues)

			sFieldValue2 = split(aFieldValues(iFld), ";")
'			sWhereFields = sWhereFields & "("
			for MPKL = 0 to ubound(aFieldNames)
				sFieldName = trim(aFieldNames(MPKL))
				lFieldType = CLng(aFieldTypes(MPKL))
				sFieldValue = trim(sFieldValue2(MPKL))

				if MPKL > 0 then
					Logic = " AND"
				else
					Logic = " OR"
				end if

				select case lFieldType
					case adDate, adDBDate, adDBTime, adDBTimeStamp
						if isDate(sFieldValue) then 
							sFieldValue = cDate(sFieldValue)
							sFieldValue = month(sFieldValue) & "/" & day(sFieldValue) & "/" & year(sFieldValue)
						end if
						
						if sWhereFields = "" then
							sWhereFields = "([" & sFieldName & "]=" & sDateSeperator & sFieldValue & sDateSeperator & ")"
						else
							sWhereFields = sWhereFields & Logic & " ([" & sFieldName & "]=" & sDateSeperator & sFieldValue & sDateSeperator & ")"
						end if
					case adTinyInt, adSmallInt, adInteger, adBigInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adUnsignedBigInt, adSingle, adDouble, adCurrency, adDecimal, adNumeric, adBoolean
						'Added by Hakan
						'Convert decimal point to dot if it's a comma
						sFieldValue = replace(sFieldValue, ",", ".")
						if sWhereFields = "" then
							sWhereFields = "([" & sFieldName & "]=" & sFieldValue & ")"
						else
							sWhereFields = sWhereFields & Logic & " ([" & sFieldName & "]=" & sFieldValue & ")"
						end if
					case else
						'Added by Hakan
						'Prepare SQL value by replacing single quote with two single quotes
						sFieldValue = replace(sFieldValue, "'", "''")
						if sWhereFields = "" then
							sWhereFields = "([" & sFieldName & "]='" & sFieldValue & "')"
						else
							sWhereFields = sWhereFields & Logic & " ([" & sFieldName & "]='" & sFieldValue & "')"
						end if
				end select
			next ' MPKL
'			sWhereFields = sWhereFields & ")"
		next ' iFld
		if sWhereFields <> "" then sWhere = " WHERE " & sWhereFields


	if request.form("excel_ordering") <> "" then
		sOrderBy = " ORDER BY [" & request.form("excel_ordering") & "] "
		select case request.form("excel_ordering_dir") 
			case "DESC"
				sOrderBy = sOrderBy & " DESC"
			case else
				sOrderBy = sOrderBy & " ASC"
		end select
	end if

	
	if instr(lcase(sTableName), "order by") <> 0 then
		sOrderBy = ""
	end if

	'Added by Danival
	'Modified by Hakan
	bProc = request.querystring("proc")
	if instr(1, ucase(sTableName), "SELECT") then
		if sWhereFields <> "" then
			if instr(1, ucase(sTableName), "WHERE") then
				sWhereQuery = " AND " & sWhereFields
			else
				sWhereQuery = " WHERE " & sWhereFields
			end if
		end if
		sSQL =  sTableName & sWhereQuery & sOrderBy
	else
		if bProc <> "" then
			bRecAdd = False
			bRecEdit = False
			bRecDel = False
			sParamString = request.querystring("paramstring")
			sProcURL = "&proc=1&paramstring=" & sParamString
			sSQL = "EXEC [" & sTableName & "] " & sParamString
		else
			sSQL = "SELECT * FROM [" & sTableName & "]" & sWhere & sOrderBy
		end if
	end if
	on error resume next
	
'	response.write "<BR>" & sSQL & "<BR>"

	rs.CursorLocation = adUseServer
	rs.Open sSQL, conn, adOpenStatic
	
	if err <> 0 then
		response.write "<Error>" & err.description
		if bQuery then
			response.write "<SQL>" & sSQL & "</SQL>" & vbCrLf
			response.write "</Error>" & vbCrLf
		else
			response.write "<SQL>" & sSQL & "</SQL>" & vbCrLf
			response.write "</Error>" & vbCrLf
		end if
		if XMLExportSchema then	response.write "</root>"

		CloseRS
		response.end
	end if

	if XMLExportSchema then
		DatarootNameSpace = "xmlns:xsi=""http://www.w3.org/2000/10/XMLSchema-instance"""
	else
		DatarootNameSpace = "xmlns:od=""urn:schemas-microsoft-com:officedata"""
	end if

	lRecs = rs.RecordCount
	TimeOutAfter = int(lRecs / 600) + 60
	'on my computer (700 @ 889 MHz, 384 MB ram), it made 644 recs per second
	Server.ScriptTimeout = TimeOutAfter

	NumberOfFields = 0
	for each fld in rs.fields
		NumberOfFields = NumberOfFields + 1
	next

	' this section was added to avoid the repetative calls to format the field name
	redim fldName(NumberOfFields)
	redim fldXMLName(NumberOfFields)
	redim fldType(NumberOfFields)
	CurrentField = 0
	for each fld in rs.fields
		fldName(CurrentField) = fld.name
		fldXMLName(CurrentField) = FormatXML(fld.name)
		fldType(CurrentField) = fld.type
		CurrentField = CurrentField + 1
	next

	DoneLoops = 0

	Response.write "<dataroot " & DatarootNameSpace & ">" & vbCrLf
	do while not rs.eof 
		DoneLoops = DoneLoops + 1
		if (DoneLoops MOD 100) = 0 then Response.Flush
		if not Response.IsClientConnected then exit do
		Response.write "<" & XMLTableName & ">" & vbCrLf
		for fldidx = 0 to NumberOfFields - 1
			FieldName = fldXMLName(fldidx)
			Response.Write "<" & FieldName & ">"
				select case fldType(fldidx)
					case adBoolean
						if rs(fldName(fldidx))=true then
							response.write "1"
						else
							response.write "0"
						end if
					case adDate, adDBDate, adDBTime, adDBTimeStamp
						sVal = rs(fldName(fldidx))
						if isDate(sVal) then 
							sVal = cDate(sVal)
							response.write year(sVal) & "-" & LeadingZero(month(sVal), 2) & "-" & LeadingZero(day(sVal),2) & "T" & LeadingZero(Hour(sVal),2) & ":" & LeadingZero(Minute(sVal),2) & ":" & LeadingZero(Second(sVal),2)
						end if					
					case adVarWChar, adLongVarWChar		'Text, Memo
						sVal = rs(fldName(fldidx))
						needCDATA = DataNeedCDATA(sVal)
						if needCDATA then
							Response.write "<![CDATA[" & sVal & "]]>"
						else
							response.write sVal
						end if
					case adLongVarBinary
						Response.write "XML EXPORTER: Currently OLE Data not supported"
					case else
						Response.Write rs(fldName(fldidx))
				end select
				Response.Write "</" & FieldName & ">" & vbCrLf
		next
		response.write "</" & XMLTableName & ">"
		rs.movenext
		if Not rs.eof then response.write vbCrLf
	loop
	CloseRS

	if te_debug then response.write "<sql>" & sSQL & "</sql>"
%>
</dataroot>
<% if XMLExportSchema then response.write "</root>" %>