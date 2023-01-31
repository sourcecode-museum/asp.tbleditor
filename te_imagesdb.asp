<!--#include file="te_config.asp"-->
<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_imagesdb.asp
	' Description: Sends binary image to response stream
	' Initiated By ? on ?
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
	' # May 20, 2002 by Rami Kattan
	' Images are loaded directly from database, not session variables
	'==============================================================

lConnID = request.querystring("cid")
sTableName = request.querystring("tablename")

sFieldNames = request.querystring("fld")
sFieldValues = request.querystring("val")
sFieldTypes = request.querystring("fldtype")
sOLEField = request.querystring("olefield")

OpenRS2 arrConn(lConnID)

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
sSQL = "SELECT * FROM [" & sTableName & "]" & sWhere

rs2.Open sSQL, , , adCmdTable

Response.Expires = 0
Response.Buffer = TRUE

'display image
    
Response.ContentType = "image/bmp"
'size = Session(sessVarSize)
'blob = Session(sessVarblob)
Response.BinaryWrite rs2(sOLEField)

CloseRS2

response.End
%>