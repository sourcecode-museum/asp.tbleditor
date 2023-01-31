<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_multidelete.asp
	' Description: Deletes multiple records
	' Initiated By Hakan Eskici on Jan 11, 2001
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
	' # Mar 29, 2001 by Hakan Eskici
	' Added support for multiple primary keys
	'==============================================================
%>
<!--#include file="te_config.asp"-->
<%

	'Permission check
	if bProtected then
		if not bRecDel then 
			%><!--#include file="te_header.asp"--><%
			response.write "You don't have permission to delete!<br>"
			%><!--#include file="te_footer.asp"--><%
			response.end
		end if
	end if

	sNoJscript = request.querystring("nojs")
	if sNoJscript = "1" then
		if not ValidSecurityID("Javaless_browser", request.querystring("SecID")) then
			response.write "<p class=""smallertext"">Error: you must be <a href=""index.asp"">logged</a> on this site.</p>"
			response.end
		end if
	end if

	function DeleteRecord(sFldVal)

		sFieldNames = request.form("txtFieldName")
		sFieldTypes = request.form("txtFieldType")
		
		aFieldNames = split(sFieldNames, ";")
		aFieldTypes = split(sFieldTypes, ";")
		aFieldVals = split(sFldVal, ";")
		
		select case arrType(lConnID)
			case tedbSQLServer
				sDateSeperator = "'"
			case else
				sDateSeperator = "#"
		end select

		for iFld = 0 to ubound(aFieldNames)
			sFieldName = aFieldNames(iFld)
			lFieldType = CLng(aFieldTypes(iFld))
			sFieldValue = aFieldVals(iFld)

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
	
		
		'Added by Danival
		if instr(1, ucase(sTableName), "SELECT") then
			sSQL =  sTableName
		else
			sSQL = "SELECT * FROM [" & sTableName & "]"
		end if
		
		'Modified by Hakan
		'Open the table/query first
		rs.ActiveConnection = conn
        rs.Open sSQL, , , adCmdTable

		'Filter the records with Where Statement
		rs.Filter = sWhereFields
		
		'If able to find, delete
		if not rs.eof or rs.bof then
			on error resume next
			rs.delete
			if err <> 0 then
				bError = True
			else
				bError = False
			end if
		else
			bError = True
		end if
		rs.close
		DeleteRecord = not bError
	end function
%>
<!--#include file="te_header.asp"-->
<table border=0 cellspacing=1 cellpadding=2 bgcolor="#ffe4b5" width="100%">
	<tr>
		<td class="smallertext">
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <a href="te_listtables.asp?cid=<%=request("cid")%>"><%=arrDesc(request("cid"))%></a> » <%if bQuery then response.write "Query" else response.write "[Table :" & request("tablename") & "]"%>
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
	
	if request("chkDel") = "" then
		response.write "<p>No records to delete!</p>"
		response.write "<a href=""" & request.servervariables("http_referer") & """>Go back</a>"
		%><!--#include file="te_footer.asp"--><%
		response.end
	else

		dim lConnID
		dim sTableName
		
		lConnID = request("cid")
		sTableName = request("tablename")
		sQuery = request("q")
		
		conn.open arrConn(lConnID)
	
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

		if request("cmdYes") = "" then
		%>
		<p class="smallerheader">Are you sure that you want to delete <%=ubound(aFieldValues)+1%> records?</p>
		
		<table border=0>
			<tr><td>
				<form action="te_multidelete.asp?cid=<%=lConnID%>&tablename=<%=server.urlencode(sTableName)%>" method="post">
					<input type="hidden" name="txtFieldName" value="<%=sFieldNames%>">
					<input type="hidden" name="txtFieldType" value="<%=sFieldTypes%>">
					<%
					for i=0 to ubound(aFieldValues)
					%>
					<input type="hidden" name="chkDel" value="<%=trim(aFieldValues(i))%>">
					<%
					next
					
					%>
					<input type="submit" name="cmdYes" value=" Yes " class="cmdflat">
				</form>
			</td><td>
				<form action="<%=request.servervariables("http_referer")%>" method="post">
					<input type="submit" name="cmdNo" value="  No  " class="cmdflat">
				</form>
			</td></tr>
		</table>
		<%
		else

			for iRec = 0 to ubound(aFieldValues)
				if DeleteRecord(aFieldValues(iRec)) <> True then
					bErr = True
				end if
			next
		
			if bErr = True then
				%><!--#include file="te_header.asp"--><%
				response.write "<p class=""smallerheader"">Cannot delete the record.</p>"
				response.write "<strong>Error Reported</strong>: " & err.description & "<br>"
				response.write "<strong>SQL</strong>: " & sSQL & "<br>"
				response.write "<strong>Filter</strong>: " & sWhereFields & "<br>"
				%><!--#include file="te_footer.asp"--><%
				response.end
			end if
			
			response.write "<br>" & ubound(aFieldValues) + 1 & " records deleted from the table '" & sTableName & "'.<br><br>"
			response.write "<a href="""& TableViewerCompat & "?cid=" & lConnID & "&tablename=" & sTableName & """>Go back</a>"
		end if
	end if
%>
<!--#include file="te_footer.asp"-->