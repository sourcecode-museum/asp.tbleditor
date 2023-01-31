<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_deleterecord.asp
	' Description: Deletes a record given key field name and value
	' Initiated By Hakan Eskici on Nov 01, 2000
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
	' # Nov 16, 2000 by Hakan Eskici
	' Added Navigation bar
	' # Mar 27, 2001 by Hakan Eskici
	' Changed the method for finding records to be deleted
	' Added support for deleting from user specified select queries
	' Added support for multiple keys
	' # Nov 14, 2001 by Jeff Wilkinson (jwilkinson@mail.com)
	' security fix entered per Dilyias suggested fix 10/29/01
	' prevents nonadmin users from accessing the admin db (conn=0)
	'==============================================================
%>
<!--#include file="te_config.asp"-->
<%

	'Check if the user approved delete and has permissions to delete
	if (request("sure") <> "") and bRecDel then

		'disallow connections to admin table for nonadmin users 11/14/01
		lConnID = request("cid")
		if not( (lConnID = 0) and bAdmin = False) then	

			sTableName = request("tablename")
			sFieldNames = request.querystring("fld")
			sFieldValues = request.querystring("val")
			sFieldTypes = request.querystring("fldtype")
			
			OpenRS arrConn(lConnID)

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
		
			'Added by Danival
			if instr(1, ucase(sTableName), "SELECT") then
				sSQL =  sTableName
			else
				sSQL = "SELECT * FROM [" & sTableName & "]"
			end if
			
			'Modified by Hakan
			'Open the table/query first
			rs.Open sSQL, , , adCmdTable

			'Filter the records with Where Statement
			rs.Filter = sWhereFields
			
			'If able to find, delete
			if not rs.eof or rs.bof then
				on error resume next
				rs.delete
				if err <> 0 then
					'Display the error
					%><!--#include file="te_header.asp"--><%
					response.write "<p class=""smallerheader"">Cannot delete the record.</p>"
					response.write "<strong>Error Reported</strong>: " & err.description & "<br>"
					response.write "<strong>SQL</strong>: " & sSQL & "<br>"
					response.write "<strong>Filter</strong>: " & sWhereFields & "<br>"
					%><!--#include file="te_footer.asp"--><%
					response.end
				else
					if bPopUps then
						response.write "<script language=""JavaScript"" type=""text/javascript"">window.close()</script>"
					else
						response.redirect TableViewerCompat & "?cid=" & lConnID & "&tablename=" & sTableName & "&ipage="  & request("ipage")
					end if
				end if
			else
				%><!--#include file="te_footer.asp"--><%
				response.write "Cannot filter records to be deleted.<br><br>"
				response.write "Table: " & sSQL & "<br>"
				response.write "Filter: " & sWhereFields & "<br>"
				%><!--#include file="te_footer.asp"--><%
				response.end
			end if
		end if
	end if

	if bPopUps then
		NoButtonAction = "javascript:window.close()"
	else
		NoButtonAction = request.servervariables("http_referer")
	end if
%>
<!--#include file="te_header.asp"-->
<table border=0 cellspacing=1 cellpadding=2 bgcolor = "#ffe4b5" width=100%>
	<tr>
		<td class="smallertext">
		<% if bPopUps then %>
			<%=arrDesc(request("cid"))%> » Table [<%=request("tablename")%>] » Delete Record
		<% else %>
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <a href="te_listtables.asp?cid=<%=request("cid")%>"><%=arrDesc(request("cid"))%></a> » <a href="<% =TableViewerCompat %>?cid=<%=request("cid")%>&tablename=<%=server.urlencode(request("tablename"))%>&ipage=<%response.write(request("ipage"))%><% if bQuery then response.write("&q=1") end if %>">Table [<%=request("tablename")%>]</a> » Delete Record
		</td>
		<td class="smallerheader" width=130 align=right>
		<%
			if bProtected then 
				response.write session("teFullName")
				response.write " (<a href=""te_logout.asp"">logout</a>)" 
			end if
		end if ' For popups
		%>
		
		</td>
	</tr>
</table>
<br>

	<p class="smallerheader">Are you sure that you want to delete the record?</p>
	<a href="te_deleterecord.asp?<%=request.querystring%>&sure=1">Yes</a>&nbsp;
	<a href="<% =NoButtonAction %>">No</a>
			
<!--#include file="te_footer.asp"-->