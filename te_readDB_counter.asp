<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_readDB_counter.asp
	' Description: Generate CSV file giving total records in a table
	' Initiated By Rami Kattan on Apr 22, 2002
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

	Response.ContentType = "text/csv"

' Get the requested number of records per page
%>
<!--#include file="te_config.asp"-->
<%
	lConnID = request("cid")
	sTableName = request("tablename")
	sQuery = request("q")
	
	'------------------------------
	'added 8/10/01 by j.wilkinson, jwilkinson@mail.com
	'added a check for nonAdmin users trying to view the admin table
	'This is just checking that the connection ID = 0, assumes that
	'non-admin users have no legitimate reason to get to that db at all.
	'  note that this may not protect against using queries to view
	'  this db and table
	if lConnID=0 and not bAdmin then
		response.end
	end if
	'------------------------------

    const csvchar = ","
	
	if sQuery <> "" then
		bQuery = True
		sTableName = replace(sTableName, """", "'")
	end if

	OpenRS arrConn(lConnID)
	
	'Added by Danival
	'Modified by Hakan
	bProc = request.querystring("proc")
	if instr(1, ucase(sTableName), "SELECT") then
		sSQL =  sTableName & sOrderBy
	else
		if bProc <> "" then
			sParamString = request.querystring("paramstring")
			sSQL = "EXEC [" & sTableName & "] " & sParamString
		else
			sSQL = "SELECT * FROM [" & sTableName & "]" & sWhere & sOrderBy
		end if
	end if
	
	on error resume next
	
	rs.CursorLocation = adUseServer
	rs.Open sSQL, conn, adOpenStatic
	
	if err <> 0 then
		response.write "Error: " & err.description & "<br><br>"
		if bQuery then
			response.write "SQL: " & sSQL & "<br><br>"
		end if
		CloseRS
		response.end
	end if

	on error goto 0
	
	'Performance Issue:
	'Getting the recordset properties may take long time for tables with many records
	lRecs = rs.RecordCount
	
	response.write "TotalRecords" & VbCrLf
	response.write lRecs

	CloseRS
%>