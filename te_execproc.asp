<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_execproc.asp
	' Description: Executes a parameterized query
	' Initiated By Hakan Eskici on Mar 27, 2001
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
<%

	lConnID = request("cid")
	sProcName = request("procname")
	
	conn.Open arrConn(lConnID)

	set cat = server.createobject("adox.catalog")
	cat.ActiveConnection = conn

	'Check if the user approved delete and has permissions
	if (request("cmdExec") <> "") and bQueryExec then
	
		select case arrType(lConnID)
			case tedbAccess
				sDateDelimiter = "#" 
			case tedbSQLServer
				sDateDelimiter = "'"
		end select

		for each p in request.form
			if p <> "cmdExec" then
				if sParamString <> "" then sParamString = sParamString & ","
				'select case cat.Procedures(sProcName).command.parameters(p).type
				'	case adDate, adDBDate, adDBTime, adDBTimeStamp
				'		sParamString = sParamString & sDateDelimiter & request(p) & sDateDelimiter
				'	case adTinyInt, adSmallInt, adInteger, adBigInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adUnsignedBigInt, adSingle, adDouble, adCurrency, adDecimal, adNumeric, adBoolean
						sParamString = sParamString & request(p)
				'	case else
				'		sParamString = sParamString & "'" & request(p) & "'"
				'end select
			end if
		next
		
		'response.write server.htmlencode(sParamString)
		response.redirect TableViewerCompat & "?cid=" & lConnID & "&tablename=" & sProcName & "&proc=1&paramstring=" & server.htmlencode(sParamString)
	else

		
		sParams = ""
		for each p in cat.Procedures(sProcName).command.parameters
			select case p.type
				case adDate, adDBDate, adDBTime, adDBTimeStamp
					sPType = "Date/Time"
				case adTinyInt, adSmallInt, adInteger, adBigInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adUnsignedBigInt, adSingle, adDouble, adCurrency, adDecimal, adNumeric, adBoolean
					sPType = "Numeric"
				case else
					sPType = "Text or Unknown"
			end select
			sPType = sPType & " (" & p.type & ")"
			sParams = sParams & "<tr><td class=""smallertext"">" & p.name & "</td><td class=""smallertext""><input type=""text"" name=""" & p.name & """ class=""tbflat"">&nbsp;" & sPType & "</td></tr>"
		next


%>
<!--#include file="te_header.asp"-->
<table border=0 cellspacing=1 cellpadding=2 bgcolor = "#ffe4b5" width="100%">
	<tr>
		<td class="smallertext">
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <a href="te_listtables.asp?cid=<%=request("cid")%>"><%=arrDesc(request("cid"))%></a> » execute procedure
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

	<p class="smallerheader">Please specify the required parameters</p>
	<form action="te_execproc.asp?cid=<%=lConnID%>&procname=<%=server.htmlencode(sProcName)%>&cmdExec=1" method="post">
		<table border=0>
		<%=sParams%>
		</table>
		<input type="submit" name="cmdExec" class="cmdFlat" value=" Execute ">
	</form>
<%
	end if
%>	
<!--#include file="te_footer.asp"-->