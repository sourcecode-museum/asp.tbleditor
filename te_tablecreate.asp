<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_tablecreate.asp
	' Description: Creates a new table
	' Initiated By Hakan Eskici on Nov 17, 2000
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
	' # May 30, 2002 by Rami Kattan
	' Security check if user can add tables
	'==============================================================
%>
<!--#include file="te_config.asp"-->
<%
	sub ShowForm
%>
	<!--#include file="te_header.asp"-->
	<table border=0 cellspacing=1 cellpadding=2 bgcolor="#ffe4b5" width="100%">
		<tr>
			<td class="smallertext">
				<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <a href="te_listtables.asp?cid=<%=request("cid")%>"><%=arrDesc(request("cid"))%></a> » Create Table
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
	
	<br><%
		if bTableAdd then
	%><p class="smallerheader">
		Enter the table name.<br><br><%=sErr%>
	</p>
	<form action="te_tablecreate.asp?cid=<%=lConnID%>" method="post">
	<table border=0>
		<tr>
			<td class="smallerheader">Table Name</td>
			<td><input type="text" name="txtTableName" class="tbflat"></td>
		</tr>
		<tr>
			<td></td>
			<td><input type="submit" name="cmdCreate" value=" Create " class="cmdflat"></td>
		</tr>
	</table>
	</form><%
		else
			response.write "<p class=""smallerheader"">You have no permission to add tables.</p>"
		end if
	%><!--#include file="te_footer.asp"-->
	<%
	end sub

	lConnID = request("cid")
	sTableName = request("txtTableName")
	
	on error resume next
	if request("cmdCreate") <> "" and bTableAdd then
	
		OpenRS arrConn(lConnID)
		
		if sTableName = "" then
			sErr = "Error : <br>Please specify a table name."
			ShowForm
			response.end
		end if
		
		sSQL = "CREATE TABLE [" & sTableName & "]"
		conn.execute sSQL
		if err <> 0 then
			sErr = "Error : <br>" & err.description
			bErr = True
			if te_debug then sErr = sErr  & "<br>" & sSQL

			conn.close
			set conn=nothing
			set rs=nothing

			ShowForm
		else
			conn.close
			set conn=nothing
			set rs=nothing

			response.redirect "te_fieldedit.asp?cid=" & lConnID & "&tablename=" & sTableName
		end if
	
	else
		ShowForm
	end if
%>