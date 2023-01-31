<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_queryedit.asp
	' Description: Creates or edits a query
	' Initiated By Hakan Eskici on Nov 22, 2000
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
	' Nov 23, 2000 by Hakan Eskici
	' Changed the query name assignment which caused a bug in
	' creating and editing queries.
	'==============================================================
%>
<!--#include file="te_config.asp"-->
<%

	sub ShowForm
%>
	<!--#include file="te_header.asp"-->
	<table border=0 cellspacing=1 cellpadding=2 bgcolor = "#ffe4b5" width=100%>
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
	
	<br>
	<p class="smallerheader">
		<%=sAction%><br><br><%=sErr%>
	</p>
	<form action="te_queryedit.asp?cid=<%=lConnID & sAdd%>&queryname=<%=sQueryName%>" method="post">
	<table border=0>
		<tr>
			<td class="smallerheader">Query Name</td>
			<td><input type="text" name="txtQueryName" class="tbflat" value="<%=sQueryName%>"></td>
		</tr>
		<tr>
			<td class="smallerheader">SQL</td>
			<td>
			<textarea cols="60" rows="8" name="txtSQL" id="txtSQL" class="tbflat" value="SELECT * FROM r"><%=sSQL%></textarea>
			</td>
		</tr>
		<tr>
			<td></td>
			<td><input type="submit" name="cmdSave" value=" Save " class="cmdflat"><%
				if te_debug then %>
			<button onclick="OpenSQLBuilder()" class="cmdflat">SQL Builder</button>
			<% end if %></td>
		</tr>
	</table>
	</form>
	<!--#include file="te_footer.asp"-->
	<%
	end sub

	lConnID = request("cid")
	sQueryName = request("txtQueryName")
	sSQL = request("txtSQL")

	sQueryName = request("queryname")
	sNewQueryName = request("txtQueryName")

	on error resume next

	if request("add") <> "" then
		bAdd = True
		sAdd = "&add=1"
		sAction = "Create new Query"
	else
		bAdd = False
		sAction = "Edit Query"
		if request("cmdSave") <> "" then
			sSQL = request("txtSQL")
		else

			OpenRS arrConn(lConnID)
	
			set cmd = server.createobject("adodb.command")
			set cat = server.createobject("adox.catalog")
	
			set cat.ActiveConnection = conn
			set cmd.ActiveConnection = conn
			
			set views = cat.views
			
			set cmd = views(sQueryName).Command
			sSQL = cmd.CommandText
	
			conn.close
			set conn=nothing
			set rs=nothing
			set cmd=nothing
			set cat=nothing

		end if
	end if
	
	if request("cmdSave") <> "" then
	
		OpenRS arrConn(lConnID)

		set cmd = server.createobject("adodb.command")
		set cat = server.createobject("adox.catalog")

		set cat.ActiveConnection = conn
		set cmd.ActiveConnection = conn
		
		cmd.CommandType = adCmdText
		cmd.CommandText = sSQL

		set views = cat.views
		if bAdd then
			if sQueryName = "" then
				sQueryName = sNewQueryName
			end if
			views.append sQueryName, cmd
		else
			views(sQueryName).Command = cmd
			'Cannot rename
			'views(sQueryName).Name = sNewQueryName
		end if
		
		if err <> 0 then
			sErr = "Error : <br>" & err.description
			bErr = True

			conn.close
			set conn=nothing
			set rs=nothing
			set cmd=nothing
			set cat=nothing

			ShowForm
		else
			conn.close
			set conn=nothing
			set rs=nothing
			
			response.redirect "te_listtables.asp?cid=" & lConnID
		end if
	
	else
		ShowForm
	end if
%>