<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_queryedit.asp
	' Description: Edits or creates a query
	' Initiated By Hakan Eskici on Nov 22, 2000
	'--------------------------------------------------------------
	' Copyright (c) 2001, 2eNetWorX/dev.
	'
	' TableEditoR is distributed with General Public License.
	' Any derivatives of this software must remain OpenSource and
	' must be distributed at no charge.
	' (See license.txt for additional information)
	'
	' See Credits.txt for the list of contributors.
	'
	' Change Log:
	'-------------------------------------------------------------
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
				<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <a href="te_listtables.asp?cid=<%=request("cid")%>"><%=arrDesc(request("cid"))%></a> » <%=sAction%>
			</td>
			<td class="smallerheader" width=130 align=right>
			<%
			if bProtected then 
				response.write session("teFullName")
				response.write "<a href=""te_logout.asp""> (logout)</a>" 
			end if
			%>
			</td>
		</tr>
	</table>
	
	<br>
	<p class="smallerheader">
		<%=sAction%><br><br><%=sErr%>
	</p>
	<form action="te_procedit.asp?cid=<%=lConnID & sAdd%>&queryname=<%=sQueryName%>" method="post">
	<table border=0>
		<tr>
			<td class="smallerheader">Procedure Name</td>
			<td><input type="text" name="txtQueryName" class="tbflat" value="<%=sQueryName%>"></td>
		</tr>
		<tr>
			<td class="smallerheader">SQL</td>
			<td>
			<textarea cols=60 rows=8 name="txtSQL" class="tbflat"><%=sSQL%></textarea>
			</td>
		</tr>
		<tr>
			<td class="smallerheader" colspan=2>Syntax:
				Write "PARAMETERS 1st_par_name 1st_par_type, 2nd_par_name 2nd_par_type, ... ;"<br> and then the Sql query text followed by ";"<br>
				ex. PARAMETERS pippo Long; SELECT * FROM users WHERE id_user > pippo;<br>
				<font color="red"><u><b>ATTENTION: The Sql Query created/modified in this way<br> will not be readable(nor visible, nor editable) by Ms Access 2000<a href="http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnacc2k/html/adocreateq.asp">(Read More)</a>.</b></u></font>
			</td>
		</tr>

		<tr>
			<td></td>
			<td><input type="submit" name="cmdSave" value=" Save " class="cmdflat"></td>
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

	'on error resume next

	if request("add") <> "" then
		bAdd = True
		sAdd = "&add=1"
		sAction = "Create new Stored Procedure"
	else
		bAdd = False
		sAction = "Edit Stored Procedure"
		if request("cmdSave") <> "" then
			sSQL = request("txtSQL")
		else

			OpenRS arrConn(lConnID)
	
			set cmd = server.createobject("adodb.command")
			set cat = server.createobject("adox.catalog")
	
			set cat.ActiveConnection = conn
			set cmd.ActiveConnection = conn
			
			set views = cat.procedures
			
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
		
		cmd.CommandText = sSQL
		
		set views = cat.procedures
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
			bErr = True'

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