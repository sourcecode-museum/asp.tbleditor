<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_compactdb.asp
	' Description: Compacts an Access Mdb file
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
	' # April 18, 2002 by Rami Kattan
	' Temporary files are in the same folder as the database
	' as some ISP doesn't allow write to normal html folders
	' # May 21, 2002 by Rami Kattan
	' Fixes few bugs, when filename has ; for password
	' # May 30, 2002 by Rami Kattan
	' Backup now makes a copy of the original file itself, not 
	' a compacted version
	'==============================================================

bOnlyBackup = (request.querystring("onlybackup") = "true")

%>
<!--#include file="te_config.asp"-->
<!--#include file="te_header.asp"-->
<br><br>
<table border=0 cellspacing=1 cellpadding=2 bgcolor="#ffe4b5" width="100%">
	<tr>
		<td class="smallertext">
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <% if bOnlyBackup then response.write "Backup" else response.write "Compact" %> Database
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
<img src="images/compact_db.gif" align="right">
<%
	lConnID = request("cid")
	on error resume next
	sub ShowForm
	%>
		<p class="smallertext">
		<% if not bOnlyBackup Then %>
			The original database "<b><%=sFileFrom%></b>" will be compacted.<br>
			After compacting, the original file will be renamed to ".(date/time).bak" extension as a backup.<br><br>
			During the compact process, a ".temp.mdb" file will be created, and after the compact is successful, it will be deleted.
			<br><br>
			Compacting may take some time depending on your database size. The changes made to the database during the compact will not be saved.
		<% else %>
			The original database "<b><%=sFileFrom%></b>" will be backed-up.<br>
			The backup file will have the name with ".(date/time).bak" extension as a backup.
			<br><br>
			Backing up may take some time depending on your database size. The changes made to the database during the backup will not be saved.			
		<% end if %>
		</p>
		<p class="smallerheader">
			<% If not bOnlyBackup Then %>
			Are you sure that you want to compact the database?
			<% Else %>
				Are you sure that you want to backup the database?
			<% End If %>
		</p>
		<a href="te_compactdb.asp?<%=request.querystring%>&sure=1">Yes</a>&nbsp;
		<a href="<%=request.servervariables("http_referer")%>">No</a>	
	<%
	end sub

	ShortFileName = arrDBs(lConnID)
	sFileFromParts = split(ShortFileName, ";")
	ShortFileName = sFileFromParts(0)

	Sub RenameFile(sFrom, sTo)
	   Dim fso, f
	   Set fso = CreateObject("Scripting.FileSystemObject")
	   sDateTimeStamp = "." & year(now) & LeadingZero(month(now),2) & LeadingZero(day(now),2) & "-" & LeadingZero(hour(now),2) & LeadingZero(minute(now),2) & LeadingZero(second(now),2)
	   If not bOnlyBackup Then
			fso.MoveFile sTo, Server.MapPath(ShortFileName & sDateTimeStamp & ".bak")
			fso.MoveFile sFrom, sTo
	   Else
			fso.CopyFile sTo, sTo & sDateTimeStamp & ".bak"
	   End If
	End Sub

	sFileFrom = Server.MapPath(ShortFileName)
	sFileTo = left(sFileFrom, InStr(sFileFrom, ".mdb")-1) & ".temp.mdb"
	
	sConnFrom = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & sFileFrom
	sConnTo   = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & sFileTo

	if request("sure")<>""  then

		if bTableEdit then
			If not bOnlyBackup Then
				set jro = server.createobject("jro.JetEngine")
				jro.CompactDatabase sConnFrom, sConnTo
			
				if err <> 0 then
					bError = True
					response.write "Error:<br>" & err.description
				end if
			end if
			
			RenameFile sFileTo, sFileFrom
			
			if err <> 0 then
				bError = True
				response.write "Error:<br>" & err.description
			end if
			
			if not bError then
				If not bOnlyBackup Then
				response.write "<br><br>Compacting successful."
				Else
					response.write "<br><br>Backing up successful."
				End If
			else
				response.write "<br><br>Errors occured."
			end if
			
		else
			response.write "You dont have permissions to compact databases."
		end if
	else
		ShowForm
	end if
%>
<br><br>
<!--#include file="te_footer.asp"-->