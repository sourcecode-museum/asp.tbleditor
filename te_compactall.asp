<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_compactall.asp
	' Description: Bulk Compact databases
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
<% 
	if bAdmin = False then	
		response.write "Not authorized to view this page."
		%><!--#include file="te_footer.asp"--><%
		response.end
	end if

if not bBulkCompact then
%>
<center><h3><b>Bulk compact/backup is not allowed !!!</b></h3></center>
<% else %>
<br><br>
<table border="0" cellspacing="1" cellpadding="2" bgcolor="#ffe4b5" width="100%">
	<tr>
		<td class="smallertext">
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <% if bOnlyBackup then response.write "Backup" else response.write "Compact" %> Databases
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
	dim lConnID
		sub ShowForm
	%>
		<p class="smallertext">
		<% If not bOnlyBackup Then %>
			All Access Databases will be compacted.<br>
			After compacting, the original files will be renamed to ".(date/time).bak" extension as a backup.<br><br>
			During the compact process, ".temp.mdb" files will be created, and after the compact is successful, they will be deleted.
			<br><br>
			Compacting may take some time depending on your databases
			sizes. The changes made to the database during the compacting
			will not be saved.
		<% Else %>
			All Access Databases will be backed-up.<br>
			The backup files will have the name with ".(date/time).bak" extension as a backup.
			<br><br>
			Backing up may take some time depending on your databases sizes. The changes made to the database during the backup will not be saved.			
		<% End If %>
		</p>
		<p class="smallerheader">
			<% If not bOnlyBackup Then %>
				Are you sure that you want to compact all databases?
			<% Else %>
				Are you sure that you want to backup all databases?
			<% End If %>
		</p>
		<a href="te_compactall.asp?<% =request.querystring %>&sure=1">Yes</a>&nbsp;
		<a href="<%=request.servervariables("http_referer")%>">No</a>	
	<%
	end sub

	if request("sure")<>""  then
	%>
	<h1><a name="top">Compact Databases</a></h1>
	<p>Please wait, compacting loaded connections...</p>
	<p>&nbsp;</p>
	<%
		For lConnID = 1 to iTotalConnections
			if arrType(lConnID) = tedbAccess Then
				Compact lConnID	
			end if
		Next
		Response.Write "<P><B>Complete</B>"
	else
		ShowForm
	end if
	
	Sub RenameFile(sFrom, sTo)
		Dim fso, f
		Set fso = CreateObject("Scripting.FileSystemObject")
		sDateTimeStamp = "." & year(now) & LeadingZero(month(now),2) & LeadingZero(day(now),2) & "-" & LeadingZero(hour(now),2) & LeadingZero(minute(now),2) & LeadingZero(second(now),2)

		ShortFileName = arrDBs(lConnID)
		sFileFromParts = split(ShortFileName, ";")
		ShortFileName = sFileFromParts(0)

		If not bOnlyBackup Then
			fso.MoveFile sTo, left(Server.MapPath(ShortFileName), InStr(Server.MapPath(ShortFileName), ".mdb")-1) & sDateTimeStamp & ".bak"
			fso.MoveFile sFrom, sTo
		Else
			fso.CopyFile sTo, sTo & sDateTimeStamp & ".bak"
		End If
	   
	End Sub
	
	Sub Compact(lConnID)

		ShortFileName = arrDBs(lConnID)
		sFileFromParts = split(ShortFileName, ";")
		ShortFileName = sFileFromParts(0)

		If not bOnlyBackup Then
			response.write "<p>Compacting '<b>" & arrDesc(lConnID) & "</b>'..."
		else
			response.write "<p>Backing up '<b>" & arrDesc(lConnID) & "</b>'..."
		end if

		sFileFrom = Server.MapPath(ShortFileName)
		sFileTo = left(sFileFrom, InStr(sFileFrom, ".mdb")-1) & ".temp.mdb"		
			If not bOnlyBackup Then
				sConnFrom = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & sFileFrom
				sConnTo   = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & sFileTo
				set jro = server.createobject("jro.JetEngine")
				jro.CompactDatabase sConnFrom, sConnTo
				
				if err <> 0 then
					bError = True
					response.write "Error Compacting Database:<br>" & err.description
				end if
				err.clear
			end if

			RenameFile sFileTo, sFileFrom
			
			if err <> 0 then
				bError = True
				response.write "Error Renaming Tempory File:<br>" & err.description 
			end if
			
			if not bError then
				If not bOnlyBackup Then
					response.write "<br>Compacting successful."
				else
					response.write "<br>Backing up successful."
				end if
			else
				response.write "<br>Errors occured."
			end if
	
			response.write "</p>" & vbCrLf

	end sub

%>
<br><br>
<% end if %>
<!--#include file="te_footer.asp"-->