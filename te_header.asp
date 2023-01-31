<!--#include file="te_count_active_users.asp"--><%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_header.asp
	' Description: Display the header on all pages
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
	' # May 20, 2002 by Rami Kattan
	' Display in the title of the browser a description of the current page
	' Adds for some pages to the body the "onload" call.
	'==============================================================

if bOnPageLoad then
	if bIEAdvancedMode then	ExtraLoadString = " onload=""check_paging();"""
	if bOnPageLoadPage = "index" then ExtraLoadString = " onload=""enableJS();"""
end if

sScript = lcase(Request.ServerVariables("SCRIPT_NAME"))
title_sep = " &raquo;"   'change this according to what seperator you want to use

if instr(sScript, "te_admin_options") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " Table Editor Administration"
elseif instr(sScript, "te_admin") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " Home"
elseif instr(sScript, "te_adddb") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " Add Database"
elseif instr(sScript, "te_compactall") then
		Select Case request("onlybackup")
			case "1"  strAdditionalTitle = strAdditionalTitle & title_sep & " Backup All"
			case else strAdditionalTitle = strAdditionalTitle & title_sep & " Compact All"
		end select
elseif instr(sScript, "te_compactdb") then
		Select Case request("onlybackup")
			case "true"  strAdditionalTitle = strAdditionalTitle & title_sep & " Backup Database"
			case else strAdditionalTitle = strAdditionalTitle & title_sep & " Compact Database"
		end select
elseif instr(sScript, "te_config_help") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " Config Help"
elseif instr(sScript, "te_deleterecord") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " Delete Record"
elseif instr(sScript, "te_dynamic_config") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " Configuraton"
elseif instr(sScript, "te_execproc") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " Stored Procedure Manager " & title_sep & " Execute"
elseif instr(sScript, "te_fieldedit") then
		Select Case len(request("fldname"))
			case 0		strAdditionalTitle = strAdditionalTitle & title_sep & " Add Field"
			case else	strAdditionalTitle = strAdditionalTitle & title_sep & " Edit Field"
		end select
elseif instr(sScript, "te_fieldremove") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " Remove Field"
elseif instr(sScript, "te_help") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " Release Notes"
elseif instr(sScript, "te_indexremove") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " Remove Index"
elseif instr(sScript, "te_listtables") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " " & arrDesc(request("cid"))
elseif instr(sScript, "te_multidelete") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " Multi Delete"
elseif instr(sScript, "te_procmanager") then
		Select Case Request("action")
			case "new"		strAdditionalTitle = strAdditionalTitle & title_sep & " Stored Procedure Manager " & title_sep & " New"
			case "remove"	strAdditionalTitle = strAdditionalTitle & title_sep & " Stored Procedure Manager " & title_sep & " Remove"
			case else		strAdditionalTitle = strAdditionalTitle & title_sep & " Stored Procedure Manager"
		end select
elseif instr(sScript, "te_procedit") then
		select case request("add")
			case 1		strAdditionalTitle = strAdditionalTitle & title_sep & " Stored Procedure Manager " & title_sep & " Add"
			case else	strAdditionalTitle = strAdditionalTitle & title_sep & " Stored Procedure Manager " & title_sep & " Edit"
		end select
elseif instr(sScript, "te_runquery") then
			strAdditionalTitle = strAdditionalTitle & title_sep & " Run Query"
elseif instr(sScript, "te_queryedit") then
		Select Case Request("add")
			case "1"	strAdditionalTitle = strAdditionalTitle & title_sep & " Add Query"
			case else	strAdditionalTitle = strAdditionalTitle & title_sep & " Edit Query"
		end select
elseif instr(sScript, "te_queryinfo") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " Query Information"
elseif instr(sScript, "te_query") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " Run Query"
elseif instr(sScript, "te_searchtable") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " Search Table"
elseif instr(sScript, "te_showrecord") then
		Select Case request.querystring("add")
			case "1"	strAdditionalTitle = strAdditionalTitle & title_sep & " Add Record"
			case else	strAdditionalTitle = strAdditionalTitle & title_sep & " Edit Record"
		end Select
elseif instr(sScript, "te_showschema") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " Schema List"
elseif instr(sScript, "te_showtable") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " " & arrDesc(request("cid")) & title_sep & " " & request("tablename")
elseif instr(sScript, "te_tablecreate") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " Create Table"
elseif instr(sScript, "te_tableedit") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " Edit Table [" & request("tablename") & "]"
elseif instr(sScript, "te_tableremove") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " Remove Table [" & request("tablename") & "]"
elseif instr(sScript, "te_view_active_users") then
        strAdditionalTitle = strAdditionalTitle & title_sep & " View Active Users"
end if

%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>Table Editor 0.8 Beta<% =strAdditionalTitle %></title>
	<link rel="stylesheet" type="text/css" href="te.css">
	<script language="JavaScript" type="text/javascript" src="functions.js"></script>
</head>

<body bgcolor="#ffffff"<% =ExtraLoadString %>>
<table border=0 width="100%">
	<tr><td class="smallheader">
		Table Editor 0.8 Beta
	</td>
	<td class="smallertext" align=right>
		<a href="http://www.2enetworx.com/dev/projects/reportbug.asp?pid=2" style="color: #0000FF; text-decoration: underline;">Report a Bug</a> |
		<a href="http://www.2enetworx.com/dev/projects/recommend.asp?pid=2" style="color: #0000FF; text-decoration: underline;">Recommend a feature</a> |
		<a href="http://www.2enetworx.com/dev/projects/question.asp?pid=2" style="color: #0000FF; text-decoration: underline;">Ask a question</a> |
		<a href="http://www.2enetworx.com/dev/projects/submitsite.asp?pid=2" style="color: #0000FF; text-decoration: underline;">Submit a site</a>
		<% if bAdmin then %> | <a href="te_admin_options.asp" style="color: #0000FF; text-decoration: underline;">Administration</a><% end if %>

	</td></tr>
</table>
<hr class="hrBar" size="1" noshade>