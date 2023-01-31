<!--#include file="te_includes.asp"-->
<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_formaction.asp
	' Description: multidelete & exports for non-javascript browser
	' Initiated By Rami Kattan on May 16, 2002
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
	' Send security check if TE is sending the request or a user.
	'==============================================================

	Action = Request.form("action")
	Table  = Request.form("mainFrmExt")
	recs   = Request.form("chkDel")

	sFieldNames = request.form("txtFieldName")
	sFieldTypes = request.form("txtFieldType")

	
	recs   = replace(recs, " ", "")
'	response.write Action

	Select case Action
		case "Delete Selected"			TargetPage = "te_multidelete.asp"
		case "Export to XML"			TargetPage = "te_xml.asp"
		case "Export to Excel"			TargetPage = "te_excel.asp"
	end select

	SecID = GetSecurityID("Javaless_browser")

	
	NextPage = TargetPage & Table & "&nojs=1&txtFieldName=" & sFieldNames & "&txtFieldType=" & sFieldTypes & "&SecID=" & SecID

	if recs <> "" then NextPage = NextPage & "&chkDel=" & recs

	if Action <> "" then response.redirect(NextPage)
'	response.write NextPage

 %>