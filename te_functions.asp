<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_functions.asp
	' Description: function for TableEditor
	' Initiated By Rami Kattan on May 31, 2002
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

	sub allTablesCombo()	'Database Selector
		if bComboTables and bJSEnable then
			iCID = cint(request("cid"))
			response.write "<select id=""allDBs"" name=""allDBs"" class=""smallertext"" onchange=""cboChangeDB()"">"
			if bAdmin and iCID = 0 then response.write "<option value=""0"">Administrator Database"
			for lv = 1 to iTotalConnections
				response.write "<option value=""" & lv & """"
				if iCID = lv then response.write " selected"
				response.write ">" & arrDesc(lv) & vbCrLf
			next
			response.write "</select>"
		else
			response.write "<a href=""te_listtables.asp?cid=" & request("cid") &""">" & arrDesc(request("cid")) & "</a>"
		end if
	end sub

	sub allTablesCombo2(lConnID)	'Table Selector
		if bComboTables and bJSEnable then
			conn.Open arrConn(lConnID)
			set rs = conn.OpenSchema(adSchemaTables)
			response.write "<select name=""allTables"" id=""allTables"" class=""smallertext"" onchange=""ChangeTable()"">"
			response.write "<option value='<<'><- Back"

			while not rs.eof
				if rs("table_type") = "TABLE" then
					response.write "<option value=""" & rs("table_name") & """"
					if request("tablename") = rs("table_name") then response.write " selected"
					response.write ">" & rs("table_name") & vbCrLf
				end if
				rs.movenext
			wend
			response.write "</select>"
			rs.close
			conn.close
		else
			response.write sTableName
		end if
	end sub

	Sub isPerPage(InValue, ThisValue)
		if cint(InValue) = cint(ThisValue) then Response.Write " Selected"
	End sub

	function isSelected(InValue, ThisValue)
		if cint(InValue) = cint(ThisValue) then isSelected = " Selected"
	end function

	function MakeURL(Data)
		UrlData = Data
		if ConvertURL then
			UrlData = edit_hrefs(UrlData, 1)
			UrlData = edit_hrefs(UrlData, 2)
			'UrlData = edit_hrefs(UrlData, 3)
			UrlData = edit_hrefs(UrlData, 4)
			UrlData = edit_hrefs(UrlData, 5)
			UrlData = edit_hrefs(UrlData, 6)
		end if
		MakeURL = UrlData
	end function

	function SQLEncode(strTheText)
		SQLEncode = "'" & Replace (strTheText, "'", "''" ) & "'"
	end function

	function GetRandomChars(width)
		Randomize
		data = ""
		while len(data) < width
			data = data & chr(Int((80 * Rnd) + 47))
		wend
		GetRandomChars = data
	end function

	function GetSecurityID(ForAction)
		session("TableEditor_" & ForAction) = GetRandomChars(6) & right(Session.SessionID,4) & GetRandomChars(6)
		GetSecurityID = session("TableEditor_" & ForAction)
	end function

	function ValidSecurityID(ForAction, SecID)
		if session("TableEditor_" & ForAction) = SecID then
			ValidSecurityID = true
		else
			ValidSecurityID = false
		end if
	end function

	function FormatXML(data)
		if isNumeric(left(data,1)) then
			data = FormatNumericXML(data)
		end if
		data = replace(data, "?", "_x003F_")
		data = replace(data, " ", "_x0020_")
		data = replace(data, "/", "_x002F_")
		data = replace(data, "=", "_x003D_")
		data = replace(data, "%", "_x0025_")
		slash = "\"
		data = replace(data, slash, "_x005C_")
		data = replace(data, "~", "_x007E_")
		data = replace(data, "@", "_x0040_")
		data = replace(data, "#", "_x0023_")
		data = replace(data, "$", "_x0024_")
		data = replace(data, "%", "_x0025_")
		data = replace(data, "^", "_x005E_")
		data = replace(data, "&", "_x0026_")
		data = replace(data, "*", "_x002A_")
		data = replace(data, "(", "_x0028_")
		data = replace(data, ")", "_x0029_")
		data = replace(data, "+", "_x002B_")
		data = replace(data, "{", "_x007B_")
		data = replace(data, "}", "_x007D_")
		data = replace(data, "|", "_x007C_")
		data = replace(data, "'", "_x0027_")
		data = replace(data, "<", "_x003C_")
		data = replace(data, ">", "_x003E_")
		data = replace(data, ",", "_x002C_")
		data = replace(data, ";", "_x003B_")
   		FormatXML = data
	end function

	function FormatNumericXML(data)
		StrLeft = Left(data, 1)
		StrRight = Right(data, (len(data) - 1))
		ReturnValue = "_x003" & StrLeft & "_" & StrRight
		FormatNumericXML = ReturnValue
	end function

	function FormatXMLRev(data)
		if left(data,5) = "_x003" then
			if isNumeric(mid(data,6,1)) then
				data = mid(data,6,1) & right(data, len(data)-7)
			end if
		end if

		data = replace(data, "_x003F_", "?")
		data = replace(data, "_x0020_", " ")
		data = replace(data, "_x002F_", "/")
		data = replace(data, "_x003D_", "=")
		data = replace(data, "_x0025_", "%")
		slash = "\"
		data = replace(data, "_x005C_", slash)
		data = replace(data, "_x007E_", "~")
		data = replace(data, "_x0040_", "@")
		data = replace(data, "_x0023_", "#")
		data = replace(data, "_x0024_", "$")
		data = replace(data, "_x0025_", "%")
		data = replace(data, "_x005E_", "^")
		data = replace(data, "_x0026_", "&")
		data = replace(data, "_x002A_", "*")
		data = replace(data, "_x0028_", "(")
		data = replace(data, "_x0029_", ")")
		data = replace(data, "_x002B_", "+")
		data = replace(data, "_x007B_", "{")
		data = replace(data, "_x007D_", "}")
		data = replace(data, "_x007C_", "|")
		data = replace(data, "_x0027_", "'")
		data = replace(data, "_x003C_", "<")
		data = replace(data, "_x003E_", ">")
		data = replace(data, "_x002C_", ",")
		data = replace(data, "_x003B_", ";")
   		FormatXMLRev = data
	end function

	function DataNeedCDATA(data)
		need = false
		if instr(data, "<") then need = true
		if instr(data, "&") then need = true
		DataNeedCDATA = need
	end function

	function LeadingZero(data, numdigits)
		while len(data) < numdigits
			data = "0" & data
		wend
		LeadingZero = data
	end function

	function IsObjInstalled(strClassString)
		Err.clear
		set oTest2 = Server.CreateObject(strClassString)
		if err = 0 then
			IsObjectInstalled = true
		else 
			err.Clear
			IsObjectInstalled = false
		end if
		set oTest2 = nothing
		IsObjInstalled = IsObjectInstalled
	end function

%>
<script language="javascript1.2" runat=server>
function edit_hrefs(s_html, type){
    s_str = new String(s_html);
	if (type == 1) { // Start with http://
     	s_str = s_str.replace(/\b(http\:\/\/[\w+\.]+[\w+\.\:\/\_\?\=\&\-\'\#\%\~\;\,\$\!\+\*]+)/gi,
		  "<a href=\"$1\" target=\"_blank\">$1<\/a>");
	} 
	if (type == 2) { // Start with https://

		s_str = s_str.replace(/\b(https\:\/\/[\w+\.]+[\w+\.\:\/\_\?\=\&\-\'\#\%\~\;\,\$\!\+\*]+)/gi,
		  "<a href=\"$1\" target=\"_blank\">$1<\/a>");
	}
	if (type == 3) { // Start with file://
		s_str = s_str.replace(/\b(file\:\/\/\/\w\:\\[\w+\/\w+\.\:\/\_\?\=\&\-\'\#\%\~\;\,\$\!\+\*]+)/gi,
		  "<a href=\"$1\" target=\"_blank\">$1<\/a>");
	}
	if (type == 4) { // Start with www.

		s_str = s_str.replace(/\b(www\.[\w+\.\:\/\_\?\=\&\-\'\#\%\~\;\,\$\!\+\*]+)/gi,
 		  "<a href=\"http://$1\" target=\"_blank\">$1</a>");
	}
	if (type == 5) { // email
		s_str = s_str.replace(/\b([\w+\-\'\#\%\.\_\,\$\!\+\*]+@[\w+\.?\-\'\#\%\~\_\.\;\,\$\!\+\*]*)/gi,
 		  "<a href=\"mailto\:$1\">$1</a>");
	}
	if (type == 6) { // Start with ftp://
     	s_str = s_str.replace(/\b(ftp\:\/\/[\w+\.]+[\w+\.\:\/\_\?\=\&\-\'\#\%\~\;\,\$\!\+\*]+)/gi,
		  "<a href=\"$1\" target=\"_blank\">$1<\/a>");
	} 
		  	  
    return s_str;
}
</script>