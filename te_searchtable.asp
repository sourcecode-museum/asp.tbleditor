<%
	'==============================================================
	' TableEditoR 0.81 Beta
	' http://www.2enetworx.com/dev/projects/tableeditor.asp
	'--------------------------------------------------------------
	' File: te_searchtable.asp
	' Description: Displays a search form for a given table
	' Initiated By Kevin Yochum on Nov 07, 2000
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
<!--#include file="te_header.asp"-->
<table border=0 cellspacing=1 cellpadding=2 bgcolor = "#ffe4b5" width=100%>
	<tr>
		<td class="smallertext">
			<a href="index.asp">Home</a> » <a href="te_admin.asp">Connections</a> » <a href="te_listtables.asp?cid=<%=request("cid")%>"><%=arrDesc(request("cid"))%></a> » <a href="<% =TableViewerCompat %>?cid=<%=request("cid")%>&tablename=<%=request("tablename")%>&ipage=<%=request("ipage")%>">Table [<%=request("tablename")%>]</a> » Search Table
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
	    
<form action="<% =TableViewerCompat %>?<%=request.querystring%>&q=1" method="post" name="kayitform">
	<table border=0 cellspacing=2 cellpadding=3 bgcolor = "#ffe4b5" width=500>

<%
	    lConnID = request("cid")
	    sTableName = request("tablename")
	    sFieldName = request("fld")
	    sFieldValue = request("val")
	    lFieldType = CLng(request("fldtype"))
	    iPage = request("ipage")
	    sQuery = request("q")
    	
	    OpenRS arrConn(lConnID)
    	
	    sSQL = "SELECT * FROM [" & sTableName & "]" 

	    rs.Open sSQL, , , adCmdTable
    		
	    on error resume next
	    for each fld in rs.fields
		    response.write "<tr>"
		    response.write "<td class=""smallerheader"">" & fld.name & "</td>"
		    select case fld.type
			    case adLongVarWChar	'memo
				    response.write "<td><textarea class=""tbflat"" cols=60 rows=6 name=""" & fld.name & """></textarea></td>"
			    case adLongVarBinary	'ole
				    'response.write "<td><img src=""te_getole.asp?cid=" & lConnID & "&tablename=" & sTableName & "&fld=" & fld.name & """></td>"
				    response.write "<td></td>"
			    case adBoolean
				    response.write "<td><input type=""checkbox"" name=""" & fld.name & """>"
			    case else
				    'response.write "<td><input class=""tbflat"" size=64 type=""text"" name=""" & fld.name &  """></td>"
				    
				'-[3]- Added by Daniele
				'copy&Paste from showrecord.asp
					select case fld.type
							case adSmallInt, adInteger, adCurrency, adUnsignedTinyInt, adDate, adDBDate, adDBTime, adDBTimeStamp
							
							if arrType(lConnID) <> tedbDsn then

								'-[2]- Added by Danival
								flag = false
								set rs2 = conn.OpenSchema(adSchemaForeignKeys)
								do while not rs2.eof	
									if rs2("FK_TABLE_NAME") = sTableName and rs2("FK_COLUMN_NAME") = fld.name  and relation then
										Set Consulta = Server.CreateObject("ADODB.Command")
										Set Consulta.ActiveConnection = conn
										consulta.CommandText = "SELECT * FROM " & rs2("PK_TABLE_NAME")
										set RSConsulta = consulta.execute
										response.write "<td><select class=""tbflat""  name=""" & fld.name &  """>"
										Response.Write "<option value="""">        </option>"
										do while not RSConsulta.eof
											if len(svalue) > 0 and len (RsConsulta(0)) > 0 then
												if cint(sValue) = cint(RSConsulta(0)) then
												   flagnbh = " selected"
												else
													flagnbh = ""
												end if
											end if
											Response.Write "<option value=""" & RSConsulta(trim(rs2("PK_COLUMN_NAME"))) & """" & flagnbh & ">" & RSConsulta(trim(rs2("PK_COLUMN_NAME"))) & " - " & RSConsulta(1) & "</option>"
											RSConsulta.movenext
										loop
										set rsconsulta = nothing
										set consulta = nothing
										Response.Write "</select></td>"									
										flag = true
									end if
									rs2.movenext
								loop
							end if
								
							if flag = false then
								response.write "<td><input class=""tbflat"" size=64 type=""text"" name=""" & fld.name &  """ value=""" & sValue & """></td>"
							end if	
							'-[2]- End Addition

							case else
								response.write "<td><input class=""tbflat"" size=64 type=""text"" name=""" & fld.name &  """ value=""" & sValue & """ maxlength=""" & fld.definedsize & """></td>"
						end select
				'-[3]- End Addition
		    end select
		    response.write "</tr>"
	    next
    			
%>
		<tr>
			<td></td>
			<td class="smallerheader"><input type="checkbox" name="chkSubstring"> Substring Search (applies to non-numeric fields only)</td>
		</tr>
	    
	    <tr>
		    <td></td>
		    <td>
			    <input type="submit" name="cmdSearch" value="  Search  " class="cmdflat">
		    </td>
	    </tr>
    	</table>
</form>
<!--#include file="te_footer.asp"-->