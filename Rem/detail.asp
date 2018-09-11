<%@ Language=VBScript%>
<%if Session("user")<>"sbc" then
	Session("exp")="1"
	Response.Redirect("sbc_pwd.asp")
else
Application.Lock()
lname=LCase(Request.QueryString("lname"))
id=Request.QueryString("id")
if lname="" then
	Response.Redirect("errpage.htm")
else
	set conn=Server.CreateObject("ADODB.connection")
	DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("RemDB.MDB")
	conn.open DSN
	set rs=Server.CreateObject("ADODB.Recordset")
	''''''
	'Get login name 'lname' & temporary id 'id' to admit the page requested
	''''''
	rs.Open "SELECT AccID FROM Keyin WHERE AccID='" & lname & "' AND id='" & id & "'", conn
	if rs.eof or rs.bof then 
		rs.Close()
		conn.Close()
		Response.Redirect("errpage.htm")
	else
		rs.Close()
		''''''
		'Get recordset from Querying the Table RecData with REF_No chosen
		''''''
		rs.Open "SELECT * FROM RecData WHERE AccID='" & lname & "' AND REF_No='" & Request.QueryString("ref") & "'", conn
		status=rs.Fields("Status")
%>
<HTML>
<title>APPLICATION FORM DETAILS</title>
<script language="JavaScript" src="images/button.js"></script>

<STYLE TYPE="text/css">
@import url("images/sbc.css");
BODY {
	SCROLLBAR-BASE-COLOR: #E8FFFF
}
</style> 

<body bgcolor="#666699" onload="window.name='detail'">
<div align="center">
<FORM name=f1 method=post action="reorder.asp" target=_self>
<table border=2 width="613" id="tb1" cellspacing="0" bgcolor="#deebff" bordercolor="#4D4E79">
	<tr>
		<td colspan="2">
		<p style="MARGIN-TOP: 0px; margin-bottom:0" align="center" class=headeng>&nbsp;</p>
		<p style="MARGIN-TOP: 0px" align="center" class=headeng><b>APPLICATION FORM 
		DETAILS</b></p>
		<p align="right" style="MARGIN-TOP: 0px" class="td10pt"><b>SBC Bank Co., LTD of Cambodia</b>
		<HR size=2 color=maroon style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px"></td>
	</tr>
	<tr>
		<td align="middle" bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-right-style:none; border-right-width:medium" nowrap width="326">
		<p style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 6px" class=thw10pt align="left"> 
		<span style="letter-spacing: 1px">&nbsp;</span> <a onmouseover="status='Close this form... '; return true" onmouseout="status=''" href="javascript:window.close()"><font color="#FFFFFF">
		CLOSE FORM</font></a>
		<%if status=4 then
			Session("REFNo")=Request.QueryString("ref")%>
			<input type=hidden name="lname" value="<%=lname%>">
			<input type=hidden name="id" value="<%=id%>">
			|  <a onmouseover="status='Order the form again... '; return true" onmouseout="status=''" href="javascript:document.f1.submit()"><font color="#FFFFFF">Reorder The Form</font></a>
		<%end if%>
		<td align="middle" bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-left-style:none; border-left-width:medium" nowrap width="277">
		<IMG onmouseup="FP_swapImg(0,0,/*id*/'img16',/*url*/'images/button4A.gif')" onmousedown="FP_swapImg(1,0,/*id*/'img16',/*url*/'images/button4B.gif')" id=img16 onmouseover="FP_swapImg(1,0,/*id*/'img16',/*url*/'images/button4A.gif')" onmouseout="FP_swapImg(0,0,/*id*/'img16',/*url*/'images/button49.gif')" height=25 alt =" HELP!" src="images/button49.gif" width =125 border=0  fp-title=" HELP!" fp-style="fp-btn: Jewel 1; fp-font-style: Bold; fp-font-size: 16; fp-font-color-normal: #800000; fp-font-color-hover: #FFFF00; fp-font-color-press: #FF0000; fp-justify-horiz: 0; fp-transparent: 1; fp-orig: 0"><tr bgcolor="#bdc3ce" class="tdb10pt" nowrap>
	<td bordercolor="#000080" bgcolor="#FFFFF0" colspan="2">
	
	<table border="1" width="100%" id="tbl1" bgcolor="#FFEBD6"  cellspacing="0" bordercolor="#bdc3ce" class="td10pt" cellpadding="6" style="letter-spacing: 1px">
            <tr>
              <%if status<>4 then
              	Response.Write("<td width='10%'></td>")
              end if%>
              <td bgcolor="#DEEBFF" style="font-family: Verdana" width="80%">
				<table border="1" cellpadding="6" cellspacing="0" width="100%" bordercolor="#BDC3CE" bgcolor="#FFFFF7" id="table1" class="tdv10pt">
					<tr>
						<td width="80%" colspan="2">This form was dated on <%=rs.Fields("Dated")%></td>
						<%if status=4 then%>
							<td width="14%"><p align="center"><b>Status</b></td>
						<%end if%>
					</tr>
					<tr>
	              		<td colspan="2">Amount Applied : <b> <%=rs.Fields("Currency") & " " & FormatNumber(rs.Fields("Amount"),2)%></b></td>
	              		<%if status=4 then
		              		if rs.Fields("Currency")="USD" then%><td align="center">
		              		<%else
			              		Response.Write("<td width='20%' rowspan='2' align='center'>")
		              		end if
		              		set rs1=Server.CreateObject("ADODB.Recordset")
							rs1.Open "SELECT * FROM ErrLog WHERE EREFNo='" & Request.QueryString("ref") & "'", conn
							%>
		              		<b><%if rs1.Fields("EAmount")="" then
		              			Response.Write("OK")
		              		else
		              			Response.Write("<font color=#FF0000>" & rs1.Fields("EAmount") & "</font>")
		              		end if%>
		              		</b></td>
		              	<%end if%>
            		</tr>
					<%if rs.Fields("Currency")<>"USD" then
		                Response.Write("<tr><td>&nbsp;</td><td width='60%'>")
		                Response.Write("In US$ : <b>" & FormatNumber(rs.Fields("Amount")/rs.Fields("Rate"),2) & "</b>, Rate: <b>" & FormatNumber(rs.Fields("Rate"),2) & "*</b>")
		                Response.Write("</td></tr>")
				        end if%>
				<tr>
	              <td colspan="2">LOCAL BANK CHARGES FOR :</td>
	              <%if status="4" then%>
	              <td width="14%" rowspan="2" align="center"><b>OK</b></td>
	              <%end if%>
           		</tr>
				<tr>
				  <td>&nbsp;</td>
	              <td nowrap><b>
                  <%if rs.Fields("Charge_for")="0" then
					Response.Write("<input type='radio' name='R3' disabled value=0 checked> My/Our Account")
					else%>
	             	   <input type="radio" name="R3" disabled checked>Beneficiary's Account (RECIPIENT) 
                 	<%end if%>
                 	</b>
					</td>
           		</tr>
				<tr>
	              <td colspan="2"> Beneficiary's Name (RECIPIENT) :</td>
	              <%if status="4" then%>
		              <td width="14%" rowspan="2" align="center" nowrap><b>
		              	<%if rs1.Fields("EBName")="" then
		              		Response.Write("OK")
		              	else
		              		Response.Write("<font color=#FF0000>" & rs1.Fields("EBName") & "</font>")
		              	end if%>
		              	</b>
		              </td>
	              <%end if%>
           		</tr>
				<tr>
	              <td> &nbsp;</td>
	              <td width="69%"><b><%=rs.Fields("BFName") & " " & rs.Fields("BLName")%></b></td>
           		</tr>
				<tr>
					<%if Status="4" then%>
		            	<td colspan="3">
		            <%else 
		            	Response.Write("<td colspan='2'>")	
		            end if%>
	              Beneficiary's Address : </td>
           		</tr>
				<tr>
	              	<td rowspan="3">&nbsp;</td> 
					<td width="69%">Street : <b> <%=rs.Fields("BStreet")%></b></td>
		            <%if status=4 then%>
		            <td width="14%" rowspan="3" align="center"><b>
						<%if rs1.Fields("EBAddress")="" then
		              		Response.Write("OK")
		              	else
		              		Response.Write("<font color='#FF0000'>" & rs1.Fields("EBAddress") & "</font>")
		              	end if%>
		             </b>
		            </td>
		            <%end if%>
				</tr>
				<tr>
					<td width="69%">State/City : <b> <%=rs.Fields("BCity")%></b></td>
				</tr>
				<tr>
					<td width="69%">Country : <b> <%=rs.Fields("BCountry")%></b></td>
				</tr>
				<tr>
		            <%if Status="4" then%>	              	
	              	<td colspan="3">
	              	<%else
	              		Response.Write("<td colspan='2'>")
	              	end if%>	
	              	Beneficiary's BANK/City/Country :
	              	</td>
					<tr>
	              	<td rowspan="4">&nbsp;</td> 
					<td width="69%">Beneficiary's BANK : <b> <%=rs.Fields("BBank")%></b></td>
					<%if status=4 then%>
		              	<td width="14%" align="center" nowrap><b>
			              	<%if rs1.Fields("EBBank")="" then
			              		Response.Write("OK")
			              	else
			              		Response.Write("<font color=#FF0000>" & rs1.Fields("EBBank") & "</font>")
			              	end if%>
		              	</b>
		              	</td>
	              	<%end if%>
				</tr>
				<tr>
					<td width="69%">State/City : <b> <%=rs.Fields("bbCity")%></b></td>
		              	<%if status=4 then%>
		              	<td width="14%" rowspan="2" align="center" nowrap><b>
			              	<%if rs1.Fields("EBLocation")="" then
			              		Response.Write("OK")
			              	else
			              		Response.Write("<font color=#FF0000>" & rs1.Fields("EBLocation") & "</font>")
			              	end if%>
			            </b>
		              	</td>
		              	<%end if%>
               	</tr>
				<tr>
					<td width="69%">Country : <b> <%=rs.Fields("bbCountry")%></b>
					</tr>
				<tr>
					<td width="69%">Account No. : <b> <%=rs.Fields("bAccNo")%></b></tr>
		            <%if Status="4" then%>
	              	<td width="14%" align="center" nowrap><b>
						<%if rs1.Fields("EBAccNo")="" then
		              		Response.Write("OK")
		              	else
		              		Response.Write("<font color=#FF0000>" & rs1.Fields("EBAccNo") & "</font>")
		              	end if%>
		            </b>
	              	</td>
					<%end if%>
					</table>
			<%if status=4 then%>
              <td width="15%" onclick="document.f1.submit();"><a onmouseover="status='Order the form again... '; return true" onmouseout="status=''" href="javascript:document.f1.submit();">
				<p align="center"><b>ORDER</b></p>
				<p align="center"><b>THIS FORM</b></p>
				<p align="center"><b>&nbsp;AGAIN</b></p></a></td>
			<%else
				Response.Write("<td></td>")
			end if%>
			</td>
            </tr>
          </table>
          
	<%if rs.Fields("Currency")<>"USD" then%>
	<p style="margin-top: 6px; margin-bottom: 12px">
	<span style="letter-spacing: 1px" class="td10pt">*Please be informed that this exchange rate was valued due to the date you had applied! 
	</span> </p>
	<%end if
	if rs.Fields("Status")="0" then conn.Execute "UPDATE RecData SET Status='1' WHERE REF_No='" & Request.QueryString("ref") & "'", recaffected
	rs.Close()
	conn.Close()%>
	<p style="margin-top: 12px; margin-bottom: 12px" align="center">
	<font color="#0000CD"><b>
	<a onmouseover="status='Close this form... '; return true" onmouseout="status=''" href="javascript:window.close()">CLOSE THIS FORM</a></b></font></p>
	<HR size=2 color=maroon style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px"></td>
</tr>
</table>
</FORM>
</div>
</body>
</HTML>
<%end if
end if
end if
set rs=nothing
set conn=nothing%>