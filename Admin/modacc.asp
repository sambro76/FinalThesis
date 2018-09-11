<%@ Language=VBScript%>
<%Application.Lock()
Response.Buffer=true
tmpid=Request.QueryString("tmpid")
if Session("admin")<>"sbc" then
	Session("exp")="1"
	Response.Redirect("adlogin.asp")
else
	if tmpid="" then
		Response.Redirect("errpage.htm")
	else
		set conn=Server.CreateObject("ADODB.connection")
		DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("BankDB.MDB")
		conn.open DSN
		set rs=Server.CreateObject("ADODB.Recordset")
		rs.Open "SELECT AID FROM AdmInfos WHERE tmpid='" & tmpid & "'", conn
		if rs.eof or rs.bof then
			rs.Close()
			conn.Close()
			Response.Redirect("errpage.htm")
		else
			rs.Close()
			rs.open "SELECT * FROM AppInfos WHERE AccID='" & Request.QueryString(".accid") & "'" , conn
			if rs.BOF or rs.EOF then
				rs.Close()
				if Session("Search")="" then
					Response.Redirect("addmodacc.asp?tmpid=" & tmpid)
				elseif Session("Search")=1 then
					frow=1
					rs.Open "SELECT * FROM AppInfos WHERE AccID='" & Session("accid") & "'" , conn
				elseif cint(Session("Search"))>1 then
					frow=cint(Session("Search"))
					rs.Open "SELECT * FROM AppInfos WHERE " & Session("accid"), conn
				end if
			else
				frow=1
			end if
			Randomize()
%>
<HTML>
<title>Modify Remittance Account... </title>
<script language="javascript" src="../admin/images/fieldval.js"></script>
<STYLE TYPE="text/css">
@import url("images/admin.css");
BODY {
	SCROLLBAR-BASE-COLOR: #E8FFFF
}
</style> 

<body bgcolor="#474545" background="images/cordurouy.gif" onload="window.name='modacc'">
<div align="center">
<FORM name=f1 method=get target=_self>
<table border=1 width="600" id="tb1" cellspacing="0" bgcolor="#e7fbfe" bordercolor="#666699">
	<tr>
		<td colspan="2">
		<p style="MARGIN-TOP: 0px; margin-bottom:0" align="center" class=headeng>&nbsp;</p>
		<p style="MARGIN-TOP: 0px" align="center" class="headeng">
		<span style="letter-spacing: 2px">MODIFY ACCOUNT INFORMATION</span></p>
		<p align="right" style="MARGIN-TOP: 0px" class="td10pt"><b>SBC Bank Co., LTD of Cambodia</b>
		<p align="right" style="MARGIN-TOP: 0px" class="td10pt">You are login in as : <b><%=Session("AdminName")%></b>
		<HR size=2 color=maroon style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px"></td>
	</tr>
	<tr>
		<td align="middle" bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-right-style:none; border-right-width:medium" nowrap width="408">
			<p style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 6px" class=thw10pt align="left"> &nbsp;
				<a onmouseover="status='Create new account... '; return true;" onmouseout="status=''" href="javascript:window.open('createacc.asp?tmpid=<%=tmpid%>&.<%=rnd()%>','modacc')">
					<font color="#FFFFFF">Create New</font></a>&nbsp;&nbsp;|&nbsp;&nbsp; 
				<a onmouseover="status='Modify other account... '; return true;" onmouseout="status=''" href="javascript:window.open('addmodacc.asp?tmpid=<%=tmpid%>&.<%=rnd()%>','modacc')">
					<font color="#FFFFFF">Modify other account</font></a>&nbsp;&nbsp;
			<%if frow=1 then%>
				|&nbsp;&nbsp;
				<a onmouseover="status='Delete current displayed account... '; return true;" onmouseout="status=''" href="javascript:window.open('delAcc.asp?tmpid=<%=tmpid%>&AccID=<%=rs.Fields("AccID")%>','modacc')">
					<font color="#FFFFFF">Delete Account</font>
				</a>
			<%end if%>
				
			<td align="middle" bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-left-style:none; border-left-width:medium" nowrap width="184">
			<p class="thw10pt">
				<a onmouseover="status='Back to main menu... '; return true" onmouseout="status=''" href="javascript:window.open('manacc.asp?tmpid=<%=tmpid%>','modacc')">
				<font color="#FFFFFF">Main Menu</font>
				</a>&nbsp;&nbsp; |&nbsp;&nbsp;&nbsp;
				<a onmouseover="status='Log out... '; return true" onmouseout="status=''" href="javascript:window.open('adlogout.asp?tmpid=<%=tmpid%>','modacc')">
				<font color="#FFFFFF">Log out</font></a>
			</p>
		</td>
	<tr bgcolor="#bdc3ce" class="tdb10pt" nowrap>
		<td bordercolor="#000080" bgcolor="#FFFFF0" colspan="2">
	
	<table border="1" width="100%" id="tbl1" bgcolor="#ffffff" cellspacing="0" bordercolor="#bdc3ce" class="td10pt" cellpadding="6" style="letter-spacing: 1px">
		<tr>
			<td width="19%" bgColor=#FFEBD6 rowspan="2" nowrap onclick="window.open('createacc.asp?tmpid=<%=tmpid%>&.<%=rnd()%>','modacc')">
				<a onmouseover="status='Create new account... '; return true;" onmouseout="status=''" href="javascript:window.open('createacc.asp?tmpid=<%=tmpid%>&.<%=rnd()%>','modacc')">
					<p align="center" style="margin-bottom: 0"><b>Create New</b></p>
					<p align="center" style="margin-top: 3px"><b>Account</b></p>
				</a>
			</td>
			<td bgColor=#FFFFF7>Personal Information of APPLICANT...
			<%if Session("Modify")="1" then
				Response.Write("<b>(Modified successfully...)</b>")
			elseif Session("Modify")="ErrAccID" then%>
				<p style="margin-top: 3px"><%
				Response.Write("<font color='#FF0000'>(Account Identification has not been modified due to exiting one!)</font>")
			elseif Session("Modify")="ErrAccNo" then%>
				<p style="margin-top: 3px"><%
				Response.Write("<font color='#FF0000'>(Bank Account Number has not been modified due to exising one!)</font>")
			end if%>
			</td>
		</tr>

		<tr>
			<td bgColor=#FFFFD0 nowrap>
			<%if frow>1 then%>
				<p style="margin-top: 6px; margin-bottom: 0"><font color="#000080"><b><%=frow%> records have been found...</b></font></p>
				<p style="margin-top: 6px"><font color="#000080"><b>Please choose the nearest match on label: &quot;Record No&quot; or 
				&quot;SELECT&quot;</b></p>
			<%end if%>
			<%do until rs.EOF
			found=found+1
			%>
			<table border="1" id="table<%=found%>" bgcolor="#ffffff" cellspacing="0" bordercolor="#bdc3ce" class="td10pt" cellpadding="6" style="letter-spacing: 1px">
			<%if frow>1 then%>
				<tr>
					<td bgColor=#FFFFF7 width="42%" style="border-right-style: none; border-right-width: medium; border-bottom-style: solid; border-bottom-width: 1px" colspan="4">
						<p style="margin-top: 0; margin-bottom:0"><a href="javascript:window.open('modacc.asp?tmpid=<%=tmpid%>&.accid=<%=rs.Fields("AccID")%>&'+math.random(),'modacc')"><b>Record No: <%=found%></b></a></p></td>
					<td bgColor= #037EAD width="8%" class="thw10pt" onclick="window.open('modacc.asp?tmpid=<%=tmpid%>&.accid=<%=rs.Fields("AccID")%>&'+ Math.random(),'modacc')" style="border-bottom-style: solid; border-bottom-width: 1px; border-left-style:none; border-left-width:medium">
					<p align="center"><a href="javascript:window.open('modacc.asp?tmpid=<%=tmpid%>&.accid=<%=rs.Fields("AccID")%>&'+ Math.random(),'modacc')">
					<font color="#FFFFFF">SELECT</font></a></td>
				</tr>
			<%end if%>
			<tr>
				<td bgColor=#FFFFF7 width="25%" style="border-right-style: none; border-right-width: medium; border-bottom-style: none; border-bottom-width: medium; border-top-style:solid; border-top-width:1px" colspan="2">
					<p style="margin-top: 0; margin-bottom:0">&nbsp;
					<%if Session("Modify")="ErrAccID" then
						Response.Write("<font color='#FF0000'>Account Identification:</font>")
					else%>Account Identification:
					<%end if%>
					</p></td>
				<td width="25%" bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="3">
					<p style="margin-top: 0; margin-bottom: 0"><input type="text" name="AccID" size="20" maxlength="18" value="<%=rs.Fields("AccID")%>"></p></td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="2">
					<p style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;Password:</p></td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="3">
					<p style="margin-top: 0; margin-bottom: 0">
						<%set conn1=Server.CreateObject("ADODB.connection")
						DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & left(Server.MapPath("BankDB.MDB"),len(Server.MapPath("BankDB.MDB"))-17) & "\Rem\RemDB.MDB"
						conn1.Open DSN
						set rs1=Server.CreateObject("ADODB.Recordset")
						rs1.Open "SELECT pwd FROM Keyin WHERE AccID='" & rs.Fields("AccID") & "'", conn1
						pwd=rs1.Fields("pwd")
						rs1.Close()
						conn1.Close()
						set rs1=nothing
						set conn1=nothing
						%><input type="text" name="pwd" size="25" maxlength="24" value="<%=pwd%>">
					</p></td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="2">
					<p style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;Confirm Password:</p></td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="3">
					<p style="margin-top: 0; margin-bottom: 0">
					<input type="text" name="conpwd" size="25" maxlength="24" value="<%=pwd%>">
					</p></td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="2">
					<p style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;Initial 
					Credit Amount (In USD):</p></td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="3">
					<p style="margin-top: 0; margin-bottom: 0">
					<input type="text" name="InitCredit" size="20" maxlength="18" value="<%=rs.Fields("InitCredit")%>" disabled></p></td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="2">
					<p style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;First Name:</p></td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="3">
					<p style="margin-top: 0; margin-bottom: 0"><input type="text" name="FName" size="20" maxlength="18" value=<%=rs.Fields("FName")%>></p></td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="2">
				<p style="margin-top: 0; margin-bottom:0">&nbsp;&nbsp;Last Name (Family):</p></td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="3">
				<p style="margin-top: 0; margin-bottom: 0"><input type="text" name="LName" size="20" maxlength="18" value=<%=rs.Fields("LName")%>></td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style:none; border-bottom-width:medium" colspan="2">
					<p style="margin-top: 0; margin-bottom: 0">&nbsp;
					<%if Session("Modify")="ErrAccNo" then
						Response.Write("<font color='#FF0000'>Bank Account Number:</font>")
					else%>Bank Account Number:
					<%end if%>
					</td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style:none; border-bottom-width:medium" colspan="3">
					<p style="margin-top: 0; margin-bottom: 0"><input type="text" name="AccNo" size="20" maxlength="18" value=<%=rs.Fields("AccNo")%>></td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium" colspan="2">
					<p style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;Contact Phone No.:</td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium" colspan="3">
					<p style="margin-top: 0; margin-bottom: 0"><input type="text" name="Telephone" size="20" maxlength="18" value=<%=rs.Fields("Telephone")%>></td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 width="13%" nowrap style="border-right-style: none; border-right-width: medium; border-bottom-style: none; border-bottom-width: medium">
					<p style="margin-top: 0; margin-bottom:0">&nbsp;&nbsp;House/Building 
					No.:</p></td>
				<td bgColor=#FFFFF7 width="12%" style="border-right-style: none; border-right-width: medium; border-bottom-style: none; border-bottom-width: medium; border-left-style:none; border-left-width:medium">
					<input type="text" name="HB" size="7" maxlength="5" value=<%=rs.Fields("HB")%>></td>
				<td bgColor=#FFFFF7 width="14%" style="border-right-style: none; border-right-width: medium; border-bottom-style: none; border-bottom-width: medium; border-left-style:none; border-left-width:medium">
					<p style="margin-top: 0; margin-bottom:0">Street:</p></td>
				<td bgColor=#FFFFF7 width="16%" style="border-right-style: none; border-right-width: medium; border-bottom-style: none; border-bottom-width: medium; border-left-style:none; border-left-width:medium">
					<input type="text" name="Street" size="7" maxlength="5" value=<%=rs.Fields("Street")%>></td>
				<td bgColor=#FFFFF7 width="15%" style="border-right-style: solid; border-right-width: 1px; border-bottom-style: none; border-bottom-width: medium; border-left-style:none; border-left-width:medium">
					&nbsp;</td>
			</tr>
			<tr>
				<%city=""
				cty=rs.Fields("City")
				m=InStr(1,cty," ",vbTextCompare)
				if m>0 then
					do while(m>0)
						m=InStr(1,cty," ",vbTextCompare)
						if m>0 then
							city=city & left(cty,m-1) & "&nbsp;"
							cty=right(cty,len(cty)-m)
						else
							city=city & cty
						end if
					loop
				else
					city=cty
				end if
						
				country=""
				coun=rs.Fields("Country")
				m=InStr(1,coun," ",vbTextCompare)
				if m>0 then
					do while(m>0)
						m=InStr(1,coun," ",vbTextCompare)
						if m>0 then
							country=country & left(coun,m-1) & "&nbsp;"
							coun=right(coun,len(coun)-m)
						else
							country=country & coun
						end if
					loop
				else
					country=coun
				end if
				PayDetail=""
				PD=rs.Fields("Payment_Details")
				m=InStr(1,PD," ",vbTextCompare)
				if m>0 then
					do while(m>0)
						m=InStr(1,PD," ",vbTextCompare)
						if m>0 then
							PayDetail=PayDetail & left(PD,m-1) & "&nbsp;"
							PD=right(PD,len(PD)-m)
						else
							PayDetail=PayDetail & PD
						end if
					loop
				else
					PayDetail=PD
				end if
				spInst=""
				sp=rs.Fields("SpIns_MPay")
				m=InStr(1,sp," ",vbTextCompare)
				if m>0 then
					do while(m>0)
						m=InStr(1,sp," ",vbTextCompare)
						if m>0 then
							spInst=spInst & left(sp,m-1) & "&nbsp;"
							sp=right(sp,len(sp)-m)
						else
							spInst=spInst & sp
						end if
					loop
				else
					spInst=sp
				end if
				%>
				<td bgColor=#FFFFF7 style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="2">
					<p style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;City/State:</p></td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="3">
					<p style="margin-top: 0; margin-bottom: 0"><input type="text" name="City" size="20" maxlength="18" value=<%=city%>></p></td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="2">
				<p style="margin-top: 0; margin-bottom:0">&nbsp;&nbsp;Country:</p></td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="3">
				<p style="margin-top: 0; margin-bottom: 0"><input type="text" name="Country" size="20" maxlength="18" value=<%=country%>></td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style:none; border-bottom-width:medium" colspan="2">
					<p style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;Payment Details:</td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style:none; border-bottom-width:medium" colspan="3">
					<p style="margin-top: 0; margin-bottom: 0"><input type="text" name="PayDetail" size="30" value=<%=PayDetail%>></td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 nowrap style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium" colspan="2">
					<p style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;Special Instruction/Multiple Payments:</td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium" colspan="3">
					<p style="margin-top: 0; margin-bottom: 0"><input type="text" name="SpIns" size="30" value=<%=spInst%>></td>
			</tr>
			</table>
			<input type=hidden name="AccID1" value="<%=rs.Fields("AccID")%>">
			<%if frow>1 then%>
				<p style="margin-top: 0; margin-bottom: 0">
			<%end if
			rs.MoveNext()
			loop%>
			</td>
		</tr>
	</table>
	<%if frow=1 then%>
		<input type=hidden name="tmpid" value=<%=tmpid%>>
		<p style="margin-top: 12px; margin-bottom: 12px" align="center">
		<input type=button name="bSearch" value="Commit Changes" onclick="if(validate()==1){document.f1.action='commitMod.asp'; document.f1.submit();}" style="font-family: Verdana; font-size: 12px; font-weight: bold"></p>
	<%end if%>
	<HR size=2 color=maroon style="MARGIN-TOP: 12px; MARGIN-BOTTOM: 0px"></td>
</tr>
</table>
</FORM>
</div>
</body>
</HTML>
<%end if
end if
end if
Session("Modify")=""
'Session("Search")=""
set rs=nothing
set conn=nothing%>