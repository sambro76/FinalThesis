<%@ Language=VBScript%>
<%Application.Lock()
if Session("admin")<>"sbc" then
	Session("exp")="1"
	Response.Redirect("adlogin.asp")
else

Response.Buffer=true
tmpid=Request.QueryString("tmpid")
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
		if Session("Create")="1" then
			created=1
			rs.Open "SELECT * FROM AppInfos WHERE AccID='" & Session("AccCreated") & "'", conn
			Session("Search")="1"
			Session("accid")=rs.Fields("AccID")
		else
			created=0
		end if
%>
<HTML>
<title>Create Remittance Account... </title>
<script language="javascript" src="../admin/images/fieldval.js"></script>

<STYLE TYPE="text/css">
@import url("images/admin.css");
BODY {
	SCROLLBAR-BASE-COLOR: #E8FFFF
}
</style>

<body bgcolor="#474545" background="images/cordurouy.gif" onload="window.scroll(0,20); window.name='createacc'">
<div align="center">
<FORM name=f1 method=get target=_self>
<table border=1 width="600" id="tb1" cellspacing="0" bgcolor="#e7fbfe" bordercolor="#666699">
	<tr>
		<td colspan="2">
		<p style="MARGIN-TOP: 0px; margin-bottom:0" align="center" class=headeng>
		&nbsp;</p>
		<p style="MARGIN-TOP: 0px; margin-bottom:3px" align="center" class="headeng">
		<span style="letter-spacing: 2px">CREATE ACCOUNT</span></p>
		<p align="right" style="MARGIN-TOP: 0px" class="td10pt"><b>SBC Bank Co.,
		LTD of Cambodia</b>
		<p align="right" style="MARGIN-TOP: 0px" class="td10pt">You are login as
		: <b><%=Session("AdminName")%></b>
		<HR size=2 color=maroon style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px"></td>
	</tr>
	<tr>
		<td align="middle" bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-right-style:none; border-right-width:medium" nowrap width="408">
			<p style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 6px" class=thw10pt align="left">&nbsp;
			<%if created=1 then%>
				<a onmouseover="status='Create another account... '; return true;" onmouseout="status=''" href="javascript:window.open('createacc.asp?tmpid=<%=tmpid%>','createacc')"><font color="#FFFFFF">
			&nbsp;Create New</font></a>
				&nbsp;&nbsp;|&nbsp;&nbsp;
				<a onmouseover="status='Modify this account... '; return true;" onmouseout="status=''" href="javascript:window.open('modacc.asp?tmpid=<%=tmpid%>','createacc')"><font color="#FFFFFF">
			Modify this account</font></a>
			<%else%>
				Please fill information in the following form.
			<%end if%>
			<td align="middle" bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-left-style:none; border-left-width:medium" nowrap width="184">
			<p class="thw10pt">
				<a onmouseover="status='Back to main menu... '; return true" onmouseout="status=''" href="javascript:window.open('manacc.asp?tmpid=<%=tmpid%>','createacc')">
				<font color="#FFFFFF">Main Menu</font>
				</a>&nbsp;&nbsp; |&nbsp;&nbsp;&nbsp;
				<a onmouseover="status='Log out... '; return true" onmouseout="status=''" href="javascript:window.open('adlogout.asp?tmpid=<%=tmpid%>','createacc')">
				<font color="#FFFFFF">Log out</font></a>
			</p>
		</td>
	<tr bgcolor="#bdc3ce" class="tdb10pt" nowrap>
		<td bordercolor="#000080" bgcolor="#FFFFF0" colspan="2">

	<table border="1" width="100%" id="tbl1" bgcolor="#ffffff" cellspacing="0" bordercolor="#bdc3ce" class="td10pt" cellpadding="6" style="letter-spacing: 1px">
		<tr>
			<td width="19%" bgColor=#FFEBD6 rowspan="2" nowrap>
				<p align="center">&nbsp;</td>
			<td bgColor=#FFFFF7>Personal Information of APPLICANT...
				<%if Session("AccCreated")="ErrAccID" then
					Response.Write("<font color='#FF0000'> (Account Identification is duplicated)</font>")
				elseif Session("AccCreated")="ErrAccNo" then
					Response.Write("<font color='#FF0000'> (Bank Account is duplicated)</font>")
				elseif Session("Create")="1" then
					Response.Write("<b> (Account created successfully...)</b>")
				end if%>
			</td>
		</tr>

		<tr>
			<td bgColor=#FFFFD0 nowrap>
			<font color="#000080">
			<table border="1" id="table1" bgcolor="#ffffff" cellspacing="0" bordercolor="#bdc3ce" class="td10pt" cellpadding="6" style="letter-spacing: 1px">
			<tr>
				<td bgColor=#FFFFF7 width="25%" style="border-right-style: none; border-right-width: medium; border-bottom-style: none; border-bottom-width: medium; border-top-style:solid; border-top-width:1px" colspan="2" nowrap>
					<p style="margin-top: 0; margin-bottom:0">&nbsp;
					<b>
					<%if Session("AccCreated")="ErrAccID" then
						Response.Write("<font color='#FF0000'>Account Identification (Login account):</font>")
					else%>Account Identification (Login account):
					<%end if%>
					</b>
					</p>
				</td>
				<td width="25%" bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="3" nowrap>
					<p style="margin-top: 0; margin-bottom: 0">
					<%if created=1 then%>
						<input type="text" name="AccID" size="20" maxlength="18" value="<%=rs.Fields("AccID")%>" disabled>
					<%else%>
						<input type="text" name="AccID" size="20" maxlength="18" value="<%=Request.QueryString("AccID")%>">
					<%end if%>
					</p></td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="2">
					<p style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;Password:</p></td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="3">
					<p style="margin-top: 0; margin-bottom: 0">
					<%if created=1 then
						Response.write("<input type=text name='pwd' size='25' maxlength='24' value='" & Session("pwd") & "' disabled>")
					else%>
						<input type="text" name="pwd" size="25" maxlength="24" value="<%=Request.QueryString("pwd")%>">
					<%end if%>
					</p></td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="2">
					<p style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;Confirm
					Password:</p></td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="3">
					<p style="margin-top: 0; margin-bottom: 0">
					<%if created=1 then
						Response.write("<input type=text name='conpwd' size=25 maxlength=24 value='" & Session("pwd") & "' disabled>")
					else%>
						<input type="text" name="conpwd" size="25" maxlength="24" value="<%=Request.QueryString("pwd")%>">
					<%end if%>
					</p></td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="2">
					<p style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;Initial credit
					amount (In USD):</p></td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="3">
					<p style="margin-top: 0; margin-bottom: 0">
					<%if created=1 then%>
						<input type="text" name="InitCredit" size="20" maxlength="18" value="<%=rs.Fields("InitCredit")%>" disabled>
					<%else%>
						<input type="text" name="InitCredit" size="20" maxlength="18" value="<%=Request.QueryString("InitCredit")%>">
					<%end if%>
					</p></td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="2">
					<p style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;First Name:</p></td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="3">
					<p style="margin-top: 0; margin-bottom: 0">
					<%if created=1 then%>
						<input type="text" name="FName" size="20" maxlength="18" value="<%=rs.Fields("FName")%>" disabled>
					<%else%>
						<input type="text" name="FName" size="20" maxlength="18" value="<%=Request.QueryString("FName")%>">
					<%end if%>
					</p></td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="2">
				<p style="margin-top: 0; margin-bottom:0">&nbsp;&nbsp;Last Name (Family):</p></td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="3">
				<p style="margin-top: 0; margin-bottom: 0">
					<%if created=1 then%>
						<input type="text" name="LName" size="20" maxlength="18" value="<%=rs.Fields("LName")%>" disabled>
					<%else%>
						<input type="text" name="LName" size="20" maxlength="18" value="<%=Request.QueryString("LName")%>">
					<%end if%>
				</td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style:none; border-bottom-width:medium" colspan="2">
					<p style="margin-top: 0; margin-bottom: 0">&nbsp;
					<b>
					<%if Session("AccCreated")="ErrAccNo" then
						Response.Write("<font color='#FF0000'>Bank Account Number:</font>")
					else%>Bank Account Number:
					<%end if%>
					</b>
					</p>
				</td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style:none; border-bottom-width:medium" colspan="3">
					<p style="margin-top: 0; margin-bottom: 0">
					<%if created=1 then%>
						<input type="text" name="AccNo" size="20" maxlength="18" value="<%=rs.Fields("AccNo")%>" disabled>
					<%else%>
						<input type="text" name="AccNo" size="20" maxlength="18" value="<%=Request.QueryString("AccNo")%>">
					<%end if%>
				</td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium" colspan="2">
					<p style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;Contact Phone
					No.:</td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium" colspan="3">
					<p style="margin-top: 0; margin-bottom: 0">
					<%if created=1 then%>
						<input type="text" name="Telephone" size="20" maxlength="18" value="<%=rs.Fields("Telephone")%>" disabled>
					<%else%>
						<input type="text" name="Telephone" size="20" maxlength="18" value="<%=Request.QueryString("Telephone")%>">
					<%end if%>
				</td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 width="13%" nowrap style="border-right-style: none; border-right-width: medium; border-bottom-style: none; border-bottom-width: medium">
					<p style="margin-top: 0; margin-bottom:0">&nbsp;&nbsp;House/Building
					No.:</p></td>
				<td bgColor=#FFFFF7 width="12%" style="border-right-style: none; border-right-width: medium; border-bottom-style: none; border-bottom-width: medium; border-left-style:none; border-left-width:medium">
				<%if created=1 then%>
					<input type="text" name="HB" size="7" maxlength="5" value="<%=rs.Fields("HB")%>" disabled>
				<%else%>
					<input type="text" name="HB" size="7" maxlength="5" value="<%=Request.QueryString("HB")%>">
				<%end if%>
				</td>
				<td bgColor=#FFFFF7 width="14%" style="border-right-style: none; border-right-width: medium; border-bottom-style: none; border-bottom-width: medium; border-left-style:none; border-left-width:medium">
					<p style="margin-top: 0; margin-bottom:0">Street:</p></td>
				<td bgColor=#FFFFF7 width="16%" style="border-right-style: none; border-right-width: medium; border-bottom-style: none; border-bottom-width: medium; border-left-style:none; border-left-width:medium">
				<%if created=1 then%>
					<input type="text" name="Street" size="7" maxlength="5" value="<%=rs.Fields("Street")%>" disabled>
				<%else%>
					<input type="text" name="Street" size="7" maxlength="5" value="<%=Request.QueryString("Street")%>">
				<%end if%>
				</td>
				<td bgColor=#FFFFF7 width="15%" style="border-right-style: solid; border-right-width: 1px; border-bottom-style: none; border-bottom-width: medium; border-left-style:none; border-left-width:medium">
					&nbsp;</td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="2">
					<p style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;City/State:</p></td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="3">
					<p style="margin-top: 0; margin-bottom: 0">
				<%if created=1 then
				city=""
				city=InStr(1,rs.Fields("City")," ",vbTextCompare)
					if city>0 then
						city=left(rs.Fields("City"),city-1) & "&nbsp;" & mid(rs.Fields("City"),city+1)
					else
						city=rs.Fields("City")
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
				end if%>

					<input type="text" name="City" size="20" maxlength="18" value="<%=city%>" disabled>
				<%else%>
					<input type="text" name="City" size="20" maxlength="18" value="<%=Request.QueryString("City")%>">
				<%end if%>
				</p></td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="2">
				<p style="margin-top: 0; margin-bottom:0">&nbsp;&nbsp;Country:</p></td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" colspan="3">
				<p style="margin-top: 0; margin-bottom: 0">
				<%if created=1 then%>
					<input type="text" name="Country" size="20" maxlength="25" value="<%=country%>" disabled>
				<%else%>
					<input type="text" name="Country" size="20" maxlength="25" value="<%=Request.QueryString("Country")%>">
				<%end if%>
				</td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style:none; border-bottom-width:medium" colspan="2">
					<p style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;Payment
					Details:</td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style:none; border-bottom-width:medium" colspan="3">
					<p style="margin-top: 0; margin-bottom: 0">
				<%if created=1 then%>
					<input type="text" name="PayDetail" size="30" value="<%=PayDetail%>" disabled>
				<%else%>
					<input type="text" name="PayDetail" size="30" value="<%=Request.QueryString("PayDetail")%>">
				<%end if%>
				</td>
			</tr>
			<tr>
				<td bgColor=#FFFFF7 nowrap style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium" colspan="2">
					<p style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;Special
					Instruction/Multiple Payments:</td>
				<td bgColor=#FFFFF7 style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium" colspan="3">
					<p style="margin-top: 0; margin-bottom: 0">
				<%if created=1 then%>
					<input type="text" name="SpIns" size="30" value="<%=spInst%>" disabled>
				<%else%>
					<input type="text" name="SpIns" size="30" value="<%=Request.QueryString("SpIns")%>">
				<%end if%>
				</td>
			</tr>
			</table>
			</td>
		</tr>
	</table>
	<input type=hidden name="tmpid" value=<%=tmpid%>>
	</script>
	<p style="margin-top: 12px; margin-bottom: 12px" align="center">
	<%if created=1 then%>
		<input type=button name="bModify" value="Modify Account" onclick="window.open('modacc.asp?tmpid=<%=tmpid%>','createacc')" style="font-family: Verdana; font-size: 12px; font-weight: bold">
	<%else%>
		<input type=button name="bCreate" value="Create Account" onclick="if(validate()==1) {document.f1.action='cAcc.asp'; document.f1.submit();}" style="font-family: Verdana; font-size: 12px; font-weight: bold">
	<%end if%>
	</p>
	<HR size=2 color=maroon style="MARGIN-TOP: 12px; MARGIN-BOTTOM: 0px"></td>
</tr>
</table>
</FORM>
</div>
</body>
</HTML>
<%Session("Create")=""
end if
end if
end if
set rs=nothing
set conn=nothing%>