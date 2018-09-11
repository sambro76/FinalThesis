<%@ Language=VBScript%>
<%Application.Lock()
if Session("admin")<>"sbc" then
	Session("exp")="1"
	Response.Redirect("adlogin.asp")
else
tmpid=Request.QueryString("tmpid")
if tmpid="" then
	Response.Redirect("errpage.htm")
else
	set conn1=Server.CreateObject("ADODB.connection")
	DSN1="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("BankDB.MDB")
	conn1.open DSN1
	set rs1=Server.CreateObject("ADODB.Recordset")
	rs1.Open "SELECT AID FROM AdmInfos WHERE tmpid='" & tmpid & "'", conn1
	if rs1.eof or rs1.bof then
		rs1.Close()
		conn1.Close()
		Response.Redirect("errpage.htm")
	else
		rs1.Close()
		rs1.Open "SELECT * FROM LogDataApp WHERE Report='0'"
		if rs1.EOF or rs1.BOF then
			report=1
		end if
		rs1.Close()
		set conn=Server.CreateObject("ADODB.connection")
		DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & left(Server.MapPath("BankDB.MDB"),len(Server.MapPath("BankDB.MDB"))-17) & "\Rem\RemDB.MDB"
		conn.Open DSN
		set rs=Server.CreateObject("ADODB.Recordset")
		rs.Open "SELECT * FROM RecData WHERE Status='3' ORDER BY Dated ASC, AccID ASC", conn
		if rs.EOF or rs.BOF then
			norec=1
		else
			if Session("oForms")="" then
				do until rs.EOF
					Session("oForms")=Session("oForms")+1
					rs.MoveNext()
				loop
				rs.MoveFirst()
				Session("Bookmark")=1
				Session("Approver")=Session("Approver")
			end if
			
			debit=rs.Fields("Amount")/rs.Fields("Rate")
			if rs.Fields("Charge_For")=0 then
				rs1.Open "SELECT * FROM SBCCharges WHERE CostPoint>" & debit, conn1
				if rs1.EOF or rs1.BOF then
					rs1.Close()
					rs1.Open "SELECT * FROM SBCCharges ORDER BY CostPoint DESC", conn1
				end if
				cSum=rs1.Fields("Commission")+rs1.Fields("CCable")+rs1.Fields("AgentC")
				rs1.Close()
			else
				cSum=0
			end if
			
			rs1.Open "SELECT Balance, AccTrans.AccNo FROM AccTrans, AppInfos WHERE AccTrans.AccNo=AppInfos.AccNo AND AccID='" & rs.Fields("AccID") & "' ORDER BY Dated DESC", conn1
			Session("Balance")=rs1.Fields("Balance")
			Session("AccNo")=rs1.Fields("AccNo")
		end if
		conn1.Close()
	mincredit=50
%>
<HTML>
<head>
<title>APPLICATION FORM APPROVAL... </title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<STYLE TYPE="text/css">
@import url("images/admin.css");
BODY {
	SCROLLBAR-BASE-COLOR: #E8FFFF
}
</style>

</head>
<body background="images/cordurouy.gif" onload="window.name='formapp'">
<div align="center">
<FORM name=f1 method=post action="frmseen.asp">
<table border=2 width="640" id="tb1" cellspacing="0" bgcolor="#e7fbfe" bordercolor="#9696CD">
	<tr>
		<td colspan="2" class="td10pt">
		<p style="MARGIN-TOP: 12px; margin-bottom:0" align="center" class="headeng"><span style="letter-spacing: 2px">FORM APPROVAL</span></p>
		<p align="right" style="MARGIN-TOP: 6px"><b>SBC Bank Co., LTD of Cambodia</b>
		<p align="right" style="MARGIN-TOP: 0px">You are login as : <b><%=Session("AdminName")%></b>
		<HR size=2 color="#800000" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px"></td>
	</tr>
	<tr class="thw10pt">
		<td bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-right-style:none; border-right-width:medium" nowrap width="497">
			<%if report=0 then%>
				&nbsp; <a onmouseover="status='Some Approved Form(s) Not Yet Made Report... '; return true" onmouseout="status=''" href="frmReport.asp?tmpid=<%=tmpid%>">
			<font color="#FFFFFF">Some approved form(s) attending the report
			</font>
			</a> <p>
			<%end if%> 
			&nbsp; Recent submitted forms are listed as the following pages: 
		<td bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-left-style:none; border-left-width:medium" nowrap width="126">
			<p align="center"><a onmouseover="status='Logout from this system...'; return true;" onmouseout="status=''" href="javascript:window.open('adlogout.asp?tmpid=<%=tmpid%>','formapp')">
			<font color="#FFFFFF">LOG OUT</font></a><tr bgcolor="#bdc3ce" nowrap>
	<td bordercolor="#000080" bgcolor="#FFFFF0" colspan="2">
	
	<table border="1" width="100%" id="tbl1" bgcolor="#800080" cellspacing="0" bordercolor="#bdc3ce" class="td10pt" cellpadding="6" style="letter-spacing: 1px">
		<tr>
			<td width="30%" bgColor=#FFEBD6 nowrap rowspan="2">
			<%if report=0 then%>
				<a onmouseover="status='Some Approved Form(s) Not Yet Made Report... '; return true" onmouseout="status=''" href="frmReport.asp?tmpid=<%=tmpid%>">
				<p align="center" style="margin-top: 6px; margin-bottom: 0"><b>Report</b></p>
			<p align="center" style="margin-top: 6px; margin-bottom: 0"><b>Approved</b></p>
			<p align="center" style="margin-top: 6px"><b>Forms</b></p></a>
			<%end if%>
			</td>
		</tr>
		<tr>
			<td bgColor=#E7FBFF>
			<%if norec=1 then
				Response.Write("There is no new or more form to be approved...<p><p><p><p><p>")
			else%>
			<table border="1" cellpadding="6" cellspacing="0" width="100%" bordercolor="#BDC3CE" bgcolor=#FFFFF7 id="table2" class="td10pt">
				<tr>
					<td nowrap style="border-right-style: none; border-right-width: medium" colspan="2" bgcolor="#0394CB">
			<p><font color="#FFFFFF"><b>&nbsp;
				<%if Session("numApp")>0 then
					if Session("numApp")=1 then
						Response.Write("1 Form Recently Approved")
					else 
						Response.Write(Session("numApp") & " Forms Recently Approved")
					end if
				end if%>
					</b></font></td>
					<td nowrap style="border-right-style: solid; border-right-width: 1px; border-left-style:none; border-left-width:medium" colspan="2" bgcolor="#0394CB">
			<p align="right"><font color="#FFFFFF"><b>Form No.: <%=Session("Bookmark")%> of <%=Session("oForms")%></b></font></td>
				</tr>
				<tr>
					<input type="hidden" name="tmpid" value="<%=tmpid%>">
					<input type="hidden" name="refno" value="<%=rs.Fields("REF_No")%>">
					<input type="hidden" name="AccID" value="<%=rs.Fields("AccID")%>">
					<td nowrap style="border-right-style: solid; border-right-width: 1px" bgcolor="#FFEBD6">
                      <b>Bank REF. No: <%=rs.Fields("REF_No")%></b></td>
					<td width="21%" nowrap align="center" bgcolor="#FFEBD6"><b>Field Requested</b></td>
					<td width="5%" nowrap align="center" bgcolor="#FFEBD6"><b>OK?</b></td>
					<td width="24%" align="center" nowrap bgcolor="#FFEBD6"><b>Comment</b></td>
				</tr>
				<tr>
					<td width="43%" nowrap>
					<b>Applicant ID:</b></td>
					<td width="18%"><b><%=rs.Fields("AccID")%></b></td>
					<td width="17%">
					<p align="center">
					<input type="checkbox" name="cAccID" value="1" disabled checked></td>
					<td width="17%">&nbsp;</td>
					</tr>
				<tr>
					<td width="43%" nowrap>Local Bank Charge For:</td>
					<td width="21%" nowrap><b>
						<%if rs.Fields("Charge_for")=0 then%><%="Me/Our Acount"%>
						<%else%><%="Beneficiary's Account"%>
						<%end if%>
						</b></td>
					<td width="5%" align="center">
					<input type="checkbox" name="cCharge" value="1" checked disabled></td>
					<td width="24%" align="center">
					&nbsp;</td>
				</tr>
				<tr>
					<td width="43%" nowrap>
					<b>Amount Requested (In USD):</b></td>
					<td width="21%"><b><%=debit%></b></td>
					<td width="5%" align="center" rowspan="2">
					<%if debit+cSum > Session("Balance")-mincredit then
						Response.Write("<input type='hidden' name='cAmount' value=0>")
						Response.Write("<input type='checkbox' name='cAmnt' disabled>")
					else%>
						<input type=hidden name="cAmount" value="1">
						<input type="checkbox" name="cAmnt" value="1" disabled checked>
					<%end if%>
					</td>
					<td width="24%" align="center" rowspan="2">
					<%if debit+cSum > Session("Balance")-mincredit then
						Response.Write("<input type='hidden' name='eAmount' value='Not Enough'>")
						Response.Write("<input type='text' name='eAmnt' size='21' value='Not Enough' class='td10pt' disabled>")
					else%>
						<input type="text" name="eAmount" size="21" class="td10pt" value="Amount is ok" disabled>
					<%end if%>
					</td>
				</tr>
				<tr>
					<td width="43%" nowrap>
					<b>Last Balance (In USD):</b></td>
					<td width="21%"><b><%=FormatNumber(Session("Balance"),2)%></b></td>
				</tr>
				<tr>
					<td width="43%" nowrap>Beneficiary's Name (First, Last):</td>
					<td width="21%" nowrap><b><%=rs.Fields("BFName") & " " & rs.Fields("BLName")%></b></td>
					<td width="5%" align="center">
					<input type="checkbox" name="cBName" value="1"></td>
					<td width="24%" align="center">
					<input type="text" name="eBName" size="21" class="td10pt"></td>
				</tr>
				<tr>
					<td nowrap style="border-bottom-style: none; border-bottom-width: medium">Beneficiary's Address:</td>
					<td nowrap><b><%=rs.Fields("BHB") & ", Street " & rs.Fields("BStreet")%></b></td>
					<td width="5%" align="center" rowspan="2">
					<input type="checkbox" name="cAddress" value="1"></td>
					<td width="24%" align="center" rowspan="2">
					<input type="text" name="eAddress" size="21" class="td10pt"></td>
				</tr>
				<tr>
					<td nowrap style="border-top-style: none; border-top-width: medium">&nbsp;</td>
					<td nowrap><b><%=rs.Fields("BCity") & ", " & rs.Fields("BCountry")%></b></td>
				</tr>
				<tr>
					<td nowrap>Beneficiary's Bank:</td>
					<td nowrap><b><%=rs.Fields("BBank")%></b></td>
					<td width="5%" align="center">
					<input type="checkbox" name="cBBank" value="1"></td>
					<td width="24%" align="center">
					<input type="text" name="eBBank" size="21" class="td10pt"></td>
				</tr>
				<tr>
					<td nowrap>
					<p align="right">Bank Location:</td>
					<td nowrap><b><%=rs.Fields("bbCity") & ", " & rs.Fields("bbCountry")%></b></td>
					<td width="5%" align="center">
					<input type="checkbox" name="cLocation" value="1"></td>
					<td width="24%" align="center">
					<input type="text" name="eBLocation" size="21" class="td10pt"></td>
				</tr>
				<tr>
					<td nowrap>
					<p align="right">Beneficiary Account Number:</td>
					<td nowrap><b><%=rs.Fields("bAccNo")%></b></td>
					<td width="5%" align="center">
					<input type="checkbox" name="cBAccNo" value="1"></td>
					<td width="24%" align="center">
					<input type="text" name="eBAccNo" size="21" class="td10pt"></td>
				</tr>
				<tr>
					<td nowrap>Date of Request: </td>
					<td nowrap colspan="3"><b><%=rs.Fields("Dated")%></b></td>
				</tr>
			</table>
			<%rs.Close()
			conn.Close()
			end if%>
			</td>
		</tr>
		</table>
	<script language="javascript">
		function seen(){
		if(document.f1.cBName.checked==0&&document.f1.eBName.value.length<3){window.alert("You have not verified the Beneficiary's Name... "); document.f1.eBName.select();}
		else if(document.f1.cAddress.checked==0&&document.f1.eAddress.value.length<3){window.alert("You have not verified the Beneficiary's Address... "); document.f1.eAddress.select();}
		else if(document.f1.cBBank.checked==0&&document.f1.eBBank.value.length<3){window.alert("You have not verified the Beneficiary's Bank... "); document.f1.eBBank.select();}
		else if(document.f1.cLocation.checked==0&&document.f1.eBLocation.value.length<3){window.alert("You have not verified the Bank Location... "); document.f1.eBLocation.select();}
		else if(document.f1.cBAccNo.checked==0&&document.f1.eBAccNo.value.length<3){window.alert("You have not verified the Beneciary Account Number... "); document.f1.eBAccNo.select();}
		else{document.f1.submit();}
		}
	</script>	
	<p style="margin-top: 12px; margin-bottom: 12px" align="center">
	<%if norec=1 then%>
		<input type="button" value="Logout" name="bLogout" onclick="window.open('adlogout.asp?tmpid=<%=tmpid%>','formapp')" style="font-family: Verdana; font-weight: bold">
	<%else%>	
		<input type="button" value="Seen This Form" name="bSeen" onclick="seen()" style="font-family: Verdana; font-weight: bold">
	<%end if%>
	</p>
	<HR size=2 color="#800000" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px"></td>
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