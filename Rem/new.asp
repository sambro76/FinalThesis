<%@ Language=VBScript%>
<%Application.Lock()
if Session("user")<>"sbc" or Session("lout")="1" then
	Session("exp")="1"
	Response.Redirect("sbc_pwd.asp")
else
lname=LCase(Request.QueryString("lname"))
id=Request.QueryString("id")
if lname="" then
	Response.Redirect("errpage.htm")
else
	set conn=Server.CreateObject("ADODB.connection")
	DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("RemDB.MDB")
	conn.open DSN
	set rs=Server.CreateObject("ADODB.Recordset")
	sql="SELECT AccID FROM Keyin WHERE AccID='" & lname & "' AND id='" & id & "'"
	rs.Open sql, conn
	if rs.eof or rs.bof then
		rs.Close()
		Response.Redirect("errpage.htm")
	else
		rs.Close()
		rs.Open "SELECT Curr_Type, Country FROM Rate", conn
		dim rArr()
		dim cArr()
		do until rs.EOF
			lenArr=lenArr+1
			rs.MoveNext()
		loop
		rs.MoveFirst()
		redim rArr(lenArr-1)
		redim cArr(lenArr-1)
		for i=0 to lenArr-1
			rArr(i)=rs.Fields("Curr_Type")
			cArr(i)=rs.Fields("Country")
			rs.MoveNext()
		next
		rs.Close()
		conn.Close()

		DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & left(Server.MapPath("RemDB.MDB"),len(Server.MapPath("RemDB.MDB"))-13) & "\Admin\BankDB.MDB"
		conn.Open DSN
		sql="SELECT * FROM AppInfos WHERE AccID='" & lname & "'"
		rs.open sql, conn
%>
<HTML>
<head>
<title>APPLICATION FORM</title>
<SCRIPT language="JavaScript" src="images/valid.js"></SCRIPT>
<script language="JavaScript" src="images/button.js"></script>
<script language="JavaScript" src="images/order.js"></script>

<STYLE TYPE="text/css">
@import url("images/sbc.css");
BODY {
	SCROLLBAR-BASE-COLOR: #E8FFFF
}
</style>
</head>
<body bgcolor="#666699" onkeyup="keyAction();" onload="FP_preloadImgs(/*url*/'images/button5D.gif', /*url*/'images/button5E.gif', /*url*/'images/button62.gif', /*url*/'images/button63.gif', /*url*/'images/button4A.gif', /*url*/'images/button4B.gif')">
<div align="center">
<FORM name=f1 method=post style="FONT-FAMILY: Tahoma" action="order.asp">
			<input type=hidden name="id" value=<%=Request.QueryString("id")%>>
			<input type=hidden name="lname" value=<%=Request.QueryString("lname")%>>
<table border=0 width="660" id="tb1" cellspacing="0" bgcolor="#deebff" bordercolor="#666699">
	<tr>
		<td>
		<p style="MARGIN-TOP: 0px; margin-bottom:0" align="center" class=headkh>
		sUmbMeBjTMrg;lixitesñIsMu</p>
		<p style="MARGIN-TOP: 0px" align="center" class=headeng><b>APPLICATION
		FORM FOR ONLINE-REMITTANCE</b></p>
		<p align="right" style="MARGIN-TOP: 0px" class="td10pt"><b>SBC Bank Co.,
		LTD of Cambodia</b>
		<HR size=2 color=maroon style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px"></td>
	</tr>
	<tr>
		<td align="middle" bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px" nowrap>
		<p style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 6px" align="left" class=thw10pt>
		<span style="letter-spacing: 1px">&nbsp;Please complete all available fields
		here, and be valid...&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>
		<IMG onmouseup="FP_swapImg(0,0,/*id*/'img16',/*url*/'images/button4A.gif')" onmousedown="FP_swapImg(1,0,/*id*/'img16',/*url*/'images/button4B.gif')" id=img16 onmouseover="FP_swapImg(1,0,/*id*/'img16',/*url*/'images/button4A.gif')" onmouseout="FP_swapImg(0,0,/*id*/'img16',/*url*/'images/button49.gif')" height=25 alt =" HELP!" src="images/button49.gif" width =125 border=0  fp-title=" HELP!" fp-style="fp-btn: Jewel 1; fp-font-style: Bold; fp-font-size: 16; fp-font-color-normal: #800000; fp-font-color-hover: #FFFF00; fp-font-color-press: #FF0000; fp-justify-horiz: 0; fp-transparent: 1">
		<tr bgcolor="#bdc3ce" class="tdb10pt" nowrap>
	<td bordercolor="#000080" bgcolor="#FFFFF0">
	<table border="1" width="100%" id="tbl1" bgcolor="#ffffff" cellspacing="0" bordercolor="#bdc3ce" class="thb10pt" cellpadding="6">
		<tr>
			<td width="50%" bgColor=#FFFFF7><p>Applicant Information :
				</p>
				<p align="right" style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 3px">
				First Name :
					<input name="fname" size="19" disabled style="background-color: #FFFFF7; font-weight:bold" value=<%=rs.fields("FName")%>></p>
				<p align="right" style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 3px">
				Last Name (Family) :
					<input name="lname" size="19" disabled style="background-color: #FFFFF7; font-weight:bold" value=<%=rs.fields("LName")%>></p>
				<p align="right" style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 3px">
				Address :
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
					%>
					<input name="lbladd1" size="29" disabled style="background-color: #FFFFF7; font-weight:bold" value="<%=rs.Fields("HB") & ",&nbsp;Street&nbsp;" & rs.Fields("Street")%>">
					<input name="lbladd2" size="29" disabled style="background-color: #FFFFF7; font-weight:bold" value="<%=city & ",&nbsp;" & country%>"></p>
				<p align="right" style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 3px">
				Contact Phone No.:
					<input name="lbltel" size="15" disabled style="background-color: #FFFFF7; font-weight:bold" value=<%=rs.Fields("Telephone")%>>
			</td>
			<td width="50%" bgColor=#FFFFF7>
			<p style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px" align="left">CURRENCY
			: <select size="1" name="D2">
				<option selected>Choose one:</option>
				<%for i=0 to lenArr-1%>
					<option><%=rArr(i)%></option>
				<%next%>
				<option>USD</option>
			</select></p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px" align="left">AMOUNT :&nbsp;&nbsp;&nbsp;
			<input name="amnt" size="16" style="text-align: right" maxlength="15"><p style="MARGIN-TOP: 12px; MARGIN-BOTTOM: 0px">
			LOCAL BANK CHARGES FOR :</p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="radio" name="R3" value="0"> My/Our Account</p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="radio" name="R3" value="1"> Beneficiary's Account
			(RECIPIENT)</p></td>
		</tr>
		<tr>
			<td width="50%" bgColor=#FFEBD6>
			<p style="MARGIN-BOTTOM: 6px; margin-top:0px">Payment Details : <p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 6px" align="right">
			<TEXTAREA name=S1 rows=3 cols=36 disabled style="background-color: #FFFFF7; font-weight:bold" nowrap><%=rs.Fields("Payment_Details")%></TEXTAREA>
			<p style="MARGIN-TOP: 12px; MARGIN-BOTTOM: 6px">Special
			Instructions/Multiple Payments : <p align="right" style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 3px">
			<TEXTAREA name=S2 rows=3 cols=36 disabled style="background-color: #FFFFF7; font-weight:bold" wrap><%=rs.Fields("SpIns_MPay")%></TEXTAREA></p>
			<p align="right" style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 3px">
			&nbsp;</p>
			</td>
			<td width="50%" bgColor=#FFEBD6>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">Beneficiary's Name
			(RECIPIENT) :</p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">&nbsp;&nbsp; First
			Name : <input name="bfname" size="17" maxlength="15"></p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">&nbsp;&nbsp; Last
			Name : <input name="blname" size="17" maxlength="15"></p></p>
			<p style="MARGIN-TOP: 12px; MARGIN-BOTTOM: 0px" align="left">
			Beneficiary's Address :</p>
			<p style="MARGIN-TOP: 0; MARGIN-BOTTOM: 0px">&nbsp;&nbsp;
			House/Building No. : <input name="bhb" size="5" maxlength="5"></p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">&nbsp;&nbsp; Street : <input name="bstreet" size="16" maxlength="16"></p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">&nbsp;&nbsp;
			State/City :
				<select size="1" name="D4" onchange="if(this.selectedIndex==4) {document.f1.bcityo.style.backgroundColor='#FFFFFF'; document.f1.bcityo.disabled=false; document.f1.bcityo.focus();} else {document.f1.bcityo.value=''; document.f1.bcityo.style.backgroundColor='#FFFFF7'; document.f1.bcityo.disabled=true;}">
				<option selected>Choose one:</option>
				<option>Phnom Penh</option>
				<option>Bangkok</option>
				<option>Tokyo</option>
				<option value=0>Other Specified</option>
			</select><input name="bcityo" size="12" maxlength="15" disabled nowrap style="background-color: #FFFFF7">
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">&nbsp;&nbsp; Country
			:
			<select size="1" name="D5" onchange="if(this.selectedIndex==5) {document.f1.bcountryo.style.backgroundColor='#FFFFFF'; document.f1.bcountryo.disabled=false; document.f1.bcountryo.focus();} else {document.f1.bcountryo.value=''; document.f1.bcountryo.style.backgroundColor='#FFFFF7'; document.f1.bcountryo.disabled=true;}">
				<option selected>Choose one:</option>
				<%for i=0 to lenArr-1%>
					<option><%=cArr(i)%></option>
				<%next%>
				<option>USA</option>
				<option value=0>Other Specified:</option>
			</select><input name="bcountryo" size="14" maxlength="29" disabled nowrap style="background-color: #FFFFF7">
			</p>
			</td>
		</tr>
		<tr>
			<td width="50%" bgColor=#FFFFF7>Settlement by Debiting (Savings or
			Current) :<p align="right" style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">
			<font color="#800000"><span style="FONT-SIZE: 8pt">(In order to
			apply this on-line form, you need</span></font></p>
			<p align="right" style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">
			<font color="#800000"><span style="FONT-SIZE: 8pt">a Bank Account
			registered before... )</span></font></p>
			<p align="right" style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 3px">Your
			account number :
			<input name="accn" size="16" maxlength="15" disabled style="background-color: #FFFFF7; font-weight:bold" value=<%=rs.fields("AccNo")%>></td>
			<%rs.Close()
			conn.Close()%>
			<td width="50%" bgColor=#FFFFF7>
			<p style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px">Beneficiary's
			BANK/City/Country :</p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">
			&nbsp;&nbsp; Beneficiary's BANK : <input name="bbank" size="17" maxlength="16" onmouseover="status='Recipient bank here...'" onmouseout="status=''"></p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">
			&nbsp;&nbsp; State/City :<select size="1" name="D6" onchange="if(this.selectedIndex==4) {document.f1.bbcityo.style.backgroundColor='#FFFFFF'; document.f1.bbcityo.disabled=false; document.f1.bbcityo.focus();} else {document.f1.bbcityo.value=''; document.f1.bbcityo.style.backgroundColor='#FFFFF7'; document.f1.bcityo.disabled=true;}">
				<option selected>Choose one:</option>
				<option>Phnom Penh</option>
				<option>Bangkok</option>
				<option>Tokyo</option>
				<option value=0>Other Specified</option>
			</select><input name="bbcityo" size="12" maxlength="15" disabled nowrap style="background-color: #FFFFF7"> </p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 3px">&nbsp;&nbsp; Country
			:&nbsp;<select size="1" name="D7" onchange="if(this.selectedIndex==5) {document.f1.bbcountryo.style.backgroundColor='#FFFFFF'; document.f1.bbcountryo.disabled=false; document.f1.bbcountryo.focus();} else {document.f1.bbcountryo.value=''; document.f1.bbcountryo.style.backgroundColor='#FFFFF7'; document.f1.bbcountryo.disabled=true;}">
				<option selected>Choose one:</option>
				<%for i=0 to lenArr-1%>
					<option><%=cArr(i)%></option>
				<%next%>
				<option>USA</option>
				<option value=0>Other Specified:</option>
			</select><input name="bbcountryo" size="14" maxlength="29" disabled nowrap style="background-color: #FFFFF7"></p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 3px">&nbsp;&nbsp; Account
			No. : <input name="baccn" size="17" maxlength="16" onmouseover="status='Recipient account number here...'" onmouseout="status=''"></p>
			</td>
		</tr>
	</table>
		<p class=thb10pt style="margin-top: 0; margin-bottom: 0px">
		<input type="checkbox" name="agree" value="ON"> <font color="#800000">
		I/we agree that that you may at your discretion confirm this application
		for remittance with</font></p>
		<p class=thb10pt style="margin-top: 0; margin-bottom: 0px"><font color="#800000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		me/us before acting on it. I/WE HAVE READ and AGREED to the Conditions
		appeared on The</font></p>
		<p class=thb10pt style="margin-top: 0; margin-bottom: 0px">
		<font color="#B00000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Help's Button above.</font></p>
	<p style="margin-top: 12px; margin-bottom: 12px" align="center">
	<img border="0" id="img17" src="images/button58.gif" height="23" width="115" alt="Submit Form" onclick="valid()" fp-style="fp-btn: Metal Capsule 1; fp-font: Verdana; fp-font-style: Bold; fp-font-color-normal: #00008B; fp-font-color-hover: #037EAD; fp-font-color-press: #B00000; fp-transparent: 1" fp-title="Submit Form" onmouseover="FP_swapImg(1,0,/*id*/'img17',/*url*/'images/button5D.gif')" onmouseout="FP_swapImg(0,0,/*id*/'img17',/*url*/'images/button58.gif')" onmousedown="FP_swapImg(1,0,/*id*/'img17',/*url*/'images/button5E.gif')" onmouseup="FP_swapImg(0,0,/*id*/'img17',/*url*/'images/button5D.gif')">
	<img border="0" id="img18" src="images/button61.gif" height="23" width="90" alt="Reset" onclick="resetbtn();" onmouseover="FP_swapImg(1,0,/*id*/'img18',/*url*/'images/button62.gif')" onmouseout="FP_swapImg(0,0,/*id*/'img18',/*url*/'images/button61.gif')" onmousedown="FP_swapImg(1,0,/*id*/'img18',/*url*/'images/button63.gif')" onmouseup="FP_swapImg(0,0,/*id*/'img18',/*url*/'images/button62.gif')" fp-style="fp-btn: Metal Capsule 1; fp-font: Verdana; fp-font-style: Bold; fp-font-color-normal: #00008B; fp-font-color-hover: #037EAD; fp-font-color-press: #B00000; fp-transparent: 1; fp-proportional: 0" fp-title="Reset"></p>
	<HR size=2 color=maroon style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px"></td>
</tr>
</table>
</FORM>
</div>
</body>
</HTML>
<%end if
end if
set rs=nothing
set conn=nothing
end if%>