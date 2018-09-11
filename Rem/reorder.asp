<%@ Language=VBScript%>
<%Application.Lock()
if Session("user")<>"sbc" then
	Session("exp")="1"
	Response.Redirect("sbc_pwd.asp")
else
lname=LCase(Request.Form("lname"))
id=Request.Form("id")
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
		rs.Open "SELECT * FROM RecData WHERE REF_No='" & Session("REFNo") & "'", conn
%>		
<HTML>
<head>
<title>APPLICATION FORM</title>
<SCRIPT language="JavaScript" src="images/valid.js"></SCRIPT>
<script language="JavaScript" src="images/button.js"></script>
<script language="JavaScript" src="images/reorder.js"></script>

<STYLE TYPE="text/css">
@import url("images/sbc.css");
BODY {
	SCROLLBAR-BASE-COLOR: #E8FFFF
}
</style> 
</head>
<body bgcolor="#666699" onkeyup="keyAction();" onload="FP_preloadImgs(/*url*/'images/button5D.gif', /*url*/'images/button5E.gif', /*url*/'images/button62.gif', /*url*/'images/button63.gif')">
<div align="center">
<FORM name=f1 method=post style="FONT-FAMILY: Tahoma" action="ModOrd.asp">
			<input type=hidden name="id" value="<%=id%>">
			<input type=hidden name="lname" value="<%=lname%>">
<table border=0 width="660" id="tb1" cellspacing="0" bgcolor="#deebff" bordercolor="#666699">
	<tr>
		<td colspan="2">
		<p style="MARGIN-TOP: 0px; margin-bottom:0" align="center" class=headkh>
		sUmbMeBjTMrg;lixitesñIsMu</p>
		<p style="MARGIN-TOP: 0px" align="center" class=headeng><b>APPLICATION 
		FORM FOR ONLINE-REMITTANCE</b></p>
		<p align="right" style="MARGIN-TOP: 0px" class="td10pt"><b>SBC Bank Co., 
		LTD of Cambodia</b>
		<HR size=2 color=maroon style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px"></td>
	</tr>
	<tr>
		<td align="middle" bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px" nowrap width="525">
		<p style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 6px" align="left" class=thw10pt> 
		<span style="letter-spacing: 1px">&nbsp;Please complete all available fields 
		here, and be valid...</span>&nbsp;
		<td align="middle" bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px" nowrap width="131" class="td10pt">
		<p>
		<b>&nbsp;<a onmouseover="status='Close this form... '; return true" onmouseout="status=''" href="javascript:window.close()"><font color="#FFFFFF">CLOSE FORM</font></a></b><tr bgcolor="#bdc3ce" class="tdb10pt" nowrap>
	<td bordercolor="#000080" bgcolor="#FFFFF0" colspan="2">
	<table border="1" width="100%" id="tbl1" bgcolor="#ffffff" cellspacing="0" bordercolor="#bdc3ce" class="thb10pt" cellpadding="6">
		<tr>
			<td width="25%" bgColor=#FFFFF7>
			<p style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px" align="left">CURRENCY 
			: <select size="1" name="D2">
				<option>Choose one:</option>
				<%for i=0 to lenArr-1
					if rArr(i)=rs.Fields("Currency") and rs.Fields("Currency")<>"USD" then%>
						<option selected><%=rArr(i)%></option>
					<%else%>
						<option><%=rArr(i)%></option>
					<%end if%>
				<%next
				if rs.Fields("Currency")="USD" then%>
					<option selected>USD</option>
				<%else
					Response.Write("<option>USD</option>")
				end if%>
			</select></p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px" align="left">AMOUNT :&nbsp;&nbsp;&nbsp;
			<input type=text name="amnt" size="16" style="text-align: right" maxlength="15" value="<%=rs.Fields("Amount")%>">
			<p style="MARGIN-TOP: 12px; MARGIN-BOTTOM: 0px">
			LOCAL BANK CHARGES FOR :</p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<%if rs.Fields("Charge_for")="0" then
					Response.Write("<input type='radio' name='R3' value='0' checked>")
				else%>
					<input type="radio" name="R3" value="0">
				<%end if%>
				My/Our Account</p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<%if rs.Fields("Charge_for")="1" then
					Response.Write("<input type='radio' name='R3' value='1' checked>")
				else%>
					<input type="radio" name="R3" value="1"> 
				<%end if%>
				Beneficiary's Account (RECIPIENT)
				</p></td>
			<td width="25%" bgColor=#FFFFF7 rowspan="2">
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">Beneficiary's Name (RECIPIENT) :</p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">&nbsp;&nbsp; First 
			Name : <input name="bfname" size="17" maxlength="15" value="<%=rs.Fields("BFName")%>"></p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">&nbsp;&nbsp; Last 
			Name : <input name="blname" size="17" maxlength="15" value="<%=rs.Fields("BLName")%>"></p></p>
			<p style="MARGIN-TOP: 12px; MARGIN-BOTTOM: 0px" align="left">
			Beneficiary's Address :</p>
			<p style="MARGIN-TOP: 0; MARGIN-BOTTOM: 0px">&nbsp;&nbsp; 
			House/Building No. : <input name="bhb" size="5" maxlength="5" value="<%=rs.Fields("BHB")%>"></p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">&nbsp;&nbsp; Street : <input name="bstreet" size="16" maxlength="16" value="<%=rs.Fields("BStreet")%>"></p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">&nbsp;&nbsp; State/City : <input name="bcity" size="21" maxlength="20" nowrap value="<%=rs.Fields("BCity")%>">
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">&nbsp;&nbsp; Country : <input name="bcountry" size="25" maxlength="20" nowrap value="<%=rs.Fields("BCountry")%>">
			</p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">&nbsp;</p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">&nbsp;</p></td>
		</tr>
		<tr>
			<td width="25%" bgColor=#FFEBD6>
			<p style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px">Beneficiary's 
			BANK/City/Country :</p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">
			&nbsp;&nbsp; Beneficiary's BANK : <input name="bbank" size="17" maxlength="16" value="<%=rs.Fields("BBank")%>"></p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 0px">
			&nbsp;&nbsp; State/City :<input name="bbcity" size="12" maxlength="15" value="<%=rs.Fields("bbCity")%>"> </p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 3px">&nbsp;&nbsp; Country 
			:&nbsp;<input name="bbcountry" size="14" maxlength="29" nowrap value="<%=rs.Fields("bbCountry")%>"></p>
			<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 3px">&nbsp;&nbsp; Account 
			No. : <input name="baccn" size="17" maxlength="16" value="<%=rs.Fields("BAccNo")%>"></p>
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
	<img border="0" id="img18" src="images/button61.gif" height="23" width="90" alt="Reset" onclick="document.f1.reset();" onmouseover="FP_swapImg(1,0,/*id*/'img18',/*url*/'images/button62.gif')" onmouseout="FP_swapImg(0,0,/*id*/'img18',/*url*/'images/button61.gif')" onmousedown="FP_swapImg(1,0,/*id*/'img18',/*url*/'images/button63.gif')" onmouseup="FP_swapImg(0,0,/*id*/'img18',/*url*/'images/button62.gif')" fp-style="fp-btn: Metal Capsule 1; fp-font: Verdana; fp-font-style: Bold; fp-font-color-normal: #00008B; fp-font-color-hover: #037EAD; fp-font-color-press: #B00000; fp-transparent: 1; fp-proportional: 0" fp-title="Reset"></p>
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