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
		rs.Open "SELECT * FROM LogDataApp WHERE Report='0' ORDER BY DateApp ASC, REF_No ASC"
		Session("oForms")=""
		Session("Bookmark")=""
		Session("Balance")=""
		Session("numApp")=""
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
<FORM name=f1 method=post action="adlogout.asp">
<table border=2 width="680" id="tb1" cellspacing="0" bgcolor="#e7fbfe" bordercolor="#333366">
	<tr>
		<td colspan="2" class="td10pt">
		<p style="MARGIN-TOP: 12px; margin-bottom:0" align="center" class="headeng"><span style="letter-spacing: 2px">FORM 
            REPORT </span></p>
		<p align="right" style="MARGIN-TOP: 6px"><b>SBC Bank Co., LTD of Cambodia</b>
		<p align="right" style="MARGIN-TOP: 0px">You are login as : <b><%=Session("AdminName")%></b>
		<HR size=2 color="#800000" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px"></td>
	</tr>
	<tr class="thw10pt">
		<td bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-right-style:none; border-right-width:medium" nowrap width="500">
			&nbsp; Recent approved forms are listed as the following table: 
		<td bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-left-style:none; border-left-width:medium" nowrap width="149">
			<p align="center"><a onmouseover="status='Logout from this system...'; return true;" onmouseout="status=''" href="javascript:window.open('adlogout.asp?tmpid=<%=tmpid%>','formapp')">
			<font color="#FFFFFF">LOG OUT</font></a><tr bgcolor="#bdc3ce" nowrap>
	<td bordercolor="#000080" bgcolor="#FFFFF0" colspan="2">
	
    <table width="100%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolor="#333333" bgcolor=#FFFFF7 class="td10pt" id="table2">
            <tr bgcolor="#037ead" class="thw10pt"> 
              <td width="5%" rowspan="2" nowrap align="center">#</td>
              <td width="10%" rowspan="2" nowrap align="center">Bank REF</td>
              <td width="15%" rowspan="2" nowrap align="center">Login Account</td>
              <td width="15%" rowspan="2" nowrap align="center">Charge For</td>
              <td width="15%" rowspan="2" nowrap align="center">Amount<p style="margin-top: 3px">
				(In US$)</td>
              <td width="15%" rowspan="2" nowrap align="center">Amount 
				<p style="margin-top: 3px">To Send</td>
              <td colspan="4" nowrap align="center">BENEFICIARY'S INFORMATION</td>
            </tr>
            <tr bgcolor="#037ead" class="thw10pt"> 
              <td width="15%" align="center" nowrap>Name</td>
              <td width="15%" align="center" nowrap> Account No</td>
              <td width="15%" align="center" nowrap> Bank</td>
              <td width="10%" align="center" nowrap>Bank Location</td>
            </tr>
         <%i=0
         do until rs.EOF%>    	
         	<input type=hidden name="ref<%=i%>" value="<%=rs.Fields("REF_No")%>">
            <tr>
              <td align="center" nowrap><%count=count+1%><%=count%></td>
              <td align="center" nowrap><%=rs.Fields("REF_No")%></td>
              <td align="center" nowrap><%=rs.Fields("AccID")%></td>
              <td align="center" nowrap>
               	<%if rs.Fields("Charge_for")="0" then%><%="Me/Our Account"%>
              	<%else%><%="Beneficiary's Account"%>
              	<%end if%>
				</td>
              <td align="center" nowrap><%=FormatNumber(rs.Fields("Amount")/rs.Fields("Rate"),2)%></td>
              <td align="center" nowrap>
              	<%if rs.Fields("Charge_for")="1" then
					set rs1=Server.CreateObject("ADODB.Recordset")
					rs1.Open "SELECT Commission, CCable, AgentC FROM AccTrans WHERE REFNo='" & rs.Fields("REF_No") & "'", conn
					Response.Write(FormatNumber((rs.Fields("Amount")/rs.Fields("Rate"))-rs1.Fields("Commission")-rs1.Fields("CCable")-rs1.Fields("AgentC"),2))
					rs1.Close()
				else
					Response.Write(FormatNumber(rs.Fields("Amount")/rs.Fields("Rate"),2))
				end if%>
              	</td>
              <td nowrap align="center"><%=rs.Fields("BFName") & " " & rs.Fields("BLName")%></td>
              <td nowrap align="center"><%=rs.Fields("bAccNo")%></td>
              <td nowrap align="center"><%=rs.Fields("BBank")%></td>
              <td nowrap align="center"><%=rs.Fields("bbCity") & ", " & rs.Fields("bbCountry")%></td>
            </tr>
          <%rs.MoveNext()
          i=i+1
          loop%>
          </table>
	<p style="margin-top: 12px; margin-bottom: 12px" align="center">
		&nbsp;</p>
	<p style="margin-top: 12px; margin-bottom: 12px" align="center">
		<input type=hidden name="tmpid" value="<%=tmpid%>">
		<input type=hidden name="dreport" value="1">
		<input type="button" value="Done Report" name="bDone" onclick="document.f1.submit()" style="font-family: Verdana; font-weight: bold"></p>
	<HR size=2 color="#800000" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px"></td>
            </tr>
     </table>
	<p></td>
</tr>
</table>
</p>
</FORM>
</div>
</body>
</HTML>
<%end if
end if
end if
set rs=nothing
set conn=nothing%>