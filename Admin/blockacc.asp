<%@ Language=VBScript%>
<%Application.Lock()
tmpid=Request.Form("tmpid")
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
%>
<HTML>
<title>Block Online Remittance Account... </title>

<STYLE TYPE="text/css">
@import url("images/admin.css");
BODY {
	SCROLLBAR-BASE-COLOR: #E8FFFF
}
</style> 

<body bgcolor="#474545" background="images/cordurouy.gif">
<div align="center">
<FORM name=f1>
<table border=1 width="600" id="tb1" cellspacing="0" bgcolor="#e7fbfe" bordercolor="#666699">
	<tr>
		<td colspan="2">
		<p style="MARGIN-TOP: 0px; margin-bottom:0" align="center" class=headeng>&nbsp;</p>
		<p style="MARGIN-TOP: 0px" align="center" class="headeng"><span style="letter-spacing: 2px">FORM APPROVAL</span></p>
		<p align="right" style="MARGIN-TOP: 0px" class="td10pt"><b>SBC Bank Co., LTD of Cambodia</b>
		<p align="right" style="MARGIN-TOP: 0px" class="td10pt">You are logging in as : <b><%=Session("AdminName")%></b>
		<HR size=2 color=maroon style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px"></td>
	</tr>
	<tr>
		<td align="middle" bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-right-style:none; border-right-width:medium" nowrap width="292">
			<p style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 6px" class=thw10pt align="left"> &nbsp;
			Recent submitted forms are listed below:
		<td align="middle" bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-left-style:none; border-left-width:medium" nowrap width="298">
			<p align="right" class="thw10pt">Previous Form | Next Form
		
		<tr bgcolor="#bdc3ce" class="tdb10pt" nowrap>
	<td bordercolor="#000080" bgcolor="#FFFFF0" colspan="2">
	
	<table border="1" width="100%" id="tbl1" bgcolor="#ffffff" cellspacing="0" bordercolor="#bdc3ce" class="td10pt" cellpadding="6" style="letter-spacing: 1px">
		<tr>
			<td width="20%" bgColor=#FFEBD6 nowrap>&nbsp;</td>
			<td bgColor=#FFFFF7>
			&nbsp;</td>
			<td width="20%" bgColor=#FFEBD6>&nbsp;</td>
		</tr>
		</table>
	<p style="margin-top: 12px; margin-bottom: 12px" align="center">
		&nbsp;</p>
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
set conn=nothing%>