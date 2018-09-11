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
		Session("Modify")=""
		Session("accid")=""
		Session("Search")=""
		Session("Create")=""
		Session("AccCreated")=""
%>
<HTML>
<title>CUSTOMERS ACCOUNT MANAGEMENT... </title>

<STYLE TYPE="text/css">
@import url("images/admin.css");
BODY {
	SCROLLBAR-BASE-COLOR: #E8FFFF
}
</style> 

<body bgcolor="#474545" background="images/cordurouy.gif" onload="window.name='manacc'">
<div align="center">
<FORM name=f1 method=post target=_self>
<input type=hidden value="<%=tmpid%>" name="tmpid">
<table border=2 width="600" id="tb1" cellspacing="0" bgcolor="#e7fbfe" bordercolor= #BFBAAE>
	<tr>
		<td colspan="2">
		<p style="MARGIN-TOP: 0px; margin-bottom:0" align="center" class=headeng>&nbsp;</p>
		<p style="MARGIN-TOP: 0px" align="center" class="headeng">
		<span style="letter-spacing: 2px">ACCOUNT MANAGEMENT</span></p>
		<p align="right" style="MARGIN-TOP: 0px" class="td10pt"><b>SBC Bank Co., LTD of Cambodia</b>
		<p align="right" style="MARGIN-TOP: 3px; margin-bottom:6px" class="td10pt">You are login as : <b><%=Session("AdminName")%></b>
		</td>
	</tr>
	<tr>
		<td align="middle" bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-right-style:none; border-right-width:medium" nowrap width="535">
			<p style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 6px" class=thw10pt align="left"> 
			&nbsp;Administration activities are made confidentially. Select an item to be occupied...
		<td align="bottom" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-left-style:none; border-left-width:medium" nowrap class="td10pt" onclick="document.f1.action='logout.asp';document.f1.submit();">
			<p align="right"><a onmouseover="status='Log out... '; return true" onmouseout="status=''" href="javascript:window.open('adlogout.asp?tmpid=<%=tmpid%>','manacc')"><b>Log out</b></a></p>
	<tr bgcolor="#bdc3ce" class="tdb10pt" nowrap>
	<td bordercolor="#000060" bgcolor="#FFFFF0" colspan="2">
	
	<table border="1" width="100%" id="tbl1" bgcolor="#ffffff" cellspacing="0" bordercolor="#bdc3ce" class="td10pt" cellpadding="6" style="letter-spacing: 1px">
		<tr>
			<td width="25%" bgColor=#FFEBD6 onclick="document.f1.action='charges.asp';document.f1.submit();">
				<a onmouseover="status='Add/Modify Charges... '; return true" onmouseout="status=''" href="javascript:document.f1.action='charges.asp';document.f1.submit();">			
				<font color="#000084">
					<p align="center" style="margin-bottom: 0">Add/Modify SBC</p>
					<p align="center" style="margin-top: 3px"><b>Charges</b></p>
				</font>
				</a>				
					<%if Session("cCharge")="Charged" then
						Session("cCharge")="Charging"%>
						<font color="#0000FF"><p align="center" style="margin-top: 12px"><b>Successful...</b></p></font>
					<%elseif Session("cCharge")="Error" then
						Session("cCharge")="Charging"%>			
						<font color="#FF0000"><p align="center" style="margin-top: 12px"><b>Error occurs!</b></p></font>
					<%end if%>
			</td>
			<td bgColor=#FFFFF7 colspan="2" onclick="document.f1.action='addmodacc.asp';document.f1.submit();">
				<p align="center">&nbsp;
					<a onmouseover="status='Add/Modify Logging Account for Online Remittance... '; return true" onmouseout="status=''" href="javascript:document.f1.action='addmodacc.asp';document.f1.submit();">
				<p align="center">Create/Modify a <b>login Account</b> for
				<p align="center" style="margin-top: 3px; margin-bottom: 0">Online Remittance Service</p>
				<p align="center" style="margin-top: 3px; margin-bottom: 0">Including : 
				&lt;&lt;Credit Amount&gt;&gt;...</p>
					</a>
				<p align="center">&nbsp;</p>
			
			</td>
			<td width="25%" bgColor=#FFEBD6 onclick="document.f1.action='approver.asp';document.f1.submit();">
			<a onmouseover="status='Add Approver(s)... '; return true" onmouseout="status=''" href="javascript:document.f1.action='approver.asp';document.f1.submit();">
			<font color="#000084">
				<p align="center" style="margin-bottom: 0">Add</p>
				<p align="center" style="margin-top: 3px; margin-bottom: 0"><b>Approver(s)</b></p>
				<p align="center" style="margin-top: 3px"><b>Information</b></p>
			</font>
			</a>	
			</td>
		</tr>
		<tr>
			<td bgColor=#FFFFF7 colspan="2" width="50%" onclick="document.f1.action='search.asp';document.f1.submit();">
			<p align="center"><a onmouseover="status='Search old remittance activities... '; return true" onmouseout="status=''" href="javascript:document.f1.action='search.asp';document.f1.submit();"><b>Search old Remittance Activities</b></a></p></td>
			<td bgColor=#FFFFF7 colspan="2" onclick="document.f1.action='blockacc.asp';document.f1.submit();">
				<p align="center" style="margin-bottom: 0">&nbsp;</p>
				<a onmouseover="status='Block hacked account... '; return true" onmouseout="status=''" href="javascript:document.f1.action='blockacc.asp';document.f1.submit();">
					<p align="center" style="margin-top: 3px; margin-bottom: 0"><b>Block Hacked Account</b> for Online</p>
					<p align="center" style="margin-top: 3px; margin-bottom: 0">Remittance from</p>
					<p align="center" style="margin-top: 3px; margin-bottom: 0">customers complaining</p>
				</a>	
				<p align="center" style="margin-top: 3px; margin-bottom: 0">&nbsp;</p>
				</td>
		</tr>
		<tr>
			<td colspan=4 onclick="window.open('adlogout.asp?tmpid=<%=tmpid%>','manacc')" class="tx12pt">
				<p align="center"><a onmouseover="status='Log out... '; return true" onmouseout="status=''" href="javascript:window.open('adlogout.asp?tmpid=<%=tmpid%>','manacc')">
					<b>Log out</b></a>
				</p>
			</td>
		</tr>
		</table>
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