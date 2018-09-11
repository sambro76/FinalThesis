<%@ Language=VBScript%>
<%Application.Lock()
if Session("admin")<>"sbc" then
	Session("exp")="1"
	Response.Redirect("adlogin.asp")
else

tmpid=Request.Form("tmpid")
if tmpid="" then tmpid=Request.QueryString("tmpid")
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
<head>
<title>Create/Modify/Delete Account... </title>
<script language="JavaScript" src="../admin/images/button.js"></script>
<script language="JavaScript" src="../rem/images/valid.js"></script>

<STYLE TYPE="text/css">
@import url("images/admin.css");
BODY {
	SCROLLBAR-BASE-COLOR: #E8FFFF
}
</style>
</head>
<body bgcolor="#474545" background="images/cordurouy.gif" onload="window.name='addmodacc'">
<div align="center">
<FORM method=get name=f1 target="_self">
<table border=1 width="600" id="tb1" cellspacing="0" bgcolor="#e7fbfe" bordercolor="#666699">
	<tr>
		<td colspan="2">
		<p style="MARGIN-TOP: 0px; margin-bottom:0" align="center" class=headeng>&nbsp;</p>
		<p style="MARGIN-TOP: 0px" align="center" class="headeng">
		<span style="letter-spacing: 2px">REMITTANCE ACCOUNT INFORMATION</span></p>
		<p align="right" style="MARGIN-TOP: 0px" class="td10pt"><b>SBC Bank Co., LTD of Cambodia</b>
		<p align="right" style="MARGIN-TOP: 0px" class="td10pt">You are login as : <b><%=Session("AdminName")%></b>
		<HR size=2 color="#800000"></td>
	</tr>
	<tr>
		<td align="middle" bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-right-style:none; border-right-width:medium" nowrap width="408">
			<p style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 6px" class=thw10pt align="left"> &nbsp;
			<a onmouseover="status='Create new account... '; return true" onmouseout="status=''" style="color: #FFFFFF" href="javascript:window.open('createacc.asp?tmpid=<%=tmpid%>','addmodacc')">Create New
			</a>
		<td align="middle" bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-left-style:none; border-left-width:medium" nowrap width="184">
			<p class="thw10pt">
				<A style="color: #FFFFFF" onmouseover="status='Back to main menu... '; return true" onmouseout="status=''" href="javascript:window.open('manacc.asp?tmpid=<%=tmpid%>','addmodacc')">
				Main Menu</A>&nbsp;&nbsp; |&nbsp;&nbsp;&nbsp;
				<A style="color: #FFFFFF" onmouseover="status='Log out... '; return true" onmouseout="status=''" href="javascript:window.open('adlogout.asp?tmpid=<%=tmpid%>','addmodacc')">
				Log out</A>
			</p>
		</td>
	<tr bgcolor="#bdc3ce" class="tdb10pt" nowrap>
		<td bordercolor="#000080" bgcolor="#FFFFF0" colspan="2">

	<table border="1" width="100%" id="tbl1" bgcolor="#ffffff" cellspacing="0" bordercolor="#bdc3ce" class="td10pt" cellpadding="6" style="letter-spacing: 1px">
		<tr>
			<td width="19%" bgColor="#FFEBD6" rowspan="5" nowrap onclick="window.open('createacc.asp?tmpid=<%=tmpid%>','addmodacc')">
				<a onmouseover="status='Create new account... '; return true" onmouseout="status=''" href="javascript:window.open('createacc.asp?tmpid=<%=tmpid%>','addmodacc')">
				<p align="center" style="margin-bottom: 0"><b>Create New</b></p>
				<p align="center" style="margin-top: 3px; margin-bottom: 0"><b>Account</b></p>
				</a>
			</td>
			<td bgColor="#FFFFF7" colspan="2">
				<%if Session("Delete")="1" then%>
					<b>One record has been deleted...</b>
				<%end if%>
				<%if Session("Search")="0" then%>
					<p style="margin-top: 6px; margin-bottom:0"><font color="#FF0000">There is no record with this information... </font></p>
					<p style="margin-top: 6px; margin-bottom:0">Please retry again:</p>
				<%else%>
					<p style="margin-top: 0px; margin-bottom:0">Input a Personal Information of APPLICANT...</p>
					<p style="margin-top: 6px; margin-bottom:0">Search by:</p>
				<%end if%>
			</td>
		</tr>
		<tr>
				<td bgColor="#FFFFF7" width="40%" style="border-right-style: none; border-right-width: medium; border-bottom-style: none; border-bottom-width: medium">
					<p style="margin-top: 0; margin-bottom:0">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input type="checkbox" name="PIA1" value="AccID" onclick="if(this.checked) {document.f1.txtPIA1.disabled=false;document.f1.txtPIA1.focus();} else {document.f1.txtPIA1.value=''; document.f1.txtPIA1.disabled=true;}"> Account Identification
				</td>
				<td width="35%" bgColor="#FFFFF7" style="border-left-style: none; border-left-width: medium; border-bottom-style: none; border-bottom-width: medium">
					<p style="margin-top: 0; margin-bottom: 0">
					<input type="text" name="txtPIA1" size="20" disabled maxlength="18"></p>
				</td>
			</tr>
			<tr>
				<td bgColor="#FFFFF7" width="40%" style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium">
					<p style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input type="checkbox" name="PIA2" value="FName" onclick="if(this.checked) {document.f1.txtPIA2.disabled=false;document.f1.txtPIA2.focus();} else {document.f1.txtPIA2.value=''; document.f1.txtPIA2.disabled=true;}">
					First Name</td>
				<td width="35%" bgColor="#FFFFF7" style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium">
					<p style="margin-top: 0; margin-bottom: 0">
					<input type="text" name="txtPIA2" size="20" disabled maxlength="18"></td>
			</tr>
			<tr>
				<td bgColor="#FFFFF7" width="40%" style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium">
				<p style="margin-top: 0; margin-bottom:0">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="checkbox" name="PIA3" value="LName" onclick="if(this.checked) {document.f1.txtPIA3.disabled=false;document.f1.txtPIA3.focus();} else {document.f1.txtPIA3.value=''; document.f1.txtPIA3.disabled=true;}"> Last Name
				(Family)</td>
				<td width="35%" bgColor="#FFFFF7" style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium">
				<p style="margin-top: 0; margin-bottom: 0">
				<input type="text" name="txtPIA3" size="20" disabled maxlength="18"></td>
			</tr>
			<tr>
				<td bgColor="#FFFFF7" width="40%" style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium">
				<p style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="checkbox" name="PIA4" value="AccNo" onclick="if(this.checked) {document.f1.txtPIA4.disabled=false;document.f1.txtPIA4.focus();} else {document.f1.txtPIA4.value=''; document.f1.txtPIA4.disabled=true;}"> Bank Account Number
				</td>
				<td width="35%" bgColor="#FFFFF7" style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium">
				<p style="margin-top: 0; margin-bottom: 0">
				<input type="text" name="txtPIA4" size="20" disabled maxlength="18"></td>
			</tr>
		</table>
	<p style="margin-top: 12px; margin-bottom: 12px" align="center">
		<input type=hidden name="tmpid" value="<%=tmpid%>">
		<script language="javascript">
			function valid(){
				if(document.f1.txtPIA1.value==""&&document.f1.txtPIA2.value==""&&document.f1.txtPIA3.value==""&&document.f1.txtPIA4.value==""){
					window.alert("Please select and fill in at least one text box... ");
				}
				else {
					document.f1.action='searchacc.asp';
					document.f1.submit();
				}
			}
		</script>
		<input type=button name="bSearch" value="Search Now" onclick="valid()" style="font-family: Verdana; font-size: 12px; font-weight: bold"></p>
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
Session("Search")=false
Session("accid")=false
Session("Modify")=false
Session("Delete")=false
set rs=nothing
set conn=nothing%>