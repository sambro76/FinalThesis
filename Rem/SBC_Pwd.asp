<%@ Language=VBScript%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="javascript" src="../rem/images/button.js"></script>
<script language="javascript" src="../rem/images/func.js"></script>

<title>WELCOME TO REMITTANCE SERVICE, SBC Bank Of Cambodia</title>
<style type=text/css>
@import url("images/sbc.css");
BODY {SCROLLBAR-BASE-COLOR: "#dff7fd"}
</style>

</head>
<body style="TEXT-ALIGN: center" bgcolor=#dff7fd onload="window.name='main'; window.scroll(0,40); FP_preloadImgs(/*url*/'images/buttonC4.gif', /*url*/'images/buttonC5.gif', /*url*/'images/buttonE3.gif', /*url*/'images/buttonE4.gif'); document.f1.lname.focus()">
<form method=post name=f1>
<input type=hidden name=id>
<table border="0" width="100%" id="tb">
	<tr>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td class=10pt nowrap>
		<p align="center" style="MARGIN-BOTTOM: 0px" class="headkh"><a href="http://m.1asphost.com/civsit/service">
		<span style="TEXT-DECORATION: none">sUmsVaKmn_esvaepÞrR)ak;</span></a></p>
		<p style="MARGIN-TOP: 0px" align="center" class="headeng"><b>WELCOME TO
		APPLICATION FOR REMITTANCE SERVICE</b></p>
		<p align="right" style="margin-top: 0" class="td10pt"><b>SBC Bank Co.,
		LTD of Cambodia</b>
		<HR size=2 color=maroon style="margin-top: 0; margin-bottom: 0"></td>
	</tr>
	<tr>
		<td width="100%" align="middle" nowrap>
		<p align="right" class="subtitle">(This page requires a login account in
		order to take the advantage)</p>
		<p align="right">&nbsp;</p></td>
	</tr>
	<tr class="10login">
		<td width="100%" align="middle">
		<table width="60%" id="table3" style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 1px; BORDER-TOP-STYLE: outset; PADDING-TOP: 1px; BORDER-RIGHT-STYLE: outset; BORDER-LEFT-STYLE: outset; BORDER-COLLAPSE: collapse; BORDER-BOTTOM-STYLE: outset"
      border="2" bordercolor="#800000">
			<tr>
				<td>
			<div align="center">
<table width="95%" id="tb2" style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 1px; BORDER-TOP-STYLE: none; PADDING-TOP: 1px; BORDER-RIGHT-STYLE: none; BORDER-LEFT-STYLE: none; BORDER-BOTTOM-STYLE: none" onkeypress="actKey()">
	<tr class="td10login">
		<td width="25%" align="middle" bordercolorlight="#0000ff" bordercolordark="#6666ff" nowrap>
		<p align="right" style="MARGIN-TOP: 3px; MARGIN-BOTTOM: 3px">Login code/name :</p></td>
		<td width="25%" align="middle" bordercolorlight="#0000ff" bordercolordark="#6666ff">
		<p align="left" style="MARGIN-TOP: 3px; MARGIN-BOTTOM: 3px"><INPUT type=text name="lname" title="Your Account Name Here!" class=10pt onclick="document.f1.pwd.value='';"></p></td>
		</tr>
	<tr class="td10login">
		<td width="25%" align="middle" bordercolor="#0000ff" bordercolorlight="#0000ff" bordercolordark="#6666ff" nowrap>
		<p align="right" style="MARGIN-TOP: 3px; MARGIN-BOTTOM: 3px">Certified Password :</p></td>
		<td width="25%" align="middle" bordercolor="#0000ff" bordercolordark="#6666ff">
		<p align="left" style="MARGIN-TOP: 3px; MARGIN-BOTTOM: 3px" class=10pt><INPUT type=password name="pwd" size="32" style="float: left" title="Your Certified Password Here!"></font></p></td>
	</tr>
</table>
		<p align="center" style="margin-top: 3px; margin-bottom: 3px" class="td10pt">
		<font color="#FA8072"><span style="letter-spacing: 1px; font-weight: 700">
		<%if Session("exp")="1" then
			Response.Write("Your session has expired... ")
		elseif Session("lout")="1" then
			Response.Write("You have been logout successfully... ")
		end if
		Session.Abandon()
		%>
		</span></font>
					</div>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="middle">
		<p style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px"
     >&nbsp;</p>
		<p style="MARGIN-TOP: 0px" nowrap><font face="Tahoma" color="#ff00ff" size="2" nowrap>
		<img border="0" id="keyin" src="images/buttonC3.gif" height="25" width="125" alt="Key In" onclick="login()" onmouseover="FP_swapImg(1,0,/*id*/'keyin',/*url*/'images/buttonC4.gif')" onmouseout="FP_swapImg(0,0,/*id*/'keyin',/*url*/'images/buttonC3.gif')" onmousedown="FP_swapImg(1,0,/*id*/'keyin',/*url*/'images/buttonC5.gif')" onmouseup="FP_swapImg(0,0,/*id*/'keyin',/*url*/'images/buttonC4.gif')" fp-style="fp-btn: Embossed Capsule 9; fp-font: Verdana; fp-font-style: Bold; fp-font-size: 11; fp-font-color-normal: #DFF7FD; fp-font-color-hover: #FFFF00; fp-font-color-press: #800000; fp-transparent: 1" fp-title="Key In">
		<img border="0" id="clear" src="images/buttonE2.gif" height="25" width="125" alt="Reset" onclick="cls()" onmouseover="FP_swapImg(1,0,/*id*/'clear',/*url*/'images/buttonE3.gif')" onmouseout="FP_swapImg(0,0,/*id*/'clear',/*url*/'images/buttonE2.gif')" onmousedown="FP_swapImg(1,0,/*id*/'clear',/*url*/'images/buttonE4.gif')" onmouseup="FP_swapImg(0,0,/*id*/'clear',/*url*/'images/buttonE3.gif')" fp-style="fp-btn: Embossed Capsule 9; fp-font: Verdana; fp-font-style: Bold; fp-font-size: 11; fp-font-color-normal: #DFF7FD; fp-font-color-hover: #FFFF00; fp-font-color-press: #800000; fp-transparent: 1" fp-title="Reset"></font></p></td>
	</tr>
	<tr>
		<td width="100%" align="middle">&nbsp;</td>
	</tr>
	</table>
<HR size=3 color=maroon>
&nbsp;</form>
</body>
</HTML>