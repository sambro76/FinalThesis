<%@ Language=VBScript%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="javascript" src="../rem/images/button.js"></script>
<script language="javascript" src="../admin/images/adfunc.js"></script>
<title>SBC Bank Of Cambodia...</title>
<style type=text/css>
@import url("images/admin.css");
BODY {SCROLLBAR-BASE-COLOR: #DFF7FD}
</style>

</head>
<body style="TEXT-ALIGN: center" bgcolor="#dff7fd" onload="FP_preloadImgs(/*url*/'images/log02.gif', /*url*/'images/log03.gif', /*url*/'images/reset02.gif', /*url*/'images/reset03.gif'); window.name='main'; window.scroll(0,40); document.f1.lname.focus()">
<form method=post name="f1" target="_self">
  <input type=hidden name="tmpid">
  <table border="0" width="100%" id="tb">
    <tr>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td class=10pt nowrap> <p align="center" style="MARGIN-BOTTOM: 0px" class="headkh">
		kar<span style="TEXT-DECORATION: none">sMercelITMrg;ep&THORN;rR)ak;</span></p>
        <p style="MARGIN-TOP: 0px" align="center" class="headeng"><b>ONLINE REMITTANCE
          FORM APPROVAL</b></p>
        <p align="right" style="margin-top: 0" class="td10pt"><b>SBC Bank Co.,
          LTD of Cambodia</b>
        <HR size=2 color=maroon style="margin-top: 0; margin-bottom: 0"></td>
    </tr>
    <tr>
      <td width="100%" align="middle" nowrap> <p align="right" class="subtitle">(This
          page requires a login account in order to take the advantage)</p>
        <p align="right">&nbsp;</p></td>
    </tr>
    <tr class="10login">
      <td width="100%" align="middle"> <table width="43%" id="table3" style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 1px; BORDER-TOP-STYLE: outset; PADDING-TOP: 1px; BORDER-RIGHT-STYLE: outset; BORDER-LEFT-STYLE: outset; BORDER-COLLAPSE: collapse; BORDER-BOTTOM-STYLE: outset"
      border="2" bordercolor="#800000">
          <tr>
            <td> <div align="center">
                <table width="65%" id="tb2" class="td10login" style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 1px; BORDER-TOP-STYLE: none; PADDING-TOP: 1px; BORDER-RIGHT-STYLE: none; BORDER-LEFT-STYLE: none; BORDER-BOTTOM-STYLE: none" onkeypress="actKey()">
                  <tr>
                    <td width="25%" align="middle" nowrap>
                      <p align="right" style="MARGIN-TOP: 3px; MARGIN-BOTTOM: 3px">Login
                        account :</p></td>
                    <td width="65%" align="middle">
                      <p align="left" style="MARGIN-TOP: 3px; MARGIN-BOTTOM: 3px">
                        <INPUT type=text name="lname" title="Your Account Name Here!" class=10pt onclick="document.f1.pwd.value='';">
                      </p></td>
                  </tr>
                  <tr>
                    <td width="25%" align="middle" bordercolor="#0000ff" bordercolorlight="#0000ff" bordercolordark="#6666ff" nowrap>
                      <p align="right" style="MARGIN-TOP: 3px; MARGIN-BOTTOM: 3px">Password
                        :</p></td>
                    <td width="67%" align="middle" bordercolor="#0000ff" bordercolorlight="#0000ff" bordercolordark="#6666ff">
                      <p align="left" style="MARGIN-TOP: 3px; MARGIN-BOTTOM: 3px" class=10pt>
                        <INPUT type=password name="pwd" size="32" style="float: left" title="Your Certified Password Here!"></font>
                      </p></td>
                  </tr>
                </table>
               	<p align="center" style="margin-top: 3px; margin-bottom: 3px" class="td10pt">
                	<font color="#FA8072"><span style="letter-spacing: 1px; font-weight: 700">
	                <%if Session("exp")="1" then
						Response.Write("Your session has expired... ")
	                elseif Session("rpt")="0" then
                		Response.Write("You have been logout....")
                	elseif Session("rpt")="1" then
                		Response.Write("Form(s) report has been made...")
	                end if
	                Session.Abandon()%>
	                </span></font>
              </div></td>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td width="100%" align="middle"> <p style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px">&nbsp;</p>
        <p style="MARGIN-TOP: 0px" nowrap><img border="0" id="keyin" src="images/log01.gif" height="22" width="100" alt="Log In" onclick="login()" onmouseover="FP_swapImg(1,0,/*id*/'keyin',/*url*/'images/log02.gif')" onmouseout="FP_swapImg(0,0,/*id*/'keyin',/*url*/'images/log01.gif')" onmousedown="FP_swapImg(1,0,/*id*/'keyin',/*url*/'images/log03.gif')" onmouseup="FP_swapImg(0,0,/*id*/'keyin',/*url*/'images/log02.gif')" fp-style="fp-btn: Embossed Capsule 9; fp-font: Verdana; fp-font-style: Bold; fp-font-color-normal: #DFF7FD; fp-font-color-hover: #FFFF00; fp-font-color-press: #800000; fp-transparent: 1" fp-title="Log In">
          <img border="0" id="clear" src="images/reset01.gif" height="22" width="100" alt="Reset" onclick="cls()" onmouseover="FP_swapImg(1,0,/*id*/'clear',/*url*/'images/reset02.gif')" onmouseout="FP_swapImg(0,0,/*id*/'clear',/*url*/'images/reset01.gif')" onmousedown="FP_swapImg(1,0,/*id*/'clear',/*url*/'images/reset03.gif')" onmouseup="FP_swapImg(0,0,/*id*/'clear',/*url*/'images/reset02.gif')" fp-style="fp-btn: Embossed Capsule 9; fp-font: Verdana; fp-font-style: Bold; fp-font-color-normal: #DFF7FD; fp-font-color-hover: #FFFF00; fp-font-color-press: #800000; fp-transparent: 1; fp-proportional: 0" fp-title="Reset"></p></td>
    </tr>
  </table>
  <p style="margin-top: 0px" nowrap>&nbsp;</p>
  <HR size=3 color=maroon>
  &nbsp;
</form>
</body>
</HTML>
<%Session.Abandon()%>