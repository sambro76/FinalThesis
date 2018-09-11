<%@ Language=VBScript%>
<%Session("admin")=false
Session("user")=false%>
<html><head>
<title>Welcome to Singapore Banking Corporation</title>

<STYLE TYPE="text/css">
   <!-- /* $WEFT -- Created by: Kounthea Im (kountheaim@mobitel.com.kh) on 10/31/2004 -- */
a:link {
	background: none;
	text-decoration: none;
	color: #000099;
}
a:visited {
	background: none;
	text-decoration: none;
	color: #000099;
}
a:hover {
	background: none;
	text-decoration: underline;
	color: #ff0000;
}
INPUT.btn1 {
	PADDING-RIGHT: 0px;
	PADDING-LEFT: 0px;
	FONT-WEIGHT: bold;
	FONT-SIZE: 60%;
	BORDER-LEFT-COLOR: #b7cfeb;
	BACKGROUND: #366496;
	BORDER-BOTTOM-COLOR: #003366;
	PADDING-BOTTOM: 0px;
	WIDTH: 100%;
	COLOR: #ffffff;
	BORDER-TOP-COLOR: #cbe3ff;
	PADDING-TOP: 0px;
	FONT-FAMILY: Verdana, Geneva, Arial, Helvetica, sans-serif;
	BORDER-RIGHT-COLOR: #003366
}
-->



</STYLE>

<STYLE TYPE="text/css">
<!-- /* $WEFT -- Created by: Samnang Chay (csamnang@everyday.com.kh) on 10/20/2004 -- */
  @font-face {
    font-family: Kounthea R1;
    font-style:  normal;
    font-weight: 700;
    src: url(../KOUNTHE10.eot);
  }
  @font-face {
    font-family: Tahoma;
    font-style:  normal;
    font-weight: 700;
    src: url(../TAHOMA13.eot);
  }

-->
</STYLE>

</head>
<body style="TEXT-ALIGN: center" bgcolor="#dff7fd" onload="window.name='main'">
<div align="center">
  <table width="750" border="0" cellpadding="0" id="tb">
    <tr>
      <td colspan="3">&nbsp;</td>
    </tr>
    <tr align="center" valign="bottom">
      <td height="54" colspan="3" nowrap class=10pt>
        <p align="center" class="headkh"><font color="#000099" size="7" face="Kounthea R1">FnaKarsig&eth;burIsUmsVaKmn_</font><font color="#6666FF" size="7" face="Kounthea R1, Kounthea S1"><br>
          <font color="#000099" size="4" face="Verdana, Arial, Helvetica, sans-serif"><b>Welcome
          to Singapore Banking Corporation</b></font></font></p>
        </td>
    </tr>
    <tr>
      <td width="109" align="center" valign="middle" nowrap>&nbsp;</td>
      <td width="15" align="middle" nowrap>&nbsp;</td>
      <td width="626" align="middle" nowrap>&nbsp;</td>
    </tr>
  </table>
  <br>
</div>
<hr width="750" color="#6699CC">
<div align="center">&nbsp; </div>
<table width="750" border="0" cellspacing="2" cellpadding="0">
  <tr>
    <td width="109" rowspan="2" align="center" valign="top"><img src="manandpc1.jpg" width="109" height="85"></td>
    <td width="19">&nbsp;</td>
    <td colspan="2" valign="top"> <p align="justify"><font size="1" face="verdana"><strong>Singapore
        Banking Corporation (SBC) is very proud to bring you our new online remittance
        system. This new system a very valuable tools for you to be success in
        your businesses. </strong></font><strong><font color="#000000" size="1" face="verdana">To
        access your account, please click on the <a href="rem/sbc_pwd.asp">Customer
        Login Button</a>. The Admin Login is only for SBC officer. We hope that
        this new system will take you to the best experience of your success in
        the digital world. Thank you very much for banking with us.</font></p>
	<p>&nbsp;</p></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td width="231" align="right" valign="top"> <table width="120" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><form action="rem/sbc_pwd.asp" method="post" name="form1" class="subtitle">
              <strong><font size="3" face="verdana"><a href="rem/sbc_pwd.asp">
              </a></font></strong>
              <table width="120" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="12"><strong><font size="3" face="verdana">
					<a href="rem/sbc_pwd.asp">
                    <input name="Submit2" type="submit" class="btn1" value="Customer Login"></a></font></strong></td>
                </tr>
              </table>
            </form></td>
        </tr>
      </table>
      <div align="right"></div></td>
    <td width="381" valign="top"><table width="120" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><form name="form2" method="post" action="admin/adlogin.asp">
              <font face="verdana"><strong><a href="admin/adlogin.asp"> </a></strong></font>
              <table width="120" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="12"><font face="verdana"><strong><a href="admin/adlogin.asp">
                    <input name="Submit" type="submit" class="btn1" value="Admin Login"></a></strong></font></td>
                </tr>
              </table>
            </form></td>
        </tr>
      </table></td>
  </tr>
</table>
<hr width="750" color="#6699CC">
</body>
</HTML>