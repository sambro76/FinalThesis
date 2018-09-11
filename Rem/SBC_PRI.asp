<%@ Language=VBScript%>
<%if Session("user")<>"sbc" then
	Session("exp")="1"
	Response.Redirect("sbc_pwd.asp")
else
Application.Lock()
''''''''
'Get QueryString 'lname' or 'id' to admit the page requested
''''''''
lname=LCase(Request.QueryString("lname"))
id=Request.QueryString("id")

if lname="" or id="" then
	Response.Write "<title>Message From SBC Bank of Cambodia...</title>"
	Response.Write "<body onload=window.open('errpage.htm')></body>"
else
	set conn=Server.CreateObject("ADODB.connection")
	DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("RemDB.MDB")
	conn.open DSN
	set rs=Server.CreateObject("ADODB.recordset")
	sql="SELECT AccID FROM Keyin WHERE AccID='" & lname & "' AND id='" & id & "'"
	rs.Open sql, conn
	if rs.eof or rs.bof then
		rs.close()
		conn.close()
		Response.Write "<title>Message From SBC Bank of Cambodia...</title>"
		Response.Write "<body onload=window.open('errpage.htm')></body>"
	else
		init=Request.QueryString("init")
		rs.close()
%>
<HTML>
<HEAD>
<TITLE>Your Previous Remittance Information</TITLE>
<STYLE Type=text/css>
@import url("images/sbc.css");
</STYLE>
<STYLE type=text/css>BODY {SCROLLBAR-BASE-COLOR: #E7FBFE}</STYLE>

<SCRIPT language="JavaScript" src="images/valid.js"></SCRIPT>
<script language="JavaScript" src="images/button.js"></script>
<script language="JavaScript" src="images/page.js"></script>
<script language="JavaScript" src="images/mark.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

</HEAD>
<%
mark=Request.QueryString("mark")
hs=Request.QueryString("hs")
''''''''
'Create a query to get the recordset from database, 'RecData' Table
'Depending on QueryString 'hs'- hide/show deleted forms...
''''''''
sql="SELECT Dated, UCase(BFName & ' ' & BLName) & ', ' & BCity & ', ' & BCountry AS Ex1, REF_No, [Currency], Amount, Rate, Amount/Rate AS Ex2, Status FROM RecData WHERE AccID='" + lname + "' "
if hs="hide" then
	sql=sql & "AND Status IN ('0','1','3','4') ORDER BY "
else
	sql=sql & "ORDER BY "
end if
	''''''
	'Order codes : 0 for <<DATE>>, 1 for <<TO PAYEE>>, and 2 for <<Total In US>>
	''''''
	ordby=Request.QueryString("ordby")
	if ordby=1 then
		ord="UCase(BFName & ' ' & BLName) & ', ' & BCity & ', ' & BCountry"
	elseif ordby=2 then
		ord="Amount/Rate"
	else
		ordby=0
		ord="Dated"
	end if
set rs=Server.CreateObject("ADODB.recordset")
rs.Open sql + ord + " DESC;", conn
Randomize()
rand=right(Rnd(),len(Rnd())-1)
%>
<body bgcolor= #666699  link="#e7fbfe" alink="#ffff00" topmargin="0" onload="FP_preloadImgs(/*url*/'images/button4.gif', /*url*/'images/button5.gif', /*url*/'images/button12.gif', /*url*/'images/button13.gif'); window.scroll(0,10); window.name='main'">
<form method=get name=f1 style="FONT-FAMILY: Tahoma" enctype="multipart/form-data">

<!--condition if no record to display-->
<%if rs.BOF or rs.EOF then%>
<table border=2 width=640 id=tno cellspacing=0 bgcolor="#e7fbfe" cellpadding=0 align="center">
	<tr>
		<td colspan="3">
		<p align="center" style="MARGIN: 0px" class="headkh">
		RtYtBinitüskmµPaBcas;²</p>
		<p style="MARGIN-TOP: 0px; margin-bottom:0" align="center" class="headeng"><b>
		CHECKING YOUR OLD REMITTANCE ACTIVITIES</b></p>
		<p align="right" style="MARGIN-TOP: 6px; margin-bottom:0" class="td10pt"><b>
		SBC Bank Co., LTD of Cambodia</b></p>
		<p align="right" class="td10pt" style="margin-top: 6px; margin-bottom: 0">
		&nbsp;Welcome <b><%=Session("AppName")%>... </b></p>
		<HR size=2 color=maroon style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px">
		</td>
	</tr>
	<tr>
		<td align="middle" bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-right-style:none; border-right-width:medium" nowrap>
		<IMG border="0" id="img18" src="images/button14.gif" height="25" width="100" alt="Check Form" onclick="window.open('sbc_pri.asp?lname=<%=lname%>&id=<%=id%>&.<%=rnd()%>','main')" fp-style="fp-btn: Glass Tab 1; fp-font: Verdana; fp-font-style: Bold; fp-font-color-normal: #0000FF; fp-font-color-hover: #FFCC33; fp-font-color-press: #FF0000; fp-transparent: 1; fp-proportional: 0" fp-title="Check Form" align="left" onmouseover="FP_swapImg(1,0,/*id*/'img18',/*url*/'images/button18.gif'); status='Recheck the Forms... '; return true" onmouseout="FP_swapImg(0,0,/*id*/'img18',/*url*/'images/button14.gif'); status=''" onmousedown="FP_swapImg(1,0,/*id*/'img18',/*url*/'images/button19.gif')" onmouseup="FP_swapImg(0,0,/*id*/'img18',/*url*/'images/button18.gif')">&nbsp;&nbsp;<img border="0" id="img21" onclick="if(<%=Session("exp")%>=="1") {window.open('new.asp?lname=<%=lname%>&id=<%=id%>','main')} else {window.open('new.asp?lname=<%=lname%>&id=<%=id%>','new')}" src="images/button3.gif" height="25" width="100" alt="New" fp-style="fp-btn: Glass Tab 1; fp-font: Verdana; fp-font-style: Bold; fp-font-color-normal: #0000FF; fp-font-color-hover: #FFCC33; fp-font-color-press: #FF0000; fp-transparent: 1; fp-proportional: 0; fp-orig: 0" fp-title="New" onmouseover="FP_swapImg(1,0,/*id*/'img21',/*url*/'images/button4.gif')" onmouseout="FP_swapImg(0,0,/*id*/'img21',/*url*/'images/button3.gif')" onmousedown="FP_swapImg(1,0,/*id*/'img21',/*url*/'images/button5.gif')" onmouseup="FP_swapImg(0,0,/*id*/'img21',/*url*/'images/button4.gif')">
		<img border="0" id="img22" src="images/button11.gif" height="25" width="100" alt="Log  out" onclick="window.open('logout.asp?lname=<%=lname%>&id=<%=id%>','main')" fp-style="fp-btn: Glass Tab 1; fp-font: Staccato222 BT; fp-font-size: 18; fp-font-color-normal: #0000FF; fp-font-color-hover: #FFCC33; fp-font-color-press: #FF0000; fp-transparent: 1; fp-proportional: 0; fp-orig: 0" fp-title="Log  out" onmouseover="FP_swapImg(1,0,/*id*/'img22',/*url*/'images/button12.gif')" onmouseout="FP_swapImg(0,0,/*id*/'img22',/*url*/'images/button11.gif')" onmousedown="FP_swapImg(1,0,/*id*/'img22',/*url*/'images/button13.gif')" onmouseup="FP_swapImg(0,0,/*id*/'img22',/*url*/'images/button12.gif')"><IMG onmouseup="FP_swapImg(0,0,/*id*/'img16',/*url*/'images/button4A.gif')" onmousedown="FP_swapImg(1,0,/*id*/'img16',/*url*/'images/button4B.gif')" id=img16 onmouseover="FP_swapImg(1,0,/*id*/'img16',/*url*/'images/button4A.gif')" onmouseout="FP_swapImg(0,0,/*id*/'img16',/*url*/'images/button49.gif')" height=25 alt =" HELP!" src="images/button49.gif" width =125 border=0  fp-title=" HELP!" fp-style="fp-btn: Jewel 1; fp-font-style: Bold; fp-font-size: 16; fp-font-color-normal: #800000; fp-font-color-hover: #FFFF00; fp-font-color-press: #FF0000; fp-justify-horiz: 0; fp-transparent: 1; fp-orig: 0">
		</td>
		<%if hs="hide" then%>
			<td align="center" bgcolor="#037ead" class="thw10pt" nowrap>
			<p style="margin-top: 6px; margin-bottom: 6">&nbsp;&nbsp;&nbsp;<span style="LETTER-SPACING: 1px"><A onmouseover="status='Show all deleted forms... '; return true;" onmouseout="status=''" href="sbc_pri.asp?lname=<%=lname%>&id=<%=id%>&.<%=rand%>"><font color="#FFFFFF">Show Deleted Forms</font></A></span>
			&nbsp;&nbsp;</p></td>
		<%end if%>
	<tr bgcolor="#037ead" class="thw10pt" nowrap>
		<%if hs="hide" then%>
			<td colspan="3" align="center">There is no previous form, or your old forms have been deleted...</td>
		<%else
			Response.Write("<td colspan='3' align='center'>There is no previous form, or your old forms have been purged...</td>")
		end if%>
	</tr>
<%else
	const numRec=4
	if init=0 then
		do until rs.EOF
			count=count+1
			rs.MoveNext
		loop
		pageNo = 1
		if (count/numRec)=Int(count/numRec) then
			pages = Int(count/numRec)
		else
			pages = Int(count/numRec) + 1
		end if
	else
		pageNo=Request.QueryString("pageNo")
		pages=Request.QueryString("pages")
		count=Request.QueryString("count")
		firstp=Request.QueryString("firstp")
		prep=Request.QueryString("prep")
		nextp=Request.QueryString("nextp")
		lastp=Request.QueryString("lastp")
		selectp=Request.QueryString("selectp")
	end if

	if init=0 then
		if pages=1 then
			row=count
		else
			row=numRec
		end if
		rs.MoveFirst()
	else
		if firstp=1 then
			row=numRec
			rs.MoveFirst()
		elseif prep=1 then
			row=numRec
			rs.MoveFirst()
			for k=1 to (pageNo-1)*numRec
				rs.MoveNext()
			next
		elseif nextp=1 then
			if pageNo=pages then
				row=count-(pageNo-1)*numRec
			else
				row=numRec
			end if
			rs.MoveFirst()
			for k=1 to (pageNo-1)*numRec
				rs.MoveNext()
			next
		elseif lastp=1 then
			row=count-(pageNo-1)*numRec
			rs.MoveFirst()
			for k=1 to (pageNo-1)*numRec
				rs.MoveNext()
			next
		elseif selectp=1 then
			rs.MoveFirst()
			if pages=1 then
				row=count
			elseif pageNo=pages then
				row=count-(pageNo-1)*numRec
				for k=1 to (pageNo-1)*numRec
					rs.MoveNext()
				next
			else
				row=numRec
				for k=1 to (pageNo-1)*numRec
					rs.MoveNext()
				next
			end if
		end if
	end if
	%>
<input type=hidden name=h1>
<table border="2" width=800 id="tb1" cellspacing="0" bgcolor="#e7fbfe" cellpadding="0" align="center">
	<tr>
		<td colspan="3">
		<p align="center" style="MARGIN: 0px" class="headkh">
		RtYtBinitüskmµPaBcas;²</p>
		<p style="MARGIN-TOP: 0px; margin-bottom:0" align="center" class="headeng"><b>
		CHECKING YOUR OLD REMITTANCE ACTIVITIES</b></p>
		<p align="right" class="td10pt" style="margin-top: 6px; margin-bottom: 0"><b>
		SBC Bank Co., LTD of Cambodia</b></p>
		<p align="right" class="td10pt" style="margin-top: 6px; margin-bottom: 0">
		&nbsp;Welcome <b><%=Session("AppName")%>... </b></p>
		<HR size=2 color=maroon style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px">
		</td>
	<tr>
		<td width=240 align="middle" bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-right-style:none; border-right-width:medium" nowrap>
		<img border="0" id="img18" src="images/button14.gif" height="25" width="100" onclick="window.open('sbc_pri.asp?lname=<%=lname%>&id=<%=id%>&ordby=<%=ordby%>&mark=<%=Request.QueryString("mark")%>&hs=<%=hs%>&<%=rand%>','main','').opener=self;" alt="Check Form" fp-style="fp-btn: Glass Tab 1; fp-font: Verdana; fp-font-style: Bold; fp-font-color-normal: #0000FF; fp-font-color-hover: #FFCC33; fp-font-color-press: #FF0000; fp-transparent: 1; fp-proportional: 0" fp-title="Check Form" align="left" onmouseover="FP_swapImg(1,0,/*id*/'img18',/*url*/'images/button18.gif'); status='Recheck the forms... '" onmouseout="FP_swapImg(0,0,/*id*/'img18',/*url*/'images/button14.gif'); status=''" onmousedown="FP_swapImg(1,0,/*id*/'img18',/*url*/'images/button19.gif')" onmouseup="FP_swapImg(0,0,/*id*/'img18',/*url*/'images/button18.gif')"></td>
		<td width=240 align="center" bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-right-style:none; border-right-width:medium; border-left-style:none; border-left-width:medium" nowrap>
		<p align="right">
		<img border="0" id="img19" onclick="window.open('new.asp?lname=<%=lname%>&id=<%=id%>','new')" src="images/button3.gif" height="25" width="100" alt="New" fp-style="fp-btn: Glass Tab 1; fp-font: Verdana; fp-font-style: Bold; fp-font-color-normal: #0000FF; fp-font-color-hover: #FFCC33; fp-font-color-press: #FF0000; fp-transparent: 1; fp-proportional: 0" fp-title="New" onmouseover="FP_swapImg(1,0,/*id*/'img19',/*url*/'images/button4.gif')" onmouseout="FP_swapImg(0,0,/*id*/'img19',/*url*/'images/button3.gif')" onmousedown="FP_swapImg(1,0,/*id*/'img19',/*url*/'images/button5.gif')" onmouseup="FP_swapImg(0,0,/*id*/'img19',/*url*/'images/button4.gif')"></p>
		</td>
		<td width=240 align="middle" bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-right-style:none; border-right-width:medium; border-left-style:none; border-left-width:medium" nowrap>
		<p style="MARGIN-TOP: 6px; MARGIN-BOTTOM: 6px" align="left">
		&nbsp;<img border="0" id="img20" src="images/button11.gif" height="25" width="100" alt="Log  out" onclick="window.open('logout.asp?lname=<%=lname%>&id=<%=id%>','main')" fp-style="fp-btn: Glass Tab 1; fp-font: Staccato222 BT; fp-font-size: 18; fp-font-color-normal: #0000FF; fp-font-color-hover: #FFCC33; fp-font-color-press: #FF0000; fp-transparent: 1; fp-proportional: 0" fp-title="Log  out" onmouseover="FP_swapImg(1,0,/*id*/'img20',/*url*/'images/button12.gif')" onmouseout="FP_swapImg(0,0,/*id*/'img20',/*url*/'images/button11.gif')" onmousedown="FP_swapImg(1,0,/*id*/'img20',/*url*/'images/button13.gif')" onmouseup="FP_swapImg(0,0,/*id*/'img20',/*url*/'images/button12.gif')">
		<IMG onmouseup="FP_swapImg(0,0,/*id*/'img16',/*url*/'images/button4A.gif')" onmousedown="FP_swapImg(1,0,/*id*/'img16',/*url*/'images/button4B.gif')" id=img16 onmouseover="FP_swapImg(1,0,/*id*/'img16',/*url*/'images/button4A.gif')" onmouseout="FP_swapImg(0,0,/*id*/'img16',/*url*/'images/button49.gif')" height=25 alt =" HELP!" src="images/button49.gif" width =125 border=0  fp-title=" HELP!" fp-style="fp-btn: Jewel 1; fp-font-style: Bold; fp-font-size: 16; fp-font-color-normal: #800000; fp-font-color-hover: #FFFF00; fp-font-color-press: #FF0000; fp-justify-horiz: 0; fp-transparent: 1; fp-orig: 0"></p></td>
	</tr>
<%'Cut from here%>
		<tr bgcolor="#037ead" class="thw10pt" nowrap>
			<th width="33%" style="border-right-style: none; border-right-width: medium">
			<p align="left" style="MARGIN-TOP: 3px; MARGIN-BOTTOM: 3px">
			<span style="LETTER-SPACING: 1px">&nbsp;Activities List</span></p>
			</th>
		<th width="33%" style="border-left-style: none; border-left-width: medium; border-right-style: none; border-right-width: medium">
			<p align="center"><span style="LETTER-SPACING: 1px">
			<%if pages>1 then
				Response.Write("Page " & pageNo & " of " & pages)
			else
				Response.Write("&nbsp;")
			end if%>
			</span></p>
		</th>
		<th style="border-left-style: none" class="thw10pt" bgcolor="#037ead">
			<p align="right"><span style="LETTER-SPACING: 1px">
			<%if pages>1 then
				if pageNo=1 then
					Response.Write("1 to " & numRec*pageNo)
				elseif pageNo=pages then
					Response.Write(numRec*(pageNo-1)+1 & " to " & count)
				else
					Response.Write(numRec*(pageNo-1)+1 & " to " & numRec*(pageNo-1)+numRec)
				end if
				Response.Write(" of " & count & " Forms")
			else
				if count>1 then
					Response.Write("1 to " & count & " Forms")
				else
					Response.Write("1 Form Only")
				end if
			end if%>
			</span></p>
		</th>
	</tr>
		<tr bgcolor="#bdc3ce" nowrap>
			<td width=240 nowrap style="BORDER-RIGHT: medium none; BORDER-BOTTOM: medium none">
			<p align="left" style="MARGIN-TOP: 3px; MARGIN-BOTTOM: 3px">
			<span style="LETTER-SPACING: 1px">&nbsp;<select size="1" name="Select" style="VERTICAL-ALIGN: baseline; FONT-FAMILY: Tahoma; LETTER-SPACING: 1px" onclick="check(0)">
			<option selected value="0">Select:</option>
			<option value="1">All</option>
			<option value="2">None</option></select>
			<select size="1" name="MrkAs" style="VERTICAL-ALIGN: baseline; FONT-FAMILY: Tahoma; LETTER-SPACING: 1px" onchange="if(document.f1.MrkAs.selectedIndex!=0) {mark(<%=row%>, '<%=lname%>', '<%=id%>',document.f1.MrkAs.value, <%=pages%>, <%=count%>, <%=pageNo%>, <%=ordby%>,'<%=hs%>');}">
			<option selected>Mark As:</option>
			<option value="r">Read</option>
			<option value="u">Unread</option>
			<option value="d1">Deleted</option>
			<option value="d0">Undeleted</option>
			</select></span></p></td>
		<td width=240 align="middle" rowspan="2" nowrap style="BORDER-RIGHT: medium none; BORDER-LEFT: medium none">
		<%if pages>1 then%>
			<a title="First Page">
				<%if pageNo=1 then
				 	Response.Write("<IMG name='first' src='images/firstW.gif' height=16 width=16 align=middle>")
				else%>
				 	<IMG name="first" src='images/first.gif' onclick="page('<%=lname%>','<%=id%>','firstp',<%=pages%>,<%=count%>,1,<%=ordby%>,'<%=rand%>','<%=hs%>')" onmouseover="this.src='images/firstW.gif'" onmouseout="this.src='images/first.gif'" height=16 width=16 align=middle>
				<%end if%>
				</a><a title="Previous Page">
				<%if pageNo=1 then
			 		Response.Write("<IMG name='previous' src='images/previouW.gif' height=16 width=16 align=middle>")
				else%>
			 		<IMG name="previous" src='images/previous.gif' onclick="page('<%=lname%>','<%=id%>','prep',<%=pages%>,<%=count%>,<%=pageNo-1%>,<%=ordby%>,'<%=rand%>','<%=hs%>')" onmouseover="this.src='images/previouW.gif'" onmouseout="this.src='images/previous.gif'" height=16 width=16 align=middle>
				<%end if%>
				</a>&nbsp;<input type=text name="tbox" size="4" value=<%=pageNo%> style="TEXT-ALIGN: center" tabindex="0" onkeyup="var a=parseInt(event.keyCode); if((a==13)&&(document.f1.tbox.value!=<%=pageNo%>)) {page('<%=lname%>','<%=id%>','selectp',<%=pages%>,<%=count%>,document.f1.tbox.value,<%=ordby%>,'<%=Rnd()%>','<%=hs%>');}">&nbsp;
				<a title="Next Page">
				<%if pageNo=pages then
				 	Response.Write("<IMG name='next' src='images/nextW.gif' height=15 width=15 align=middle>")
				else%>
				 	<IMG name="next" src='images/next.gif' onclick="page('<%=lname%>','<%=id%>','nextp',<%=pages%>,<%=count%>,<%=pageNo+1%>,<%=ordby%>,'<%=rnd%>','<%=hs%>')" onmouseover="this.src='images/nextW.gif'" onmouseout="this.src='images/next.gif'" height=15 width=15 align=middle>
				<%end if%>
				</a><a title="Last Page">
				<%if pageNo=pages then
			 		Response.Write("<IMG name='lastp' src='images/lastW.gif' height=16 width=16 align=middle>")
				else%>
				 	<IMG name="lastp" src='images/last.gif' onclick="page('<%=lname%>','<%=id%>','lastp',<%=pages%>,<%=count%>,<%=pages%>,<%=ordby%>,'<%=rnd%>','<%=hs%>')" onmouseover="this.src='images/lastW.gif'" onmouseout="this.src='images/last.gif'" height=16 width=16 align=middle>
				<%end if%>
			</a></td>
		<%else%>
			<input type=hidden name="hide">
		<%end if%>
		</td>
		<td width=240 align="middle" nowrap style="BORDER-LEFT: medium none; BORDER-BOTTOM: medium none">&nbsp;</td>
	</tr>
	<tr bgcolor="#bdc3ce" class="tdb10pt" nowrap>
		<td width=240 align="middle" nowrap style="BORDER-RIGHT: medium none; BORDER-TOP: medium none">
		<p align="left"><span style="LETTER-SPACING: 1px">&nbsp;<A onmouseover="status='Delete selected forms... '; return true;" onmouseout="status=''" href="javascript:mark(<%=row%>,'<%=lname%>','<%=id%>','d1',<%=pages%>,<%=count%>,<%=pageNo%>,<%=ordby%>,'<%=hs%>');">Delete</a>
			| <A onmouseover="status='Undelete selected forms... '; return true;" onmouseout="status=''" href="javascript:mark(<%=row%>,'<%=lname%>','<%=id%>','d0',<%=pages%>,<%=count%>,<%=pageNo%>,<%=ordby%>,'<%=hs%>');">
		Undelete</a></span></p>
		</td>
		<td width=240 align="middle" nowrap style="BORDER-TOP: medium none; BORDER-LEFT: medium none">
		<p align="right"><span style="LETTER-SPACING: 1px">
		<%if hs="hide" then%>
			<A onmouseover="status='Show all deleted forms... '; return true;" onmouseout="status=''" href="sbc_pri.asp?lname=<%=lname%>&id=<%=id%>&ordby=<%=ordby%>&mark=<%=mark%>&hs=show&<%=rand%>">
			Show Deleted</A>
		<%else%>
			<A onmouseover="status='Hide all deleted forms... '; return true;" onmouseout="status=''" href="sbc_pri.asp?lname=<%=lname%>&id=<%=id%>&ordby=<%=ordby%>&mark=<%=mark%>&hs=hide&<%=rand%>">
			Hide Deleted</A>
		<%end if%>
			| <A onmouseover="status='Purge all deleted forms... '; return true;" onmouseout="status=''" href="purge.asp?lname=<%=lname%>&id=<%=id%>&ordby=<%=ordby%>&mark=<%=mark%>&hs=<%=hs%>&<%=rand%>">
		Purge Deleted</A></span></p>
		</td>
	</tr>
	<tr bgcolor="#bdc3ce" class="tdb10pt" nowrap>
		<td width=720 align="middle" colspan="3">
		<table bgcolor="#bdc3ce" border="1" width="101%" id="tb2" style="FONT-FAMILY: Tahoma; LETTER-SPACING: 1px" bordercolor="#e7fbfe" align=center class="td10pt" nowrap>
			<TBODY>
				<tr style="FONT-WEIGHT: bold; COLOR: #000000" bgcolor="#ffffff" align="left" class="thb10pt">
					<th width="30px">
					<input type="checkbox" name="chkAll" onclick="check(2)" value="chkAll">#</th>
			<script language="JavaScript">
function chk(){
	for(i=1;i<=<%=row%>;i++){document.f1.elements[i+4].checked = document.f1.chkAll.checked;}
}
function check(m){
	if(m==0) {
		if(document.f1.Select.selectedIndex==1) {document.f1.Select0.selectedIndex=1; document.f1.chkAll.checked=1;chk();}
		else if(document.f1.Select.selectedIndex==2) {document.f1.Select0.selectedIndex=2; document.f1.chkAll.checked=0;document.f1.reset();}
		else {document.f1.Select0.selectedIndex=0;}
		}
	else if(m==1){
		if(document.f1.Select0.selectedIndex==1) {document.f1.Select.selectedIndex=1; document.f1.chkAll.checked=1;chk();}
		else if(document.f1.Select0.selectedIndex==2) {document.f1.Select.selectedIndex=2; document.f1.chkAll.checked=0;document.f1.reset();}
		else {document.f1.Select.selectedIndex=0;}
		}
	else if(m==2){
		if(document.f1.chkAll.checked==0) {document.f1.reset();}
		else {document.f1.Select.selectedIndex=document.f1.Select0.selectedIndex=1; chk();}
		}
	}
</script>
					<th width="72px" nowrap>
					<a onmouseover="status='Click here to sort by DATE... '; return true;" onmouseout="status=''" href="javascript:if(<%=count%>!=1&&<%=ordby%>!=0) page('<%=lname%>','<%=id%>','selectp',<%=pages%>,<%=count%>,1,0,'<%=mark%>','<%=hs%>')">
					<IMG height=14 src="images/ArrUp.jpg" width=14 border=0></a>
					DATE</th>
					<th width="265px" nowrap>
					<a onmouseover="status='Click here to sort by PAYEE... '; return true;" onmouseout="status=''"  href="javascript:if(<%=count%>!=1&&<%=ordby%>!=1) page('<%=lname%>','<%=id%>','selectp',<%=pages%>,<%=count%>,1,1,'<%=mark%>','<%=hs%>')">
					<IMG height=14 src="images/ArrUp.jpg" width=14 border=0></a>
					To PAYEE (Name, Address)</th>
					<th width="83px" align="center" nowrap>Bank REF.</TH>
					<th width="131" align="center" nowrap>AMOUNT</th>
					<th width="39" align="center" nowrap>Rate</th>
					<th width="87" nowrap>
					<a onmouseover="status='Click here to sort by Amount In US$... '; return true;" onmouseout="status=''" href="javascript:if(<%=count%>!=1&&<%=ordby%>!=2) page('<%=lname%>','<%=id%>','selectp',<%=pages%>,<%=count%>,1,2,'<%=mark%>','<%=hs%>')">
					<IMG height=14 src="images/ArrUp.jpg" width=14 border=0></a>
					In US$</th>
					<th width="67" align="center">Approval</th>
				</tr>
	<%for j=1 to row
	status=rs.Fields("Status")
	if status=0 then%>
		<tr bgcolor="#DEEBFF" class="td10ptb" id="t<%=j%>" onmouseover="this.bgColor='#C7DFFE'" onmouseout="this.bgColor='#DEEBFF'">
	<%elseif status=1 then%>
		<tr bgcolor="#FFFFFF" class="td10pt" id="t<%=j%>" onmouseover="this.bgColor='#C7DFFE'" onmouseout="this.bgColor='#FFFFFF'">
	<%elseif status=2 then%>
		<tr bgcolor="#BDC3AD" class="td10pt" style="text-decoration: line-through" id="t<%=j%>" onmouseover="this.bgColor='#C7DFFE'" onmouseout="this.bgColor='#BDC3AD'">
	<%elseif status=3 then%>
		<tr bgcolor="#FFF7D9" class="td10ptb" id="t<%=j%>" onmouseover="this.bgColor='#C7DFFE'" onmouseout="this.bgColor='#FFF7D9'">
		<%napp=1
	elseif status=4 then%>
		<tr bgcolor="#FFF4F4" class="td10ptb" id="t<%=j%>">
		<%errfrm=1
	end if%>
		 <td nowrap><input type="checkbox" name="cb" value=<%=rs.fields("REF_No")%> onclick="if(this.checked==0||document.f1.chkAll.checked==1) {document.f1.chkAll.checked=0;document.f1.Select.selectedIndex=0;document.f1.Select0.selectedIndex=0;}">
			<a onmouseover="status='Click for detail...'; return true" onmouseout="status=''" href="detail.asp?lname=<%=lname%>&id=<%=id%>&ref=<%=rs.fields("REF_No")%>&<%=rand%>" target=_blank>
			<%=j+(pageNo-1)*numRec%></a></td>
		 <td nowrap><a onmouseover="status='Click for detail...'; return true" onmouseout="status=''" href="detail.asp?lname=<%=lname%>&id=<%=id%>&ref=<%=rs.fields("REF_No")%>&<%=rand%>" target=_blank>
			<%if FormatDateTime(rs.Fields("Dated"),2)=FormatDateTime(Now(),2) then%>
				<%="Today " & FormatDateTime(rs.fields("Dated"),3)%>
			<%else%>
				<%=FormatDateTime(rs.fields("Dated"),2)%>
			<%end if%>
			</a></td>
		 <td nowrap>
		 	<%if len(rs.fields("Ex1"))>30 then%><span style="font-size:8pt">
			 	<a onmouseover="status='Click for detail...'; return true" onmouseout="status=''" href="detail.asp?lname=<%=lname%>&id=<%=id%>&ref=<%=rs.fields("REF_No")%>&<%=rand%>" target=_blank>
				<%=rs.fields("Ex1")%></a></span>
			<%else %>
			 	<a onmouseover="status='Click for detail...'; return true" onmouseout="status=''" href="detail.asp?lname=<%=lname%>&id=<%=id%>&ref=<%=rs.fields("REF_No")%>&<%=rand%>" target=_blank>
			 	<%=rs.fields("Ex1")%></a>
			<%end if%></td>
		 <td align="center" nowrap><a onmouseover="status='Click for detail...'; return true;" onmouseout="status=''" href="detail.asp?lname=<%=lname%>&id=<%=id%>&ref=<%=rs.fields("REF_No")%>&<%=rand%>" target=_blank>
			<%=rs.fields("REF_No")%></a></td>
		 <td align="right" nowrap>
			<%Curr = rs.fields("Currency") + " " + FormatNumber(rs.fields("Amount"),2)
				if len(Curr)>15 then%><span style="font-size:8pt">
				<a onmouseover="status='Click for detail...'; return true;" onmouseout="status=''" href="detail.asp?lname=<%=lname%>&id=<%=id%>&ref=<%=rs.fields("REF_No")%>&<%=rand%>" target=_blank>
				<%=Curr%></a></span>
				<%else%><a onmouseover="status='Click for detail...'; return true;" onmouseout="status=''" href="detail.asp?lname=<%=lname%>&id=<%=id%>&ref=<%=rs.fields("REF_No")%>&<%=rand%>" target=_blank>
				<%=Curr%></a><%end if%></td>
		 <td align="right" nowrap><a onmouseover="status='Click for detail...'; return true;" onmouseout="status=''" href="detail.asp?lname=<%=lname%>&id=<%=id%>&ref=<%=rs.fields("REF_No")%>&<%=rand%>" target=_blank>
				<%=rs.fields("Rate")%></a></td>
		 <td align="right" nowrap><a onmouseover="status='Click for detail...'; return true;" onmouseout="status=''" href="detail.asp?lname=<%=lname%>&id=<%=id%>&ref=<%=rs.fields("REF_No")%>&<%=rand%>" target=_blank>
			<%=FormatNumber(rs.fields("Amount")/rs.fields("Rate"),2)%>
			</a></td>
		 <td align="center" nowrap><a onmouseover="status='Click for detail...'; return true;" onmouseout="status=''" href="detail.asp?lname=<%=lname%>&id=<%=id%>&ref=<%=rs.fields("REF_No")%>&<%=rand%>" target=_blank>
		 	<%if status=3 then
		 		Response.Write("Not Yet")
			elseif status=4 then
				Response.Write("Error!")
			else
				Response.Write("Yes")
			end if%>
			</a>
		 </td>
		</tr>
	<%rs.MoveNext
	next
	rs.Close()%>
		<tr bgcolor="#FFFFFF" class="td10pt">
		 <td nowrap colspan="4">
		 	<%if init=1 then
		 		Session("init")=1
		 	else
		 		Session("init")=0
		 		Session("notapp")=0
		 		Session("napp_amount")=0
		 	end if
		 	if Session("init")=0 then
				rs.Open "SELECT Amount, Rate FROM RecData WHERE Status='3' AND AccID='" & lname & "'", conn
				do until rs.EOF
					Session("notapp")=Session("notapp")+1
					Session("napp_amount")=Session("napp_amount")+rs.Fields("Amount")/rs.Fields("Rate")
					rs.MoveNext()
				loop
			end if
		 	if Session("notapp")>0 then
		 		if Session("notapp")=1 then
					Response.Write("<b>1 Form ")
				else
					Response.Write("<b>" & Session("notapp") & " Forms ")
				end if
				Response.Write("attending approval : " & FormatNumber(Session("napp_amount"),2) & " USD of total amount to be paid...</b>")
			else
				Response.Write("&nbsp;")
			end if
		conn.close()%>
		</td>
		 <td nowrap colspan="4">
		 	<b>&nbsp;Your Balance:</b>
		 	<%if Session("init")=1 then
				Response.Write("<b>" & FormatNumber(Session("Balance"),2) & " USD</b>")
		 	else
		 		DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & left(Server.MapPath("RemDB.MDB"),len(Server.MapPath("RemDB.MDB"))-13) & "\Admin\BankDB.MDB"
				conn.Open DSN
				rs.Open "SELECT Balance FROM AccTrans, AppInfos WHERE AccTrans.AccNo=AppInfos.AccNo AND AccID='" & lname & "' ORDER BY Dated DESC", conn
				Response.Write("<b>" & FormatNumber(rs.Fields("Balance"),2) & " USD</b>")
				Session("Balance")=rs.Fields("Balance")
				rs.Close()
				conn.Close()
		 	end if
		 	Session("init")=0%>
	 	</td>
		</tr>
		</table></TD>
	</TR>
	<tr bgcolor="#bdc3ce" class="tdb10pt" nowrap>
		<td width=720 align="middle" colspan="3">
		<div align="left">
		<TABLE cellPadding=0 border=0 class="tdb10pt">
			<TR><td width=80><span style="letter-spacing: 1px">&nbsp;Status:</span></td>
			<TD bgcolor="#DEEAFF" width=15 height=15></TD>
			<td width=60><span style="letter-spacing: 1px">&nbsp;Unread</span></td>
			<TD bgcolor="#FFFFFF" width=15 height=15></TD>
			<td width=60><span style="letter-spacing: 1px">&nbsp;Read</span></td>
			<TD bgcolor="#BDC3AD" width=15 height=15></TD>
			<td width=65><span style="letter-spacing: 1px">&nbsp;Deleted</span></td>
			<TD bgcolor="#FFF7D9" width=15 height=15></TD>
			<td width=145><span style="letter-spacing: 1px">
			<%if napp=1 then
				Response.Write("<b>&nbsp;Not Yet Approved</b>")
			else%>&nbsp;Not Approved Yet
			<%end if%>
			</span></td>
			<TD bgcolor="#FFF4F4" width=15 height=15></TD>
			<td width=165 nowrap><span style="letter-spacing: 1px">
			<%if errfrm=1 then
				Response.Write("<b>&nbsp;Error Ordered Form</b>")
			else%>&nbsp;Error Ordered Form
			<%end if%>
			</span></td>
			</TR>
		</TABLE>
		</div>
		</td>
	</tr>
	<tr bgcolor="#bdc3ce" class="tdb10pt" nowrap>
		<td width=240 align="middle" style="BORDER-RIGHT: medium none; BORDER-BOTTOM: medium none">
		<p align="left"><span style="LETTER-SPACING: 1px">&nbsp;<A onmouseover="status='Delete selected forms... '; return true;" onmouseout="status=''" href="javascript:mark(<%=row%>,'<%=lname%>','<%=id%>','d1',<%=pages%>,<%=count%>,<%=pageNo%>,<%=ordby%>,'<%=hs%>');">Delete</a>
			| <A onmouseover="status='Undelete selected forms... '; return true;" onmouseout="status=''" href="javascript:mark(<%=row%>,'<%=lname%>','<%=id%>','d0',<%=pages%>,<%=count%>,<%=pageNo%>,<%=ordby%>,'<%=hs%>');">
			Undelete</a></span></p></td>
		<td width=240 align="middle" style="BORDER-RIGHT: medium none; BORDER-LEFT: medium none; BORDER-BOTTOM: medium none">&nbsp;</td>
		<td width=240 align="middle" style="BORDER-LEFT: medium none; BORDER-BOTTOM: medium none">
			<p align="right"><span style="LETTER-SPACING: 1px">
			<%if hs="hide" then%>
				<A onmouseover="status='Show all deleted forms... '; return true;" onmouseout="status=''" href="sbc_pri.asp?lname=<%=lname%>&id=<%=id%>&ordby=<%=ordby%>&mark=<%=mark%>&hs=show&<%=rand%>">
				Show Deleted</A>
			<%else%>
				<A onmouseover="status='Hide all deleted forms... '; return true;" onmouseout="status=''" href="sbc_pri.asp?lname=<%=lname%>&id=<%=id%>&ordby=<%=ordby%>&mark=<%=mark%>&hs=hide&<%=rand%>">
				Hide Deleted</A>
			<%end if%>
			| <A onmouseover="status='Purge all deleted forms... '; return true;" onmouseout="status=''" href="purge.asp?lname=<%=lname%>&id=<%=id%>&ordby=<%=ordby%>&mark=<%=mark%>&hs=<%=hs%>&<%=rand%>">
			Purge Deleted</A></span></p>
		</td>
	</tr>
	<tr bgcolor="#bdc3ce" class="tdb10pt" nowrap>
		<td width=240 align="middle" style="BORDER-RIGHT: medium none; BORDER-TOP: medium none"  >
			<p align="left" style="MARGIN-TOP: 3px; MARGIN-BOTTOM: 3px">
			<span style="LETTER-SPACING: 1px">&nbsp;<select size="1" name="Select0" style="VERTICAL-ALIGN: baseline; FONT-FAMILY: Tahoma; LETTER-SPACING: 1px" onclick="check(1)">
			<option selected value="0">Select:</option>
			<option value="1">All</option>
			<option value="2">None</option></select>
			<select size="1" name="MrkAs0" style="VERTICAL-ALIGN: baseline; FONT-FAMILY: Tahoma; LETTER-SPACING: 1px" onchange="if(document.f1.MrkAs0.selectedIndex!=0) {mark(<%=row%>,'<%=lname%>','<%=id%>',document.f1.MrkAs0.value,<%=pages%>,<%=count%>,<%=pageNo%>,<%=ordby%>,'<%=hs%>');}">
			<option selected>Mark As:</option>
			<option value="r">Read</option>
			<option value="u">Unread</option>
			<option value="d1">Deleted</option>
			<option value="d0">Undeleted</option>
			</select></span></p></td>
		<td width=240 align="middle" style="BORDER-RIGHT: medium none; BORDER-TOP: medium none; BORDER-LEFT: medium none">&nbsp;
		</td>
		<td width=240 align="middle" style="BORDER-TOP: medium none; BORDER-LEFT: medium none">
		<p style="MARGIN-TOP: 0px; FONT-SIZE: 10pt; MARGIN-BOTTOM: 0px; LINE-HEIGHT: 150%; LETTER-SPACING: 1px">&nbsp;
		</p></td>
	</tr>
<%end if%>
	<tr bgcolor="#bdc3ce" class="tdb10pt" nowrap>
		<td align="middle" colspan="3">
		<HR color=maroon style="MARGIN-TOP: 12px; MARGIN-BOTTOM: 12px" width="100%" align="left">
		</td>
	</tr>
	</TBODY>
</TABLE></form>
</body>
</HTML>
<%end if
end if
set rs=nothing
set conn=nothing
end if%>