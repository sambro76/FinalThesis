<%@ Language=VBScript%>
<%Application.Lock()
if Session("admin")<>"sbc" then
	Session("exp")="1"
	Response.Redirect("adlogin.asp")
else
	tmpid=Request.Form("tmpid")
	if tmpid="" then
		tmpid=Request.QueryString("tmpid")
	end if
if tmpid="" then
	Response.Redirect("errpage.htm")
else
	if Session("cCharge")="" then
		Session("cCharge")="Charging"
	end if
	if Session("cCharge")<>"Charging" then
		Response.Redirect("manacc.asp?tmpid=" & tmpid)
	end if

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
		const range=5
%>
<HTML>
<head>
<title>SBC Charges Method... </title>
<script language="javascript" src="../admin/images/cValid.js"></script>
<STYLE TYPE="text/css">
@import url("images/admin.css");
BODY {
	SCROLLBAR-BASE-COLOR: #E8FFFF
}
</style> 
</head>
<body bgcolor="#474545" background="images/cordurouy.gif" onload="window.name='charges'">
<div align="center">
<FORM name=f1 method=post target=_self action="charges.asp">
<table border=1 width="640" id="tb1" cellspacing="0" bgcolor="#e7fbfe" bordercolor="#666699">
	<tr>
		<td colspan="2">
		<p style="MARGIN-TOP: 0px; margin-bottom:0" align="center" class=headeng>&nbsp;</p>
		<p style="MARGIN-TOP: 0px" align="center" class="headeng">
		<span style="letter-spacing: 2px">
		<%rs.Open "SELECT * FROM SBCCharges ORDER BY CostPoint ASC;",conn
		if rs.EOF or rs.BOF then
			rs.Close()
			exist=0
			Response.Write("ADD")
		else
			exist=1%>MODIFY
		<%end if%>CHARGES METHOD</span></p>
		<p align="right" style="MARGIN-TOP: 0px" class="td10pt"><b>SBC Bank Co., LTD of Cambodia</b>
		<p align="right" style="MARGIN-TOP: 0px" class="td10pt">You are login as : <b><%=Session("AdminName")%></b>
		<input type=hidden name=tmpid value=<%=tmpid%>>
		<HR size=2 color=maroon style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px"></td>
	</tr>
	<tr>
		<td align="middle" bgcolor="#037ead" nowrap style="border-right-style: none; border-right-width: medium" width="75%">
			<p style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 6px" class=thw10pt align="left"> 
			&nbsp;Modifications are made in US Dollars...
		<td align="bottom" class=thw10pt bgcolor="#037ead" style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px; border-left-style:none; border-left-width:medium" nowrap onclick="window.open('logout.asp','charges')">
			<p align="center">
				<a onmouseover="status='Back to main menu... '; return true" onmouseout="status=''" href="javascript:window.open('manacc.asp?tmpid=<%=tmpid%>','charges')">
				<font color="#FFFFFF">Main Menu</font></a><font color="#FFFFFF">
				</font>&nbsp;&nbsp;|&nbsp;&nbsp;
				<a onmouseover="status='Log out... '; return true" onmouseout="status=''" href="javascript:window.open('adlogout.asp?tmpid=<%=tmpid%>','charges')">
				<font color="#FFFFFF">Log out</font></a>
			</p>
		</td>
	<tr bgcolor="#bdc3ce" class="tdb10pt" nowrap>
		<td bordercolor="#000080" bgcolor="#FFFFF0" colspan="2">
	
	<table border="1" width="100%" id="tbl1" bgcolor="#ffffff" cellspacing="0" bordercolor="#bdc3ce" cellpadding="6" style="letter-spacing: 1px">
		<tr bgColor=#660033 class="thw10pt">
			<th width="20%">Cost Range<p style="margin-top: 0">(In US$)</th>
			<th width="20%">Commission<p style="margin-top: 0">Charge</th>
			<th width="20%">Cable<p style="margin-top: 0">Charge</th>
			<th width="20%">Agent<p style="margin-top: 0">Charge</th>
			<th bgColor= #800040 nowrap>Remarks</th>
			<th bgColor= #800040 nowrap>Delete<p style="margin-top: 0">Existing?</th>
		</tr>
		<%add=Request.Form("add")
		if add="" then
			add=0
		end if
		add=cint(add)
		if exist=1 then
			do until rs.EOF
				count=count+1
				rs.MoveNext()
			loop
			rs.MoveFirst()
			do until rs.EOF
				i=i+1
				cost2=rs.Fields("CostPoint")%>
			<tr bgColor=#FFFFF7 nowrap class="td10pt">
				<%if i=1 then%>
					<td nowrap><p align="right">Less or equal 
					<input type=text name="cost<%=i%>" size="8" maxlength="10" value=<%=cost2%> onchange="if(<%=count%>>1 || <%=add%>>0) {document.f1.txt<%=i+1%>.disabled=false; if(isNaN(this.value)){document.f1.txt<%=i+1%>.value='Error!'} else {document.f1.txt<%=i+1%>.value=parseInt(this.value)+0.01;} document.f1.txt<%=i+1%>.disabled=true;}" style="text-align: center"></p></td>
				<%else%>
					<td nowrap><p align="right">
					<input type=text size="8" maxlength="10" value=<%=cost1+0.01%> disabled style="background-color: #FFFFF0; text-align:center" name="txt<%=i%>"> 
					To
					<%if i=count then%> 
						<input type=text name="cost<%=i%>" size="8" maxlength="10" value=<%=cost2%> onchange="if(<%=add%>>0) {document.f1.txt<%=i+1%>.disabled=false; document.f1.txt<%=i+1%>.value=parseInt(this.value)+0.01; document.f1.txt<%=i+1%>.disabled=true;}" style="text-align: center"></p></td>
					<%else %>
						<input type=text name="cost<%=i%>" size="8" maxlength="10" value=<%=cost2%> onchange="document.f1.txt<%=i+1%>.disabled=false; document.f1.txt<%=i+1%>.value=parseInt(this.value)+0.01; document.f1.txt<%=i+1%>.disabled=true;" style="text-align: center"></p></td>
					<%end if%>
				<%end if%>
					<td><p align="center"><input type=text name="commission<%=i%>" size="6" maxlength="5" value=<%=rs.Fields("Commission")%> style="text-align: center"></td>
					<td><p align="center"><input type=text name="ccable<%=i%>" size="6" maxlength="5" value=<%=rs.Fields("CCable")%> style="text-align: center"></td>
					<td><p align="center"><input type=text name="agentc<%=i%>" size="6" maxlength="5" value=<%=rs.Fields("AgentC")%> style="text-align: center"></td>
					<td nowrap><p align="center">Existing</td>
					<td nowrap><p align="center"><a onmouseover="status='Delete existing cost range... '; return true;" onmouseout="status=''" href="javascript:window.open('delrow.asp?tmpid=<%=tmpid%>&cost<%=i%>=<%=cost2%>&numRow=<%=count%>','charges')">DELETE</a></td>
			</tr>
			<%cost1=cost2
			rs.MoveNext()%>
			<%loop
			rs.Close()
			conn.Close()
		end if%>
		<%
			if add=0 and exist=0 then%>
				<tr bgColor=#FFFFF7 nowrap class="td10pt">
					<td colspan="6"><p align="center"><font color="#FA8072">There is no cost range assigned before, 
					or all costs have been cleared... </font></td>
				</tr>
			<%else
				if add>range then 
					exceed=range
				else 
					exceed=add
				end if
				for j=1 to exceed%>
				<tr bgColor=#FFFFF7 nowrap class="td10pt">
					<td nowrap>
					<p align="right">
					<%if Request.Form("cost" & cstr(i+j-1))<>"" then
						if IsNumeric(Request.Form("cost" & cstr(i+j-1))) then 
							preValue=Request.Form("cost" & cstr(i+j-1))+0.01
						else 
							preValue="Error!"
						end if
					else
						preValue=""
					end if
					if j=1 then
						if exist=1 then%>
							<input type=text size="8" maxlength="10" disabled value="<%=cost1+0.01%>" style="background-color: #FFFFF0; text-align:center" name="txt<%=i+j%>">
							 To <input type=text name="cost<%=i+j%>" size="8" maxlength="10" value="<%=Request.Form("cost" & cstr(i+j))%>" onchange="if(<%=exceed%>>1) {document.f1.txt<%=j+i+1%>.disabled=false; if(isNaN(this.value)) {document.f1.txt<%=j+i+1%>.value='Error!';} else {document.f1.txt<%=j+i+1%>.value=parseInt(this.value)+0.01;} document.f1.txt<%=j+i+1%>.disabled=true;}" style="text-align: center">
						<%else%>
							Less or equal <input type=text name="cost<%=i+j%>"  value="<%=Request.Form("cost" & cstr(i+j))%>" size="8" maxlength="10" onchange="if(<%=exceed%>>1) {document.f1.txt<%=j+i+1%>.disabled=false; if(isNaN(this.value)) {document.f1.txt<%=j+i+1%>.value='Error!';} else {document.f1.txt<%=j+i+1%>.value=parseInt(this.value)+0.01;} document.f1.txt<%=j+i+1%>.disabled=true;}" style="text-align: center">
						<%end if%>
					<%elseif j=exceed then%>
						<input type=text size="8" maxlength="10" disabled value="<%=preValue%>" style="background-color: #FFFFF0; text-align:center" name="txt<%=i+j%>">
							To <input type=text name="cost<%=i+j%>" size="8" maxlength="10" value="<%=Request.Form("cost" & cstr(i+j))%>" style="text-align: center">
					<%else%>
						<input type=text size="8" maxlength="10" disabled value="<%=preValue%>" style="background-color: #FFFFF0; text-align:center" name="txt<%=i+j%>">
							To <input type=text name="cost<%=i+j%>" size="8" maxlength="10" value="<%=Request.Form("cost" & cstr(i+j))%>" onchange="document.f1.txt<%=j+i+1%>.disabled=false; if(isNaN(this.value)) {document.f1.txt<%=j+i+1%>.value='Error!';} else {document.f1.txt<%=j+i+1%>.value=parseInt(this.value)+0.01;} document.f1.txt<%=j+i+1%>.disabled=true;" style="text-align: center">
					<%end if%>
					</p>
					</td>
					<td nowrap><p align="center">
						<input type=text name="commission<%=i+j%>" size="6" maxlength="5" value="<%=Request.Form("commission" & cstr(i+j))%>" style="text-align: center"></p></td>
					<td><p align="center">
						<input type=text name="ccable<%=i+j%>" size="6" maxlength="5" value="<%=Request.Form("ccable" & cstr(i+j))%>" style="text-align: center"></p></td>
					<td><p align="center">
						<input type=text name="agentc<%=i+j%>" size="6" maxlength="5" value="<%=Request.Form("agentc" & cstr(i+j))%>" style="text-align: center"></p></td>
					<td><p align="center">Add</p></td>
					<td>&nbsp;</td>
				</tr>
				<%next%>
			<%end if%>
		</table>
		<%if add>range then%>
			<p style="margin-top: 6px; margin-bottom: 12px" align="center" class="td10pt"><font color="#FF0000">Only <%=range%> rows per session...</font></p>
			<p style="margin-top: 12px; margin-bottom: 12px" align="center">
			<input type="button" value="No More" name="bAdd" disabled style="font-family: Verdana; font-size: 10pt; font-weight: bold">
			<%add=range%>
		<%else
			add=add+1%>
			<p style="margin-top: 12px; margin-bottom: 12px" align="center">
			<input type="button" value="Add a row" name="bAdd" title='Add new row to cost range... ' onclick="document.f1.submit();" style="font-family: Verdana; font-size: 10pt; font-weight: bold">
		<%end if%>
		<input type=hidden name="add" value=<%=add%>>
		<%if j>1 then%>
			<input type="button" value="Delete last row" name="bDel" title="Delete last added range... " onclick="if(document.f1.bAdd.value=='No More') {document.f1.add.value='<%=add-1%>';} else {document.f1.add.value='<%=add-2%>';} document.f1.submit();" style="font-family: Verdana; font-size: 10pt; font-weight: bold">
		<%else%>
			<input type="button" value="Delete last row" name="bDel" disabled style="font-family: Verdana; font-size: 10pt; font-weight: bold">
		<%end if%>
		<%if j<=1 then%>
			<input type="button" value=" Reset " name="bReset" disabled style="font-family: Verdana; font-size: 10pt; font-weight: bold">&nbsp;&nbsp;&nbsp;&nbsp;
		<%else%>
			<input type="button" value=" Reset " name="bReset" title="Reset to orginal cost range... " onclick="document.f1.add.value='0'; document.f1.submit();" style="font-family: Verdana; font-size: 10pt; font-weight: bold">&nbsp;&nbsp;&nbsp;&nbsp;
		<%end if%>
		<input type=hidden name="numRow" value=<%=i+j-1%>>
		<%if i+j=0 then%>
			<input type="button" value="Commit Changes" name="bModify" disabled style="font-family: Verdana; font-size: 10pt; font-weight: bold">
		<%else%>
			<input type="button" value="Commit Changes" name="bModify" title="Apply modification to cost range here... " onclick="if(chargeValid()!=0) {document.f1.action='cmodify.asp'; document.f1.submit();}" style="font-family: Verdana; font-size: 10pt; font-weight: bold">
		<%end if%>
		</p>
	<HR size=2 color=maroon style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px"></td>
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