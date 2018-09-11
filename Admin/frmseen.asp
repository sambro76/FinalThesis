<%@ Language=VBScript %>
<%Application.Lock()
if Session("admin")<>"sbc" then
	Session("exp")="1"
	Response.Redirect("adlogin.asp")
else
	tmpid=Request.Form("tmpid")
	if tmpid="" then
		Response.Redirect("errpage.htm")
	else
		set conn=Server.CreateObject("ADODB.connection")
		DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & left(Server.MapPath("BankDB.MDB"),len(Server.MapPath("BankDB.MDB"))-17) & "\Rem\RemDB.MDB"
		conn.Open DSN
		set rs=Server.CreateObject("ADODB.Recordset")
		
		with Request
		if .Form("cAmount")=1 and .Form("cBName")=1 and .Form("cAddress")=1 and .Form("cBBank")=1 and .Form("cLocation")=1 and .Form("cBAccNo")=1 then
			rs.Open "SELECT * FROM RecData WHERE REF_No='" & .Form("refno") & "'", conn
			set conn1=Server.CreateObject("ADODB.connection")
			DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("BankDB.MDB")
			conn1.Open DSN
			'Create Log of Form Approved
			conn1.Execute "INSERT INTO LogDataApp VALUES ('" & rs.Fields("AccID") & "','" & _ 
				Session("Approver") & "','" & rs.Fields("REF_No") & "','" & rs.Fields("Currency") & "','" & _
				rs.Fields("Amount") & "','" & rs.Fields("Rate") & "','" & rs.Fields("Charge_for") & "','" & _
				rs.Fields("BFName") & "','" & rs.Fields("BLName") & "','" & rs.Fields("BHB") & "','" & _
				rs.Fields("BStreet") & "','" & rs.Fields("BCity") & "','" & rs.Fields("BCountry") & "','" & _
				rs.Fields("BBank") & "','" & rs.Fields("bbCity") & "','" & rs.Fields("bbCountry") & "','" & _
				rs.Fields("bAccNo") & "','" & rs.Fields("Dated") & "','" & Now() & "','0')" 
			
			debit=rs.Fields("Amount")/rs.Fields("Rate")
			'Get Bank Charges depending on this debit for each order
			set rs1=Server.CreateObject("ADODB.Recordset")
			rs1.Open "SELECT * FROM SBCCharges WHERE CostPoint>" & debit, conn1
			if rs1.EOF or rs1.BOF then
				rs1.Close()
				rs1.Open "SELECT * FROM SBCCharges ORDER BY CostPoint DESC", conn1
			end if
			commission=rs1.Fields("Commission")
			ccable=rs1.Fields("CCable")
			agentc=rs1.Fields("AgentC")
			rs1.Close()
			
			if rs.Fields("Charge_for")="0" then
				cSum=commission+ccable+agentc
			else
				cSum=0
			end if
			'Get Last Balance and AccNo
			balance=Session("Balance")
			AccNo=Session("AccNo")
						
			'Create Account Transaction from the succeeded order...
			conn1.Execute "INSERT INTO AccTrans (AccNo,REFNo,Commission,CCable,AgentC,Debit,Balance,Dated) VALUES('" &  _
				AccNo & "','" & rs.Fields("REF_No") & "'," & commission & "," & ccable & "," & agentc & "," & debit & "," & balance-debit-cSum & ",'" & Now() & "')"
			conn1.Close()
			set rs1=nothing
			set conn1=nothing
			'Set status to the user order form...
			conn.Execute "UPDATE RecData SET Status='0' WHERE REF_No='" & .Form("refno") & "'"
			Session("numApp")=Session("numApp")+1
		else
			if .Form("cAmount")=0 then eAmount=.Form("eAmount")
			if .Form("cBName")<>1 then eBName=.Form("eBName")
			if .Form("cAddress")<>1 then eAddress=.Form("eAddress")
			if .Form("cBBank")<>1 then eBBank=.Form("eBBank")
			if .Form("cLocation")<>1 then eLocation=.Form("eLocation")
			if .Form("cBAccNo")<>1 then eBAccNo=.Form("eBAccNo")
			on error resume next
			conn.Execute "INSERT INTO ErrLog VALUES ('" & .Form("refno") & "','" & eAmount & "','" & eBName & "','" & eAddress & "','" & eBBank & "','" & eLocation & "','" & eBAccNo & "')"
			if err<>0 then
				conn.Execute "UPDATE ErrLog SET EAmount='" & eAmount & "', EBName='" & eBName & "', EBAddress='" & eAddress & _ 
					"', EBBank='" & eBBank & "', EBLocation='" & eLocation & "', EBAccNo='" & eBAccNo & "' WHERE EREFNo='" & .Form("refno") & "'"
			end if
			'Set status to the user order form...
			conn.Execute "UPDATE RecData SET Status='4' WHERE REF_No='" & .Form("refno") & "'"
		end if
		end with
		rs.Close()
		conn.Close()
		set rs=nothing
		set conn=nothing
		if Session("Bookmark")=Session("oForms") then
			if Session("numApp")>0 then
				Response.Redirect("frmReport.asp?tmpid=" & tmpid)
			else
				Response.Redirect("formapp.asp?tmpid=" & tmpid)
			end if
		else
			Session("Bookmark")=Session("Bookmark")+1
			Response.Redirect("formapp.asp?tmpid=" & tmpid)
		end if
	end if
end if
%>