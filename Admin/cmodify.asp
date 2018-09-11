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
		if Session("cCharge")<>"Charging" then
			Response.Redirect("manacc.asp?tmpid=" & tmpid)
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
			on error resume next
			conn.Execute "DELETE * FROM SBCCharges"
			with Request
			for i=1 to cint(.Form("numRow"))
				conn.Execute "INSERT INTO SBCCharges (Commission, CCable, AgentC, CostPoint) VALUES (" & .Form("commission" & i) & "," & .Form("ccable" & i) & "," & .Form("agentc" & i) & "," & .Form("cost" & i) & ");"
			next
			conn.Close()
			end with
			Randomize()
			if err=0 then
				Session("cCharge")="Charged"
				Response.Redirect("manacc.asp?tmpid=" & tmpid & "&." & Rnd())
			else
				Session("cCharge")="Error"
				Response.Redirect("manacc.asp?tmpid=" & tmpid & "&." & Rnd())
			end if
		end if
		end if
	end if
end if
set rs=nothing
set conn=nothing
%>