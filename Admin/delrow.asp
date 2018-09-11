<%@ Language=VBScript %>
<%Application.Lock()
if Session("admin")<>"sbc" then
	Session("exp")="1"
	Response.Redirect("adlogin.asp")
else
	tmpid=Request.QueryString("tmpid")
	if tmpid="" then
		Response.Redirect("errpage.htm")
	else
		if Session("cCharge")<>"Charging" then
			Response.Redirect("charges.asp?tmpid=" & tmpid)
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
			for i=1 to int(Request.QueryString("numRow"))
				costi=Request.QueryString("cost" & i)
				if costi<>"" then
					conn.Execute "DELETE FROM SBCCharges WHERE CostPoint=" & costi
				end if
			next
			conn.Close()
			Randomize()
			Response.Redirect("charges.asp?tmpid=" & tmpid & "&." & Rnd())
		end if
		end if
	end if
end if
set rs=nothing
set conn=nothing
%>