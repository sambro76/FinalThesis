<%@ Language=VBScript %>
<%Application.Lock()
Response.Buffer=true
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
			Randomize()
			rs.Close()
			on error resume next
			conn.Execute "DELETE FROM AppInfos WHERE AccID='" & Request.QueryString("AccID") & "'"
			if err=0 then
				conn.Close()
				DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & left(Server.MapPath("BankDB.MDB"),len(Server.MapPath("BankDB.MDB"))-17) & "\Rem\RemDB.MDB"
				conn.Open DSN
				conn.Execute "DELETE FROM Keyin WHERE AccID='" & Request.QueryString("AccID") & "'"
				Session("Delete")="1"
				Response.Redirect("addmodacc.asp?tmpid=" & tmpid & "&." & Rnd())
			else
				Session("Delete")="0"
				Response.Redirect("modacc.asp?tmpid=" & tmpid & "&." & Rnd())
			end if
		end if
	conn.Close()
	end if
end if
set rs=nothing
set conn=nothing
%>