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
			with Request
			sql="UPDATE AppInfos SET "
			sql=sql & "AccID='" & .QueryString("AccID") & "',"
			sql=sql & "FName='" & .QueryString("FName") & "',"
			sql=sql & "LName='" & .QueryString("LName") & "',"
			sql=sql & "HB='" & .QueryString("HB") & "',"
			sql=sql & "Street='" & .QueryString("Street") & "',"
			sql=sql & "City='" & .QueryString("City") & "',"
			sql=sql & "Country='" & .QueryString("Country") & "',"
			sql=sql & "Telephone='" & .QueryString("Telephone") & "',"
			sql=sql & "Payment_Details='" & .QueryString("PayDetail") & "',"
			sql=sql & "SpIns_MPay='" & .QueryString("SpIns") & "',"
			sql=sql & "AccNo='" & .QueryString("AccNo") & "' WHERE AccID='" & .QueryString("AccID1") & "'"
			Randomize()
			on error resume next
			conn.Execute sql, recaffected
			if err=0 then
				conn.Close()
				DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & left(Server.MapPath("BankDB.MDB"),len(Server.MapPath("BankDB.MDB"))-17) & "\Rem\RemDB.MDB"
				conn.Open DSN
				conn.Execute "UPDATE Keyin SET AccID='" & .QueryString("AccID") & "',pwd='" & .QueryString("pwd") & "' WHERE AccID='" & .QueryString("AccID1") & "'"
				conn.Close()
				Session("Modify")="1"
				Session("accid")=.QueryString("AccID")
				Response.Redirect("modacc.asp?tmpid=" & tmpid & "&.accid=" & .QueryString("AccID") & "&." & Rnd())
			else
				on error resume next
				conn.Execute "UPDATE AppInfos SET AccID='" & .QueryString("AccID") & "' WHERE AccID='" & .QueryString("AccID1") & "'"
				if err=0 then 
					on error resume next
					conn.Execute "UPDATE AppInfos SET AccNo='" & .QueryString("AccNo") & "' WHERE AccID='" & .QueryString("AccID") & "'"
					if err<>0 then
						Session("Modify")="ErrAccNo"
						Response.Redirect("modacc.asp?tmpid=" & tmpid & "&.accid=" & .QueryString("AccID") & "&." & Rnd())
					end if
				else
					Session("Modify")="ErrAccID"
					Response.Redirect("modacc.asp?tmpid=" & tmpid & "&.accid=" & .QueryString("AccID1") & "&." & Rnd())
				end if
				conn.Close()
			end if
			end with
		end if
	end if
end if
set rs=nothing
set conn=nothing
%>