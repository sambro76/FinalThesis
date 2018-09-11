<%@ Language=VBScript%>
<%lname=Request.QueryString("lname")
id=Request.QueryString("id")
if lname="" or id="" then
	Response.Redirect("errpage.htm")
else
	set conn=Server.CreateObject("ADODB.Connection")
	DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("RemDB.MDB")
	conn.open DSN
	set rs=Server.CreateObject("ADODB.recordset")
	sql="SELECT AccID FROM Keyin WHERE AccID='" & lname & "' AND id='" & id & "'"
	rs.Open sql, conn
	if rs.EOF or rs.BOF then
		rs.close()
		conn.close()
		Response.Redirect("errpage.htm")
	else
		rs.Close()
		conn.Execute "DELETE * FROM RecData WHERE AccID='" & lname & "' AND Status='2'", recaffected
		conn.Close()
		Randomize()
		Response.Redirect("sbc_pri.asp?lname=" & lname & "&id=" & id & "&" & Right(Rnd(),len(Rnd())-1))
	end if
end if
set rs=nothing
set conn=nothing%>