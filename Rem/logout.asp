<%@ Language=VBScript%>
<%if Session("user")<>"sbc" then
	Session("exp")="1"
	Response.Redirect("sbc_pwd.asp")
else
''''''
'Get login name 'lname' & temporary id 'id' to admit the page requested
''''''
	lname=Request.QueryString("lname")
	id=Request.QueryString("id")
if lname="" or id="" then
	Response.Redirect("errpage.htm")
else
	set conn=Server.CreateObject("ADODB.connection")
	DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("RemDB.MDB")
	conn.open DSN
	set rs=Server.CreateObject("ADODB.recordset")
	''''''
	'Get login name 'lname' & temporary id 'id' to admit the page requested
	''''''
	sql="SELECT AccID FROM Keyin WHERE AccID='" & lname & "' AND id='" & id & "'"
	rs.Open sql, conn
	if rs.BOF or rs.EOF then
		rs.Close()
		conn.Close()
		Response.Redirect("errpage.htm")
	else
		rs.Close()
		set conn=Server.CreateObject("ADODB.connection")
		DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("RemDB.MDB")
		conn.open DSN
		sql="UPDATE Keyin SET id='' WHERE AccID='" & lname & "' AND id='" & id & "'"
		conn.Execute sql, recaffected
		conn.Close()
		set conn=nothing
		'Session.Abandon()
		Session("lout")="1"
		Session("exp")="0"
		Randomize()
		Response.Redirect("sbc_pwd.asp?.rnd=" & Right(Rnd(),6))
	end if
end if
end if
%>