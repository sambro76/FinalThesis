<title>Welcome to SBC Bank, Cambodia... </title>
<%@ Language=VBScript%>
<%Application.Lock
Response.Buffer=true
''''''
'get the login name & pwd
''''''
lname=Request.Form("lname")
pwd=Request.Form("pwd")
''''''
'call a session variable 'user' to validate the current session
''''''
Session("user")="sbc"

if lname="" or pwd="" then
	Response.Redirect("errpage.htm")
else
	''''''
	'Create ADODB Connection object to open database file
	''''''
	set conn=Server.CreateObject("ADODB.connection")
	DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("RemDB.MDB")
	conn.open DSN
	''''''
	'Create ADODB Recordset to get the record out
	''''''
	set rs=Server.CreateObject("ADODB.recordset")
	sql="SELECT AccID FROM Keyin WHERE AccID='" & lname & "' AND pwd='" & pwd & "'"
	rs.Open sql, conn
	pwd=""
	if rs.BOF or rs.EOF then
		rs.Close()
		conn.Close()
		''''
		'Return error page if no record found with that user account
		''''
		Response.Redirect("errpage.htm")
	else
		rs.Close()
		id=Request.Form("id")
		''''
		'Execute Update Query with the Temporary ID
		''''
		sql="UPDATE Keyin SET id='" & id & "' WHERE AccID='" & lname & "'"
		conn.Execute sql, recaffected
		conn.Close()
		DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & left(Server.MapPath("RemDB.MDB"),len(Server.MapPath("RemDB.MDB"))-13) & "\Admin\BankDB.MDB"
		conn.Open DSN
		''''
		'Get the users name from BankDB Database
		''''
		rs.Open "SELECT FName, LName FROM AppInfos WHERE AccID='" & lname & "'", conn
		Session("AppName")=rs.Fields("FName") & " " & rs.Fields("LName")
		rs.Close()
		conn.Close()

		Session("user")="sbc"
		Response.Redirect("sbc_pri.asp?lname=" + lname + "&id=" + id)
	end if
end if
set pwd=nothing
set id=nothing
set rs=nothing
set conn=nothing
%>