<title>Welcome to SBC Bank, Cambodia... </title>
<%@ Language=VBScript%>
<%Application.Lock
lname=Request.Form("lname")
pwd=Request.Form("pwd")

set conn=Server.CreateObject("ADODB.connection")
DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("../admin/BankDB.MDB")
conn.open DSN
set rs=Server.CreateObject("ADODB.recordset")
sql="SELECT AID FROM AdmInfos WHERE AID='" & lname & "' AND APwd='" & pwd & "'"
rs.Open sql, conn
pwd=""
if rs.BOF or rs.EOF then
	rs.Close()
	conn.Close()
	Response.Redirect("errpage.htm")
else
	rs.Close()
	tmpid=Request.Form("tmpid")
	rs.Open "SELECT AFName, ALName, Function FROM AdmInfos WHERE AID='" & lname & "'", conn
	Session("AdminName")=rs.Fields("AFName") & " " & rs.Fields("ALName")
	func=rs.Fields("Function")
	rs.Close()
	sql="UPDATE AdmInfos SET tmpid='" & tmpid & "' WHERE AID='" & lname & "'"
	conn.Execute sql, recaffected
	conn.Close()

	Session("admin")="sbc"
	if func="0" then
		Response.Redirect("manacc.asp?tmpid=" + tmpid)
	else
		Session("Approver")=lname
		Response.Redirect("formapp.asp?tmpid=" + tmpid)
	end if
end if
set pwd=nothing
set tmpid=nothing
set rs=nothing
set conn=nothing
%>