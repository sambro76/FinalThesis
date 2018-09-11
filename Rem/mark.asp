<%@ Language=VBScript%>
<%Response.Buffer=true
lname=Request.QueryString("lname")
id=Request.QueryString("id")
set conn=Server.CreateObject("ADODB.Connection")
DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("RemDB.MDB")
conn.open DSN
set rs=Server.CreateObject("ADODB.recordset")
sql="SELECT AccID FROM Keyin WHERE AccID='" & lname & "' AND id='" & id & "'"
rs.Open sql, conn
if rs.EOF or rs.BOF then
	rs.close()
	conn.close()
	Response.Redirect(errpage.htm)
else
	rs.close()
	with Request
		row=.QueryString("row")
		sel=.QueryString("sel")
		sqlapp="SELECT Status FROM RecData WHERE REF_No='"
		for i=1 to row
			if .QueryString("cb[" & i & "]")<>"" then
				rs.Open sqlapp & .QueryString("cb[" & i & "]") & "'", conn
				status=rs.Fields("Status")
				rs.Close()
				if sel=2 then 'No.2 represent code of the deleted state
					if status=3 or status=4 then
						conn.Execute "DELETE FROM RecData WHERE AccID='" & lname & "' AND REF_No='" & .QueryString("cb[" & i & "]") & "'"
						init=0
					else
						sql="UPDATE RecData SET Status='" & sel & "' WHERE REF_No='" & .QueryString("cb[" & i & "]") & "'"
						conn.Execute sql, recaffected
						init=1
					end if
				else
					if status<>3 and status<>4 then
						sql="UPDATE RecData SET Status='" & sel & "' WHERE REF_No='" & .QueryString("cb[" & i & "]") & "'"
						conn.Execute sql, recaffected
					end if
					init=1
				end if
			end if
		next
		conn.Close()
		Randomize()
		hs=.QueryString("hs")
		if hs="hide" then
			url="&init=0&hs=" & hs
		else
			if init=0 then
				url="&init=0&hs=" & hs
			else
				url="&init=" & init & "&selectp=1&pages=" & .QueryString("pages") & "&count=" & .QueryString("count") & "&pageNo=" & .QueryString("pageNo") & "&ordby=" & .QueryString("ordby") & "&hs=" & hs
			end if
		end if
	end with
	Response.Redirect("sbc_pri.asp?lname=" & lname & "&id=" & id & url & "&mark=" & Rnd())
	set rs=nothing
	set conn=nothing
end if%>