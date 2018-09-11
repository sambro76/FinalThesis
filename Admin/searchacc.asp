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
			rs.Close()
			with Request
			for i=1 to 4
				if .QueryString("txtPIA" & i)<>"" then
					j=j+1
					if j=1 then
						sql=.QueryString("PIA" & i) & "='" & .QueryString("txtPIA" & i) & "'"
					else
						sql=sql & " AND " & .QueryString("PIA" & i) & "='" & .QueryString("txtPIA" & i) & "'"
					end if
				end if
			next
			end with
			rs.Open "SELECT * FROM AppInfos WHERE " & sql, conn
			if rs.BOF or rs.EOF then
				Session("Search")="0"
				Randomize()
				Response.Redirect("addmodacc.asp?tmpid=" & tmpid & "&." & Rnd())
			else
				do until rs.EOF
					frow=frow+1
					rs.MoveNext()
				loop
				rs.MoveFirst()
				if frow=1 then
					Session("Search")="1"
					Session("accid")=rs.Fields("AccID")
					rs.Close()
					Response.Redirect("modacc.asp?tmpid=" & tmpid)
				else
					Session("Search")=frow
					Session("accid")=sql
					rs.Close()
					Response.Redirect("modacc.asp?tmpid=" & tmpid)
				end if
			end if
		end if
	conn.Close()
	end if
end if
set i=nothing
set j=nothing
set rs=nothing
set conn=nothing
%>