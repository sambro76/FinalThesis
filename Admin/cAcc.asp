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
			sql="INSERT INTO AppInfos (AccID,FName,LName,HB,Street,City,Country,Telephone,Payment_Details,SpIns_MPay,AccNo,InitCredit) VALUES ('"
			sql=sql & .QueryString("AccID") & "','"
			sql=sql & .QueryString("FName") & "','"
			sql=sql & .QueryString("LName") & "','"
			sql=sql & .QueryString("HB") & "','"
			sql=sql & .QueryString("Street") & "','"
			sql=sql & .QueryString("City") & "','"
			sql=sql & .QueryString("Country") & "','"
			sql=sql & .QueryString("Telephone") & "','"
			sql=sql & .QueryString("PayDetail") & "','"
			sql=sql & .QueryString("SpIns") & "','"
			sql=sql & .QueryString("AccNo") & "','"
			sql=sql & .QueryString("InitCredit") & "')"
			on error resume next
			conn.Execute sql, recaffected
			Randomize()
			if err=0 then
				conn.Execute "INSERT INTO AccTrans (AccNo,Credit,Balance,Dated,Description) VALUES ('" & _ 
					.QueryString("AccNo") & "'," & .QueryString("InitCredit") & "," & .QueryString("InitCredit") & ",'" & _ 
					Now() & "','" & "Deposit" & "')"
				conn.Close()
				DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Left(Server.MapPath("BankDB.MDB"),Len(Server.MapPath("BankDB.MDB"))-17) & "\Rem\RemDB.MDB"
				conn.Open DSN
				conn.Execute "INSERT INTO Keyin (AccID, pwd) VALUES ('" & .QueryString("AccID") & "','" & .QueryString("pwd") & "')"
				conn.Close()
				Session("Create")="1"
				Session("AccCreated")=.QueryString("AccID")
				Session("pwd")=.QueryString("pwd")
				Response.Redirect("createacc.asp?tmpid=" & tmpid & "&." & Rnd())
			else
				Session("Create")="0"
				Session("pwd")=.QueryString("pwd")
				rs.Open "SELECT AccID FROM AppInfos WHERE AccID='" & .QueryString("AccID") & "'", conn
				if rs.BOF or rs.EOF then
					rs.Close()
					rs.Open "SELECT AccNo FROM AppInfos WHERE AccNo='" & .QueryString("AccNo") & "'", conn
					if rs.BOF or rs.EOF then
					else
						Session("AccCreated")="ErrAccNo"
					end if
				else 
					Session("AccCreated")="ErrAccID"
				end if
				rs.Close()
				conn.Close()
				sql="&AccID=" & .QueryString("AccID") & "&FName=" & .QueryString("FName") & "&LName="
				sql=sql & .QueryString("LName") & "&HB=" & .QueryString("HB") & "&Street="
				sql=sql & .QueryString("Street") & "&City=" & .QueryString("City") & "&Country="
				sql=sql & .QueryString("Country") & "&Telephone=" & .QueryString("Telephone") & "&PayDetail="
				sql=sql & .QueryString("PayDetail") & "&SpIns=" & .QueryString("SpIns") & "&AccNo=" 
				sql=sql & .QueryString("AccNo") & "&InitCredit=" & .QueryString("InitCredit")
				Response.Redirect("createacc.asp?tmpid=" & tmpid & sql & "&." & Rnd())
			end if
			end with
			set sql=nothing
			set rs=nothing
			set conn=nothing
		end if
	end if
end if
%>