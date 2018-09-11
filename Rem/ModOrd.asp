<%@ Language=VBScript %>
<%if Session("user")<>"sbc" then
	Session("exp")="1"
	Response.Redirect("sbc_pwd.asp")
else
lname=Request.Form("lname")
id=Request.Form("id")
set conn=Server.CreateObject("ADODB.connection")
DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("RemDB.MDB")
conn.open DSN
set rs=Server.CreateObject("ADODB.recordset")
sql="SELECT AccID FROM Keyin WHERE AccID='" & lname & "' AND id='" & id & "'"
rs.Open sql, conn
if rs.BOF or rs.EOF then
	rs.Close()
	conn.Close()
	Response.Redirect("errpage.htm")
else
	rs.Close()
	with Request
		if .Form("D2")="USD" then
			rate=1
		else
			rs.Open "SELECT * FROM Rate WHERE Curr_Type='" & .Form("D2") & "'", conn
			rate=rs.Fields("Rate")
			rs.Close()
		end if
	sql="UPDATE RecData SET [Currency]='" & .Form("D2") & "', Amount='" & .Form("amnt") & _
		"', Rate='" & rate & "', Charge_for='" & .Form("R3") & "', BFName='" & .Form("bfname") & _
		"', BLName='" & .Form("blname") & "', BHB='" & .Form("bhb") & "', BStreet='" & .Form("bstreet") & _
		"', BCity='" & .Form("bcity") & "', BCountry='" & .Form("bcountry") & "', BBank='" & .Form("bbank") & _
		"', bbCity='" & .Form("bbcity") & "', bbCountry='" & .Form("bbcountry") & "', bAccNo='" & .Form("baccn") & _
		"', Dated='" & Now() & "', Status='3' WHERE REF_No='" & Session("REFNo") & "'"
	on error resume next
	conn.Execute sql, recaffected
	conn.Close()
	end with
	%>
	<HTML>
	<head>
	<style type=text/css>
		@import url("images/sbc.css");
		BODY {SCROLLBAR-BASE-COLOR: #FFFFE1}
	</style> 
	<%if err<>0 then%>
	<title>Error Occur While Sending...</title>
	<%else
		Response.Write("<title>Your Form Has Been Sent Successfully...</title>")
	end if%>
	</head>
	<BODY bgcolor='#FFFFE1'>
	<%if err<>0 then%>
		<p class="tx12pt"><b>Sorry, there was an error while sending the information... </b></p>
		<p class="tx12pt"><b>Please check your data and be sure they are properly input.</b></p>
		<p class="td10pt"><b>Please click here to <a href="javascript:window.history.go(-1);">Go Back</a> and Try Again... </b></p>
	<%else
		Session("REFNo")=""
		Response.Write("<p class='tx11pt'>Thank For Your Order... </p>")
		Response.Write("<p class='tx11pt'><font color='#800080'><b>Please always try to close the related windows which are displaying your personal information...</b></font></p>")
		Response.Write("<p class='td10pt'><b><a href='javascript:window.close()'>Close this window now</a></b></p>")
	end if%>
	</body>
	</HTML>
<%end if
end if
set rs=nothing
set conn=nothing%>