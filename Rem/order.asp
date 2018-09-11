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
	with Request
		D2=.Form("D2")
		Amount=.Form("amnt")
		R3=.Form("R3")
		bfname=.Form("bfname")
		blname=.Form("blname")
		bhb=.Form("bhb")
		bstreet=.Form("bstreet")
		bcity=.Form("D4")
			if .Form("D4")="0"  then bcity=.Form("bcityo")
		bcountry=.Form("D5")
			if .Form("D5")="0"  then bcountry=.Form("bcountryo")
		bbank=.Form("bbank")
		bbcity=.Form("D6")
			if .Form("D6")="0"  then bbcity=.Form("bbcityo")
		bbcountry=.Form("D7")
			if .Form("D7")="0"  then bbcountry=.Form("bbcountryo")
		bAccn=.Form("baccn")
	end with
	rs.Open "SELECT * FROM Unapproved"
	if rs.BOF or rs.EOF then
		tmpcount=1
	else
		tmpcount=Int(rs.Fields("numUsers"))+1
	end if
	rs.Close()
	genRef="RM" & Right(Year(now()),2) & Month(now()) & Day(now()) & tmpcount
	if D2="USD" then 
		rate=1
	else 
		rs.Open "SELECT * FROM Rate WHERE Curr_Type='" & D2 & "'", conn
		rate=rs.Fields("Rate")
		rs.Close()
	end if
	sql="INSERT INTO RecData (AccID, REF_No, [Currency], Amount, Rate, Charge_for, BFName, BLName, BHB, BStreet, BCity, BCountry, BBank, bbCity, bbCountry, bAccNo, Dated, Status) VALUES ('"
	sql=sql & lname & "','" & genRef & "','" & D2 & "','" & Amount & "','" & rate & "','" & R3 & "','" & bfname & "','" & blname & "','" & bhb & "','" & bstreet & "','" &  bcity & "','" & bcountry & "','" & bbank & "','" & bbcity & "','" & bbcountry & "','" & bAccn & "','" & now() & "','3'" & ")"
	on error resume next
	conn.Execute sql, recaffected
	''''''
	'If error free, Execute Query to increase a user Ref. no to Table 'Unapproved'
	''''''
	if err=0 then
		if tmpcount=1 then
			conn.Execute "INSERT INTO Unapproved(numUsers) VALUES('" & tmpcount & "')", recaffected
		else
			conn.Execute "UPDATE Unapproved SET numUsers='" & tmpcount & "'", recaffected
		end if
		conn.Close()
	end if
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