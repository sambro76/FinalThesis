<%@ Language=VBScript %>
<%Application.Lock()
if Session("admin")<>"sbc" then
	Session("exp")="1"
	Response.Redirect("adlogin.asp")
else
	if Request.Form("dreport")="1" then
		tmpid=Request.Form("tmpid")
	else 
		tmpid=Request.QueryString("tmpid")
	end if
	if tmpid="" then
		Response.Redirect("errpage.htm")
	else
		set conn=Server.CreateObject("ADODB.connection")
		DSN="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("BankDB.MDB")
		conn.open DSN
		if Request.Form("dreport")=1 then
			i=0
			do until Request.Form("ref" & i)=""
				conn.Execute "UPDATE LogDataApp SET Report='1' WHERE REF_No='" & Request.Form("ref" & i) & "'"
				i=i+1
			loop
		end if
		conn.Execute "UPDATE AdmInfos SET tmpid='' WHERE tmpid='" & tmpid & "'", recaffected
		conn.Close()
		'Application.Contents.Removeall
		'Session.Abandon()
		set conn=nothing
		if i>0 then
			Session("rpt")="1"
		else
			Session("rpt")="0"
		end if
		Randomize()
		Response.Redirect("adlogin.asp?.rnd=" & Right(Rnd(),6))
	end if
end if
%>