<% @ language="VBScript"%>
<% Option Explicit %>
<%  
	Dim conn,rs
	Dim username,logpass
	set conn=Server.createobject("ADODB.connection")
	conn.provider="Microsoft.Jet.OLEDB.4.0"
	conn.open "C:\\InetPub\wwwroot\Project(WT)(new)\Project(WT)\signup.mdb"
	set rs=server.createobject("ADODB.Recordset")
	rs.open "signup1",conn,0,3,2
	username=request.form("login_em1")
	logpass=request.form("logpass1")
	while not rs.eof
		if rs("Email") = username AND rs("Password") = logpass then
			response.redirect("visa_form.html")
		else
			rs.movenext
		end if
	wend
%>