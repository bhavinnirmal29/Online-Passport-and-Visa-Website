<% @ language="VBscript" %>
<% Option Explicit %>
<% 
		Dim conn,rs
		Dim da,ti
		set conn=Server.createobject("ADODB.connection")
		conn.provider="Microsoft.Jet.OLEDB.4.0"
		conn.open "C:\\InetPub\wwwroot\Project(WT)(new)\Project(WT)\appointment1.mdb"
		set rs=server.createobject("ADODB.Recordset")
		rs.open "appointment",conn,0,3,2
		da=request.form("dob3")
		ti=request.form("time2")
		rs.Addnew
		rs("Date03") = da
		rs("TIME04") = ti 
		rs.update
		response.redirect("Successfull.html")
		response.redirect("index.html")
%>