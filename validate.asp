<% @ language="vbscript" %>
<%option explicit%>
<html>
	<head>
<%  Dim conn,rs
		Dim us,pd,fn,ln,pn,dob
		set conn=Server.createobject("ADODB.connection")
		conn.provider="Microsoft.Jet.OLEDB.4.0"
		conn.open "C:\\InetPub\wwwroot\Project(WT)(new)\Project(WT)\signup.mdb"
		set rs=server.createobject("ADODB.Recordset")
		rs.open "signup1",conn,0,3,2
		rs.Addnew
		us=request.form("em0")
		pd=request.form("pwd")
		fn=request.form("fname")
		ln=request.form("lname")
		pn=request.form("mobile")
		dob=request.form("dob0")
		rs("Fname")=fn
		rs("Lname")=ln
		rs("Mobile")=pn
		rs("DOB")=dob
		rs("Email")=us
		rs("Password")=pd
		rs.update
		response.redirect("Select.html")
%>
	</head>
</html>