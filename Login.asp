<% @ language="VBScript" %>
<html lang="en">
	<head>
	  <title>Login</title>
	  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
	  <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1.0, user-scalable=no"/>

	  <!-- CSS  -->
	  <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
	  <link href="css/materialize.css" type="text/css" rel="stylesheet" media="screen,projection"/>
	  <link href="css/style.css" type="text/css" rel="stylesheet" media="screen,projection"/>
	</head>
	<body>
	<%  Dim conn,res,
		Dim us1,pd1,count
		count=
		set conn=Server.createobject("ADODB.connection")
		conn.provider="Microsoft.Jet.OLEDB.4.0"
		conn.open "C:\\InetPub\wwwroot\signup.mdb"
		set res=server.createobject("ADODB.Connection")
		res.open "signup1",0,3,2
		us1=request.form("loginname")
		pd1=request.form("loginpwd")
		while not res.eof
			if res("Email") = us1 and res("Password") = pd1 then
				count=1
				response.redirect("Login.html")
			else
				res.movenext
			end if
		loop
		