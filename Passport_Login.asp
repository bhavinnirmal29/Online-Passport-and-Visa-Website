<% @ language="VBScript"%>
<% Option Explicit %>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
		<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1.0, user-scalable=no"/>
		<title>Online Passport And Visa System</title>

		<!-- CSS  -->
		<link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
		<link href="css/materialize.css" type="text/css" rel="stylesheet" media="screen,projection"/>
		<link href="css/style.css" type="text/css" rel="stylesheet" media="screen,projection"/>
	
	</head>
<%  
	Dim conn,rs,count
	Dim username,logpass
	count=0
	set conn=Server.createobject("ADODB.connection")
	conn.provider="Microsoft.Jet.OLEDB.4.0"
	conn.open "C:\\InetPub\wwwroot\Project(WT)(new)\Project(WT)\signup.mdb"
	set rs=server.createobject("ADODB.Recordset")
	rs.open "signup1",conn,0,3,2
	username=request.form("login_em1")
	logpass=request.form("logpass1")
	while not rs.eof
		if rs("Email") = username AND rs("Password") = logpass then
			count=1
			response.redirect("Passport.html")
		else
			rs.movenext
		end if
	wend
	if count=0 then 
%>
		<a class="waves-effect waves-light btn modal-trigger" href="#modal1">Modal</a>

		<!-- Modal Structure -->
		<div id="modal1" class="modal">
			<div class="modal-content">
				<h4>Modal Header</h4>
				<p>A bunch of text</p>
			</div>
			<div class="modal-footer">
				<a href="#!" class=" modal-action modal-close waves-effect waves-green btn-flat">Agree</a>
			</div>
		</div>
<%		response.redirect("Passport.html")
	end if
%>
	<body>
		<script>
			 $(document).ready(function(){
			// the "href" attribute of .modal-trigger must specify the modal ID that wants to be triggered
			$('.modal-trigger').leanModal();
			});
		</script>

  <!--  Scripts-->
  <script src="https://code.jquery.com/jquery-2.1.1.min.js"></script>
  <script src="js/materialize.js"></script>
  <script src="js/init.js"></script>
	</body>
</html>
