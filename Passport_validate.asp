<%@ language="vbscript" %>
<%option explicit%>
<html>
	<head>

	</head>
	<body>
		<%  
			Dim conn,rs
			Dim gen1,acn1,fn1,ln1,pn1,dob1,pcn1,address1,us1,pt1
			set conn=Server.createobject("ADODB.connection")
			conn.provider="Microsoft.Jet.OLEDB.4.0"
			conn.open "C:\\InetPub\wwwroot\Project(WT)(new)\Project(WT)\passport_db.mdb"
			set rs=server.createobject("ADODB.Recordset")
			rs.open "passport",conn,0,3,2
			rs.Addnew
			us1=request.form("em1")
			fn1=request.form("fname1")
			ln1=request.form("lname1")
			pn1=request.form("mobile1")
			dob1=request.form("dob1")
			pcn1=request.form("pcn1")
			acn1=request.form("acn1")
			address1=request.form("address1")
			rs("Passport_Type1")=pt1
			rs("Fname1")=fn1
			rs("Lname1")=ln1
			rs("Mobile1")=pn1
			rs("DOB1")=dob1
			rs("Gender1")=gen1
			rs("Address1")=address1
			rs("Email1")=us1
			rs("Pan_Card_No1")=pcn1
			rs("Aadhar_Card_No1")=acn1
			rs.update
			response.redirect("index.html")
		%>
	</body>
</html>