<%@ language="vbscript" %>
<%option explicit%>
<html>
	<head>

	</head>
	<body>
		<%  
			Dim conn,rs
			Dim f2,l2,e2,a2,g2,m2,d2,v2,cou2,c2,pan2,adhar2
			set conn=Server.createobject("ADODB.connection")
			conn.provider="Microsoft.Jet.OLEDB.4.0"
			conn.open "C:\\InetPub\wwwroot\Project(WT)(new)\Project(WT)\Visa_Database.mdb"
			set rs=server.createobject("ADODB.Recordset")
			rs.open "Visa_Database",conn,0,3,2
			rs.Addnew
			f2=request.form("fname2")
			l2=request.form("lname2")
			e2=request.form("em2")
			g2=request.form("gender2")
			m2=request.form("mobile2")
			d2=request.form("dob2")
			cou2=request.form("country2")
			c2=request.form("city2")
			pan2=request.form("pcn2")
			adhar2=request.form("acn2")
			a2=request.form("address2")
			rs("Fname2")=f2
			rs("Lname2")=l2
			rs("Email2")=e2
			rs("Address2")=a2
			rs("Gender2")=g2
			rs("Mobile2")=m2
			rs("DOB2")=d2
			rs("Country2")=cou2
			rs("City2")=c2
			rs("Pan_Card_No2")=pan2
			rs("Adhar_Card_No2")=adhar2
			rs.update
			response.redirect("Visa_Appointment.asp")
		%>
	</body>
</html>