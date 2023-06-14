<html>

<head>
	<%@ page language="java" import="java.io.*" %>
	<%@ page language="java" contentType="text/html" pageEncoding="ISO-8859-1" %>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
	<title>This page for response</title>
</head>

<body>
	<%    
	if (request.getContentLength() > 0) 
	{           	
		InputStream in = request.getInputStream();
		byte b[] = new byte[1024];
		int n;
		while ((n = in.read(b)) != -1)
		{               
		}
	    in.close(); 
	}

	response.setCharacterEncoding("ISO-8859-1");
	response.addHeader("location", "http://192.168.39.18:8080/servletTest/upload_l.jsp");
	response.setStatus(302);

%>
</body>

</html>

