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
	response.addHeader("location", "https://10.90.128.241:8443/servletTest_N/HelloServlet");
	response.setStatus(302);

%>
</body>

</html>