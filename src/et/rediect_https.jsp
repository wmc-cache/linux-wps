<html>

<head>
	<%@ page language="java" import="java.io.*,java.util.*" %>
	<%@ page language="java" contentType="text/html" pageEncoding="ISO-8859-1" %>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
	<title>This page for response</title>
</head>

<body>
	<%    
	response.setCharacterEncoding("ISO-8859-1");
	response.setContentType("text/json");  
	String jsonString = "{\"status\":302,\"location\":\"https://192.168.39.6:8443/servletTest/upload_l.jsp\"}";
	out.clear();
	out.print(jsonString);
	out.close();
%>
</body>

</html>

