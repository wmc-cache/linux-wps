<html>

<head>
	<%@ page language="java" import="java.io.*" %>
	<%@ page language="java" contentType="text/html" pageEncoding="ISO-8859-1" %>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
	<title>This page for response</title>
</head>

<body>
	<%    
	response.setCharacterEncoding("ISO-8859-1");
	response.addHeader("location", "http://192.168.39.6:8080/servletTest/XLS_30MB.xls");
	response.setStatus(302);

%>
</body>

</html>