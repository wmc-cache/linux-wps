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
	response.addHeader("location", "https://192.168.39.6:8443/servletTest/DOC_100KB.doc");
	response.setStatus(302);

%>
</body>

</html>

