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
	response.addHeader("location", "http://10.90.128.241:8080/servletTest_N/http落地测试.pptx");
	response.setStatus(302);

%>
</body>

</html>