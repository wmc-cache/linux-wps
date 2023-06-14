<html>

<head>
	<%@ page language="java" import="java.io.*,java.util.*" %>
	<%@ page import="org.apache.commons.fileupload.disk.DiskFileItemFactory" %>
	<%@ page import="org.apache.commons.fileupload.servlet.ServletFileUpload" %>
	<%@ page language="java" contentType="text/html" pageEncoding="ISO-8859-1" %>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
	<title>This page for response</title>
</head>

<body>
	<%
	DiskFileItemFactory factory = new DiskFileItemFactory();
	ServletFileUpload upload = new ServletFileUpload(factory);
	upload.setHeaderEncoding("UTF-8");
	List items = upload.parseRequest(request);

	response.setCharacterEncoding("ISO-8859-1");
	response.setContentType("text/json");  
	String jsonString = "{\"status\":302,\"location\":\"http://10.90.128.241:8080/servletTest_N/HelloServlet\"}";
	out.clear();
	out.print(jsonString);
	out.close();
%>
</body>

</html>

