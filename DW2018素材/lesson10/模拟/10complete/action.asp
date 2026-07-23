<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
</head>
<body>
<%
username=request.form("username")
password=request.form("password")
response.write("username:" & " " & request.form("username") & "<br />")
response.write("password:" & " " & request.form("password") & "<br />")
%>
</body>
</html>