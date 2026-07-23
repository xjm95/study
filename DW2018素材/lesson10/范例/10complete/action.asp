<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
</head>
<body>
<%
username=request.form("username")
email=request.form("email")
tel=request.form("tel")
address=request.form("address")
data1=request.form("data1")
info=request.form("info")
response.write("username:" & " " & request.form("username") & "<br />")
response.write("email:" & " " & request.form("email") & "<br />")
response.write("tel:" & " " & request.form("tel") & "<br />")
response.write("address:" & " " & request.form("address") & "<br />")
response.write("data:" & " " & request.form("data1") & "<br />")
response.write("info:" & request.form("info") & "<br />")
%>
</body>
</html>