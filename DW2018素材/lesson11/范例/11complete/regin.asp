<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>注册处理页面</title>
</head>
<body>
<%
	set conn=server.CreateObject("adodb.connection")
	sql="provider=microsoft.jet.oledb.4.0;data source="&server.MapPath("user.mdb")
	conn.open(sql)                                                                   
	yh=request.Form("username")
	mm=request.Form("password")
        remm=request.Form("repassword")
        dh=request.Form("tel")
        yj=request.Form("email")

if yh=""or mm=""or mm<>remm then
response.write"<p align=center>注册失败,用户名及密码不能为空,且两次输入密码要一致!</p>"
response.write"<p align=center><a href=regin.html>重新注册</a></p>"
response.End
else
                                                       
	set rs=server.CreateObject("adodb.recordset")
	sqll="select * from info"
	rs.open sqll,conn,1,3                                                              
	rs.addnew
	rs("username")=yh
	rs("password")=mm
        rs("repassword")=remm
        rs("tel")=dh
        rs("email")=yj
	rs.update                                                                        
	rs.close
	set rs=nothing   
End if                                                                    
%>
<script>
{
alert("注册成功，转回登录页面");
window.location.href="index.html"
}                                                                                        
</script>
</body>