<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>登录处理页面</title>
</head>
<body>
<%
	set conn=server.CreateObject("adodb.connection")
	sql="provider=microsoft.jet.oledb.4.0;data source="&server.MapPath("user.mdb")
	conn.open(sql)                                                                   	
	yh=request.Form("username")
	mm=request.Form("password")                                                         	
	set rs=server.CreateObject("adodb.recordset")
	sqll="select * from info where username='"&yh&"' and password='"&mm&"'"                
	rs.open sqll,conn,1,3
	if rs.Eof then                                                                   
%>
<script language="javascript">
{
	alert("用户名、密码不正确");
	window.location.href="index.html"
}
</script>
<%                                                                                 
	else
	session("yh")=yh
        response.Write(session("yh"))
	response.Write("<div align= center><H1>恭喜你登录成功！<H1></div>")
	end if
%>
</body>