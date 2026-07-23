<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/conn.asp" -->
<%
Dim rs1
Dim rs1_cmd
Dim rs1_numRows

Set rs1_cmd = Server.CreateObject ("ADODB.Command")
rs1_cmd.ActiveConnection = myconn_STRING
rs1_cmd.CommandText = "SELECT * FROM shop" 
rs1_cmd.Prepared = true

Set rs1 = rs1_cmd.Execute
rs1_numRows = 0
%>
<%
Dim Repeat1_numRows
Dim Repeat1_index

Repeat1_numRows = 4
Repeat1_index = 0
rs1_numRows = rs1_numRows + Repeat1_numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rs1_total
Dim rs1_first
Dim rs1_last

' set the record count
rs1_total = rs1.RecordCount

' set the number of rows displayed on this page
If (rs1_numRows < 0) Then
  rs1_numRows = rs1_total
Elseif (rs1_numRows = 0) Then
  rs1_numRows = 1
End If

' set the first and last displayed record
rs1_first = 1
rs1_last  = rs1_first + rs1_numRows - 1

' if we have the correct record count, check the other stats
If (rs1_total <> -1) Then
  If (rs1_first > rs1_total) Then
    rs1_first = rs1_total
  End If
  If (rs1_last > rs1_total) Then
    rs1_last = rs1_total
  End If
  If (rs1_numRows > rs1_total) Then
    rs1_numRows = rs1_total
  End If
End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs= rs1
MM_rsCount= rs1_total
MM_size= rs1_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rs1_first = MM_offset + 1
rs1_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rs1_first > MM_rsCount) Then
    rs1_first = MM_rsCount
  End If
  If (rs1_last > MM_rsCount) Then
    rs1_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = Server.HTMLEncode(MM_keepMove) & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<html lang="en">
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Untitled Document</title>
<!-- Bootstrap -->
<link href="css/bootstrap.css" rel="stylesheet">
<link href="css/main.css" rel="stylesheet" type="text/css">

<!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
<!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
<!--[if lt IE 9]>
		  <script src="https://oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
		  <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
		<![endif]-->
</head>
<body>
<div class="container-fluid back">
  <div class="table-responsive">
    <table class="table back1">
      <tbody>
        <tr>
          <td align="center" valign="middle"><img src="images/logo.png" class="img-responsive" width="400" height="200" alt=""/></td>
          <td align="left" valign="middle"><h1>我的购物车</h1></td>
        </tr>
      </tbody>
    </table>
    <table class="table back1" >
      <tbody >
        <tr>
          <td></td>
        </tr>
        <tr class="ys">
          <td width="10%" align="center" valign="middle" >编号</td>
          <td width="35%" align="center" valign="middle">商品</td>
          <td width="35%" align="center" valign="middle">详细信息</td>
          <td width="20%" align="center" valign="middle">单价（元）</td>
        </tr>
        <% 
While ((Repeat1_numRows <> 0) AND (NOT rs1.EOF)) %>
        <tr>
          <td width="10%" align="center" valign="middle"><p class="bh"><%=(rs1.Fields.Item("ID").Value)%></p></td>
          <td width="35%" align="center" valign="middle"><img src="<%=(rs1.Fields.Item("pic").Value)%>" class="img-responsive" alt=""/></td>
          <td width="35%" align="center" valign="middle"><p class="text"><%=(rs1.Fields.Item("info").Value)%></p></td>
          <td width="20%" align="center" valign="middle"><p class="price"><%=(rs1.Fields.Item("price").Value)%></p></td>
        </tr>
        <% 
  Repeat1_index=Repeat1_index+1
  Repeat1_numRows=Repeat1_numRows-1
  rs1.MoveNext()
Wend
%>
      </tbody>
    </table>
    <table class="table" >
      <tbody>
        <tr>
          <td  align="center" valign="middle"><button class="btn btn-default">
            <a href="<%=MM_moveFirst%>">第一页</a>
            </button></td>
          <td  align="center" valign="middle"><button class="btn btn-default">
            <a href="<%=MM_movePrev%>">上一页</a>
            </button></td>
          <td  align="center" valign="middle"><button class="btn btn-default">
            <a href="<%=MM_moveNext%>">下一页</a>
            </button></td>
          <td  align="center" valign="middle"><button class="btn btn-default">
            <a href="<%=MM_moveLast%>">最后一页</a></td>
        </tr>
      </tbody>
    </table>
  </div>
</div>
</body>
</html>