<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/conn.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_conn_STRING
Recordset1_cmd.CommandText = "SELECT * FROM pizzaria.tb_bairro" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Documento sem título</title>
</head>

<body>
<table border="1" cellpadding="2" cellspacing="2">
  <tr>
    <td>baiID</td>
    <td>baiData</td>
    <td>usuLogin</td>
    <td>baiNome</td>
    <td>baiFrete</td>
    <td>baiAtivo</td>
    <td>cidID</td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) %>
    <tr>
      <td><%=(Recordset1.Fields.Item("baiID").Value)%></td>
      <td><%=(Recordset1.Fields.Item("baiData").Value)%></td>
      <td><%=(Recordset1.Fields.Item("usuLogin").Value)%></td>
      <td><%=(Recordset1.Fields.Item("baiNome").Value)%></td>
      <td><%=(Recordset1.Fields.Item("baiFrete").Value)%></td>
      <td><%=(Recordset1.Fields.Item("baiAtivo").Value)%></td>
      <td><%=(Recordset1.Fields.Item("cidID").Value)%></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
