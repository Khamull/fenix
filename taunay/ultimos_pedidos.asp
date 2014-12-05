<%@LANGUAGE="VBSCRIPT" CODEPAGE="28592"%>

<%option explicit%>

<!--#include file="inc/inc_conexao.inc"-->

<%
 Call abreConexao()
%>


<%'RECUPERA ID DO CLIENTE
 Dim cliID

 cliID = Request.QueryString("cliID")
%>


<%'SELECIONA CLIENTE
Dim rs00
Dim sql00

set rs00 = Server.CreateObject("ADODB.Recordset")
sql00 = "SELECT cliNome FROM tb_cliente WHERE cliID = '"&cliID&"'"
set rs00 = conn.execute(sql00)
%>

<%'SELECIONA PEDIDOS
Dim rs01
Dim sql01


set rs01 = Server.CreateObject("ADODB.Recordset")
sql01 = "SELECT * FROM tb_venda WHERE cliID = '"&cliID&"' ORDER BY tb_venda.venID DESC LIMIT 5"
set rs01 = conn.execute(sql01)
%>

<%
Dim rs02
Dim sql02
Dim venID

set rs02 = Server.CreateObject("ADODB.Recordset")
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-2" />
<title>Ultimos Pedidos</title>
</head>

<body background="img/fundo.jpg">

<table border="0" cellpadding="2" cellspacing="0" width="350">

 <tr>
  <td colspan="3" align="center">
  <b>
  <font face="Arial, Helvetica, sans-serif" size="3">
  CLIENTE </font>
  <font face="Arial, Helvetica, sans-serif" size="3" color="#FF0000">
  <%=rs00.fields.item("cliNome").value%>
  </font>
  </b>
  </td>
 </tr>

<% While Not rs01.EoF %>
 <tr bgcolor="#000000">
  <td width="95">
  <font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="2">
  <b><%=rs01.fields.item("venData").value%></b>
  </font>
  </td>
  <td width="147"></td>
  <td width="130">
  <font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="2">
  <b>PEDIDO: <%=rs01.fields.item("venID").value%></b>
  </font>
  </td>
 </tr>
 
 <%
  venID = rs01.fields.item("venID").value
  
  sql02 = "SELECT tb_produto.proID, tb_produto.proDescricao, tb_itemvenda.proID, tb_itemvenda.venID FROM tb_itemvenda INNER JOIN tb_produto ON tb_produto.proID = tb_itemvenda.proID WHERE tb_itemvenda.venID = '"&venID&"'"
  set rs02 = conn.execute(sql02)
 %>
 
 <% While Not rs02.EoF %>
 
 <tr>
  <td colspan="3" align="left" bgcolor="#eeeeee">
  <font face="Arial, Helvetica, sans-serif" size="2" color="#000066">
  <%=rs02.fields.item("proDescricao").value%>
  </font>
  </td>
 </tr>
 
 <%
  rs02.MoveNext
 Wend
 %>
 
 <tr>
  <td colspan="3" height="10"></td>
 </tr>
 
 <%
  rs01.MoveNext
 Wend
 %> 
</table>

</body>
</html>
