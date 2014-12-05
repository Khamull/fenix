<%@LANGUAGE="VBSCRIPT" CODEPAGE="28592"%>

<%option explicit%>

<!--#include file="inc/inc_conexao.inc"-->

<%
 Call abreConexao()
%>



<%'Seleciona Através da Data
Dim rsVendas
Dim sqlVendas

set rsVendas = Server.CreateObject("ADODB.Recordset")
sqlVendas = "SELECT * FROM tb_venda WHERE venData = '"&Date()&"' AND tipVendaID = '1' OR tipVendaID = '4'"
set rsVendas = conn.execute(sqlVendas)
%>




<%'Hora em que o Produto sai da Pizzaria para ser ENTREGUE
Dim rsEntrega
Dim sqlEntrega

set rsEntrega = Server.CreateObject("ADODB.Recordset")
sqlEntrega = "SELECT * FROM tb_entrega WHERE entData = '"&Date()&"'"
set rsEntrega = conn.execute(sqlEntrega)
%>





<%'Horario da Compra
Dim rsVenda
Dim sqlVenda

set rsVenda = Server.CreateObject("ADODB.Recordset")
'sqlVenda = "SELECT * FROM tb_venda WHERE venData = '"&Date()&"' AND tipVendaID = '1' AND venID = (SELECT venID FROM tb_entrega WHERE tb_venda.venID = tb_entrega.venID ORDER BY tb_venda.venID DESC LIMIT 1)"

sqlVenda = "SELECT * FROM tb_venda WHERE venData = '"&Date()&"' AND (tipVendaID = '1' OR tipVendaID = '4') AND venID = (SELECT venID FROM tb_entrega WHERE tb_venda.venID = tb_entrega.venID ORDER BY tb_venda.venID DESC LIMIT 1) "
set rsVenda = conn.execute(sqlVenda)
%>





<%'Verifica se tem Registros
if (Not rsEntrega.EoF And Not rsVenda.EoF) Then
%>

	<%'Soma Tempo de Saida
    Dim x
    Dim horaSaida
    
    While Not rsEntrega.EoF
     x = rsEntrega.fields.item("entHoraS").value
     x = CDate(x)
     horaSaida = horaSaida + x
     horaSaida = CDate(horaSaida)
	 contador = (contador + 1)
    rsEntrega.MoveNext
    Wend
    %>
    
    <%'Soma Tempo de Compra
    Dim contador
    Dim y
    Dim horaPedido
    
    While Not rsVenda.EoF
     y = rsVenda.fields.item("venHoraF").value
     y = CDate(y)
     horaPedido = horaPedido + y
     horaPedido = CDate(horaPedido)
    rsVenda.MoveNext
    Wend
    %>
    
    <%'Atribui Valor e Soma
    Dim pedido
    Dim saida
    Dim tempoMedio
    
    pedido = CDate(horaPedido)
    saida = CDate(horaSaida)
    
    tempoMedio = (pedido - saida)
    tempoMedio = (tempoMedio/contador)
    tempoMedio = CDate(tempoMedio)
    
    %>
    
    
    
    
<% end if %>    


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-2" />
<title>Entrega</title>
</head>

<center/>

<body bgcolor="#000000">
<p><font face="Verdana, Geneva, sans-serif" color="#FFFFFF">
  Tempo m&eacute;dio entrega: </font>
  <br />
  <br />
  
  <font size="7"  face="Courier New, Courier, monospace" color="#FFFFFF">
  <strong><%=tempoMedio%></strong>
  </font></p>
<p><input type="button" onclick="javascript: window.close();" value="FECHAR" /><br />
</p>
</body>
</html>

<%
 Call fechaConexao()
%>
