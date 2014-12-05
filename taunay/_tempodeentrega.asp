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
sqlVendas = "SELECT * FROM tb_venda WHERE venData = '"&Date()&"' AND tipVendaID = '1'"
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
sqlVenda = "SELECT * FROM tb_venda WHERE venData = '"&Date()&"' AND tipVendaID = '1' AND venID = (SELECT venID FROM tb_entrega WHERE tb_venda.venID = tb_entrega.venID)" 'WHERE ven
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

<body>
<p>Hora do Pedido: <%=horaPedido%> - <b><%=pedido%></b><br />
  Hora da Saida: <%=horaSaida%>  - <b><%=saida%></b><br />
  
  <br />
  Tempo medio entrega = <b><%=tempoMedio%></b>
  <br />
  <br />
  <br />
  HOJE
  <br />
 
 <%While Not rsVendas.EoF%>
  <%=rsVendas.fields.item("venID").value%> - <%=rsVendas.fields.item("venData").value%> <br />
 <%
  rsVendas.MoveNext
 Wend
 %>

</body>
</html>

<%
 Call fechaConexao()
%>
