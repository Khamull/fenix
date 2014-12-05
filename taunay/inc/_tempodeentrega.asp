
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
    Dim x01
    Dim horaSaida
    
    While Not rsEntrega.EoF
     x01 = rsEntrega.fields.item("entHoraS").value
     x01 = CDate(x01)
     horaSaida = horaSaida + x01
     horaSaida = CDate(horaSaida)
	 contador = (contador + 1)
    rsEntrega.MoveNext
    Wend
    %>
    
    <%'Soma Tempo de Compra
    Dim contador
    Dim y01
    Dim horaPedido
    
    While Not rsVenda.EoF
     y01 = rsVenda.fields.item("venHoraF").value
     y01 = CDate(y01)
     horaPedido = horaPedido + y01
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


<%
 Call fechaConexao()
%>
