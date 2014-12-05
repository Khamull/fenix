<%@LANGUAGE="VBSCRIPT" CODEPAGE="28592"%>



<%option explicit%>



<!--#include file="inc/inc_conexao.inc"-->



<!--#include file="inc/inc_formato_data.inc"-->



<!--#include file="inc/inc_acesso.inc" -->









<%

 Call abreConexao()

%>





<%'RECUPERA O ID DESSE FECHAMENTO DE CAIXA TRAZIDO NA URL DO NAVEGADOR

Dim caixaID

caixaID = Request.QueryString("caixaID")

%>









<%'PESQUISA NA TABELA tb_caixa AS INFORMAÇÕES PRINCIPAIS

Dim rs00

Dim sql00



set rs00 = Server.CreateObject("ADODB.Recordset")

sql00 = "SELECT * FROM tb_caixa WHERE caixaID = '"&caixaID&"'"

set rs00 = conn.execute(sql00)



'Numero da Venda Inicial / Numero da Venda Final

Dim vendaIni

Dim vendaFim



vendaIni = rs00.fields.item("vendaInicial").value

vendaFim = rs00.fields.item("vendaFinal").value

%>











<%'Pesquisa Comissão

Dim rs01

Dim sql01



set rs01 = Server.CreateObject("ADODB.Recordset")

sql01 = "SELECT SUM(venValorA) as comissao FROM tb_venda WHERE venID BETWEEN '"&vendaIni&"' AND '"&vendaFim&"'"

set rs01 = conn.execute(sql01)

%>









<%'Pesquisa quanto vendeu em DINHEIRO

Dim rs02

Dim sql02



set rs02 = Server.CreateObject("ADODB.Recordset")

sql02 = "SELECT SUM(venValorT) as totalDinheiro FROM tb_venda WHERE forPgtoID = '1' AND venID BETWEEN '"&vendaIni&"' AND '"&vendaFim&"'"

set rs02 = conn.execute(sql02)

%>









<%'Pesquisa quanto vendeu em CARTÃO

Dim rs03

Dim sql03



set rs03 = Server.CreateObject("ADODB.Recordset")

sql03 = "SELECT SUM(venValorT) as totalCartao FROM tb_venda WHERE forPgtoID = '2' AND venID BETWEEN '"&vendaIni&"' AND '"&vendaFim&"'"

set rs03 = conn.execute(sql03)

%>









<%'Pesquisa quanto vendeu em CHEQUE

Dim rs04

Dim sql04



set rs04 = Server.CreateObject("ADODB.Recordset")

sql04 = "SELECT SUM(venValorT) as totalCheque FROM tb_venda WHERE forPgtoID = '3' AND venID BETWEEN '"&vendaIni&"' AND '"&vendaFim&"'"

set rs04 = conn.execute(sql04)

%>









<%'Pesquisa quanto vendeu em OUTRAS FORMAS

Dim rs05

Dim sql05



set rs05 = Server.CreateObject("ADODB.Recordset")

sql05 = "SELECT SUM(venValorT) as totalOutras FROM tb_venda WHERE forPgtoID = '4' AND venID BETWEEN '"&vendaIni&"' AND '"&vendaFim&"'"

set rs05 = conn.execute(sql05)

%>









<%'Pesquisa quantas Vendas foram feitas por Telefone

Dim rs06

Dim sql06



set rs06 = Server.CreateObject("ADODB.Recordset")

sql06 = "SELECT SUM(venValorT) as vendasTelefone, COUNT(*) as tel FROM tb_venda WHERE tipVendaID = '1' AND venID BETWEEN '"&vendaIni&"' AND '"&vendaFim&"'"

set rs06 = conn.execute(sql06)

%>







<%'Pesquisa quantas Vendas foram feitas por Mesa

Dim rs07

Dim sql07



set rs07 = Server.CreateObject("ADODB.Recordset")

sql07 = "SELECT SUM(venValorT) as vendasMesa, COUNT(*) as mesa FROM tb_venda WHERE tipVendaID = '2' AND venID BETWEEN '"&vendaIni&"' AND '"&vendaFim&"'"

set rs07 = conn.execute(sql07)

%>









<%'Pesquisa quantas Vendas foram feitas por Balcao

Dim rs08

Dim sql08



set rs08 = Server.CreateObject("ADODB.Recordset")

sql08 = "SELECT SUM(venValorT) as vendasBalcao, COUNT(*) as balcao FROM tb_venda WHERE tipVendaID = '3' AND venID BETWEEN '"&vendaIni&"' AND '"&vendaFim&"'"

set rs08 = conn.execute(sql08)

%>











<%'SOMA das vendas de Telefone + Mesa + Balcao

Dim rs09

Dim sql09



set rs09 = Server.CreateObject("ADODB.Recordset")

sql09 = "SELECT SUM(venValorT) as totalVendas, COUNT(*) as vendas FROM tb_venda WHERE venID BETWEEN '"&vendaIni&"' AND '"&vendaFim&"'"

set rs09 = conn.execute(sql09)

%>







<%'TOTAL DE DESITENCIA

Dim rs010

Dim sql010



set rs010 = Server.CreateObject("ADODB.Recordset")

sql010 = "SELECT SUM(mesa) as mesa1, SUM(balcao) as balcao1, SUM(telefone) as telefone1, COUNT(*) as total1 FROM tb_cancelados WHERE caixaID = '"&caixaID&"'"

set rs010 = conn.execute(sql010)



Dim desisTelefone

Dim desisMesa

Dim desisBalcao

Dim desisTotal



desisTelefone = rs010.fields.item("telefone1").value

desisMesa	  = rs010.fields.item("mesa1").value

desisBalcao   = rs010.fields.item("balcao1").value

desistotal    = rs010.fields.item("total1").value

%>



<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">

<head>

<meta http-equiv="content-type" content="text/html; charset=iso-8859-1" />

<title>FECHAMENTO</title>

</head>





<body>



<table cellpadding="2" cellspacing="2" border="0" width="200">

 <tr>

  <td width="142">

   <%if(Request.QueryString("acao") = "1")Then%>

   <a href="javascript: window.close();">FECHAR</a>

   <%else%>

   <a href="pg_caixa_fechado.asp">VOLTAR</a>

   <%end if%>

  </td>

  <td width="79"></td>

  <td width="99"></td>

  <td width="24" align="center"><input type="image" src="ico/ico_printer.gif" border="0" title="Imprimir" onclick="window.print();" /></td>

 </tr>

</table>



<br />

<strong>FECHAMENTO DE CAIXA</strong>

<br />

DATA............................................<%=CDate(rs00.fields.item("data").value)%></td>

<br />

<br />

<strong>MOVIMENTO DO DIA</strong>

<br />

<br />

INICIAL........................................<%=Replace(FormatCurrency(rs00.fields.item("valorInicial").value,2),",",".")%><br />

VENDAS.......................................<%=Replace(FormatCurrency((rs00.fields.item("valorFinal").value - rs00.fields.item("valorInicial").value),2),",",".")%><br />

TOTAL CAIXA.............................<%=Replace(FormatCurrency(rs00.fields.item("valorFinal").value),",",".")%><br />

 

<%if(rs01.fields.item("comissao").value = "")Then%>

COMISSAO (-)............................. R$ 0,00<br />

<%else%>

COMISSAO (-).............................<%=Replace(FormatCurrency(rs01.fields.item("comissao").value,2),",",".")%><br />

<%end if%>



<br />

<strong>MOV. LIQUIDO..........................<%=Replace(FormatCurrency((rs00.fields.item("valorFinal").value - rs01.fields.item("comissao").value),2),",",".")%></strong><br />

<br />

<br /> 

<strong>MOVIMENTO DE CAIXA</strong><br />

<br />



<%if(rs02.fields.item("totalDinheiro").value <> "")Then%>

DINHEIRO................................... <%=Replace(FormatCurrency(rs02.fields.item("totalDinheiro").value,2),",",".")%><br />

<%end if%>

 

<%if(rs03.fields.item("totalCartao").value <> "")Then%>

CARTAO......................................<%=Replace(FormatCurrency(rs03.fields.item("totalCartao").value,2),",",".")%><br />

<%end if%>

 

<%if(rs04.fields.item("totalCheque").value <> "")Then%>

CHEQUE......................................<%=Replace(FormatCurrency(rs04.fields.item("totalCheque").value,2),",",".")%><br />

<%end if%>



<%if (rs05.fields.item("totalOutras").value <> "") Then%>

OUTROS......................................<%=Replace(FormatCurrency(rs05.fields.item("totalOutras").value,2),",",".")%><br />

<%end if%>



<br />

<strong>PEDIDOS RECEBIDOS</strong><br />

<br />

.................................VALOR / QUANTIDADE<br />



<%if (rs06.fields.item("vendasTelefone").value <> "") Then%>

TELEFONE..................................<%=Replace(FormatCurrency(rs06.fields.item("vendasTelefone").value,2),",",".")%> / <%=rs06.fields.item("tel").value%><br />

<%end if%>

 

<%if (rs07.fields.item("vendasMesa").value <> "") Then%>

MESA...........................................<%=Replace(FormatCurrency(rs07.fields.item("vendasMesa").value,2),",",".")%> / <%=rs07.fields.item("mesa").value%><br />

<%end if%>

 

<%if (rs08.fields.item("vendasBalcao").value <> "") Then%>

BALCAO......................................<%=Replace(FormatCurrency(rs08.fields.item("vendasBalcao").value,2),",",".")%> / <%=rs08.fields.item("balcao").value%><br />

<%end if%>

 

<%if (rs09.fields.item("totalVendas").value <> "") Then%>



<br />



<strong>TOTAL.........................................<%=Replace(FormatCurrency(rs09.fields.item("totalVendas").value,2),",",".")%> / <%=rs09.fields.item("vendas").value%></strong><br />

<%end if%>





<br />

<br />

 

<strong>PEDIDOS CANCELADOS</strong><br />

<br />

TELEFONE..................................<%=desisTelefone%><br />

MESA...........................................<%=desisMesa%><br />

BALCAO......................................<%=desisBalcao%><br />

<strong>TOTAL.........................................<%=desisTotal%></strong><br />



<br />

<br />

<br />
<font face="Arial, Helvetica, sans-serif" size="1"><b>NOME DO RESTAURANTE - Rua Endere&ccedil;o, 1<br />
Bairro - Cidade -- SP -

      CEP: <br />
Telefone: (11) <br />
Site: www.site.com.br</b></font><br />

<br />

</body>

</html>



<%

 Call fechaConexao

%>