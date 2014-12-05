<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%option explicit%>

<!--#include file="inc/inc_conexao.inc"-->

<!--#include file="inc/inc_formato_data.inc"-->

<!--#include file="inc/inc_acesso.inc" -->




<%
call abreConexao()
%>



<%
Dim venID

venID = Request.QueryString("venID")
%>


<%'----------------------- VOLTA OS VALORES PARA O ESTOQUE ------------------------------------------

'Ao entrar nessa tela os produtos referentes a essa venda voltarao para o estoque,
'Sendo assim o usuario pode excluir o pedido ou fechalo novamente sem q interfira no estoque atual.
'Isso é necessário pois todas as vezes que pedido é fechado, automaticamnete o estoque é dado baixa.


'SELECIONA ITENS DO PEDIDO
Dim rs011
Dim sql011

set rs011 = Server.CreateObject("ADODB.Recordset")
sql011 = "SELECT * FROM tb_itemvenda WHERE venID = '"&venID&"'"
set rs011 = conn.execute(sql011)

'VERIFICA QUANTO TEM EM ESTOQUE
Dim rs012
Dim sql012
Dim prodID
Dim estoqueAnterior
Dim decrescimo
Dim estoqueAtual

While Not rs011.EoF

prodID = rs011.fields.item("proID").value

set rs012 = Server.CreateObject("ADODB.Recordset")
sql012 = "SELECT * FROM tb_produto WHERE proID = '"&prodID&"'"
set rs012 = conn.execute(sql012)

 '--- Calcula Direfença ----
 estoqueAnterior = rs012.fields.item("proEstoque").value
 decrescimo = rs011.fields.item("iteQtde").value
 estoqueAtual = (estoqueAnterior + decrescimo)
 estoqueAtual = CInt(estoqueAtual)
 '--------------------------
 
 
	'-------------- DEVOLVE O PRODUTO AO ESTOQUE -------------------------
	 Dim rs013
	 Dim sql013
	 
	 set rs013 = Server.CreateObject("ADODB.Recordset")
	 sql013 = "UPDATE tb_produto SET proEstoque = '"&estoqueAtual&"' WHERE proID = '"&prodID&"'"
	 set rs013 = conn.execute(sql013)
	'----------------------------------------------------------------------

rs011.MoveNext
Wend

'-----------------------------------------------------------------------------------------------------

Response.Redirect("pg_insert_itemVendaBalcao1.asp?venID="&venID)
%>

<%
 Call fechaConexao()
%>
