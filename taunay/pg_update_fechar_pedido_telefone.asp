<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>



<%option explicit%>



<!--#include file="inc/inc_conexao.inc"-->



<!--#include file="inc/inc_formato_data.inc"-->



<!--#include file="inc/inc_acesso.inc" -->



<%

call abreConexao()

%>



<%

Dim acao

Dim venID

Dim venLocalidade

Dim venObs

Dim venValorS

Dim venValorF

Dim venValorD

Dim venValorA

Dim venValorT

Dim forPgtoID

Dim venValorR

Dim venValorTc



venID			= Request.QueryString("venID")

venLocalidade	= Request.Form("venLocalidade")

venObs			= Request.Form("venObs")

venValorS		= Request.Form("venValorS")

venValorF		= Request.Form("venValorF")

venValorD		= Request.Form("venValorD")

venValorA		= Request.Form("venValorA")

venValorT		= Request.Form("venValorT")

forPgtoID		= Request.Form("forPgtoID")

venValorR		= Request.Form("venValorR")

venValorTc		= Request.Form("venValorTc")



acao = request.querystring("acao")

if (acao = "fechar") then



dim rs8

dim sql8

set rs8 = server.CreateObject("adodb.recordset")

sql8 = "UPDATE tb_venda SET venLocalidade = '"&venLocalidade&"', venObs='"&venObs&"', venValorS='"&venValorS&"', venValorF='"&venValorF&"', venValorD='"&venValorD&"', venValorA='"&venValorA&"', venValorT='"&venValorT&"', forPgtoID='"&forPgtoID&"', venValorR='"&venValorR&"', venValorTc='"&venValorTc&"', staID = 10 WHERE venID = '"&venID&"'"

sql8 = conn.execute(sql8)



'---------------------------------------------------------------------------------------------|

	'ATUALIZA ESTOQUE DO PRODUTO COM OS NOVOS ITENS VENDIDOS

	Dim rs010

	Dim sql010

	

	Dim rs011

	Dim sql011

	

	Dim rs012

	Dim sql012

	

	

	set rs010 = Server.CreateObject("ADODB.Recordset")

	sql010 = "SELECT * FROM tb_itemvenda WHERE venID = '"&venID&"'"

	set rs010 = conn.execute(sql010)





	Dim prodID

	Dim estoqueAtual

	Dim compra

	Dim estoqueFinal

	Dim proPrecoCusto1

	

	While Not rs010.EoF

	

	 prodID = rs010.fields.item("proID").value

	 compra = rs010.fields.item("iteQtde").value

	 proPrecoCusto1 = rs010.fields.item("itePreco").value

	 

	 'CONVERTENDO PARA FLOAT NO PADRÃO DO BANDO DE DADOS

	 

	 set rs012 = Server.CreateObject("ADODB.Recordset")

	 sql012 = "SELECT * FROM tb_produto WHERE proID = '"&prodID&"'"

	 set rs012 = conn.execute(sql012)

	 

	 estoqueAtual = rs012.fields.item("proEstoque").value

	 

	 estoqueFinal = (estoqueAtual - compra)

	

	 set rs011 = Server.CreateObject("ADODB.Recordset")

	 sql011 = "UPDATE tb_produto SET proEstoque = '"&estoqueFinal&"' WHERE proID = '"&prodID&"' AND tipID <> '1' "

	 set rs011 = conn.execute(sql011)

	

	rs010.MoveNext

	Wend



'---------------------------------------------------------------------------------------------|





response.redirect("pg_print_pedido.asp?venID="&venID)



end if



if(venID = "") then

response.redirect("pg_select_pedidos_telefone.asp")

end if

%>



<%

Dim rs00

Dim sql00

set rs00 = server.CreateObject("adodb.recordset")

sql00 = "SELECT tb_itemvenda.iteObs, tb_itemvenda.iteID, tb_itemvenda.venID, tb_itemvenda.proID, tb_itemvenda.iteqtde, tb_itemvenda.itepreco, tb_itemvenda.itesubtotal, tb_produto.proDescricao, tb_produto.proCodEmpresa, tb_produto.proUnidade FROM tb_itemvenda INNER JOIN tb_produto ON tb_produto.proID = tb_itemvenda.proID WHERE tb_itemvenda.venID = '"&venID&"' GROUP BY tb_itemvenda.iteID"

set rs00 = conn.execute(sql00)

%>



<%

Dim rs01

Dim sql01

set rs01 = server.CreateObject("adodb.recordset")

sql01 = "SELECT SUM(iteSubTotal) AS subTotal FROM tb_itemvenda WHERE venID='"&venID&"'"

set rs01 = conn.execute(sql01)

%>



<%

'SELECIONA DETALHES DO PEDIDO

Dim rs02

Dim sql02

set rs02 = Server.CreateObject("ADODB.Recordset")

sql02 = "SELECT "

sql02 = sql02 & "tb_tipovenda.tipVendaID, tb_tipovenda.tipVendaDescricao, " 'TB_TIPOVENDA

sql02 = sql02 & "tb_cliente.cliID, tb_cliente.cliNome, tb_cliente.cliTelefone, " 'TB_CLIENTE 1°

sql02 = sql02 & "tb_cliente.cliEndereco, tb_cliente.baiID, tb_cliente.cidID, tb_cliente.cliPreferencias, "	 'TB_CLIENTE 2°

sql02 = sql02 & "tb_bairro.baiID, tb_bairro.baiNome, tb_bairro.baiFrete, "	 'TB_BAIRRO

sql02 = sql02 & "tb_cidade.cidID, tb_cidade.cidNome, "	 'TB_CIDADE

sql02 = sql02 & "tb_numerovenda.numerovendaID, tb_numerovenda.numerovenda as numero, tb_numerovenda.venID, " 'TB_NUMEROVENDA

sql02 = sql02 & "tb_venda.* " 'TB_VENDA

sql02 = sql02 & "FROM tb_venda " 'TABELA PRINCIPAL

sql02 = sql02 & "INNER JOIN tb_tipovenda ON tb_tipovenda.tipVendaID = tb_venda.tipVendaID " 'INNER JOIN com TIPO DE VENDA

sql02 = sql02 & "INNER JOIN tb_cliente ON tb_cliente.cliID = tb_venda.cliID " 'INNER JOIN com CLIENTE

sql02 = sql02 & "INNER JOIN tb_bairro ON tb_bairro.baiID = tb_cliente.baiID " 'INNER JOIN com BAIRRO

sql02 = sql02 & "INNER JOIN tb_cidade ON tb_cidade.cidID = tb_cliente.cidID " 'INNER JOIN com CIDADE

sql02 = sql02 & "INNER JOIN tb_numerovenda ON tb_numerovenda.venID = tb_venda.venID " 'INNER JOIN com NUMERO DA VENDA

sql02 = sql02 & "WHERE tb_venda.venID = '"&venID&"'" 'CONDIÇÃO

set rs02 = conn.execute(sql02)

%>





<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">

<head>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

<title>SISTEM FORTE EM MÍDIA</title>

<link href="css/css1.css" rel="stylesheet" type="text/css" />



<script language="javascript" type="text/javascript">

function calcular()

{



var venValorS		=	parseFloat(document.form1.venValorS.value);

var venValorF		=	parseFloat(document.form1.venValorF.value);

var venValorD		=	parseFloat(document.form1.venValorD.value);

var venValorA		=	parseFloat(document.form1.venValorA.value);



if (document.form1.taxa.checked == true) {

venValorA 			= (venValorS * 10/100)

document.form1.venValorA.value = venValorA.toFixed(2);

}

else 

{

document.form1.venValorA.value = "0.00";	

}

//var venValorT		=	parseFloat(document.form1.venValorT.value);

//venValorT 			= ((venValorS + venValorF + venValorA) - venValorD)

var venValorT = ((venValorS + venValorF + venValorA) - venValorD); 

document.form1.venValorT.value = venValorT.toFixed(2);

 

var forPgtoID		=	parseFloat(document.form1.venValorD.value);

var venValorR		=	parseFloat(document.form1.venValorR.value);

var venValorTc		=	parseFloat(document.form1.venValorTc.value);



venValorTc 			= (venValorR - venValorT)



if(venValorTc <= "0.00"){

	venValorTc = "0.00";

}

else{

	venValorTc = venValorTc;

}



document.form1.venValorTc.value = venValorTc.toFixed(2);

document.form1.venValorT.value = venValorT.tofixed(2);

}



function verForm(form1) {

if(confirm("Tem certeza que deseja fechar o pedido?"))

{

	form1.submit();

}

else{

	return false;

}

}



</script>



</head>

<body onload="calcular()">

<!--LAYOUT-->

<div id="container">

<!-- -->

<div id="topo"></div>

<div id="tituloBar"><img src="img/img_titulo_mp.gif" width="200" height="30" /></div>

<div id="corpo">

<!-- -->

<div id="areaConteudo">



	<div id="areaMenuVerfical">

	<div style="height:25px; line-height:25px; background:#ccc">Menu</div>    

	  <ul>

	    <li><a href="pg_menu.asp">Menu Principal</a></li>

        <li><a href="javascript: window.open('pg_tempodeentrega.asp', 'Entrega' , 'height = 180 , width = 320');">Tempo Medio Entrega</a></li      

	  ></ul>

	</div>

	<div id="areaPrincipal">

    <div style="height:25px; line-height:25px; background:#ccc"></div>

    <table width="96%" border="0" align="left" cellpadding="3" cellspacing="3">

  <tr>

    <td width="50" align="center"><img src="ico/ico_calculadora.gif" width="60" height="60" class="icone" /></td>

    <td width="751" align="center" class="titulo">FECHAR O PEDIDO</td>

    </tr>

  <tr>

    <td colspan="2" align="center"><form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?venID=<%=venID%>&acao=fechar" onsubmit="return verForm(this)">

      <table width="777" border="0" cellpadding="1" cellspacing="1">

        <tr>

          <td height="25" align="right">Data:&nbsp;</td>

          <td height="25" align="left"><%=rs02.fields.item("venData").value%></td>

          <td height="25" align="right">Atendente:&nbsp;</td>

          <td height="25" align="left"><input name="usuLogin" type="text" id="usuLogin" size="40" maxlength="30" value="<%=rs02.fields.item("usuLogin").value%>" readonly="readonly" /></td>

        </tr>

        <tr>

          <td width="109" height="25" align="right">Tipo de Pedido:</td>

          <td width="267" height="25" align="left"><label>

            <input name="tipVenDescr" type="text" id="tipVenDescr" size="20" maxlength="20" value="<%=rs02.fields.item("tipVendaDescricao").value%>" readonly="readonly" />

          </label></td>

          <td width="110" height="25" align="right">Nº do Pedido:</td>

          <td width="263" height="25" align="left">

          <input name="numerovenda" type="text" id="numerovenda" size="10" maxlength="10" value="<%=rs02.fields.item("numero").value%>" readonly="readonly" />

          <input name="venID" type="hidden" id="venID" size="10" maxlength="10" value="<%=rs02.fields.item("venID").value%>" readonly="readonly" />

          </td>

        </tr>

        <tr>

          <td height="25" align="right">Cliente:&nbsp;</td>

          <td height="25" align="left"><input name="cliNome" type="text" id="cliNome" size="40" maxlength="30" value="<%=rs02.fields.item("cliNome").value%>" readonly="readonly" /></td>

          <td height="25" align="right">Telefone:&nbsp;</td>

          <td height="25" align="left"><input name="cliTelefone" type="text" id="cliTelefone" size="10" maxlength="10" value="<%=rs02.fields.item("cliTelefone").value%>" readonly="readonly" /></td>

        </tr>

        <tr>

          <td height="25" align="right">Local de Entrega:&nbsp;</td>

          <td height="25" align="left"><label>

            <textarea name="venLocalidade" id="venLocalidade" cols="45" rows="3"><%=rs02.fields.item("cliEndereco").value%> - <%=rs02.fields.item("BaiNome").value%> - <%=rs02.fields.item("cidNome").value%></textarea>

          </label></td>

          <td height="25" align="right">Observações:&nbsp;</td>

          <td height="25" align="left"><textarea name="venObs" id="venObs" cols="45" rows="3"><%=rs02.fields.item("cliPreferencias").value%></textarea></td>

        </tr>

        <tr>

          <td height="25" colspan="4" align="center" class="titulo">DETALHE DO PEDIDO</td>

        </tr>

        <tr>

          <td height="25" colspan="4" align="center" class="titulo">

          <table width="100%" border="0" align="center" cellpadding="2" cellspacing="2">

            <tr class="caixaPreta">

              <td width="7%">COD.</td>

              <td width="61%">DESCRIÇÃO</td>

              <td width="4%" align="left">UN</td>

              <td width="6%" align="left">QTDE</td>

              <td width="11%" align="left">VALOR</td>

              <td width="11%" align="left">TOTAL</td>

            </tr>

            

            <%'VARIAVEL QUE IRÁ APRESENTAR O TOTAL DA COMPRA

			 Dim totaldaCompra

			%>

            

            <%			

			if not rs00.eof then

			do while not rs00.eof

			%>

            

            

            <tr class="textoComum">

              <td align="left"><%=rs00.fields.item("proCodEmpresa").value%></td>

              <td align="left"><%=rs00.fields.item("proDescricao").value%><font color="#990000">&nbsp;&nbsp;<%=rs00.fields.item("iteObs").value%></font></td>

              <td align="left"><%=rs00.fields.item("proID").value%></td>

              <td align="left">

			  

			  <%

			   Dim x

			   x = rs00.fields.item("iteqtde").value

			   if (x = "") then

			   x = "0.00"

			   Response.write(x)

			   else

			   x = Replace(FormatNumber(x),",",".")

			   Response.write(x)

			   end if

			   %>

              

              </td>

              <td align="left">

			  			  <%

			   Dim y

			   y = rs00.fields.item("itePreco").value

			   if (y = "") then

			   y = "0.00"

			   Response.write(y)

			   else

			   y = Replace(FormatNumber(y),",",".")

			   Response.write(y)

			   end if

			   %>

               </td>

              <td align="left">

			  			  <%

			   Dim z

			   z = rs00.fields.item("iteSubTotal").value

			   if (z = "") then

			   z = "0.00"

			   Response.write(z)

			   else

			   z = Replace(FormatNumber(z),",",".")

			   Response.write(z)

			   end if

			   %>

              </td>

            </tr>

            

            <%'SOMA TODOS OS SUBTOTAIS POR PRODUTO

			 totaldaCompra = totaldaCompra + rs00.fields.item("iteSubTotal").value

			%>

            

            <%

			rs00.moveNext

			loop

			end if

			%>

            

            <%'SOMA COM O FRETE

			 totaldaCompra = totaldaCompra + rs02.fields.item("baiFrete").value

			 totaldaCompra = Replace(totaldaCompra,",",".")

			%>

            

          </table></td>

          </tr>

        <tr>

          <td height="25" colspan="4" align="right"><table width="100%" border="0" align="right" cellpadding="2" cellspacing="2">

            <tr>

                <td width="531" rowspan="8" align="right" valign="top" bgcolor="#FFFFFF"><img src="img/fundo2.png" width="447" height="149" /></td>

                <td width="136" align="right" bgcolor="#eeeeee">Sub Total:</td>

                <td width="96" align="right" bgcolor="#eeeeee"><input name="venValorS" type="text" class="caixaPreta" id="venValorS" value="<%

				Dim s

				s = rs01.fields.Item("subTotal").value

				if(s = "") then				

				s = "0.00"

				Response.write(s)

				else

				s = Replace(FormatNumber(s),",",".")

				Response.write(s)

				end if

				



				%>" size="10" maxlength="10" readonly="readonly" /></td>

                </tr>

              <tr>

                <td align="right" bgcolor="#eeeeee">Frete:</td>

                <td align="right" bgcolor="#eeeeee"><input name="venValorF" type="text" id="venValorF" value="<%=Replace(rs02.fields.item("baiFrete").value,",",".")%>" size="10" maxlength="10" onblur="calcular()" onkeydown="calcular()" onfocus="calcular" onkeyup="return verificaNumero()"/></td>

                </tr>

              <tr>

                <td align="right" bgcolor="#eeeeee">Desconto:</td>

                <td align="right" bgcolor="#eeeeee"><input name="venValorD" type="text" id="venValorD" value="0.00" size="10" maxlength="10" onblur="calcular()" onkeydown="calcular()" onfocus="calcular"/></td>

                </tr>

              <tr>

                <td align="right" bgcolor="#eeeeee">Acréssimo 10%:

                  <input name="taxa" type="checkbox" id="taxa" value="true" onclick="calcular()" onchange="calcular()" onblur="calcular()" /></td>

                <td align="right" bgcolor="#eeeeee"><input name="venValorA" type="text" class="caixaPreta" id="venValorA" onfocus="calcular" onblur="calcular()" onkeydown="calcular()" value="0.00" size="10" maxlength="10" readonly="readonly"/></td>

                </tr>

              <tr>

                <td align="right" bgcolor="#eeeeee">Total:</td>

                <td align="right" bgcolor="#eeeeee"><input name="venValorT" type="text" class="caixaPreta" id="venValorT" size="10" maxlength="10" readonly="readonly" onfocus="calcular" value="0.00"/></td>

              </tr>

              <tr>

                <td align="right" bgcolor="#eeeeee">Forma de Pgto:&nbsp;</td>

                <td align="right" bgcolor="#eeeeee"><select name="forPgtoID" id="forPgtoID"/>

                  <option value="1" selected="selected">Dinheiro</option>

                  <option value="2">Cartão </option>

                  <option value="3">Cheque</option>

                  <option value="4">Outra</option>

                  </select>

                </td>

                </tr>

              <tr>

                <td align="right" bgcolor="#eeeeee">Recebido:</td>

                <td align="right" bgcolor="#eeeeee"><input name="venValorR" type="text" id="venValorR" value="0.00" size="10" maxlength="10" onblur="calcular()" onkeydown="calcular()"  onfocus="calcular"/></td>

              </tr>

              <tr>

                <td align="right" bgcolor="#eeeeee">Troco:</td>

                <td align="right" bgcolor="#eeeeee"><input name="venValorTc" type="text" class="caixaPreta" id="venValorTc" value="0.00" size="10" maxlength="10"  onfocus="calcular"/></td>

                </tr>

            </table></td>

          </tr>

        <tr>

          <td height="25" align="right">&nbsp;</td>

          <td height="25">&nbsp;</td>

          <td height="25" align="right"><input name="usuLogin" type="hidden" id="usuLogin" value="<%=Session("usuLogin")%>" /></td>

          <td height="25" align="center">

           <input name="concluir" type="submit" class="botao" id="concluir" value="Concluir" style="color:white" />&nbsp;

           <input name="imprimir" type="button" onclick="javascript: window.open('pg_print_pedido.asp?venID='+<%=rs02.fields.item("venID").value%>, 'Imprimir', 'width=600px, height=650px');" class="botao" id="imprimir" value="Imprimir" style="color:white" />

          </td>

        </tr>

      </table>

    </form></td>

  </tr>

  </table>

	</div>

</div>

<!-- -->

</div>

<div id="rodape"><br /><!--#include file="inc/inc_status.inc"--><br /></div>

</div>

<!--FIM DO LAYOUT-->



</body>

</html>

<%

call fechaConexao()

%>