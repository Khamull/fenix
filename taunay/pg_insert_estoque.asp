<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%option explicit%>

<!--#include file="inc/inc_conexao.inc"-->

<!--#include file="inc/inc_formato_data.inc"-->

<!--#include file="inc/inc_acesso.inc" -->

<%
call abreConexao()
%>

<%'PESQUISA TODOS OS PRODUTOS CADASTRADOS NO SISTEMA
Dim rs01
Dim sql01

set rs01 = Server.CreateObject("ADODB.Recordset")
sql01 = "SELECT tb_fornecedor.forID, tb_fornecedor.forNome, tb_produto.* FROM tb_produto INNER JOIN tb_fornecedor ON tb_fornecedor.forID = tb_produto.forID WHERE tb_produto.tipID = '2' ORDER BY tb_produto.proID DESC"
set rs01 = conn.execute(sql01)
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>SISTEM FORTE EM MÍDIA</title>
<link href="css/css1.css" rel="stylesheet" type="text/css" />

<script language="javascript" type="text/javascript">

function verForm(form1){
	
var forID		= document.form1.forID.value;
var comDataC	= document.form1.comDataC.value;
var comFormPgto = document.form1.comFormPgto.value;
var comNumParc  = document.form1.comNumParc.value;
var comDataV    = document.form1.comDataV.value;
var comValorF   = document.form1.comValorF.value;
var comVendedor	= document.form1.comVendedor.value;

if (forID == ""){
	alert("Favor selecionar o Fornecedor!");
	document.form1.forID.focus();
	return false;
	}

if (comFormPgto == " "){
	alert("Favor selecionar a Forma de Pagamento!");
	document.form1.comFormPgto.focus();
	return false;
	}

if (comNumParc == ""){
	alert("Favor informar o numero de parcelas!");
	document.form1.comNumParc.focus();
	return false;
	}

if (comDataV == ""){
	alert("Favor informar o data de vencimento!");
	document.form1.comDataV.focus();
	return false;
	}
	
}
function Excluir(id)
{
	if(confirm("Tem certeza que deseja excluir o pedido de compra?\nNão será possível recuperar os dados!"))
	{
		window.location.href = "pg_insert_compra.asp?acao=excluir&id="+id;
	}
}
</script>

</head>
<body>
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
        <li><a href="pg_insert_compra.asp">Comprar</a></li>               
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc"></div>
    <table width="96%" border="0" align="left" cellpadding="3" cellspacing="3">
  <tr>
    <td width="50" align="center"><img src="ico/ico_estoque.gif" width="60" height="60" class="icone" /></td>
    <td width="751" align="center" class="titulo">ESTOQUE POR PRODUTO</td>
    </tr>
  <tr>
    <td colspan="2" align="center">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" align="center"><table width="777" border="0" cellpadding="2" cellspacing="2">
      <tr class="textoBranco">
        <td width="179" height="20" align="left" bgcolor="#9E231B">Produto</td>
        <td width="210" height="20" align="left" bgcolor="#9E231B">Fornecedor</td>
        <td width="101" height="20" align="left" bgcolor="#9E231B">Quantidade</td>
        <td width="110" height="20" align="left" bgcolor="#9E231B">Preço de Custo</td>
        <td width="117" height="20" align="left" bgcolor="#9E231B">Preço de Venda</td>
        </tr>
      <%
	   Dim cor(2)
	   Dim i
	   
	   cor(0) = "#DDDDDD"
	   cor(1) = "#FFFFFF"
	   
	   i = 0
	  %>
      <% While not rs01.EoF%>
      <tr bgcolor="<%=cor(i)%>" >
        <td align="left"><%=rs01.fields.item("proDescricao").value%></td>
        <td align="left"><%=rs01.fields.item("forNome").value%></td>
        <td align="left"><%=rs01.fields.item("proEstoque").value%></td>
        <td align="left"><%=FormatCurrency(rs01.fields.item("proPrecoCusto").value)%></td>
        <td align="left"><%=rs01.fields.item("proPrecoVenda").value%></td>
        </tr>
      
      <%
	   if (i = 0)Then
	    i = 1
	   else
	    i = 0
	   end if
	  %>
      
      <%
	   rs01.MoveNext
	  Wend
	  %>
    </table></td>
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
call FechaConexao()
%>