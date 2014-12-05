<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%option explicit%>

<!--#include file="inc/inc_conexao.inc"-->

<!--#include file="inc/inc_formato_data.inc"-->

<!--#include file="inc/inc_acesso.inc" -->

<%
Call abreConexao()
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>SISTEM FORTE EM MÍDIA</title>
<link href="css/css1.css" rel="stylesheet" type="text/css" />

<script language="javascript" type="text/javascript">

function verForm(form1){
	
var login
var senha

Login = document.form1.Login.value;
Senha = document.form1.Senha.value;

if (Login.length < 3) {
	alert("Login ou Senha Inválido!");
	document.form1.Login.focus();
	return false;
	}
	
if (Senha.length < 3) {
	alert("Login ou Senha Inválido!");
	document.form1.Senha.focus();
	return false;

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
	    <li><a href="default.asp">Login</a></li>
	    <li><a href="pg_menu.asp?funcao=sair">Sair</a></li>
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc">Menu Principal
</div>

<%
Dim rs015
Dim sql015

set rs015 = Server.CreateObject("ADODB.Recordset")
sql015 = "SELECT * FROM tb_caixa WHERE status = 'A' "
set rs015 = conn.execute(sql015)

if rs015.EOF THEN
%>

    <table border="0" align="center" cellpadding="5" cellspacing="5">
  <tr>
    <td width="85" align="center"><a href="pg_insert_cliente.asp"><img src="ico/ico_pessoafisica.gif" width="60" height="60" border="0" class="icone" /></a></td>
    <td width="85" align="center"><a href="pg_insert_fornecedor.asp"><img src="ico/ico_pessoajuridica.gif" width="60" height="60" border="0" class="icone" /></a></td>
    <td width="85" align="center"><a href="pg_insert_funcionario.asp"><img src="ico/ico_garcon.gif" width="8" height="60" border="0" class="icone" /></a></td>
    <td width="85" align="center"><a href="pg_insert_produto.asp"><img src="ico/ico_pizza.gif" alt="" width="60" height="60" border="0" class="icone" /></a></td>
    <td width="85" align="center"><a href="pg_insert_cidade.asp"><img src="ico/ico_mapa.gif" alt="" width="60" height="60" border="0" class="icone" /></a></td>
    <td width="85" align="center"><a href="pg_insert_bairro.asp"><img src="ico/ico_casabairro.gif" alt="" width="60" height="60" border="0" class="icone" /></a></td>
    <td width="85" align="center"><a href="pg_select_usuario.asp"><img src="ico/ico_usuario.gif" alt="" width="60" height="60" class="icone" /></a></td>
  </tr>
  <tr>
    <td align="center">Cliente</td>
    <td align="center">Fornecedor</td>
    <td align="center">Funcionário</td>
    <td align="center">Produtos</td>
    <td align="center">Cidade</td>
    <td align="center">Bairro</td>
    <td align="center">Usuário</td>
  </tr>
  <tr>
    <td align="center"><a href="pg_insert_fornecedor.asp"></a><a href="pg_insert_compra.asp"><img src="ico/ico_calculadora.gif" alt="" width="60" height="60" border="0" class="icone" /></a><a href="pg_caixa.asp"></a></td>
    <td align="center"><a href="pg_insert_compra.asp"></a><a href="pg_caixa.asp"><img src="ico/ico_caixa.png" alt="" width="246" height="246" border="0" class="icone" /></a><a href="pg_insert_produto.asp"></a></td>
    <td align="center"><a href="sis_update_empresa.asp"></a></td>
    <td align="center">&nbsp;</td>
    <td align="center"><a href="pg_insert_funcionario.asp"></a></td>
    <td align="center"><a href="pg_insert_cidade.asp"></a></td>
    <td align="center"><a href="pg_insert_bairro.asp"></a></td>
  </tr>
  <tr>
    <td align="center">Comprar</td>
    <td align="center">Caixa</td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
  </tr>
  </table>
    
<%else%>

 <table border="0" align="center" cellpadding="5" cellspacing="5">
  <tr>
    <td width="85" align="center"><a href="pg_insert_cliente.asp"><img src="ico/ico_pessoafisica.gif" width="60" height="60" border="0" class="icone" /></a></td>
    <td width="85" align="center"><a href="pg_menu_pedidos.asp"><img src="ico/ico_lembrete.gif" width="60" height="60" border="0" class="icone" /></a></td>
    <td width="85" align="center"><a href="pg_select_pedidos_telefone.asp"><img src="ico/ico_telefone.gif" width="60" height="60" class="icone" /></a></td>
    <td width="85" align="center"><a href="pg_select_pedidos_mesa.asp"><img src="ico/ico_mesa.gif" width="60" height="60" class="icone" /></a></td>
    <td width="85" align="center"><a href="pg_select_pedidos_balcao.asp"><img src="ico/ico_cesta.gif" width="60" height="60" class="icone" /></a></td>
    <td width="85" align="center"><a href="pg_select_entrega.asp"><img src="ico/ico_moto.gif" alt="" width="60" height="60" border="0" class="icone" /></a></td>
    <td width="85" align="center"><a href="pg_select_usuario.asp"><img src="ico/ico_usuario.gif" alt="" width="60" height="60" class="icone" /></a><a href="pg_select_entrega.asp"></a></td>
  </tr>
  <tr>
    <td align="center">Cliente</td>
    <td align="center">Pedido</td>
    <td align="center">Venda Tel</td>
    <td align="center">Venda Mesa</td>
    <td align="center">Venda Balcão</td>
    <td align="center">Entregas</td>
    <td align="center">Usuário</td>
  </tr>
  <tr>
    <td align="center"><a href="pg_insert_fornecedor.asp"><img src="ico/ico_pessoajuridica.gif" width="60" height="60" border="0" class="icone" /></a></td>
    <td align="center"><a href="pg_insert_produto.asp"><img src="ico/ico_pizza.gif" width="60" height="60" border="0" class="icone" /></a></td>
    <td align="center"><a href="pg_insert_compra.asp"><img src="ico/ico_calculadora.gif" width="60" height="60" border="0" class="icone" /></a></td>
    <td align="center"><img src="ico/ico_money.gif" width="60" height="60" border="0" class="icone" /></td>
    <td align="center"><a href="pg_insert_funcionario.asp"><img src="ico/ico_garcon.gif" width="8" height="60" border="0" class="icone" /></a></td>
    <td align="center"><a href="pg_insert_cidade.asp"><img src="ico/ico_mapa.gif" width="60" height="60" border="0" class="icone" /></a></td>
    <td align="center"><a href="pg_insert_bairro.asp"><img src="ico/ico_casabairro.gif" width="60" height="60" border="0" class="icone" /></a></td>
  </tr>
  <tr>
    <td align="center">Fornecedor</td>
    <td align="center">Produtos</td>
    <td align="center">Comprar</td>
    <td align="center">Vendas</td>
    <td align="center">Funcionário</td>
    <td align="center">Cidade</td>
    <td align="center">Bairro</td>
  </tr>
  <tr>
    <td align="center"><a href="pg_caixa.asp"><img src="ico/ico_caixa.png" alt="" width="246" height="246" border="0" class="icone" /></a></td>
    <td align="center"><a href="sis_update_home.asp"><img src="ico/ico_home.jpg" alt="Home" width="246" height="246" border="0" class="icone" /></a></td>
    <td align="center"><a href="sis_update_empresa.asp"><img src="ico/ico_empresa.png" alt="Home" width="246" height="246" border="0" class="icone" /></a></td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td align="center"><a href="pg_insert_usuario.asp"></a></td>
  </tr>
  <tr>
    <td align="center">Caixa</td>
    <td align="center">Home</td>
    <td align="center">Empresa</td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
  </tr>
    </table>

<%end if%>

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
Call fechaConexao()
%>