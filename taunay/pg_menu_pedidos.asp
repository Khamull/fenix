<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%option explicit%>



<!--#include file="inc/inc_conexao.inc"-->



<!--#include file="inc/inc_formato_data.inc"-->



<!--#include file="inc/inc_acesso.inc" -->



<%

call abreConexao()

%>



<%

Dim imagem

Dim tipo

Dim cliTelefone

Dim msg



imagem = "ico_telefone.gif"



cliTelefone = Request.QueryString("cliTelefone")

tipo = request.querystring("tipo")



if (cliTelefone <> "") then



	Dim rs00

	Dim sql00

	set rs00 = server.CreateObject("ADODB.Recordset")

	sql00 = "SELECT cliID, cliNome, cliTelefone, cliAtivo, cliEndereco FROM tb_cliente WHERE cliTelefone = '"&cliTelefone&"' OR cliTelefone2 = '"&cliTelefone&"' OR cliTelefone3 = '"&cliTelefone&"' OR cliTelefone4 = '"&cliTelefone&"' OR cliTelefone5 = '"&cliTelefone&"'"

	set rs00 = conn.execute(sql00)


	if (not rs00.eof) then

	

	dim telefone

	dim nome

	dim situacao

	dim endereco

	dim cliID

	dim incluir

	

	cliID = rs00.fields.item("cliID").value

	'telefone = rs00.fields.item("cliTelefone").value

	telefone = request.QueryString("cliTelefone")

	nome = rs00.fields.item("cliNome").value

	situacao = rs00.fields.item("cliAtivo").value

	endereco = rs00.fields.item("cliEndereco").value

	

	incluir = "<a href='pg_update_cliente2.asp?cliID="&cliID&"'><b><u>Incluir Telefone</u></b></a>"

	

	tipo = "telefone"

	imagem = "ico_confirmado.gif"



	else

	

	tipo = "telefone"

	msg = "Cliente não encontrado!<br><b><a href='pg_insert_cliente.asp?cliTelefone="&Request.Querystring("cliTelefone")&"'>Clique aqui para cadastrar!</a></b>"

	imagem = "ico_cancelado.gif"

	

	end if



end if



%>



<%



	Dim rs01

	Dim sql01

	set rs01 = server.CreateObject("ADODB.Recordset")

	sql01 = "SELECT * FROM tb_mesa WHERE mesAtiva = 'S' AND NOT EXISTS(SELECT * FROM tb_venda WHERE tb_venda.staID = 1 AND tb_venda.mesID = tb_mesa.mesID)"

	set rs01 = conn.execute(sql01)

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">

<head>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

<title>SISTEM FORTE EM MÍDIA</title>

<link href="css/css1.css" rel="stylesheet" type="text/css" />



<script language="javascript" type="text/javascript">

function verForm(form1)

{

	

var cliTelefone = document.form1.cliTelefone.value;



if (cliTelefone.length < 8) {

	alert("Telefone inválido!");

	document.form1.cliTelefone.value = "";

	document.form1.cliTelefone.focus();

	return false;

	}

}

function verPedido()

{

var telefone = document.form1.cliTelefone.value;

var cliID = document.form1.cliID.value;



if(cliID == "")

{

alert("Não é possível efetuar a venda!\nÉ necessário selecionar o clientes.");

document.form1.cliTelefone.focus();

}

else

{

	if(confirm("Tem certeza que deseja cadastrar um pedido?"))

		{

			window.location.href = "pg_insert_venda.asp?cliID=<%=cliID%>&tipVendaID=1&pedido=ok";

		}

		

	else

		{

			window.location.href="pg_menu_pedidos.asp?tipo=telefone";

			}

	}



	

}



function verPedidoMesa()

{



var mesID = document.form1.mesID.value;



if(mesID == ""){

alert("Não é possível efetuar a venda!\nÉ necessário selecionar uma mesa.");

document.form1.mesID.focus();

}

else

{

	if(confirm("Tem certeza que deseja cadastrar um pedido?"))

		{

			window.location.href = "pg_insert_venda.asp?mesID="+mesID+"&tipVendaID=2&pedido=ok";

		}

		

	else

	

		{

			window.location.href="pg_menu_pedidos.asp?tipo=mesa";

		}

	}

}





function verPedidoBalcao()

{

	if(confirm("Tem certeza que deseja cadastrar um pedido?"))

		{

			window.location.href = "pg_insert_venda.asp?tipVendaID=3&pedido=ok";

		}

		

	else

	

		{

			window.location.href="pg_menu_pedidos.asp?tipo=balcao";

		}

}



function cancelar()

{

	window.location.href="pg_menu_pedidos.asp";

}

</script>



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

                    <li><a href="pg_select_pedidos_telefone.asp">Vendas Telefone</a></li>

                              <li><a href="pg_select_pedidos_mesa.asp">Vendas Mesa</a></li>

                                        <li><a href="pg_select_pedidos_balcao.asp">Vendas Balcão</a></li>

                                                  <li><a href="pg_select_pedidos_internet.asp">Vendas Internet</a></li>

                                                  		<li><a href="javascript: window.open('pg_tempodeentrega.asp', 'Entrega' , 'height = 180 , width = 320');">Tempo Medio Entrega</a></li>

        </ul>

      </div>

      <div id="areaPrincipal">

        <div style="height:25px; line-height:25px; background:#ccc">Menu Principal </div>

        <table border="0" align="center" cellpadding="5" cellspacing="5">

          <tr>

            <td width="85" align="center"><a href="pg_menu_pedidos.asp?tipo=mesa"><img src="ico/ico_mesa.gif" width="60" height="60" border="0" class="icone" /></a></td>

            <td width="85" align="center"><a href="pg_menu_pedidos.asp?tipo=telefone"><img src="ico/ico_telefone.gif" width="60" height="60" border="0" class="icone" /></a></td>

            <td width="85" align="center"><a href="pg_menu_pedidos.asp?tipo=balcao"><img src="ico/ico_cesta.gif" width="60" height="60" border="0" class="icone" /></a></td>

             <td width="85" align="center"><a href="pg_select_pedidos_internet.asp"><img src="ico/ico_site.gif" width="60" height="60" border="0" class="icone" /></a></td>

          </tr>

          <tr>

            <td align="center">Mesas</td>

            <td align="center">Telefone</td>

            <td align="center">Balcão</td>

            <td align="center">Internet</td>

          </tr>

          <tr>

        </table>



<%if (tipo = "telefone") then %>



                    <table width="400" border="0" align="center" cellpadding="3" cellspacing="3">

                      <form name="form1" action="pg_menu_pedidos.asp?tipo=telefone" method="get" onsubmit="return verForm(this)">

                        <tr>

                          <td height="92" align="left" valign="top"><img src="ico/<%=imagem%>" width="60" height="60" class="icone" /></td>

                          <td height="100" align="left" valign="top"><%=msg%><br /><%=nome%><br /><%=telefone%><br /><%=endereco%><br /><br /><%=incluir%></td>

                        </tr>

                        <tr>

                          <td width="53" align="right">Telefone:</td>

                          <td width="326" height="27" align="left"><label>

                            <input name="cliTelefone" type="text" id="cliTelefone" size="9" maxlength="9"/>

                            <input name="cliID" type="hidden" id="cliID" value="<%=cliID%>" />

                          </label></td>

                        </tr>

                        <tr>

                          <td>&nbsp;</td>

                          <td height="32" align="left">

<input type="image" src="img/bot_pesquisar.png" border="0"/>&nbsp;

<a href="javascript:verPedido()"><img src="img/bot_pedido.png" width="58" height="19" border="0"/></a>&nbsp;

<a href="javascript:cancelar()"><img src="img/bot_cancelar.png" width="70" height="19" border="0"/></a></td>

                        </tr>

                      </form>

                    </table>

                    

<%elseif (tipo = "mesa") then%>





<table width="400" border="0" align="center" cellpadding="3" cellspacing="3">

                      <form name="form1" action="pg_menu_pedidos.asp?tipo=mesa" method="get" onsubmit="return verForm(this)">

                        <tr>

                          <td height="92" align="left" valign="top"><a href="pg_menu_pedidos.asp?tipo=mesa"><img src="ico/ico_mesa.gif" width="60" height="60" border="0" class="icone" /></a></td>

                          <td height="100" align="left" valign="top">Selecione a mesa</td>

                        </tr>

                        <tr>

                          <td width="53" align="right">Mesa:</td>

                          <td width="326" height="27" align="left"><label>

                            <select name="mesID" id="mesID">

                            <option value="" selected></option>

                            <% 

							if not rs01.eof then

							do while not rs01.eof

							%>

                             <option value="<%=rs01.fields.item("mesID").value%>"><%=rs01.fields.item("mesNumero").value%></option>

                            <%

							rs01.moveNext

							loop

							end if

							%>

                            </select>

                          </label></td>

                        </tr>

                        <tr>

                          <td>&nbsp;</td>

                          <td height="32" align="left"><a href="javascript:verPedidoMesa()"><img src="img/bot_pedido.png" width="58" height="19" border="0"/></a>&nbsp;

<a href="javascript:cancelar()"><img src="img/bot_cancelar.png" width="70" height="19" border="0"/></a></td>

                        </tr>

                      </form>

        </table>

<%elseif tipo="balcao" then%>



<table width="400" border="0" align="center" cellpadding="3" cellspacing="3">

                      <form name="form1" action="pg_menu_pedidos.asp?tipo=telefone" method="get" onsubmit="return verForm(this)">

                        <tr>

                          <td width="53" height="92" align="left" valign="top"><a href="pg_menu_pedidos.asp?tipo=balcao"><img src="ico/ico_cesta.gif" width="60" height="60" border="0" class="icone" /></a></td>

                          <td width="326" height="100" align="left" valign="top">Venda Balcão</td>

                        </tr>

                        <tr>

                          <td>&nbsp;</td>

                          <td height="32" align="left"><a href="javascript:verPedidoBalcao()"><img src="img/bot_pedido.png" width="58" height="19" border="0"/></a>&nbsp;

<a href="javascript:cancelar()"><img src="img/bot_cancelar.png" width="70" height="19" border="0"/></a></td>

                        </tr>

                      </form>

        </table>

                                        

<%end if%>

        </tr>

        <tr>

          <td colspan="3" align="center">&nbsp;</td>

        </tr>

        </table>

      </div>

    </div>

    <!-- -->

  </div>

  <div id="rodape"><br />

    <!--#include file="inc/inc_status.inc"-->

    <br />

  </div>

</div>

<!--FIM DO LAYOUT-->

</body>

</html>



<%

call fechaConexao()

%>