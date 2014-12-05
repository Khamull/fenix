<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%option explicit%>

<!--#include file="inc/inc_conexao.inc"-->

<!--#include file="inc/inc_formato_data.inc"-->

<!--#include file="inc/inc_acesso.inc" -->

<%
call abreConexao()

Dim rs00
Dim sql00

	set rs00 = server.CreateObject("adodb.recordset")
		sql00 = "SELECT *, tb_mesa.mesNumero FROM tb_venda INNER JOIN tb_mesa ON tb_mesa.mesID = tb_venda.mesID WHERE (tb_venda.staID = 1 OR tb_venda.staID = 4 OR tb_venda.staID = 5 OR tb_venda.staID = 6) AND tipVendaID = 2"
	set rs00 = conn.Execute(sql00)
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>SISTEM FORTE EM MÍDIA</title>
<link href="css/css1.css" rel="stylesheet" type="text/css" />

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
	    <li><a href="pg_menu_pedidos.asp?tipo=mesa">Novo Pedido</a></li>
	    <li><a href="pg_insert_mesa.asp">Nova Mesa</a></li>
		<li><a href="pg_select_mesa.asp">Listar Mesa</a></li>
		<li><a href="pg_pedidos_fechados_mesa.asp">Pedidos Fechados</a></li>         
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc">Menu Principal
</div>
    <table width="100%" border="0" align="center" cellpadding="2" cellspacing="2">
  <tr>
    <td height="242" align="left" valign="top">
     <%
	 if not rs00.eof then
	 do while not rs00.eof
	 %> 
      <div id="mesa" style="width:120px; border:1px solid black; margin:6px; float:left;">
        <table width="120">
          <tr>
            <!-- !!!!!CODIGO ANTIGO QUE CHAMAVA O PEDIDO PARA OUTRA TELA!!!!!
            <td width="52" rowspan="2"><a href="pg_insert_itemVendaMesa.asp?venID=<'%=rs00.fields.item("venID").value%>"><img src="ico/ico_mesa.gif" alt="" width="60" height="60" border="0" class="icone" /></a></td>
            -->
            
            <td width="52" rowspan="2"><a href="pg_insert_itemVendaMesa1.asp?venID=<%=rs00.fields.item("venID").value%>"><img src="ico/ico_mesa.gif" alt="" width="60" height="60" border="0" class="icone" /></a></td>
            <td width="66" align="center" valign="top">Mesa</td>
          </tr>
          <tr>
            <td align="center" valign="middle"><span style="text-align:center; font-family:arial; color:#333; font-size:20px; font-weight:bold"><%=rs00.fields.item("mesNumero").value%></span></td>
          </tr>
          <tr>
            <td height="16" colspan="2" align="center">Hora: <%=rs00.fields.item("venHoraA").value%></td>
          </tr>
        </table>
        </div>
         <%
		 rs00.movenext
		 loop
		 end if
		 %>
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
