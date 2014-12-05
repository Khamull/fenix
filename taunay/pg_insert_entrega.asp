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
Dim funID
Dim entLocal
Dim entHoraS
Dim entHoraR
Dim entData
Dim staID
Dim usuLogin
Dim entObs

venID		=	Request.QueryString("venID")
funID		=	Request.Form("funID")
entLocal	=	Request.Form("entLocal")
entHoraS	=	Request.Form("entHoraS")
entHoraR	=	Request.Form("entHoraR")
entData		=	Request.Form("entData")
staID		=	Request.Form("staID")
usuLogin	=	Request.Form("usuLogin")
entObs		=	Request.Form("entObs")
%>



<%
if not isEmpty(Request.Form("cadastrar")) then

		Dim rs03
		Dim sql03
		set rs03 = server.CreateObject("adodb.recordset")
			sql03 = "INSERT INTO tb_entrega (venID, funID, entLocal, entHoraS, entHoraR, entData, staID, usuLogin, entObs) VALUES ('"&venID&"','"&funID&"','"&entLocal&"','"&entHoraS&"','"&entHoraR&"','"&entData&"','6','"&usuLogin&"','"&entObs&"')"
		set rs03 = conn.execute(sql03)
		set rs03 = nothing
		
		
		dim rs04
		dim sql04
		set rs04 = server.CreateObject("adodb.recordset")
		sql04 = "UPDATE tb_venda SET staID = '6' WHERE venID = '"&venID&"'"
		sql04 = conn.execute(sql04)
		
		Response.redirect("pg_select_entrega.asp")
		
		elseif not isEmpty(Request.Form("cancelar")) then
		
		Response.redirect("pg_select_entrega.asp")
end if
%>


<%
'SELECIONA DETALHES DO PEDIDO
Dim rs02
Dim sql02
set rs02 = Server.CreateObject("ADODB.Recordset")
sql02 = "SELECT "
sql02 = sql02 & "tb_tipovenda.tipVendaID, tb_tipovenda.tipVendaDescricao, " 'TB_TIPOVENDA
sql02 = sql02 & "tb_cliente.cliID, tb_cliente.cliNome, tb_cliente.cliTelefone, " 'TB_CLIENTE 1°
sql02 = sql02 & "tb_cliente.cliEndereco, tb_cliente.baiID, tb_cliente.cidID, "	 'TB_CLIENTE 2°
sql02 = sql02 & "tb_bairro.baiID, tb_bairro.baiNome, tb_bairro.baiFrete, "	 'TB_BAIRRO
sql02 = sql02 & "tb_cidade.cidID, tb_cidade.cidNome, "	 'TB_CIDADE
sql02 = sql02 & "tb_venda.* " 'TB_VENDA
sql02 = sql02 & "FROM tb_venda " 'TABELA PRINCIPAL
sql02 = sql02 & "INNER JOIN tb_tipovenda ON tb_tipovenda.tipVendaID = tb_venda.tipVendaID " 'INNER JOIN com TIPO DE VENDA
sql02 = sql02 & "INNER JOIN tb_cliente ON tb_cliente.cliID = tb_venda.cliID " 'INNER JOIN com CLIENTE
sql02 = sql02 & "INNER JOIN tb_bairro ON tb_bairro.baiID = tb_cliente.baiID " 'INNER JOIN com BAIRRO
sql02 = sql02 & "INNER JOIN tb_cidade ON tb_cidade.cidID = tb_cliente.cidID " 'INNER JOIN com CIDADE
sql02 = sql02 & "WHERE tb_venda.venID = '"&venID&"'" 'CONDIÇÃO
set rs02 = conn.execute(sql02)


%>

<%
Dim rs01
Dim sql01
set rs01 = server.CreateObject("adodb.recordset")
sql01 = "SELECT * FROM tb_funcionario WHERE carID = 1 AND funAtivo = 'S'"
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
	
var cidNome
var funID

cidNome = document.form1.cidNome.value;
funID = document.form1.funID.value;


if (cidNome < 3) {
	alert("Favor informar o nome da Cidade!");
	document.form1.cidNome.focus();
	return false;
	}
	
if (funID == "") {
	alert("Favor informar o nome do Entregador!");
	document.form1.funID.focus();
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
	    <li><a href="pg_menu.asp">Menu Principal</a></li>
	    <li><a href="pg_select_entrega.asp">Listar Entregas</a></li>        
	    <li><a href="pg_insert_bairro.asp">Bairro</a></li>            
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc"></div>
    <table width="96%" border="0" align="left" cellpadding="3" cellspacing="3">
  <tr>
    <td width="50" align="center"><img src="ico/ico_moto.gif" width="60" height="60" class="icone" /></td>
    <td width="751" align="center" class="titulo">ENTREGAS</td>
    </tr>
  <tr>
    <td colspan="2" align="center"><form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?venID=<%=venID%>" onsubmit="return verForm(this)">
      <table width="777" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td height="25" colspan="4" align="center">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" align="right">Pedido:&nbsp;</td>
          <td width="267" height="25" align="left"><label>
            <input name="venID" type="text" id="venID" value="<%=rs02.fields.item("venID").value%>" size="10" maxlength="10" readonly="readonly" />
          </label></td>
          <td width="110" height="25" align="right">Tipo de Venda:</td>
          <td width="263" height="25" align="left"><input name="tipVendaID" type="text" id="tipVendaID" value="<%=rs02.fields.item("tipVendaDescricao").value%>" size="20" maxlength="20" readonly="readonly" /></td>
        </tr>
        <tr>
          <td height="25" align="right">Cliente:&nbsp; </td>
          <td height="25" align="left"><input name="cidNome" type="text" id="cidNome" value="<%=rs02.fields.item("cliNome").value%>" size="50" maxlength="50" readonly="readonly" /></td>
          <td height="25" align="right">Atendente:&nbsp; </td>
          <td height="25" align="left"><input name="usuLogin" type="text" id="usuLogin" value="<%=rs02.fields.item("usuLogin").value%>" size="50" maxlength="50" readonly="readonly" /></td>
          </tr>
        <tr>
          <td width="109" height="25" align="right">Localidade:</td>
          <td height="25" align="left"><label>
            <textarea name="entLocal" id="entLocal" cols="45" rows="3"><%=rs02.fields.item("cliEndereco").value%> - <%=rs02.fields.item("BaiNome").value%> - <%=rs02.fields.item("cidNome").value%></textarea>
          </label></td>
          <td height="25" align="right">Observações:&nbsp;</td>
          <td height="25" align="left"><textarea name="entObs" id="entObs" cols="45" rows="3"></textarea></td>
          </tr>
        <tr>
          <td height="25" align="right">Hora da Saida:&nbsp;</td>
          <td height="25" align="left"><input name="entHoraS" type="text" id="entHoraS" value="<%=time()%>" size="10" maxlength="10" readonly="readonly" /></td>
          <td height="25" align="right">Valor do Frete:&nbsp;</td>
          <td height="25" align="left"><input name="baiFrete" type="text" id="baiFrete" value=<%=rs02.fields.item("baiFrete").value%> size="20" maxlength="20" readonly="readonly" /></td>
        </tr>
        <tr>
          <td height="25" align="right">Entregador:&nbsp;</td>
          <td height="25" align="left">
          <select name="funID" id="funID">
          <option selected></option>
			<%
			if not rs01.eof then
			do while not rs01.eof
			%>
          <option value="<%=rs01.fields.item("funID").value%>"><%=rs01.fields.item("funNome").value%></option>
			<%
			rs01.moveNext
			loop
			end if
			%>
          </select>
          
          </td>
          <td height="25" align="right">Valor do Pedido:&nbsp;</td>
          <td height="25" align="left"><input name="tipVendaDescr3" type="text" id="tipVendaDescr3" value="<%=rs02.fields.item("venValorT").value%>" size="20" maxlength="20" readonly="readonly" /></td>
        </tr>
        <tr>
          <td height="25" align="right">&nbsp;</td>
          <td height="25"><input name="cadastrar" type="submit" class="botao" id="cadastrar" value="Cadastrar" /></td>
          <td height="25" align="right"><input name="entData" type="hidden" id="entData" value="<%=Date()%>" />            <input name="usuLogin" type="hidden" id="usuLogin" value="<%=Session("usuLogin")%>" /></td>
          <td height="25"><input name="cancelar" type="submit" class="botao" id="cancelar" value="Cancelar" /></td>
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
call FechaConexao()
%>