<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%option explicit%>

<!--#include file="inc/inc_conexao.inc"-->

<!--#include file="inc/inc_formato_data.inc"-->

<!--#include file="inc/inc_acesso.inc" -->

<%
call abreConexao()
%>

<%
'DECLARA A VARIÁVEL DE CONSULTA
Dim carID
'VERIFICA O PARAMETRO
carID = REQUEST.QUERYSTRING("carID")
if(carID = "") THEN
'REDIRECIONADA PARA O MENU PRINCIPAL
RESPONSE.REDIRECT("pg_menu.asp")
END IF

'VARIÁVEL PADRÃO
Dim msg
Dim usuLogin

'VERIFICA A MENSAGEM
msg = REQUEST.QUERYSTRING("msg")
IF(msg = "1") THEN
msg = "Os Dados foram atualizados com sucesso!"
ELSE
msg = ""
END IF

'-**********

'DECLARA AS VARIÁVEIS
Dim carDescr
Dim carAtivo


'RECUPERA OS VALORES
usuLogin 	=	Request.Form("usuLogin")
carAtivo	=	Request.Form("carAtivo")
carDescr		=	UCase(Trim(Replace(Request.Form("carDescr"),"'","")))


'SELECIONA OS REGISTROS

		Dim rs00
		Dim sql00
		set rs00 = server.CreateObject("adodb.recordset")
		sql00="SELECT "		
		sql00=sql00&" tb_cargo.carID,"
		sql00=sql00&" tb_cargo.carDescr,"
		sql00=sql00&" tb_cargo.carAtivo"
		sql00=sql00&" FROM"
		sql00=sql00&" tb_cargo" 
		sql00=sql00&" WHERE"
		sql00=sql00&" tb_cargo.carID='"&carID&"'"
		set rs00=conn.Execute(sql00)
%>

<%
'VERIFICA SE O FORME FOI DISPARADO
if (not isEmpty(Request.Form("atualizar"))) then

		Dim rs01
		Dim sql01
		set rs01 = server.CreateObject("adodb.recordset")
		sql01="UPDATE tb_cargo SET"
		sql01=sql01&" carDescr='"&carDescr&"',"
		sql01=sql01&" usuLogin='"&usuLogin&"'," 
		sql01=sql01&" carAtivo='"&carAtivo&"'" 
		sql01=sql01&" WHERE" 
		sql01=sql01&" carID='"&carID&"'"
		set rs01 = conn.execute(sql01)
		
		'RETORNA UMA MENSAGEM DO RESULTADO DA OPERAÇÃO
		response.redirect("pg_update_cargo.asp?msg=1&carID="&carID)
end if
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>SISTEM FORTE EM MÍDIA</title>
<link href="css/css1.css" rel="stylesheet" type="text/css" />

<script language="javascript" type="text/javascript">

function verForm(form1){
	

carDescr = document.form1.carDescr.value;

if (carDescr.length < 3) {
	alert("Favor Informar o nome do Cargo!");
	document.form1.carDescr.focus();
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
	    <li><a href="pg_select_cargo.asp">Lista cargos</a></li>
      	<li><a href="pg_insert_cargo.asp">Nova cargo</a></li>           
	    <li><a href="pg_insert_funcionario.asp">Funcionário</a></li>              
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc"></div>
    <table width="96%" border="0" align="left" cellpadding="3" cellspacing="3">
  <tr>
    <td width="50" align="center"><img src="ico/ico_M.gif" width="60" height="60" class="icone" /></td>
    <td width="751" align="center" class="titulo">ATUALIZAR CARGO</td>
    </tr>
  <tr>
    <td colspan="2" align="center"><form id="form1" name="form1" method="post" action="pg_update_cargo.asp?carID=<%=carID%>" onsubmit="return verForm(this)">
      <table width="777" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td height="25" colspan="4" align="center"><%=msg%></td>
        </tr>
        <tr>
          <td width="109" height="25" align="right">cargo:&nbsp;</td>
          <td width="267" height="25" align="left"><input name="carDescr" type="text" id="carDescr" value="<%=rs00.fields.item("carDescr").value%>" size="40" maxlength="30" /></td>
          <td width="110" height="25" align="right">&nbsp;</td>
          <td width="263" height="25" align="left">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" align="right">Ativo:</td>
          <td height="25" align="left"><label>
            <select name="carAtivo" id="carAtivo">
              <option value="S" selected="selected">SIM</option>
              <option value="N">NÃO</option>
  </select>
            </label></td>
          <td height="25" align="right">&nbsp;</td>
          <td height="25">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left"><input name="atualizar" type="submit" class="botao" id="atualizar" value="Atualizar" />
            <input name="usuLogin" type="hidden" id="usuLogin" value="<%=Session("usuLogin")%>" /></td>
          <td height="25" align="right">&nbsp;</td>
          <td height="25">&nbsp;</td>
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