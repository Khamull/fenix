<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%option explicit%>

<!--#include file="inc/inc_conexao.inc"-->

<!--#include file="inc/inc_formato_data.inc"-->

<!--#include file="inc/inc_acesso.inc" -->

<%
call abreConexao()
%>



<%'RECUPERA O ID DA ÚLTIMA VENDA
Dim vendaInicial
Dim rs0
Dim sql0

set rs0 = Server.CreateObject("ADODB.Recordset")
sql0 = "SELECT MAX(venID) as ultima FROM tb_venda"
set rs0 = conn.execute(sql0)

vendaInicial = rs0.fields.item("ultima").value
vendaInicial = (vendaInicial + 1)
%>




<%'CADASTRA UMA ABERTURA DE CAIXA
Dim valorInicial
Dim rs01
Dim sql01

valorInicial = Replace(Request.Form("valorInicial"),",",".")

if(Not isEmpty(Request.Form("cadastrar")))Then

set rs01 = Server.CreateObject("ADODB.Recordset")
sql01 = "INSERT INTO tb_caixa (valorInicial, vendaInicial, data) VALUES ('"&valorInicial&"', '"&vendaInicial&"', '"&Data&"')"
set rs01 = conn.execute(sql01)


	'RECUPERA O ID DO ÚLTIMO CAIXA QUE FOI ABERTO
	Dim rs011
	Dim sql011
	
	set rs011 = Server.CreateObject("ADODB.Recordset")
	sql011 = "SELECT caixaID as id FROM tb_caixa ORDER BY caixaID DESC"
	set rs011 = conn.execute(sql011)


	'CADASTRA UM NUMERO DE VENDA "0" PARA ZERAR A CONTAGEM DO CAIXA
	Dim rs012
	Dim sql012
	
	set rs012 = Server.CreateObject("ADODB.Recordset")
	sql012 = "INSERT INTO tb_numerovenda (caixaID, numerovenda) VALUES ('"&rs011.fields.item("id").value&"', '0')"
	set rs012 = conn.execute(sql012)


Response.Redirect("pg_menu.asp")
end if
%>



<%'VERIFICA SE JÁ EXISTE UM DIA ABERTO
Dim rs02
Dim sql02

set rs02 = Server.CreateObject("ADODB.Recordset")
sql02 = "SELECT * FROM tb_caixa WHERE status = 'A'"
set rs02 = conn.execute(sql02)
%>



<%'FECHAMENTO DE CAIXA
Dim fechar

fechar = Request.QueryString("fechar")

if (fechar = "ok") Then

Dim venIni
Dim vendaFinal
Dim caixaID
Dim valorIni
Dim valorFinal
Dim x
Dim rs03
Dim sql03
Dim rs04
Dim sql04

venIni = rs02.fields.item("vendaInicial").value
vendaFinal = (vendaInicial - 1)
caixaID = rs02.fields.item("caixaID").value
valorIni = rs02.fields.item("valorInicial").value

	'SOMA TODAS AS VENDAS ENTRE O ABRIMENTO DE O FECHAMENTO DO CAIXA
	set rs03 = Server.CreateObject("ADODB.Recordset")
	sql03 = "SELECT SUM(venValorT) as total FROM tb_venda WHERE venID BETWEEN '"&venIni&"' AND '"&vendaFinal&"'"
	set rs03 = conn.execute(sql03)
	
	
	valorFinal = rs03.fields.item("total").value
	valorFinal = CDbl(valorFinal)
	valorIni   = CDbl(valorIni)
	valorFinal = (valorFinal + valorIni)
	ValorFinal = Replace(valorFinal,",",".")
	

'ATUALIZA O CAIXA
set rs04 = Server.CreateObject("ADODB.Recordset")
sql04 = "UPDATE tb_caixa SET valorFinal = '"&valorFinal&"', vendaFinal = '"&vendaFinal&"', status = 'F' WHERE caixaID = '"&caixaID&"'"
set rs04 = conn.execute(sql04)

Response.Redirect("pg_view_fechamento.asp?caixaID="&caixaID)

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
	
var valorInicial	= document.form1.valorInicial.value;


if (valorInicial.length = "") {
	alert("Favor informar o Valor Inicial do Caixa!");
	document.form1.valorInicial.focus();
	return false;
	}

	
}

function fechar(){
	if(confirm("Tem certeza que Deseja Encerrar as Vendas AGORA?")){
		window.location.href="pg_caixa.asp?fechar=ok";
	}else{
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
	    <li><a href="pg_caixa_fechado.asp">Fechamentos</a></li>        
               
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc"></div>
    <table width="96%" border="0" align="left" cellpadding="3" cellspacing="3">
  <tr>
    <td width="50" align="center"><img src="ico/ico_caixa.png" width="246" height="246" class="icone" /></td>
    <td width="751" align="center" class="titulo">
    <%
	if rs02.EoF Then
     Response.Write("ABERTURA DE CAIXA")
	else
	 Response.Write("FECHAMENTO DE CAIXA")
	end if 
	%>
    </td>
    </tr>
  <tr>
    <td align="center">&nbsp;</td>
    <td align="center"><form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>"
    onsubmit="return verForm(this)">
      <table width="270" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left">&nbsp;</td>
        </tr>
        <tr>
          <td width="118" height="25" align="right" bgcolor="#EEEEEE"><strong>DATA:</strong></td>
          <td width="146" height="25" align="left" bgcolor="#EEEEEE"><%=CDate(Data)%></td>
        </tr>
        <%'só irá aparecer "ABERTURA DE CAIXA" Caso não tenha caixa Aberto
		 if (rs02.EoF) Then
		%>
        <tr>
          <td height="25" align="right" bgcolor="#EEEEEE"><strong>VALOR INICIAL:</strong></td>
          <td height="25" align="center"><input name="valorInicial" type="text" value="0.00" maxlength="10" /></td>
          </tr>
        <tr>
          <td height="25" align="left">&nbsp;</td>
          <td height="25" align="left"><input name="cadastrar" type="submit" class="botao" id="cadastrar" value="ABRIR CAIXA" /></td>
        </tr>
        
        <%else%>
        
        <tr>       
          <td height="35" align="center" colspan="2"><u><a href="javascript: fechar()">FECHAR CAIXAR</a></u></td>
        </tr>
        
        <%end if%>
        
        <tr>
          <td height="25" align="right"><input name="usuLogin2" type="hidden" id="usuLogin2" value="<%=Session("usuLogin")%>" /></td>
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