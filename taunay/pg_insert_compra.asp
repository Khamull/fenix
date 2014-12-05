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
Dim id

id = Request.QueryString("id")
acao = Request.QueryString("acao")

if (acao = "excluir") then

Dim rs04
Dim sql04
set rs04 = server.CreateObject("adodb.recordset")
	sql04 = "DELETE FROM tb_compra WHERE comID = '"&id&"'"
set rs04 = conn.execute(sql04)	

msg = "O Pedido de compra foi excluido com sucesso!"

elseif (acao = "inserir") then

msg = "REDIRECIONA!"

end if
%>


<%
Dim msg
Dim usuLogin
Dim forID
Dim comDataC
Dim comFormPgto
Dim comNumParc
Dim comDataV
Dim comValorF
Dim comVendedor

usuLogin 		=	Request.Form("usuLogin")
forID			=	UCase(Trim(Replace(Request.Form("forID"),"'","")))
comDataC		=	UCase(Trim(Replace(Request.Form("comDataC"),"'","")))
comFormPgto		=	UCase(Trim(Replace(Request.Form("comFormPgto"),"'","")))
comNumParc		=	UCase(Trim(Replace(Request.Form("comNumParc"),"'","")))
comDataV		=	UCase(Trim(Replace(Request.Form("comDataV"),"'","")))
comValorF		=	UCase(Trim(Replace(Request.Form("comValorF"),"'","")))
comVendedor		= 	UCase(Trim(Replace(Request.Form("comVendedor"),"'","")))

'comID
'comDataC
'comDataR
'comNumParc
'usuLogin
'forID
'comFormaPgto
'comDataVenc
'comValorT
'comValorF
'staID
'comVendedor

%>

<%
if (not isEmpty(Request.Form("cadastrar"))) then

	Dim fornecedor
	
	fornecedor = Request.Form("forID")

	Dim rs00
	Dim sql00
	set rs00 = server.CreateObject("adodb.recordset")
		sql00 = "SELECT * FROM tb_compra WHERE staID = 9 AND forID = '"&fornecedor&"'"
	set rs00 = conn.Execute(sql00)
	
	if (not rs00.eof) then
	
	'Mensagem
	rs00.close
	set rs00 = nothing
	msg = "Existe um pedido de compra em aberto para esse Fornecedor!"	
	
	else
	
	'Grava
		Dim rs01
		Dim sql01
		set rs01 = server.CreateObject("adodb.recordset")
			sql01 = "INSERT INTO tb_compra (usuLogin, forID, comDataC, comFormPgto, comNumParc, comDataV, comValorF, comVendedor) VALUES "
			sql01 = sql01 & "('"&usuLogin&"','"&forID&"','"&comDataC&"','"&comFormPgto&"','"&comNumParc&"','"&comDataV&"','"&comValorF&"','"&comVendedor&"')"
		set rs01 = conn.execute(sql01)
		set rs01 = nothing
		
		'****************************
		
		rs00.close
		set rs00 = nothing
		msg = "Os Dados foram cadastrados com sucesso!"
		
	end if
	
end if

%>

<%
Dim rs02
Dim sql02
set rs02 = server.CreateObject("adodb.recordset")
	sql02 = "SELECT * FROM tb_fornecedor WHERE forAtivo = 'S' AND forID <> '300' ORDER BY forNome"
set rs02 = conn.execute(sql02)	
%>

<%
Dim rs03
Dim sql03
set rs03 = server.CreateObject("adodb.recordset")
	sql03 = "SELECT *, tb_fornecedor.forNome FROM tb_compra INNER JOIN tb_fornecedor ON tb_fornecedor.forID = tb_compra.forID where staID = 9"
set rs03 = conn.execute(sql03)	
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
        <li><a href="pg_insert_estoque.asp">Ver Estoque</a></li>                  
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc"></div>
    <table width="96%" border="0" align="left" cellpadding="3" cellspacing="3">
  <tr>
    <td width="50" align="center"><img src="ico/ico_casabairro.gif" width="60" height="60" class="icone" /></td>
    <td width="751" align="center" class="titulo">CADASTRO DE COMPRA</td>
    </tr>
  <tr>
    <td colspan="2" align="center"><form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>" onsubmit="return verForm(this)">
      <table width="777" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td height="25" colspan="4" align="center"><font color="#FF0000"><%=msg%></font></td>
        </tr>
        <tr>
          <td width="109" height="25" align="right">Fornecedor:</td>
          <td width="267" height="25" align="left">
            
            <select name="forID" id="forID">
              <option value=""></option>
              <%
		  if (not rs02.eof) then
          do while not rs02.eof
		  %>
              <option value="<%=rs02.fields.item("forID").value%>"><%=rs02.fields.item("forNome").value%></option>
              
              <%
		  rs02.moveNext 
		  Loop
		  end if
		  %>
              </select>
            * </td>
          <td width="110" height="25" align="right">Data da Compra:</td>
          <td width="263" height="25" align="left"><label>
            <input name="comDataC" type="text" id="comDataC" value="<%=date()%>" size="10" maxlength="10" readonly="readonly" />
          </label></td>
        </tr>
        <tr>
          <td height="25" align="right"><label for="checkbox_row_7">Forma de Pgto</label>
            :</td>
          <td height="25" align="left"><label>
            <select name="comFormPgto" id="comFormPgto">
              <option value="1">À VISTA</option>
              <option value="2">PARCELADO</option>
              <option selected="selected"> </option>
            </select>
          *</label></td>
          <td height="25" align="right">N° de Parcelas:</td>
          <td height="25" align="left"><input name="comNumParc" type="text" id="comNumParc" value="1" size="2" maxlength="2" /></td>
        </tr>
        <tr>
          <td height="25" align="right">Data Vencimento:</td>
          <td height="25" align="left"><input name="comDataV" type="text" id="comDataV" value="<%=date()%>" size="10" maxlength="10" /></td>
          <td height="25" align="right">Valor do Frete:</td>
          <td height="25" align="left"><input name="comValorF" type="text" id="comValorF" value="0.00" size="10" maxlength="10" /></td>
        </tr>
        <tr>
          <td height="25" align="right">Vendedor:</td>
          <td height="25" align="left"><label>
            <input name="comVendedor" type="text" id="comVendedor" value="Não informado" size="40" maxlength="50" />
          </label></td>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" align="right">&nbsp;</td>
          <td height="25"><input name="cadastrar" type="submit" class="botao" id="cadastrar" value="Cadastrar" /></td>
          <td height="25" align="right"><input name="usuLogin" type="hidden" id="usuLogin" value="<%=Session("usuLogin")%>" /></td>
          <td height="25">&nbsp;</td>
        </tr>
      </table>
    </form></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><table width="777" border="0" cellpadding="2" cellspacing="2">
      <tr class="textoBranco">
        <td width="28" height="20" align="left" bgcolor="#9E231B">ID</td>
        <td width="70" height="20" align="left" bgcolor="#9E231B">Data</td>
        <td width="220" height="20" align="left" bgcolor="#9E231B">Fornecedor</td>
        <td width="80" height="20" align="left" bgcolor="#9E231B">Forma Pgto</td>
        <td width="70" height="20" align="left" bgcolor="#9E231B">Vencimento</td>
        <td width="148" height="20" align="left" bgcolor="#9E231B">Vendedor</td>
        <td width="70" align="left" bgcolor="#9E231B">Valor  Frete</td>
        <td width="15" height="20" align="left" bgcolor="#9E231B">Ex</td>
        <td width="20" height="20" align="left" bgcolor="#9E231B">Ver</td>
      </tr>
      <% if not rs03.eof then%>
      <% While Not rs03.EoF %>
      <tr>
        <td align="left"><%=rs03.fields.item("comID").value%></td>
        <td align="left"><%=rs03.fields.item("comDataC").value%></td>
        <td align="left"><%=rs03.fields.item("forNome").value%></td>
        <td align="left">
		
		<%
		
		Dim forma
		
		forma = rs03.fields.item("comFormPgto").value
		if forma = "1" then
		Response.write("À VISTA")
		end if
		
		if forma = "2" then
		Response.write("À PRAZO")
		end if
		
		%>
        
        </td>
        <td align="left"><%=rs03.fields.item("comDataV").value%></td>
        <td align="left"><%=rs03.fields.item("comVendedor").value%></td>
        <td align="left"><%=FormatNumber(rs03.fields.item("comValorF").value)%></td>
        <td align="left"><a href="javascript:Excluir(<%=rs03.fields.item("comID").value%>)"><img src="ico/ico_excluir.gif" border="0" width="15" height="15" /></a></td>
        <td align="left"><a href="pg_insert_itemCompra.asp?id=<%=rs03.fields.item("comID").value%>"><img src="ico/ico_olho.gif" border="0" width="15" height="15" /></a></td>
      </tr>
      
      <%
	   rs03.MoveNext
	  Wend
	  %>
      
      <%end if%>
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
rs02.close
set rs02 = nothing
%>
<%
rs03.close
set rs03 = nothing
%>
<%
call FechaConexao()
%>