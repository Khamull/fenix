<%@LANGUAGE="VBSCRIPT" CODEPAGE="28592"%>
<%option explicit%>
<!--#include file="inc/inc_conexao.inc"-->

<!--#include file="Connections/conn.asp" -->

<!--#include file="inc/inc_formato_data.inc"-->

<!--#include file="inc/inc_acesso.inc" -->
<%
Call abreConexao
%>
<%
Dim Recordset2
Dim Recordset2_cmd
Dim Recordset2_numRows

Set Recordset2_cmd = Server.CreateObject ("ADODB.Command")
Recordset2_cmd.ActiveConnection = MM_conn_STRING
Recordset2_cmd.CommandText = "SELECT * FROM tb_imagemempresa ORDER BY imagemID ASC" 
Recordset2_cmd.Prepared = true

Set Recordset2 = Recordset2_cmd.Execute
Recordset2_numRows = 0
%>
<%
Dim acao
Dim rs1
Dim sql01
Dim rs02
Dim sql02
Dim imagem
Dim imagemID
Dim statusID

    acao = Request.QueryString("acao")
    imagemID = Request.QueryString("imagemID")
	statusID = Request.QueryString("statusID")
	
if(acao = "1") then 

set rs02 = Server.CreateObject("ADODB.Recordset")
sql02 = "SELECT * FROM  tb_imagemempresa WHERE imagemID = '"&imagemID&"'"
set rs02 = conn.execute(sql02)

While Not rs02.EoF

imagem = rs02.fields.item("imagemCaminho").value

'---EXCLUIR A FOTO DO DIRETÓRIO
Dim path, objFSO
set objFSO = Server.CreateObject("Scripting.FileSystemObject")
path = Server.MapPath("upload_empresa")
path = path&"\"&imagem
path = objFSO.GetAbsolutePathName(path)

if objFSO.FileExists(path) = true Then
objFSO.DeleteFile path
end if

rs02.MoveNext
Wend


set rs1 = Server.CreateObject("ADODB.Recordset")
sql01   = "DELETE FROM tb_imagemempresa WHERE imagemID = '"&imagemID&"'"
set rs1 = conn.execute(sql01)
response.Redirect("sis_update_imagem_empresa.asp")

elseif(acao = "2") Then

set rs02 = Server.CreateObject("ADODB.Recordset")
sql02   = "INSERT INTO tb_imagemempresa (statusID) VALUES ('"&statusID&"')"
set rs02 = conn.execute(sql02)

Response.Redirect("sis_update_imagem_empresa.asp")
end if

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-2" />
<title>SISTEMA WEB DE GEST&Atilde;O</title>
<link href="css/css1.css" rel="stylesheet" type="text/css" />
<script type="text/javascript">
function excluir(imagemID){
	if (confirm("tem certeza que deseja excluir essa Imagem?")){
		window.location.href="sis_update_imagem_empresa.asp?imagemID="+imagemID+"&acao=1";
	}else{
		return false;
	}
}


 function abreJanela(imagemID){
	 window.open("sis_janela_ft_imagem_empresa.asp?imagemID="+imagemID, "fotos", "height = 210, width = 500");
  }
  
  function inserir(statusID){
	if (confirm("Deseja inserir uma nova foto?")){
		window.location.href="sis_update_imagem_empresa.asp?statusID="+statusID+"&acao=2";
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
<div id="topo" align="right"></div>
<div id="tituloBar">&nbsp;</div>
<div id="corpo" style="height:395px">
<!-- -->
<div id="areaConteudo">

	<div id="areaMenuVerfical">
	<div style="height:25px; line-height:25px; background:#ccc"><strong>MENU</strong></div>    
	 <ul>
	      <li><a href="pg_menu.asp">Menu Principal</a></li>
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc"><strong>EMPRESA - IMAGENS</strong></div>
    <div><!--<a href="sis_update_empresa3.asp"><input type="button" value="Quem Somos" class="botao" border="0" style="float:left;"/></a>--><a href="javascript: inserir('Novo&nbsp;titulo','Novo&nbsp;texto')"><input type="button" value="+ Nova Imagem" class="botao" border="0"/></a></div>
    
      <table border="0" align="center">
        <tr valign="baseline">
          <td colspan="3" align="center" nowrap="nowrap">&nbsp;</td>
        </tr>
        <tr valign="baseline">
          <td>&nbsp;</td>
          <td align="center" nowrap="nowrap" bgcolor="#032A6F" class="textoBranco"><strong>EDITAR IMAGENS</strong></td>
          <td align="center" nowrap="nowrap">&nbsp;</td>
        </tr>
        
        <% While Not Recordset2.EoF %>
      
        <tr valign="baseline">
          <td height="16" align="left" valign="top" nowrap="nowrap"><a href="javascript: abreJanela(<%=(Recordset2.Fields.Item("imagemID").Value)%>);"><img src="ico/ico_mfoto.gif" width="30" height="30" border="0" title="Foto" /></a></td>
          <td align="left"><img src="<%
if (Recordset2.fields.item("imagemCaminho").value <> "") Then
%>upload_empresa/<%=(Recordset2.Fields.Item("imagemCaminho").Value)%><%
else
%>img/sem_foto.png<%

end if
%>"width="217" height="143"  style=" z-index:2" title='Imagem' /></td>
          <td valign="top"><a href="javascript: excluir(<%=(Recordset2.Fields.Item("imagemID").Value)%>)"><img src="ico/ico_excluir.gif" width="15" height="15" border="0" title="Excluir"/></a></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">&nbsp;</td>
          <td align="center"><br /><hr /><br /></td><td></td>
        </tr>
     
     <% 
		     Recordset2.moveNext()
			Wend
		   %>

      </table>
          <p>&nbsp;</p>
    </div>
</div>
</div>
</div>
<div id="rodape"><br /><!--#include file="inc/inc_status.inc"--><br /></div>
<!--FIM DO LAYOUT-->
</body>
</html>
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>
