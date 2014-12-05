<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include file="inc/inc_formato_data.inc"-->

<!--#include file="inc/inc_conexao.inc"-->

<%
if not isEmpty(request.form("Enviar")) then

'Set objUpload = Server.CreateObject("Dundas.Upload.2")

'Set objUpload = nothing
end if
%>

<%
call abreConexao
%>

<%
DIm cod
cod = request.querystring("cod")
if (cod = 1) Then
msg = "A imagem foi publicada com sucesso, selecione outra ou clique no 'X' para fechar!"
else
msg = "Copyright© 2012 - Forte em Mídia Propaganda Todos os direitos reservados"
end if
%>
<%
Dim imagemID
imagemID = Request.Querystring("imagemID")

if (imagemID = "") Then
%>
<script>window.close()</script>
<%
End if
%>

<%

Dim rs01
Dim sql01

set rs01 = Server.CreateObject("ADODB.Recordset")
sql01 = "SELECT * FROM tb_imagemhome WHERE imagemID = '"&imagemID&"'"
set rs01 = conn.execute(sql01)

%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>SISTEMA FORTE EM MÍDIA - PUBLICAÇÃO DE FOTOS</title>
<link href="css/css2.css" rel="stylesheet" type="text/css" />

<script language="javascript" type="text/javascript">
function verForm(){
if (document.form1.imagem.value.length < 5)
{
	alert("Favor selecionar uma imagem JPG");
	document.form1.imagem.value.focus();
	document.form1.imagem.focus();
	return false;
	}
}
</script>
</head>

<body>
<form action="sis_script_upload_imagem_home.asp?imagemID=<%=Request.Querystring("imagemID")%>" method="post" enctype="multipart/form-data" name="form1" id="form1" onSubmit="return verForm(this)">
<table width="480" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30" colspan="3" align="center" bgcolor="#9E231B"><font color="white">Publicar Foto - <%=now()%></font></td>
  </tr>
  <tr>
    <td width="50" height="38" align="center" bgcolor="#FFFFFF"><img src="ico/ico_mfoto.gif" width="60" height="60" class="icone" /></td>
    <td width="355" align="center" bgcolor="#FFFFFF"><input name="imagemID" type="hidden" id="imagemID" value="<%=request.querystring("imagemID")%>" />
    <input name="imagemID" type="hidden" id="imagemID" value="<%=request.querystring("imagemID")%>" /></td>
    <td height="38" align="center" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="3" align="left" bgcolor="#FFFFFF">&nbsp;&nbsp;Produto: Cód(<%=rs01.fields.item("imagemID").value%>) - Título: <%=rs01.fields.item("imagemCaminho").value%></td>
  </tr>
  <tr>
    <td colspan="2" align="left" bgcolor="#FFFFFF"><input name="imagem" type="file" id="imagem" size="49" /></td>
    <td align="center" bgcolor="#FFFFFF"><input type="submit" name="Enviar" id="Enviar" value="Enviar" /></td>
  </tr>
  <tr>
    <td height="20" colspan="2" bgcolor="#FFFFFF"><label for="imagem"></label>
    &nbsp;</td>
    <td width="75" align="center" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr>
    <td height="30" colspan="3" align="center" bgcolor="#CCCCCC"><%=msg%></td>
  </tr>
  <tr>
    <td height="2" colspan="3" align="center" bgcolor="#660000"></td>
  </tr>
</table>
</form>
</body>
</html>

<%
rs01.close
set rs01 = nothing
%>

<%
call fechaConexao
%>
