<%@LANGUAGE="VBSCRIPT" CODEPAGE="28592"%>
<%option explicit%>
<!--#include file="inc/inc_conexao.inc"-->

<!--#include file="Connections/conn.asp" -->

<!--#include file="inc/inc_formato_data.inc"-->

<!--#include file="inc/inc_acesso.inc" -->
<%
Call abreConexao()
%>
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%

If (CStr(Request("MM_update")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_conn_STRING
    MM_editCmd.CommandText = "UPDATE tb_home SET homeTitulo = ?, homeTexto = ? WHERE homeID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 201, 1, 65535, Request.Form("homeTitulo")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 201, 1, 65535, Request.Form("homeTexto")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "sis_update_home.asp?cadastrado=OK"
    Response.Redirect(MM_editRedirectUrl)
  End If
End If
%>

<%
Dim Recordset2
Dim Recordset2_cmd
Dim Recordset2_numRows

Set Recordset2_cmd = Server.CreateObject ("ADODB.Command")
Recordset2_cmd.ActiveConnection = MM_conn_STRING
Recordset2_cmd.CommandText = "SELECT * FROM tb_home ORDER BY homeID ASC" 
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
Dim homeID

    acao = Request.QueryString("acao")
    homeID = Request.QueryString("homeID")
if(acao = "1") then 

set rs1 = Server.CreateObject("ADODB.Recordset")
sql01   = "DELETE FROM tb_home WHERE homeID = '"&homeID&"'"
set rs1 = conn.execute(sql01)
response.Redirect("sis_update_home.asp")

elseif(acao = "2") Then

set rs02 = Server.CreateObject("ADODB.Recordset")
sql02   = "INSERT INTO tb_home (homeTitulo, homeTexto) VALUES ('TITULO','TEXTO')"
set rs02 = conn.execute(sql02)

Response.Redirect("sis_update_home.asp")
end if

%>

<%
Dim msg

msg = request.QueryString("cadastrado")

if (msg = "OK") then
msg = "Texto atualizado com Sucesso!"
end if
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-2" />
<title>SISTEMA WEB DE GEST&Atilde;O</title>
<link href="css/css1.css" rel="stylesheet" type="text/css" />

<script src="js/nicEdit.js" type="text/javascript"></script>
<script type="text/javascript">
	bkLib.onDomLoaded(function() {
	new nicEditor({maxHeight : 250}).panelInstance('area1');
});
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
    <div style="height:25px; line-height:25px; background:#ccc"><strong>HOME</strong></div>
    <div><!--<a href="javascript: inserir()"><input type="button" value="+ Novo T&oacute;pico" class="botao" border="0"/></a>--><a href="sis_update_imagem_home.asp"><input type="button" value="+ Editar Imagens" class="botao" border="0"/></a><!--<a href="sis_slides.asp"><input type="button" value="+ Editar Banners" class="botao" border="0"/></a>--></div>
    
      
        
        <% While Not Recordset2.EoF %>
        <form  name="form1" id="form1" action="" method="post">
        <table border="0" align="center">
        <tr valign="baseline">
          <td colspan="3" align="center" nowrap="nowrap"><font color="#FF0000"><b><%=msg&"<br><br>"%></b></font></td>
        </tr>
        <tr valign="baseline">
          <td>&nbsp;</td>
          <td align="center" nowrap="nowrap" bgcolor="#FFA500">SOBRE</td>
          <td align="center" nowrap="nowrap">&nbsp;</td>
        </tr>
        <tr valign="baseline">
          <td height="16" align="left" valign="top" nowrap="nowrap">&nbsp;</td>
          <td align="left"><input type="hidden" name="homeTitulo" value="titulo" size="52"/></td>
          <td valign="top"><!--<a href="javascript: excluir(<%=(Recordset2.Fields.Item("homeID").Value)%>)"><img src="ico/ico_excluir.gif" width="15" height="15" border="0" title="Excluir"/></a>--></td>
        </tr>
        <tr valign="baseline">
          <td align="left" valign="top" nowrap="nowrap"><strong>Texto:</strong></td>
          <td align="left">
          <textarea name="homeTexto" cols="120" rows="15" id="area1"><%=(Recordset2.Fields.Item("homeTexto").Value)%></textarea></td>
          <td valign="top">&nbsp;
         
          </td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">&nbsp;</td>
          <td align="center"><input type="submit" value="Atualizar" class="botao"/></td><td></td>
        </tr>
        <input type="hidden" name="MM_update" value="form1" />
      <input type="hidden" name="MM_recordId" value="<%= Recordset2.Fields.Item("homeID").Value %>" />
    </form>
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
