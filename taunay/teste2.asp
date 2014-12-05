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
Dim Recordset2
Dim Recordset2_cmd
Dim Recordset2_numRows

Set Recordset2_cmd = Server.CreateObject ("ADODB.Command")
Recordset2_cmd.ActiveConnection = MM_conn_STRING
Recordset2_cmd.CommandText = "SELECT * FROM tb_empresa ORDER BY empresaID ASC" 
Recordset2_cmd.Prepared = true

Set Recordset2 = Recordset2_cmd.Execute
Recordset2_numRows = 0
%>
<%

If (CStr(Request("MM_update")) = "form1") Then
 
    ' execute the update
    Dim MM_editCmd
	
	Dim empresaTitulo
	Dim empresaTexto
	DIm MM_recordId
	
	    empresaTitulo = Replace(Trim(Request.Form("empresaTitulo")),"'","")
	    empresaTexto  = Replace(Trim(Request.Form("empresaTexto")),"'","")
		MM_recordId   = Request.Form("MM_recordId")
		
			
		Dim rs020
		Dim sql020
		
		set rs020 = Server.CreateObject("ADODB.Recordset")
		sql020 = "UPDATE tb_empresa SET empresaTitulo = '"&empresaTitulo&"', empresaTexto = '"&empresaTexto&"' WHERE empresaID = '"&MM_recordId&"'" 
		
		set rs020 = conn.Execute(sql020)
		
				set rs020 = nothing

			

    'Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    'MM_editCmd.ActiveConnection = MM_conn_STRING
    'MM_editCmd.CommandText = "UPDATE tb_empresa SET empresaTitulo = ?, empresaTexto = ? WHERE empresaID = ?" 
    'MM_editCmd.Prepared = true
    'MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 201, 1, 65535, Request.Form("empresaTitulo")) ' adLongVarChar
    'MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 201, 1, 65535, Request.Form("empresaTexto")) ' adLongVarChar
    'MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
   ' MM_editCmd.Execute
    'MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "teste2.asp?cadastrado=OK"
    Response.Redirect(MM_editRedirectUrl)
  
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-2">
<title>Untitled Document</title>
<link href="css/css1.css" rel="stylesheet" type="text/css" />
<script src="js/nicEdit.js" type="text/javascript"></script>
<script type="text/javascript">
	bkLib.onDomLoaded(function() {
	new nicEditor({maxHeight : 250}).panelInstance('area1');
});
</script>
</head>

<body>
<% While Not Recordset2.EoF %>
        <form  name="form1" id="form1" action="" method="post">
        <tr valign="baseline">
          <td height="16" align="left" valign="top" nowrap="nowrap"><strong>Titulo</strong></td>
          <td align="left"><input type="text" name="empresaTitulo" value="<%=(Recordset2.Fields.Item("empresaTitulo").Value)%>" size="52"/></td>
          <td valign="top"></td>
        </tr>
        <tr valign="baseline">
          <td align="left" valign="top" nowrap="nowrap"><strong>Texto:</strong></td>
          <td align="left">
          <textarea name="empresaTexto" cols="110" rows="15" id="area1"><%=(Recordset2.Fields.Item("empresaTexto").Value)%></textarea></td>
          <td valign="top">&nbsp;
         
          </td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">&nbsp;</td>
          <td align="center"><input type="submit" value="Atualizar" class="botao"/></td><td></td>
        </tr>
        <input type="hidden" name="MM_update" value="form1" />
      <input type="hidden" name="MM_recordId" value="<%= Recordset2.Fields.Item("empresaID").Value %>" />
    </form>
     <% 
		     Recordset2.moveNext()
			Wend
		   %>

      </table>
</body>
</html>
