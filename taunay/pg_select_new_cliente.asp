<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%option explicit%>

<!--#include file="Connections/conn.asp" -->

<!--#include file="inc/inc_acesso.inc" -->

<%
Dim keyword
Dim sql
keyword = Request.QueryString("keyword")
sql = "SELECT *, tb_bairro.baiNome from tb_cliente INNER JOIN tb_bairro ON tb_bairro.baiID = tb_cliente.baiID WHERE tb_cliente.cliAtivo = 'N' ORDER BY tb_cliente.cliNome, tb_cliente.baiID ASC" 
if (keyword = "") then
else
sql = "SELECT *, tb_bairro.baiNome from tb_cliente INNER JOIN tb_bairro ON tb_bairro.baiID = tb_cliente.baiID WHERE tb_cliente.cliAtivo = 'S' AND tb_cliente.cliNome Like '%"&keyword&"%' ORDER BY tb_cliente.cliNome, tb_cliente.baiID ASC" 
end if
%>

<%
Dim rs01
Dim rs01_cmd
Dim rs01_numRows

Set rs01_cmd = Server.CreateObject ("ADODB.Command")
rs01_cmd.ActiveConnection = MM_conn_STRING
rs01_cmd.CommandText = sql
rs01_cmd.Prepared = true

Set rs01 = rs01_cmd.Execute
rs01_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 12
Repeat1__index = 0
rs01_numRows = rs01_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rs01_total
Dim rs01_first
Dim rs01_last

' set the record count
rs01_total = rs01.RecordCount

' set the number of rows displayed on this page
If (rs01_numRows < 0) Then
  rs01_numRows = rs01_total
Elseif (rs01_numRows = 0) Then
  rs01_numRows = 1
End If

' set the first and last displayed record
rs01_first = 1
rs01_last  = rs01_first + rs01_numRows - 1

' if we have the correct record count, check the other stats
If (rs01_total <> -1) Then
  If (rs01_first > rs01_total) Then
    rs01_first = rs01_total
  End If
  If (rs01_last > rs01_total) Then
    rs01_last = rs01_total
  End If
  If (rs01_numRows > rs01_total) Then
    rs01_numRows = rs01_total
  End If
End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = rs01
MM_rsCount   = rs01_total
MM_size      = rs01_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rs01_first = MM_offset + 1
rs01_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rs01_first > MM_rsCount) Then
    rs01_first = MM_rsCount
  End If
  If (rs01_last > MM_rsCount) Then
    rs01_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = Server.HTMLEncode(MM_keepMove) & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<!--#include file="inc/inc_formato_data.inc"-->

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
	    <li><a href="pg_insert_cliente.asp">Novo Cliente</a></li>          
	    <li><a href="pg_insert_bairro.asp">Bairro</a></li> 
	    <li><a href="pg_insert_cidade.asp">Cidade</a></li>                      
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc"></div>
    <table width="96%" border="0" align="left" cellpadding="3" cellspacing="3">
  <tr>
    <td width="50" align="center"><img src="ico/ico_pessoafisica.gif" width="60" height="60" class="icone" /></td>
    <td width="751" align="center" class="titulo">NOVOS CLIENTES (SITE)</td>
    </tr>
  <tr>
    <td colspan="2" align="center">
    
    
       <!-- REGISTRO -->
   
    <table width="100%" border="0" align="left" cellpadding="2" cellspacing="2">
      <tr class="textoBranco">
        <td align="left" bgcolor="#FFFFFF">&nbsp;</td>
        <td colspan="3" align="left" bgcolor="#FFFFFF">
          <form name="form2" method="get">
            <input name="keyword" type="text" id="keyword" size="40" maxlength="40" />
            <input name="button" type="submit" class="botao" id="button" value="Enviar" style="height:20px;" />
            </form>
        </td>
        <td colspan="3" align="right"><input type="button" class="botao" value="Lista Clientes Ativos &rarr;" onClick="javascript: window.location.href='pg_select_cliente.asp';"/></td>
        </tr>
      <tr class="textoBranco">
        <td width="30" align="left" bgcolor="#9E231B">ID</td>
        <td width="240" align="left" bgcolor="#9E231B">cliente</td>
        <td width="60" align="left" bgcolor="#9E231B">Telefone</td>
        <td width="240" align="left" bgcolor="#9E231B">Endereço</td>
        <td width="147" align="left" bgcolor="#9E231B">Bairro</td>
        <td width="15" bgcolor="#9E231B">Ed</td>
        <td width="15" bgcolor="#9E231B">St</td>
      </tr>

      <% While ((Repeat1__numRows <> 0) AND (NOT rs01.EOF)) %>
<%
	   Dim iCount
	   Dim cor
	   iCount = iCount + 1
	   if (iCount Mod 2 = 0) Then
	   cor = "#dddddd"
	   else
	   cor = "#eaeaea"
	   end if
	  %>
      <tr>
        <td align="left" bgcolor="<%=cor%>"><%=(rs01.Fields.Item("cliID").Value)%></td>
        <td align="left" bgcolor="<%=cor%>"><%=(rs01.Fields.Item("cliNome").Value)%></td>
        <td align="left" bgcolor="<%=cor%>"><%=(rs01.Fields.Item("cliTelefone").Value)%></td>
        <td align="left" bgcolor="<%=cor%>"><%=(rs01.Fields.Item("cliEndereco").Value)%></td>
        <td align="left" bgcolor="<%=cor%>"><%=(rs01.Fields.Item("baiNome").Value)%></td>
        <td bgcolor="<%=cor%>"><a href="pg_update_cliente.asp?cliID=<%=(rs01.Fields.Item("cliID").Value)%>"><img src="ico/ico_alterar.gif" width="15" height="15" border="0"/></a></td>
        <td bgcolor="<%=cor%>"><img src="ico/ico_<%=(rs01.Fields.Item("cliAtivo").Value)%>.gif" width="15" height="15" border="0"/></td>
      </tr>   <%  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs01.MoveNext()
Wend
%>
      <tr>
        <td colspan="7" align="left"><table border="0" align="center">
            <tr>
              <td><% If MM_offset <> 0 Then %>
                  <a href="<%=MM_moveFirst%>"><img src="First.gif" border="0"></a>
                  <% End If ' end MM_offset <> 0 %></td>
              <td><% If MM_offset <> 0 Then %>
                  <a href="<%=MM_movePrev%>"><img src="Previous.gif" border="0"></a>
                  <% End If ' end MM_offset <> 0 %></td>
              <td><% If Not MM_atTotal Then %>
                  <a href="<%=MM_moveNext%>"><img src="Next.gif" border="0"></a>
                  <% End If ' end Not MM_atTotal %></td>
              <td><% If Not MM_atTotal Then %>
                  <a href="<%=MM_moveLast%>"><img src="Last.gif" border="0"></a>
                  <% End If ' end Not MM_atTotal %></td>
            </tr>
          </table></td>
        </tr>

      


    </table>
    
    
    </table>
    </td>
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
rs01.Close()
Set rs01 = Nothing
%>
