<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%option explicit%>

<!--#include file="inc/inc_conexao.inc"-->

<!--#include file="inc/inc_formato_data.inc"-->

<!--#include file="inc/inc_acesso.inc" -->

<!--#include file="Connections/conn.asp" -->

<%
Dim sql
Dim keyword
Dim filtro

keyword = Request.querystring("keyword")
filtro = Request.querystring("filtro")

if (keyword = "") then

sql = "SELECT tb_venda.venID, tb_venda.usuLogin, tb_venda.venData, tb_venda.venHoraA, tb_venda.tipVendaID, tb_venda.staID, tb_venda.cliID, tb_cliente.cliTelefone, tb_cliente.cliNome FROM tb_venda INNER JOIN tb_cliente ON tb_cliente.cliID = tb_venda.cliID WHERE (tb_venda.staID = 1 OR tb_venda.staID = 4 OR tb_venda.staID = 5 OR tb_venda.staID = 6) AND tb_venda.tipVendaID = 3 GROUP BY tb_venda.venID ORDER BY tb_venda.venID DESC" 
else 

	select case filtro
	
	case 3
	sql = "SELECT tb_venda.venID, tb_venda.usuLogin, tb_venda.venData, tb_venda.venHoraA, tb_venda.tipVendaID, tb_venda.staID, tb_venda.cliID, tb_cliente.cliTelefone, tb_cliente.cliNome FROM tb_venda INNER JOIN tb_cliente ON tb_cliente.cliID = tb_venda.cliID WHERE (tb_venda.staID = 1 OR tb_venda.staID = 4 OR tb_venda.staID = 5 OR tb_venda.staID = 6) AND tb_venda.venID '"&keyword&"' AND tb_venda.tipVendaID = 3 GROUP BY tb_venda.venID ORDER BY tb_venda.venID DESC" 
	case 4
	sql = "SELECT tb_venda.venID, tb_venda.usuLogin, tb_venda.venData, tb_venda.venHoraA, tb_venda.tipVendaID, tb_venda.staID, tb_venda.cliID, tb_cliente.cliTelefone, tb_cliente.cliNome FROM tb_venda INNER JOIN tb_cliente ON tb_cliente.cliID = tb_venda.cliID WHERE (tb_venda.staID = 1 OR tb_venda.staID = 4 OR tb_venda.staID = 5 OR tb_venda.staID = 6) AND tb_venda.usuLogin LIKE '%"&keyword&"%' AND tb_venda.tipVendaID = 3 GROUP BY tb_venda.venID ORDER BY tb_venda.venID DESC" 
	end select 
	
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

Repeat1__numRows = 15
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
	    <li><a href="pg_menu_pedidos.asp">Novo Pedido</a></li>
        <li><a href="pg_pedidos_fechados_balcao.asp">Pedidos Fechados</a></li>
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc"></div>
    <table width="96%" border="0" align="left" cellpadding="3" cellspacing="3">
      <tr>
        <td width="50" align="center"><img src="ico/ico_cesta.gif" width="60" height="60" class="icone" /></td>
        <td width="751" align="center" class="titulo">PEDIDOS - BALCÃO</td>
      </tr>
      <tr>
        <td height="134" colspan="2" align="left" valign="top"><!-- REGISTRO -->
          <table width="100%" border="0" align="left" cellpadding="2" cellspacing="2">
            <tr>
              <td colspan="6" align="left" valign="middle"><form method="get" name="form2" id="form2">
                <input name="keyword" type="text" id="keyword" size="40" maxlength="40" />
                <input name="buscar" type="submit" class="botao" id="buscar" value="Buscar" style="height:20px;" /> 
                &nbsp;
              Buscar por: 
              <label>
                <select name="filtro" id="filtro">
<option value="3">Nº do Pedido</option>
<option value="4">Atendente</option>
<option selected="selected"> </option>
                </select>
              </label>
              </form></td>
            </tr>
            <tr class="textoBranco">
              <td width="54" align="left" bgcolor="#9E231B">&nbsp;Codigo</td>
              <td width="71" align="left" bgcolor="#9E231B">Data</td>
              <td width="74" align="left" bgcolor="#9E231B">Hora</td>
              <td width="522" align="left" bgcolor="#9E231B">Atendente&nbsp;</td>
              <td width="15" bgcolor="#9E231B">Ed</td>
              <td width="17" bgcolor="#9E231B">St</td>
            </tr>
            <% 
While ((Repeat1__numRows <> 0) AND (NOT rs01.EOF)) 
%>

			<%if (rs01.Fields.Item("staID").Value = "1") Then%>

              <tr>
                <td align="left"><%=(rs01.Fields.Item("venID").Value)%></td>
                <td align="left"><%=(rs01.Fields.Item("venData").Value)%></td>
                <td align="left"><%=(rs01.Fields.Item("venHoraA").Value)%></td>
                <td align="left"><%=(rs01.Fields.Item("usuLogin").Value)%></td>
                <td><a href="pg_insert_itemVendaBalcao1.asp?venID=<%=(rs01.Fields.Item("venID").Value)%>" target="_top"><img src="ico/ico_olho.gif" width="15" height="15" border="0"/></a></td>
                <td><img src="ico/<%=(rs01.Fields.Item("staID").Value)%>.gif" width="15" height="15" border="0"/></td>
              </tr>
              
           <%end if%>
           
              <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs01.MoveNext()
Wend
%>
<tr>
  <td colspan="6" align="center">
    <table border="0">
      <tr>
        <td><% If MM_offset <> 0 Then %>
            <a href="<%=MM_moveFirst%>"><img src="First.gif"></a>
            <% End If ' end MM_offset <> 0 %></td>
        <td><% If MM_offset <> 0 Then %>
            <a href="<%=MM_movePrev%>"><img src="Previous.gif"></a>
            <% End If ' end MM_offset <> 0 %></td>
        <td><% If Not MM_atTotal Then %>
            <a href="<%=MM_moveNext%>"><img src="Next.gif"></a>
            <% End If ' end Not MM_atTotal %></td>
        <td><% If Not MM_atTotal Then %>
            <a href="<%=MM_moveLast%>"><img src="Last.gif"></a>
            <% End If ' end Not MM_atTotal %></td>
      </tr>
    </table></td>
            </tr>
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
rs01.Close()
Set rs01 = Nothing
%>
