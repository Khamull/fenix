<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%option explicit%>

<!--#include file="inc/inc_conexao.inc"-->

<!--#include file="inc/inc_formato_data.inc"-->

<!--#include file="inc/inc_acesso.inc" -->

<%
call abreConexao()
%>

<%'FAZ A CONSULTA POR MES E ANO

'variaveis
Dim diaIni
Dim diaFim
Dim mes1
Dim ano1
Dim dataIni
DIm dataFim

diaIni = 1
diaFim = 31

if(Not isEmpty(Request.Form("BUSCAR")))Then
mes1 = Request.Form("mes")
ano1 = Request.Form("ano")

else
mes1 = mes
ano1 = ano

end if

'---------- Forma a Data ---------
dataIni = ano1&"-"&mes1&"-"&diaIni
dataFim = ano1&"-"&mes1&"-"&diaFim
'---------------------------------


Dim rs01
Dim sql01

set rs01 = Server.CreateObject("ADODB.Recordset")
sql01 = "SELECT * FROM tb_caixa WHERE status = 'F' AND data BETWEEN '"&dataIni&"' AND '"&dataFim&"'"
set rs01 = conn.execute(sql01)
%>



<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>SISTEM FORTE EM MÍDIA</title>
<link href="css/css1.css" rel="stylesheet" type="text/css" />

<script language="javascript" type="text/javascript">

function verCaixa(caixaID){
	window.open("pg_view_fechamento.asp?caixaID="+caixaID+"&acao=1", "fechamento", "width=800 height=600");	
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
	    <li><a href="pg_caixa.asp">Caixa</a></li>        
               
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc"></div>
    <table width="96%" border="0" align="left" cellpadding="3" cellspacing="3">
  <tr>
    <td width="50" align="center"><img src="ico/ico_caixa.png" width="246" height="246" class="icone" /></td>
    <td width="751" align="center" class="titulo">CAIXAS FECHADOS</td>
    </tr>
  <tr>
    <td align="center">&nbsp;</td>
    <td align="center">
      <table width="437" border="0" cellpadding="2" cellspacing="1">
        <tr>
          <td height="25" colspan="5" align="center" bgcolor="#EEEEEE">
           <form method="post" action="pg_caixa_fechado.asp">
            <table width="390">
             <tr>
              <td width="136">PESQUISAR POR MÊS</td>
              <td width="109">
               <select name="mes">
                <option value="01" selected="selected">JANEIRO</option>
                <option value="02">FEVEREIRO</option>
                <option value="03">MARÇO</option>
                <option value="04">ABRIL</option>
                <option value="05">MAIO</option>
                <option value="06">JUNHO</option>
                <option value="07">JULHO</option>
                <option value="08">AGOSTO</option>
                <option value="09">SETEMBRO</option>
                <option value="10">OUTUBRO</option>
                <option value="11">NOVEMBRO</option>
                <option value="12">DEZEMBRO</option>
               </select>
              </td>
              <td width="54">
               <select name="ano">
                <option value="<%=Year(now())-4%>"><%=Year(now())-4%></option>
                <option value="<%=Year(now())-3%>"><%=Year(now())-3%></option>
                <option value="<%=Year(now())-2%>"><%=Year(now())-2%></option>
                <option value="<%=Year(now())-1%>"><%=Year(now())-1%></option>
                <option value="<%=Year(now())%>" selected="selected"><%=Year(now())%></option>
               </select>
              </td>
              <td width="71"><input type="submit" name="BUSCAR" value="Buscar" /></td>
             </tr>
            </table>
           </form>
          </td>
          </tr>
        <tr>
          <td width="84" align="right"></td>
          <td width="98" height="25" align="right"></td>
          <td width="93" align="left"></td>
          <td width="96" height="25" align="left"></td>
        </tr>
        <tr bgcolor="#333333" class="textoBranco">
          <td align="left"><strong>Data</strong></td>
          <td height="25" align="left"><strong>Valor inicial</strong></td>
          <td align="left"><strong>Vendas</strong></td>
          <td align="left"><strong>Valor Final</strong></td>
          <td width="0" height="25" align="center"></td>
        </tr>
        
        <%if (not rs01.EoF) Then%>
         
         <%Dim inicio, vendas, total%>
         
         <%Dim cor, i%>
         
        <%While not rs01.EoF%>
        
        <%
		 if(i mod 2 = 0)Then
		  cor = "#FFFFFF"
		 else
		  cor = "#DDEEFF"
		 end if
		%>
        
        <tr>
          <td height="25" align="left" bgcolor="<%=cor%>"><%=rs01.fields.item("data").value%></td>
          <td align="left" bgcolor="<%=cor%>"><%=FormatCurrency(rs01.fields.item("valorInicial").value)%></td>
          <td align="left" bgcolor="<%=cor%>">
		  <%
		  inicio = rs01.fields.item("valorInicial").value
		  total  = rs01.fields.item("valorFinal").value
		  vendas = (total - inicio)
		  
		  Response.Write(FormatCurrency(vendas))
		  %>
          </td>
          <td align="left" bgcolor="<%=cor%>"><%=FormatCurrency(rs01.fields.item("valorFinal").value)%></td>
          <td align="center" bgcolor="<%=cor%>">
           <a href="javascript: verCaixa(<%=rs01.fields.item("caixaID").value%>);">
            <img src="ico/ico_lupa.gif" width="20" height="20" border="1" title="Visualizar"/>
           </a>
          </td>
        </tr>
        
        <%i = (i+1)%>
        
        <%
		 rs01.MoveNext
		Wend
		%>
        
        <%else%>
                
        <tr>      
         <td height="35" colspan="5" align="center">Nenhum Registro para esse mês</td>
        </tr>
        
        <%end if%>
        
        <tr>
          <td align="center"></td>
          <td height="25" align="center"></td>
          <td align="center">&nbsp;</td>
          <td height="25">&nbsp;</td>
          <td></td>
          </tr>
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
call FechaConexao()
%>