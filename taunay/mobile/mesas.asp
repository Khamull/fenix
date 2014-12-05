
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%option explicit%>

<!--#include file="masterpage.asp"-->
<!--#include file="ProdutosModal.asp"-->


<% Sub HeadPlaceHolder() %>
	<link href="../css/bootstrap-modal.css" rel="stylesheet" />
   	<link href="../css/bootstrap-modal-bs3patch.css" rel="stylesheet" />
<% End Sub %>

<% Sub ContentPlaceHolder() 


	Dim rs03
	Dim sql03
	Dim idVenda
	Dim numero

	set rs03 = server.CreateObject("adodb.recordset")

	call abreConexao()

	sql03 = "SELECT venID, tb_mesa.mesNumero as mesNumero FROM tb_venda INNER JOIN tb_mesa ON tb_mesa.mesID = tb_venda.mesID WHERE (tb_venda.staID = 1 OR tb_venda.staID = 4 OR tb_venda.staID = 5 OR tb_venda.staID = 6) AND tipVendaID = 2"

	set rs03 = conn.execute(sql03) 

	 if (not rs03.eof) then

    %>


<div class="navMesas">

	<ul class="nav nav-pills nav-stacked center">
		
			<%
			   do while not rs03.eof         

          		numero = rs03("mesNumero")
          		idVenda= rs03("venID")

          		%>
          		<li id=<%=idVenda%>><a id=<%=numero%> href="#"> Mesa <%=numero%></a> </li>
          		<%
			  	rs03.moveNext 
			  
			  	Loop
		end if	
	%>
	
	</ul>

</div>


<% End Sub %>

<% Sub ContentPlaceFooterSum() %>

        
<% End Sub %>

 <% Sub  ScriptPlaceHolder() %>

    <script type="text/javascript" src="../js/mobile/mesas.js"></script>
	<script type="text/javascript" src="../js/bootstrap-modal.js"></script>
	<script type="text/javascript" src="../js/bootstrap-modalmanager.js"></script>
    <script type="text/javascript" src="../js/simpleCart.js"></script>
	<script>

		  simpleCart({
		    checkout: {
		      type: "PayPal",
		      email: "you@yours.com"
		    }, 


     	     cartColumns: [
		        { attr: "name" , label: "Produto" },
		        { attr: "price" , label: "Preço", view: 'currency' },
		        { attr: "quantity" , label: "Qty" },
		        { attr: "teste" , label: "Id" },		        
		        { attr: "total" , label: "SubTotal", view: 'currency' },
		        { view: "remove" , text: "Excluir" , label: false }
		    ],


		     currency: "BRL",

		     cartStyle: 'table'
		  
		  });

	</script>

	

<% End Sub %>


