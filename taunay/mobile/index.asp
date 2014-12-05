
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

	Dim rsVenda
	Dim sqlVenda
	set rsVenda = server.CreateObject("adodb.recordset")

	Dim idVenda
	Dim numero
	Dim defineColor

	set rs03 = server.CreateObject("adodb.recordset")

	call abreConexao()

	sql03 = "SELECT tb_mesa.mesNumero as mesNumero FROM tb_mesa WHERE tb_mesa.mesAtiva='S'order by mesNumero asc"

	set rs03 = conn.execute(sql03) 

	 if (not rs03.eof) then

    %>


	<div class="mesasAvailable">

		
		
			<%
			   do while not rs03.eof         

          		numero = rs03("mesNumero")
          		
				sqlVenda = "SELECT venID, tb_mesa.mesNumero as mesNumero FROM tb_venda INNER JOIN tb_mesa ON tb_mesa.mesID = tb_venda.mesID WHERE (tb_venda.staID = 1 OR tb_venda.staID = 4 OR tb_venda.staID = 5 OR tb_venda.staID = 6) AND tipVendaID = 2 AND tb_mesa.mesNumero = " + numero


				set rsVenda = conn.execute(sqlVenda) 

				if (not rsVenda.eof) then

					idVenda =rsVenda("venID")
					defineColor = "vendaAberta"

					%>

						<div id="mesaOpcao" class="floatRight btn-group ">
							 <button type="button" class="btn btn-default dropdown-toggle" data-toggle="dropdown">			      
						      <span class="caret"></span>
						    </button>
						    <ul class="dropdown-menu" role="menu">
						      <li><a id="<%=idVenda%>" href="#" class="btnCancelarVenda">Cancelar Pedido</a></li>	      
						    </ul>
						 </div>	


						<p  id="<%=numero%>" class="<%=defineColor%>" >
						 <label class="labelMesalll" id=<%=idVenda%>> Mesa <%=numero%> </label></p>

						<div id="detalhesVenda_<%=idVenda%>" class="collapse">

						
						</div>

          			<%

	          	else
		          	defineColor="vendaFechada"
		          	%>
						<p id="<%=numero%>" class=<%=defineColor%>><label> Mesa <%=numero%> </label> </p>
          			<%
 
				end if		

				idVenda = ""
          		rs03.moveNext 
			  
			  	Loop
		end if	
	%>
	
	</ul>

</div>
<div style="display:none">
 <div class="simpleCart_shelfItem">
    <span class="item_name" id="item_name"> </span>
    <input type="text" value="1" class="item_Quantity" id="item_Quantity">
    <span class="item_price" id="item_price"></span>
    <span class="item_teste" id="item_teste"></span>
    <a class="item_add" href="javascript:;" id="item_add"> Add to Cart </a>
    <a  class="simpleCart_empty" id="simpleCart_empty" href="javascript:;" class="simpleCart_empty">empty cart</a>
</div>
</div>


<% End Sub %>

<% Sub ContentPlaceFooterSum() %>

        
<% End Sub %>

 <% Sub  ScriptPlaceHolder() %>

    <script type="text/javascript" src="../js/mobile/mesas.js"></script>

	<script type="text/javascript" src="../js/bootstrap-modal.js"> </script>
	<script type="text/javascript" src="../js/bootstrap-modalmanager.js"> </script>

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
		        { attr: "quantity" , label: "Qtd" },
		        { attr: "teste" , label: "Id" },		        
		        { attr: "total" , label: "SubTotal", view: 'currency' },
		        { view: "remove" , text: "Excluir" , label: false }
		    ],


		     currency: "BRL",

		     cartStyle: 'table'
		  
		  });

	</script>

	

<% End Sub %>


