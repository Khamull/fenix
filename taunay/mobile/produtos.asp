
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%option explicit%>

<!--#include file="masterpage.asp"-->
<!--#include file="ProdutosModal.asp"-->

<% Sub HeadPlaceHolder() %>

	<link href="../css/bootstrap-modal.css" rel="stylesheet" />
   	<link href="../css/bootstrap-modal-bs3patch.css" rel="stylesheet" />

<% End Sub %>

<% Sub ContentPlaceHolder() %>
		
			
			<div id="foot-sumItens" class="">

		            <div class="form-group">
		                <table id="tbSumTotal" class="table footer-sum clear">
		                    <tr>
		                    	<td class="txtItens left"> Itens : </td>
		                        <td class="simpleCart_quantity txtQtyItens left  "></td>
		                        <td class="txtValorTotal right"> Valor Total : </td>
		                        <td class="right simpleCart_total txtValorTotalLabel"></td>
		                        </tr>
		                </table>
		            </div> 
	   		</div>
	   		
	   		<div id="detalhesVenda" class="collapse">

						
			</div>

	   		<div class="filter">
				
				<div class="input-group">
							
							<input type="text" class="form-control" name="txtPesquisaProduto" id="txtPesquisaProduto" placeholder="Digite seu produto" data-error="" required>			
							<span class="input-group-btn">
						        <button id="btnProcurar" class="btn btn-default" type="button">OK</button>
						    </span>		
				</div>
			</div>

		    <div class="help-block with-errors"></div>

		   

			<div class="result ">

			</div>
		
		
<% End Sub %>


<% Sub ContentPlaceFooterSum() %>
  
        
<% End Sub %>

<% Sub  ScriptPlaceHolder() %>

	<script type="text/javascript" src="../js/bootstrap-modal.js"> </script>
	<script type="text/javascript" src="../js/bootstrap-modalmanager.js"> </script>
	<script type="text/javascript" src="../js/mobile/produtos.js"> </script>
	<script type="text/javascript" src="../js/mobile/mesas.js"> </script>
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


		