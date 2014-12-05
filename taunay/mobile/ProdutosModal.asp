
<div id="mdSelecionaProdutoBody">
  <div id="mdSelecionaProduto" class="fade canceLDelete simpleCart_shelfItem" tabindex="-1" data-focus-on="input:first">
  
    <div class="modal-body ">
        <input class="hddIdProd item_id"type="hidden" name="hddIdProd" id="hddIdProd">
        <span style="display:none"class="item_teste" id="hddId">  </span>
        <div>
            <span class="txtModalSelect" > Produto : </span> 
            <span class="item_name" id="hddProdNome"> Item </span>
        </div>
        
         <div>
          <span class="txtModalSelect">Preço: </span>
          <span class="item_price" id="hddProdNPreco"></span>      
        </div>

        <div>
          <span class="txtModalSelect">Quantidade Estoque: </span>
          <span class="qtdEstoque" id="qtdEstoque"></span>    
        </div>
        <br>

        <div class="center">
          <span id="spnInformeQtd"class="txtModalSelectb">Informe a Quantidade: </span>
          <input id="txtQtySelectProduct"type="text" class="item_Quantity" maxlength="2"><br>
        </div>       
    </div>


  <div class="modal-footer">  
    
     <a class="item_add btn btn-primary"  href="javascript: confirmarQuantidade();"> Confirmar </a>
     <input id="cancelQtyProduto" type="button" class="btn " value="Cancelar" onclick="cancelar();"/>    

  </div>
</div>



<div id="mdlistaProdutosPedidos" class="fade" tabindex="-1" data-focus-on="input:first">
  <input class="hddVenId"type="hidden" name="hddVenId" id="hddVenId" value=<%=lcase(Request("venID"))%>>
  <input class="hddMesaId"type="hidden" name="hddMesaId" id="hddMesaId" value=<%=lcase(Request("mesaId"))%>>
    <div class="modal-body">      
    <table id='tbProdutosSelecionados' class='table table-hover'>

      <!-- create a checkout button -->
          <a href="javascript:;" class="simpleCart_checkout"></a>
          <!-- button to empty the cart -->
          <a href="javascript:;" class="simpleCart_empty"></a>
          <!-- show the cart -->
          <div class="simpleCart_items"></div>
          <!-- cart total (ex. $23.11)-->
          
      </table>         
    
    </div>
    <%

        if(lcase(Request("mesaId")) <> "") then

          %><div id="mesaNumber" style="display:none"> Mesa <%=lcase(Request("mesaId"))%></div>

      <% end if %>
    
  <div class="modal-footer">
    <div class="soma-total">
      <span class="left">Total : </span>
      <div class="simpleCart_grandTotal right"></div>
    </div>
    <input type="button" class="btn btn-success" id="btnSelecionaMaisProduto" value="Adicionar"/>
    <input type="button" class="btn btn-primary" id="btnConfirmaPedido" value="Confirmar"/>    
  </div>
</div>

<div id="mdSelecioneAMesa">
    <div id="confirmaMesa"  class="canceLDelete fade" tabindex="-1" data-focus-on="input:first">    
       <div class="modal-header">
          <span>Selecione a mesa :</span> 
        </div>
          <div class="modal-body">

            <div class="comboMesas">

            </div>   
          
          </div>
        <div class="modal-footer">
          <input  id="btnLancarPedido" type="button" class="btn btn-primary"  value="Lançar Pedido"/>
          <input id="cancelSelecaoMesa" type="button" class="btn" value="Cancelar"/>
        </div>
    </div>
</div>


  <div id="confirmCancelBody">
    <div id="confirmCancel"  class="canceLDelete " tabindex="-1" data-focus-on="input:first" style="display:none">
       <input class="venIdCancela"type="hidden"/>
       <div class="modal-header">      
          Deseja cancelar o pedido?
      </div>
      <div class="modal-footer">
          <input id="btnCancelarPedidoConf" type="button" class="btn btn-danger" value="Sim"/>
          <input id="btnFecharNao"type="button" class="btn btn-default"  value="Não"/>
      </div>
  </div>
</div>





<script type="text/javascript">

 


</script>