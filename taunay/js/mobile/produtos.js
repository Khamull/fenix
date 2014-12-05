$(document).ready(function() {

 if (GetURLParameter("venID")=="" && GetURLParameter("hasItem")==""){

    simpleCart.empty();
 }

	$("#btnProcurar").unbind("click").click(function(){

  $(".help-block").html("");

	 var response = new Object();

       $.ajax({
       type:"Post",
       url:"pedidos.asp?",
       data: "descProduto=" + encodeURIComponent($("#txtPesquisaProduto").val()) + "&acao=1" ,
       async:true,
       cache:false,
       success:function(result) {

            var data = result.split(",");

            response.Success = data[0].split(":");
            response.Error = data[1].split(":");
            response.Msg = data[2].split(":");
            response.Qtd = data[3].split(":");
            response.Result = data[4].split(":");

             if (response.Success[1] =="true") {

                if (response.Qtd[1] != "0")
                {
                
                  var HTMLString = $.parseHTML(response.Result[1]);

                  myResult = $("<div />").html(HTMLString).text();

                   $(".result").html("");
                   $(".result").append(myResult);


                    $("#tbProdutos tr").each(function(){

                        $(this).on("click", function(){

                       
                          $("#txtQtySelectProduct").val("");
                          $("#hddProdNome").html($(this).children('td').eq(0).html());
                          $("#hddProdNPreco").html($(this).children('td').eq(1).html());
                          $("#qtdEstoque").html($(this).children('td').eq(3).html());
                          $("#hddIdProd").val($(this).attr('id'));
                          $(".item_teste").html($(this).attr('id'));
                          


                          $("#mdlistaProdutosPedidos").css("display","none");

                          $("#confirmaMesa").css("display","none");
                          $("#mdSelecionaProduto").removeClass("fade");
                          $("#detalhesVenda").append($("#mdSelecionaProdutoBody").html());
                          $("#mdSelecionaProduto").removeClass("fade")
                          
                          $(".filter").css("display","none");
                          $(".result").css("display","none");


                         });

                           
                        });


                 }
                 else{

                    $(".help-block").html("<ul class='list-unstyled'><li>Sua busca não retornou nenhum produto</li></ul>");

                 }

             }else{

                $(".help-block").html("<ul class='list-unstyled'><li>"+ response.Msg[1] + "</li></ul>");

             }
        }
    });

	});

   
    $(".footer-sum").unbind("click").click(function()
    {
      if($("#detalhesVenda").hasClass("in")) {
          
          $("#detalhesVenda").collapse("hide");
          $(".filter").css("display","block");
          $(".result").css("display","block");

      }else{
        if ($(this).find(".txtQtyItens").html() !="0")
        {      

                $(".filter").css("display","none");
                $(".result").css("display","none");

                $("#mdlistaProdutosPedidos").removeClass("fade")
                $("#detalhesVenda").append($("#mdlistaProdutosPedidos"));
                $("#mesaNumber").css("display", "block");

                $("#btnSelecionaMaisProduto").css("display", "none");
                
                $("#btnSelecionaMaisProduto").unbind("click").click(function(){

                   window.location.href = "produtos.asp?mesaId=" + $("#hddMesaId").val() + "&venID=" + $("#hddVenId").val()

                });

                $("#btnConfirmaPedido").unbind("click").click(function(){


                   var  venId = $("#hddVenId").val()
                   var  mesaID = $("#hddMesaId").val()

                   var response = new Object();

                    if (mesaID !="undefined" &&  mesaID !="" ){

                       LancarVenda(mesaID,venId );

                    }else{

                      venId = "";
                      

                       $.ajax({
                        type:"Post",
                        url:"pedidos.asp?",
                        data: "acao=5" ,
                        async:true,
                        cache:false,
                        success:function(result) {

                            var data = result.split(",");

                            response.Success = data[0].split(":");
                            response.Error = data[1].split(":");
                            response.Msg = data[2].split(":");
                            response.Result = data[3].split(":");

                            if (response.Success[1] =="true")
                            {

                              var HTMLString = $.parseHTML(response.Result[1]);

                              myResult = $("<div />").html(HTMLString).text();

                              $(".comboMesas").html("");
                              $(".comboMesas").append(myResult);

                              $("#mdlistaProdutosPedidos").css("display","none");
                              $("#detalhesVenda").append($("#mdSelecioneAMesa").html());
                              $("#confirmaMesa").removeClass("fade")
                              
                              $(".filter").css("display","none");
                              $(".result").css("display","none");


                               $("#btnLancarPedido").unbind("click").click(function(){

                                  
                                 if ($(".simpleCart_grandTotal").html()=="R$0.00")
                                          {
                                              alert("Não é possivel confirmar um pedido sem produtos.");
                                              return ;
                                          }
                                          
                                  LancarVenda($("#slcMesa").val(),venId );
                                   
                               });

                                $("#cancelSelecaoMesa").unbind("click").click(function(){

                                  
                                  window.location.href = "produtos.asp?mnu=2&mesaId=&venID&hasItem=true" 
                                   
                               });


                            }

                          }
                        });


                    }
                     
                 });

                $("#detalhesVenda").collapse("show");

        }else{

          alert("Nenhm produto selecionado")

        }
      }

     });


   



});


function LancarVenda(mesa, venId){

        var stringIdEQty= "";
        var mesaId = mesa;
        var venId = venId;
   
        var response = new Object();

        $(".itemRow").each(function(){
            
            stringIdEQty = stringIdEQty +  $(this).find("td.item-teste").html() + "-" + $(this).find("td.item-quantity").html() + "," ;

        });

        $.ajax({
        type:"Post",
        url:"pedidos.asp?",
        data: "stringIdEQty="+ stringIdEQty + "&mesaId="+ mesaId +"&acao=2" + "&venId=" + venId ,
        async:true,
        cache:false,
        success:function(result) {

            var data = result.split(",");

            response.Success = data[0].split(":");
            response.Error = data[1].split(":");
            response.venId = data[2].split(":");
            response.Msg = data[3].split(":");  

            if (response.Success[1] =="true")
            {
              $("#hddVenId").val(response.venId[1]);

              alert("Pedido confirmado com sucesso");

              if (GetURLParameter("mnu") == "1" )
              {

                $(".mesasAvailable").find(".collapse").collapse('hide');

              }else{

                 simpleCart.empty();
                 window.location.href = "produtos.asp?mnu=2";

              }
              
            }

          }
        });

}


function GetURLParameter(sParam)
{
    var sPageURL = window.location.search.substring(1);
   
    var sURLVariables = sPageURL.split('&');
    for (var i = 0; i < sURLVariables.length; i++) 
    {
        var sParameterName = sURLVariables[i].split('=');
        if (sParameterName[0] == sParam) 
        {
            return sParameterName[1];
        }
    }
  
    return "";
  
}     

$('td.item-quantity').click(function()
  {
    var span = $(this);
    var text = span.text();

    var new_text = prompt("Altere a quantidade", text);

    if (new_text != null)
      span.text(new_text);
  });


function cancelar(){

     window.location.href = "produtos.asp?mnu=2&mesaId="+GetURLParameter("mesaId")+"&venID="+GetURLParameter("venID")+"&hasItem=true" 
  

}

function confirmarQuantidade(){


    window.location.href = "produtos.asp?mnu=2&mesaId="+GetURLParameter("mesaId")+"&venID="+GetURLParameter("venID")+"&hasItem=true" 

}