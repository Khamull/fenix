$(document).ready(function() {


	$(".mesasAvailable p label").each(function(){

        var response = new Object();

        var label = $(this).parent();

		$(this).on("click", function(){

			var detalhes = $("#detalhesVenda_"+ $(this).attr('id'));

			if(detalhes.hasClass("in"))
			{
				detalhes.collapse('hide');				
			}
			else{

				if ($(this).attr('id') != undefined)
		        {

		        	$(".mesasAvailable").find(".collapse").not(detalhes).collapse('hide');
		        	

					var venID = $(this).attr('id');
					var mesaMessage = $(this).html();
					var mesaID = $(this).parent().attr("id");

			        $.ajax({
			        type:"Post",
			        url:"pedidos.asp?",
			        data: "venID="+ venID + "&acao=3" ,
			        async:true,
			        cache:false,
			        beforeSend: function(){

			        	detalhes.collapse("show");
			        	label.addClass("onloading");

			        },
			        success:function(result) {
			         
			            var data = result.split(",");

			            response.Success = data[0].split(":");
			            response.Error = data[1].split(":");
			            response.Msg = data[2].split(":");
			            response.Result = data[3].split(":");

			            if (response.Success[1] =="true")
			            {
			      
			          		var produtos = response.Result[1].split("|");
			          		simpleCart.empty();

			          		produtos.pop();
			          		for(var i in produtos){

			          			var values = produtos[i].split("-")
			          			
				          			$("#item_name").html(values[0]);
				          			$("#item_price").html(values[1]);
				          			$("#item_Quantity").val(values[2]);
				          			$("#item_teste").html(values[3]);
				          			
				          			$("#item_add").click();

			          		}

			              	$("#tbProdutosSelecionados").html("");	

			              	$("#hddVenId").val(venID)
			              	$("#hddMesaId").val(mesaID)
			              	
			              	$('td.item-quantity').click(function()
						  	{
						    	var span = $(this);
						    	var text = span.text();
							    var new_text = prompt("Altere a quantidade", text);

							    if (new_text != null)
									span.text(new_text);
						  	});

			              	$("#mdlistaProdutosPedidos").removeClass("fade")

						  	$("#detalhesVenda_" + venID ).append($("#mdlistaProdutosPedidos"));
						  	

							$("#btnSelecionaMaisProduto").unbind("click").click(function(){

							   window.location.href = "produtos.asp?mnu=2&mesaId=" + $("#hddMesaId").val() + "&venID=" + $("#hddVenId").val()

							});

							 $("#btnConfirmaPedido").unbind("click").click(function(){


								 if ($(".simpleCart_grandTotal").html()=="R$0.00")
				                  {
				                      alert("Não é possivel confirmar um pedido sem produtos.");
				                      return ;
				                  }

						       var  venId = $("#hddVenId").val()
						       var  mesaID = $("#hddMesaId").val()

						       var response = new Object();

						        if (venId == "" && mesaID==""){

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

						                  var $modal = $('#mdSelecioneAMesa');

						                  $modal.modal();
						                  $modal.removeClass("hide");
						                }

						              }
						            });
									

						        }else{

						          LancarVenda(mesaID,venId );

						        }
						         
						     });

							label.removeClass("onloading");
							detalhes.collapse("show");


		            	}
					}		          
					});

				} else
				{
					var mesaNumero = $(this).closest('p').attr('id');
					$("#simpleCart_empty").click();
					window.location.href = "produtos.asp?mesaId=" + mesaNumero 
				}
			}

		});
	});




$('.btnCancelarVenda').click(function()
  {

		var venID = $(this).attr('id');

		$("#detalhesVenda_" + venID ).html($("#confirmCancelBody").html());
		
		$("#detalhesVenda_" + venID ).collapse("show");
		$("#confirmCancel").css("display","block");

		$('#btnFecharNao').click(function()
		{

			$(".mesasAvailable").find(".collapse").collapse('hide');
			 window.location.reload();

		});


		$('#btnCancelarPedidoConf').click(function()
		  {

		       var response = new Object();

		      $(this).each(function(){

		            $.ajax({
		            type:"Post",
		            url:"pedidos.asp?",
		            data: "venID="+ venID + "&acao=4" ,
		            async:true,
		            cache:false,
		            success:function(result) {
		             
		                var data = result.split(",");

		                response.Success = data[0].split(":");
		                response.Error = data[1].split(":");
		                response.Msg = data[2].split(":");		                

		                if (response.Success[1] =="true")
		                {
		          
		                  alert(response.Msg[1]);

		                $(".mesasAvailable").find(".collapse").collapse('hide');
		                window.location.reload();
		              }

		            }
		        });
		    });
		});

	});



 $(".navbar-nav li").each(function(){

  		$(this).click(function(){
  			var url = $(this).children(0).attr("data-link");
  			$(".navbar-nav li").removeClass("active");
  			$(this).addClass("active");

  			if  ($(this).attr("id") != "3"){
  				window.location.href = url + "?mnu=" + $(this).attr("id"); 
  			}else{

  				window.location.href = url
  			}



  		});

	});

});		
