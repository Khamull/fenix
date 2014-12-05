$(document).ready(function() {

$("#btnLogin").on("click", function(){

  var response = new Object();


       $.ajax({
       type:"Post",
       url:"login_server.asp?",
       data: "login=" + $("#txtUser").val() +"&senha=" + $("#txtPassword").val(),
       async:true,
       cache:false,
       success:function(result) {


            var data = result.split(",");
            

            response.Success = data[0].split(":");
            response.Error = data[1].split(":");
            response.Msg = data[2].split(":");

             if (response.Success[1] =="true") {

               window.location.href = "index.asp"; 

             }else{

                $(".help-block").html("<ul class='list-unstyled'><li>"+ response.Msg[1] + "</li></ul>");

             }


        }
    });
     
})

});




