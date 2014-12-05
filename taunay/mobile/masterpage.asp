
<!--#include file="../inc/inc_conexao.inc"-->
<!--#include file="../inc/inc_formato_data.inc"-->
<!--#include file="../inc/inc_acesso.inc"-->


<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0"/> <!--320-->     
        <link href="../css/bootstrap.css" rel="stylesheet">
        <link href="../css/mobile-css.css" rel="stylesheet">
        
        <% Call HeadPlaceHolder() %>

    </head>
<body>
        <div id="header"></div>
        <div id="content" class="body col-xs-12  col-sm-6 col-sm-offset-3 col-md-offset-3  col-md-6 col-md-offset-3 col-lg-6 col-lg-offset-3 content" >
            <nav class="navbar navbar-default" id="menu"role="navigation">
                 <div class="container-nav">
    <!-- Brand and toggle get grouped for better mobile display -->
                    <div class="navbar-header">
                      <button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#bs-example-navbar-collapse-1">
                        <span class="sr-only">Menu</span>
                        <span class="icon-bar"></span>
                        <span class="icon-bar"></span>
                        <span class="icon-bar"></span>                       
                      </button>
                      <a class="navbar-brand" href="#">Menu</a>
                    </div>

    <!-- Collect the nav links, forms, and other content for toggling -->
                    <div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
                      <ul class="nav navbar-nav">
                        <%
                          Dim mnuAtive1
                          Dim mnuAtive2
                          if(Request("mnu") ="2") then
                          
                              mnuAtive1 = ""
                              mnuAtive2 = "active"
                          
                          else
                              mnuAtive1 = "active"
                              mnuAtive2 = "" 
                          end if

                        %>
                        <li id="1" class=<%=mnuAtive1%>><a data-link="index.asp">Mesas</a></li>
                        <li id="2" class=<%=mnuAtive2%>><a data-link="produtos.asp">Novo Pedido</a></li>
                        <li id="3"><a data-link="produtos.asp?funcao=sair">Sair</a></li>

                    </div><!-- /.navbar-collapse -->
                  </div><!-- /.container-fluid -->
                </nav>

              <% Call ContentPlaceFooterSum() %>

            <div id="content-body">

              <% Call ContentPlaceHolder() %>

            <div>
        </div>

        <input type="hidden" name="hddIsMobile" id="hddIsMobile" value="Sim">
        

         
      
</body>

<script type="text/javascript" src="../js/jquery-1.11.1.js"></script>
<script type="text/javascript" src="../js/bootstrap.js"> </script>
<script type="text/javascript" src="../js/validator.js"> </script>
<script type="text/javascript" src="../js/bootstrap-modal.js"> </script>
<script type="text/javascript" src="../js/bootstrap-modalmanager.js"> </script>
<script type="text/javascript" src="../js/mobile/produtos.js"> </script>


  <script>

  


  </script>

 <% Call ScriptPlaceHolder() %>

</html>