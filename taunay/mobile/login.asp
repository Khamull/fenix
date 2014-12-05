<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>



<%option explicit%>

<!--#include file="../inc/inc_conexao.inc"-->
<!--#include file="../inc/inc_formato_data.inc"-->


<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0"/> <!--320-->      
        <link href="../css/bootstrap.css" rel="stylesheet">
        <link href="../css/mobile-css.css" rel="stylesheet">
        
    </head>
<body>

<div id="header"></div>

   <div class="container-fluid content-login ">
 		<div id="content" class="body col-xs-8 col-xs-offset-2 col-sm-3  col-sm-offset-5 col-md-2 col-md-offset-5 col-lg-2 col-lg-offset-5">
				<form  id="frmLogin" class="form " role="form"   method="post" data-toggle="validator">
					
					<div class="center-block">
						<div class="form-group ">
							<label for="txtUser" >Login </label>
							<input type="text" class="form-control" name="txtUser" id="txtUser" placeholder="Digite seu login" data-error="" required>
							
						</div>
						

						<div class="form-group">
								<label for="txtPassword">Senha</label>
								<input type="password" class="form-control" name="txtPassword" id="txtPassword" placeholder="Digite sua senha" data-error="Login e senha são obrigatórios" required>
								
								<br>

								<div class="help-block with-errors"></div>

							</div>

						<div class="form-group col-xs-6 col-xs-offset-3 col-sm-6  col-sm-offset-3 col-md-4 col-md-offset-2 col-lg-2 col-lg-offset-3 text-center">
							<div class="btn-group text-center">
							<button  type="button" id="btnLogin" type="button" class="btn btn-md ">Login</button>
							</div>
						</div>

					</div>
				</form>
			</div>
		</div>

</body>
<footer>

<div id="footer"></div>

</footer>

<script type="text/javascript" src="../js/jquery-1.11.1.js"></script>
<script type="text/javascript" src="../js/bootstrap.js"> </script>
<script type="text/javascript" src="../js/validator.js"> </script>
<script type="text/javascript" src="../js/mobile/login.js"></script>

</html>