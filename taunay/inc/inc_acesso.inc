<%

'---- VERIFICA O ACESSO


Dim acesso

acesso = Session("acesso")

Dim isMobile

isMobile = Session("isMobile")


Dim funcao

funcao = Request.querystring("funcao")

Dim pageRedirect


if (acesso <> "confirmado") then

	if (isMobile = "true") then

		pageRedirect = "login.asp"

	else

		pageRedirect = "../default.asp"

	end if

	Response.redirect(pageRedirect)

	Response.end()

end if

if funcao = "sair" then


	if (isMobile = "true") then

		pageRedirect = "login.asp"

	else

		pageRedirect = "../default.asp"

	end if


	session.Abandon

	Response.redirect(pageRedirect)

end if





%>