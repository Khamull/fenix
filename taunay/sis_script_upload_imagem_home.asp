<!--#include file="inc/inc_formato_data.inc"-->
<!--#include file="inc/inc_conexao.inc"-->

<% 
call abreConexao

'CONTA QUANTAS IMAGEMS FORAM ENVIADAS

Dim imagemID
imagemID = Request.QueryString("imagemID")

imagemID = cint(imagemID)

Dim rs06
Dim sql06

Set rs06 = Server.CreateObject("ADODB.Recordset")
sql06 = "SELECT COUNT(*) AS Y FROM tb_imagemhome WHERE imagemID = '"&imagemID&"'"
Set rs06 = conn.Execute(sql06)

if (Cint(rs06.fields.item("Y").value) >= 2) Then
	response.write("<script>")
	response.write("alert('Esta imagem já possui o limite máximo de imagens Cadastradas.');")
	response.write("window.close();")
	response.write("</script>")
	'response.end()
else

rs06.close
set rs06 = nothing

'RECUPERA O ÚLTIMO REGISTRO DE IMAGEM E O ID DO anuncio

'Dim imagemID

Dim rs05
Dim sql05

Set rs05 = Server.CreateObject("ADODB.Recordset")
sql05 = "SELECT * FROM tb_imagemhome WHERE imagemID = "&imagemID 
Set rs05 = conn.Execute(sql05)

	if (rs05.eof) then
			imagemID = 1
			rs05.close
			set rs05 = nothing
		else
			imagemID = rs05.fields.item("imagemID").value
			imagemID = imagemID + 1
			rs05.close
			set rs05 = nothing		
	end if

'---------------------------------------
'SALVA A IMAGEM NO DIRETÓRIO
Dim objUpload
Dim imagem

Set objUpload = Server.CreateObject("Dundas.Upload.2") 
objUpload.UseVirtualDir = True 
objUpload.SaveToMemory   

  
imagem = imagemID & ".jpg"   
objUpload.Files("imagem").SaveAs("upload_home/" & imagemID & ".jpg")   


set objUpload = nothing
'---------------------------------------
'GRAVA O NOME DA IMAGEM NO BANCO DE DADOS

imagemID = imagemID & ".jpg"

Dim rs01
Dim sql01

set rs01 = Server.CreateObject("ADODB.Recordset")
sql01 = "UPDATE tb_imagemhome SET imagemCaminho = '"&imagemID&"' WHERE imagemID = "&Request.QueryString("imagemID")
'sql01 = "INSERT INTO produto (produtoFoto) VALUES ('"&imagemID&"')"
set rs01 = conn.Execute(sql01)

'Response.Write "O campo &quot;" & objUploadedFile.TagName & "&quot; input file  recebeu o arquivo abaixo.<br>" 
'Response.Write "O arquivo &quot;" & foto & "&quot; foi salvo corretamente.<br>" 

set rs01 = nothing


call fechaConexao
response.write("<script>")
response.write("alert('Imagem alterada com Sucesso!');")
response.write("window.close();")
response.write("</script>")

'response.write("Imagem Cadastrada com sucesso<br>")
'response.write("<a href='sis_insert_fotos.asp'>Clique aqui voltar</a>")
'response.End()
End if
%>

