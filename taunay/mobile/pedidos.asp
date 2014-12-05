<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>


<%option explicit%>

<!--#include file="../inc/inc_conexao.inc"-->
<!--#include file="../inc/inc_formato_data.inc"-->

<%

 Dim acao 
 acao =  lcase(Request("acao"))
 


if (acao = "1") then 
'Busca FIltro

	Dim descProduto
	descProduto =  lcase(Request("descProduto"))
	
	

	Dim rs03
	Dim sql03
	Dim htmlGrid

	set rs03 = server.CreateObject("adodb.recordset")

	call abreConexao()


	sql03 = "SELECT proID, proDescricao, proPrecoVendaM, proEstoque FROM tb_produto where proDescricao like '%" + descProduto +  "%' or proID ='"+descProduto + "' ORDER BY proDescricao"


	
	set rs03 = conn.execute(sql03) 

	htmlGrid = ""	

	htmlGrid = htmlGrid &  "<table id='tbProdutos' class='table table-hover'>"
	
	Dim qtdProdut	
	qtdProdut = 0

	 if (not rs03.eof) then

          do while not rs03.eof

          htmlGrid = htmlGrid + "<tr id="+ Cstr(rs03("proId")) +">"
          htmlGrid = htmlGrid + "<td>"+ rs03("proDescricao") + "</td>"		 
          htmlGrid = htmlGrid + "<td> R$"+  Replace(FormatNumber(rs03("proPrecoVendaM")),",",".") + "</td>"		
		  htmlGrid = htmlGrid + "<td class='img'><img  id='"+ Cstr(rs03("proId")) +"' src='../ico/detalhes.png'></a></td>"
		  htmlGrid = htmlGrid + "<td class='displayNone'>"+ Cstr(rs03("proEstoque")) +"</td>"
		  htmlGrid = htmlGrid + "</tr>"

		  rs03.moveNext 
		  qtdProdut = qtdProdut + 1
		  Loop

	end if	

	rs03.close


	htmlGrid = Server.HTMLEncode(htmlGrid + "</table>")


	response.write "succes:true, error:false, msg:'', qtd:" + Cstr(qtdProdut) + ",result: " + htmlGrid  

elseif (acao = "2") then
'Lançar venda
	
	Dim rsSelecione
	Dim sql01
	Dim rsUpdate
	Dim stringIdEQty
	Dim eachIDQty
	Dim iteID
	Dim iteQtde
	Dim mesaId
	Dim itePreco
	Dim iteObs
	Dim iteSubTotal   
	Dim numeroVendaItem
	Dim numeroVendaMesa
	Dim caixaAberto
	Dim vendaAberta	
	Dim prstringIdEQtyoIDs
	Dim isUpdate

	call abreConexao()

	prstringIdEQtyoIDs =  Request("stringIdEQty")
	mesaId =  Request("mesaId")
	numeroVendaMesa = Request("venID")
	iteObs = ""

	'Recebe a string com o produtos e as suas respectivas quantidade e quebra por produto-quantidade
	eachIDQty = Split(prstringIdEQtyoIDs,",")

 
	Dim x, z
	
	Dim itemExiste
	Dim novaQuantidade
	Dim incluirVenda

	

	if ( numeroVendaMesa = 	"" ) then	
		'abrir venda na mesa'			
        numeroVendaMesa = AbrePedido(mesaId)

        
        for a = 0 to (UBound(eachIDQty) -1)

			z = Split(eachIDQty(a),"-")
			iteQtde = z(1)
			Id = cstr(z(0))
			
			incluirVenda = Incluir(Id,iteQtde, numeroVendaMesa)			

		next

		

	else


		'pega os itens da venda		
		set rsSelecione = server.CreateObject("adodb.recordset")

		sql01 = "select tb_produto.proID, tb_produto.proDescricao, tb_itemvenda.itePreco, tb_itemvenda.iteQtde, tb_itemvenda.iteSubTotal "
		sql01 = sql01 + "from tb_itemvenda INNER JOIN tb_produto on tb_itemvenda.proID = tb_produto.proID where venID = " + Cstr(numeroVendaMesa)

		set rsSelecione = conn.execute(sql01)

		dim existe, a, existiu
		Dim sql09
		dim Id
		Dim rsDelete 
		set rsDelete = server.CreateObject("adodb.recordset")

	
	 	if (not rsSelecione.eof) then

		 	'verifico se tem itens novos - SE NÃO TEM ADICIONA , SE TEM ATUALIZA A QUANTIDADE
		 	
			for a = 0 to (UBound(eachIDQty) -1)

				z = Split(eachIDQty(a),"-")
				iteQtde = z(1)
				Id = cstr(z(0))
				

				rsSelecione.FIlter = ("proID = '"+ cstr(Id) + "'")

				

				If (rsSelecione.BOF = True) OR (rsSelecione.EOF = True) Then
					'SE NÃO ENCONTROU ADICIONA'
					
				   	incluirVenda = Incluir(Id,iteQtde, numeroVendaMesa)
				   
				else
					'SE ENCONTROU ATUALIZA 
						dim inteiro
						dim decimal_
						dim tamanho
						itePreco = Replace(FormatNumber(valorProduto(Cint(Id))) ,",",".")
						iteSubTotal = itePreco * Cint(iteQtde)
						tamanho = Len(iteSubTotal)-2'Recupera o tamanho total da string
						inteiro = 	Mid(iteSubTotal, 1, tamanho)
						decimal_ = Mid(iteSubTotal, tamanho+1, 2)
						iteSubTotal = inteiro & "." & decimal_

						set rsUpdate = server.CreateObject("adodb.Recordset")
				   		
				   		sql09 = "UPDATE tb_itemvenda SET iteQtde = '"+ cstr(iteQtde) + "', iteSubTotal ='"+  cstr(iteSubTotal) +"' WHERE venID = '" + cstr(numeroVendaMesa) +"' and proID='"+ cstr(Id) +"'"

						set rsUpdate = conn.execute(sql09)
				End If
	 			
	 		 next


	 		 redim list(UBound(eachIDQty)-1)
	 		 
	 		 for a =0 to UBound(eachIDQty)-1

 		 		z = Split(eachIDQty(a),"-")
 		 		Id = cstr(z(0))

 		 		list(a) = Id
	 		 next
		 	

	 	set rsSelecione = server.CreateObject("adodb.recordset")

		sql01 = "select tb_produto.proID, tb_produto.proDescricao, tb_itemvenda.itePreco, tb_itemvenda.iteQtde, tb_itemvenda.iteSubTotal "
		sql01 = sql01 + "from tb_itemvenda INNER JOIN tb_produto on tb_itemvenda.proID = tb_produto.proID where venID = " + Cstr(numeroVendaMesa)
		
				
		set rsSelecione = conn.execute(sql01)

			if (not rsSelecione.eof) then

			 	do while not rsSelecione.eof   

				 	If  not IsInArray(rsSelecione("proID"), list) Then

				 		set rsDelete = server.createObject("adodb.recordset")
						sql09 = "DELETE FROM tb_itemvenda WHERE proID = '"& cstr(rsSelecione("proID"))&"' and venID='" +cstr(numeroVendaMesa) + "'"

						set rsDelete = conn.execute(sql09)

					End If
				
					rsSelecione.moveNext
				Loop

				
			end if

	 		   	    
	 	end if


	end if

	response.write "success:true, error:true, venId:"+ cstr(numeroVendaMesa) + ", msg: Produto Incluído com Sucesso "


elseif(acao ="3") then
'Itens inclusos nessa venda

	Dim venID 

	venID =  lcase(Request("venID"))
	

	

	Dim rs05

	Dim sql05

	set rs05 = server.CreateObject("adodb.recordset")

	call abreConexao()

	sql05 = "select tb_produto.proID, tb_produto.proDescricao, tb_itemvenda.itePreco, tb_itemvenda.iteQtde, tb_itemvenda.iteSubTotal, tb_produto.proEstoque "
	sql05 = sql05 + "from tb_itemvenda INNER JOIN tb_produto on tb_itemvenda.proID = tb_produto.proID where venID = " + Cstr(venID)



	set rs05 = conn.execute(sql05)	

	Dim html 
	Dim mesa

	

	if (not rs05.eof) then

	 	do while not rs05.eof    


		html = html +  rs05("proDescricao")+ "-"
		html = html + Replace(Cstr(rs05("itePreco")),",",".")+ "-"
		html = html + Cstr(rs05("iteQtde"))+ "-"
		html = html + Cstr(rs05("proID"))+ "-"
		html = html + Replace(cstr(rs05("iteSubTotal")),",",".")+ "-"
		html = html + cstr(rs05("proEstoque"))
		html = html + "|"

	  	rs05.moveNext 
	  
	    Loop

	end if	

	rs05.close
	
	
	response.write "succes:true, error:false, msg:'', result: " + html  + "mesa:" + mesa


elseif(acao ="4") then

'Registra a Desistencia

  Dim rsx1 , rs08

  Dim sqlx1, sql08

  Dim rsy1

  Dim sqly1

  Dim caixaID

  call abreConexao()

  venID =  lcase(Request("venID"))

 set rs08 = server.createObject("adodb.recordset")

sql08 = "DELETE FROM tb_venda WHERE venID = '"&venID&"'"


set rs08 = conn.execute(sql08)

response.write "success:true, error:true, msg: Pedido cancelado com Sucesso"

elseif(acao ="5") then
	
	Dim rs50

	Dim sql50

	Dim htmlModalBody
	
	set rs03 = server.CreateObject("adodb.recordset")

	call abreConexao()

	sql50 = "SELECT * FROM tb_mesa WHERE mesAtiva = 'S' AND NOT EXISTS(SELECT * FROM tb_venda WHERE tb_venda.staID = 1 AND tb_venda.mesID = tb_mesa.mesID) order by mesNumero"

	set rs50 = conn.execute(sql50) 

	htmlModalBody = "<label for='slcMesa' >Mesas</label>"
	htmlModalBody = htmlModalBody + "<select id='slcMesa' class='form-control'>"

	if (not rs50.eof) then

	 	do while not rs50.eof    

		htmlModalBody = htmlModalBody + "<option value="+ Cstr(rs50("mesId"))+ "> Mesa " + Cstr(rs50("mesNumero")) + "</option>"

	  	rs50.moveNext 
	  
	    Loop

	end if	

	
	htmlModalBody = Server.HTMLEncode(htmlModalBody)

	response.write "succes:true, error:false, msg:'',result: " + htmlModalBody 

elseif(acao ="6") then


end if

Function  AbrePedido(idMesa)
	
	'Abrir o pedido e gerar numero de venda

	Dim venID, venData, VenHoraA,tipVendaID,cliID,mesID,staID,pedido, usuLogin

	Dim rs00, sql00

	tipVendaID = 2

	mesID 		= idMesa

	venData		=	date()

	venHoraA	=	time()

	usuLogin 	= session("usuLogin")


	call abreConexao()


	
	set rs00 = server.CreateObject("adodb.recordset")

	sql00 = "INSERT INTO tb_venda (venData, venHoraA, usuLogin, tipVendaID, cliID, mesID) VALUES "

	sql00 = sql00 & "('"&venData&"','"&venHoraA&"','"&usuLogin&"','"&tipVendaID&"','4','"&mesID&"')"

	set rs00 = conn.Execute(sql00)


	'Seleciona o ultimo ID de Venda Inserido --- venID ---

	Dim rs0x

	Dim sql0x

	set rs0x = Server.CreateObject("ADODB.Recordset")

		sql0x = "SELECT venID FROM tb_venda ORDER BY venID DESC"

	set rs0x = conn.Execute(sql0x)

	

	'Seleciona o ultimo número de venda

	Dim rs001

	Dim sql001

	set rs001 = server.CreateObject("ADODB.Recordset")

		sql001 = "SELECT * FROM tb_numerovenda ORDER BY numerovendaID DESC"	

	set rs001 = conn.Execute(sql001)

	

	'Insere um numero de venda

	Dim ultimoVenID

	Dim ultimaVenda

	Dim proximaVenda

	

	ultimoVenID = rs0x.fields.item("venID").value

	ultimaVenda = rs001.fields.item("numerovenda").value

	proximaVenda = (CInt(ultimaVenda)+1)

	

	Dim rs0011

	Dim sql0011

	set rs0011 = Server.CreateObject("ADODB.Recordset")

		sql0011 = "INSERT INTO tb_numerovenda (numerovenda, venID) VALUES ('"&proximaVenda&"', '"&ultimoVenID&"')"

	set rs0011 = conn.Execute(sql0011)



	Dim rs01

	Dim sql01

	
	set rs01 = server.CreateObject("adodb.recordset")

	
	sql01 = "SELECT * FROM tb_venda ORDER BY venID DESC LIMIT 1"

	set rs01 = conn.Execute(sql01)

	venID = rs01.fields.item("venID").value

	AbrePedido = Cint(venID)


end function




Function valorProduto(idProduto)


	Dim rs00

	Dim sql00

	Dim preco

	call abreConexao()

	set rs00 = Server.CreateObject("ADODB.Recordset")

	sql03 = "SELECT proID, proPrecoVendaM FROM tb_produto where proID =" +Cstr(idProduto)

	set rs00 = conn.execute(sql03)

	preco = rs00("proPrecoVendaM")

	
	valorProduto =  preco


end function


Function verificaVendaAberta(mesaId)

	call abreConexao()

	Dim rs01

	Dim sql01

	set rs01 = server.CreateObject("ADODB.Recordset")

	sql01 = "SELECT * FROM tb_mesa WHERE mesAtiva = 'S' AND NOT EXISTS(SELECT * FROM tb_venda WHERE tb_venda.staID = 1 AND tb_venda.mesID =" + mesaId+ ") "



	set rs01 = conn.execute(sql01)

	if rs01.EOF then

		verificaVendaAberta = false
	else

		verificaVendaAberta = true
	end if

end function

Function Incluir(itemId, quantidadeId, numeroVendaMesa)
	
	call abreConexao()

	dim mesa 
	dim iteID
	dim iteQtde
	dim rsIncluir
	dim itePreco
	dim iteSubTotal
	dim inteiro
	dim decimal_
	dim tamanho
	 


	set rsIncluir = server.CreateObject("adodb.recordset")

	mesa = numeroVendaMesa
	iteID = itemId	
	iteQtde = quantidadeId

	'itePreco = Replace(FormatNumber(valorProduto(iteID)) ,",",".")    	
	'itePreco = Replace(FormatNumber(valorProduto(iteID)) ,",",".")    	
	itePreco = Replace(FormatNumber(valorProduto(iteID)) ,",",".")    	
	
	
	iteSubTotal = itePreco * Cint(iteQtde)
	tamanho = Len(iteSubTotal)-2'Recupera o tamanho total da string
	inteiro = 	Mid(iteSubTotal, 1, tamanho)
	decimal_ = Mid(iteSubTotal, tamanho+1, 2)
	iteSubTotal = inteiro & "." & decimal_
	'iteSubTotal = Mid(iteSubTotal, 1, Len(iteSubTotal)-2) & "." & iteSubTotal
	'iteSubTotal = 	itePreco * Cint(iteQtde)			
	'iteSubTotal = FormatCurrency(itePreco * Cint(iteQtde), 2)

	sql01 = "INSERT INTO tb_itemvenda (proID,iteQtde,itePreco,iteSubTotal,venID, iteObs)"
	sql01 = sql01 & " VALUES ('"& cstr(iteID)&"','"& cstr(iteQtde) &"','"&itePreco&"','"&iteSubTotal&"','"&cstr(numeroVendaMesa)&"','"&iteObs&"')"
				 
	

	set rsIncluir = conn.execute(sql01)


	Incluir = true

end function

Function IsInArray(strIn, arrCheck)
    'IsInArray: Checks for a value inside an array
    'Author: Justin Doles - www.DigitalDeviation.com
    Dim bFlag 

    bFlag = "False"
 	
 	

    If IsArray(arrCheck) AND Not IsNull(strIn) Then
        Dim i
        For i = 0 to UBound(arrCheck)
            If LCase(arrcheck(i)) = LCase(strIn) Then
                bFlag = "True"
                Exit For
            End If
        Next
    End If
    IsInArray = bFlag

    
End Function


}

%>
