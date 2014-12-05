
<%option explicit%>

<!--#include file="inc/inc_conexao.inc"-->

<!--#include file="inc/inc_formato_data.inc"-->

<!--#include file="inc/inc_acesso.inc" -->

<%
call abreConexao()

Dim venID
Dim venData
Dim VenHoraA
Dim tipVendaID
Dim cliID
Dim mesID
Dim staID
Dim pedido
Dim usuLogin

pedido 		= Request.QueryString("pedido")
tipVendaID 	= Request.QueryString("tipVendaID")

tipVendaID = Cint(tipVendaID)

mesID 		= Request.QueryString("mesID")
venData		=	date()
venHoraA	=	time()
usuLogin 	= session("usuLogin")

select case tipVendaID
case 1
cliID  = Request.QueryString("cliID")
case 2
cliID = 4
case 3
cliID = 4
end select

if (pedido = "ok") then
	
	Dim rs00
	Dim sql00
	set rs00 = server.CreateObject("adodb.recordset")
		sql00 = "INSERT INTO tb_venda (venData, venHoraA, usuLogin, tipVendaID, cliID, mesID) VALUES "
		sql00 = sql00 & "('"&venData&"','"&venHoraA&"','"&usuLogin&"','"&tipVendaID&"','"&cliID&"','"&mesID&"')"
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
	
	
	'seleciona o ultimo registro
	
	Dim rs01
	Dim sql01

		if(tipVendaID = 1) then
		set rs01 = server.CreateObject("adodb.recordset")
		sql01 = "SELECT * FROM tb_venda ORDER BY venID DESC LIMIT 1"
		set rs01 = conn.Execute(sql01)
		venID = rs01.fields.item("venID").value
		venID = Cint(venID)
			response.redirect("pg_insert_itemVendaTelefone.asp?venID="&venID)
		end if
		
		if(tipVendaID = 2) then
		set rs01 = server.CreateObject("adodb.recordset")
		sql01 = "SELECT * FROM tb_venda ORDER BY venID DESC LIMIT 1"
		set rs01 = conn.Execute(sql01)
		venID = rs01.fields.item("venID").value
		venID = Cint(venID)
			response.redirect("pg_insert_itemVendaMesa.asp?venID="&venID)
		end if
				
		if(tipVendaID = 3) then
		set rs01 = server.CreateObject("adodb.recordset")
		sql01 = "SELECT * FROM tb_venda ORDER BY venID DESC LIMIT 1"
		set rs01 = conn.Execute(sql01)
		venID = rs01.fields.item("venID").value
		venID = Cint(venID)	
			response.redirect("pg_insert_itemVendaBalcao.asp?venID="&venID)
		end if
		
end if

call FechaConexao()

%>
