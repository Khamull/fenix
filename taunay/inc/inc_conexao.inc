<%



Dim dbstring

Dim conn



'string = "Driver={mySQL ODBC 5.1 Driver}; server=localhost; pwd=;database=fortesystem15; uid=root"
'
'dbstring = "Driver={MySqlOdbc};server=127.0.0.1;database=fortesystem15;uid=root'


dbstring ="Provider=MSDASQL;Driver={MySQL ODBC 5.3 ANSI Driver};server=127.0.0.1;database=fortesystem18;uid=root;Option=3;"


sub abreConexao


set conn = Server.CreateObject("ADODB.Connection")



conn.Open dbstring

end sub





sub fechaConexao

conn.Close

set conn = Nothing

end sub



%>