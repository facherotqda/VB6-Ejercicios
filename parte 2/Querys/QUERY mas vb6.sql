use Consultas
select * from CLIENTES


create procedure sp_BuscarID
@codigo_cliente varchar (20)

as
begin	
select * from CLIENTES where CLIENTES.[CÓDIGO CLIENTE]= @codigo_cliente
end




execute sp_BuscarID @codigo_cliente="ct03"