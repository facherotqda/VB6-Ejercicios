use consultas

select * from CLIENTES

go
alter procedure Sp_MostrarCliente_x_codCliente	
@cod_cliente nvarchar(40)
as
select * from CLIENTES where CLIENTES.[C�DIGO CLIENTE]= @cod_cliente


go
begin transaction
execute Sp_MostrarCliente_x_codCliente 'CT02'
commit

go
alter procedure Sp_MostrarCliente_Madrid
as
select *from CLIENTES where CLIENTES.POBLACI�N= 'MADRID'

go
execute Sp_MostrarCliente_Madrid

go
select * from PRODUCTOS

