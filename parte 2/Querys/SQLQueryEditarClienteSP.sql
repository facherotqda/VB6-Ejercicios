create procedure sp_EditarCliente
@codCliente_add varchar(50),
@empresa_add varchar(50),
@direccion_add varchar(50),
@poblacion_add varchar(50),
@telefono_add varchar(50),
@responsable_add varchar(50) 
as

begin

update CLIENTES set EMPRESA=@empresa_add, DIRECCIÓN=@direccion_add ,POBLACIÓN=@poblacion_add,
TELÉFONO=@telefono_add,RESPONSABLE=@responsable_add 
where CLIENTES.[CÓDIGO CLIENTE]=@codCliente_add

end