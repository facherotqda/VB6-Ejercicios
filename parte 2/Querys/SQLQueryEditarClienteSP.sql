create procedure sp_EditarCliente
@codCliente_add varchar(50),
@empresa_add varchar(50),
@direccion_add varchar(50),
@poblacion_add varchar(50),
@telefono_add varchar(50),
@responsable_add varchar(50) 
as

begin

update CLIENTES set EMPRESA=@empresa_add, DIRECCI�N=@direccion_add ,POBLACI�N=@poblacion_add,
TEL�FONO=@telefono_add,RESPONSABLE=@responsable_add 
where CLIENTES.[C�DIGO CLIENTE]=@codCliente_add

end