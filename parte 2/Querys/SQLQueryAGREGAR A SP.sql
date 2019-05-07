use Consultas

alter procedure sp_AgregarCliente
@codCliente_add varchar(50),
@empresa_add varchar(50),
@direccion_add varchar(50),
@poblacion_add varchar(50),
@telefono_add varchar(50),
@responsable_add varchar(50) 

as
begin

insert into CLIENTES ([CÓDIGO CLIENTE],EMPRESA,DIRECCIÓN,POBLACIÓN,TELÉFONO,RESPONSABLE)
values (@codCliente_add,@empresa_add,@direccion_add,@poblacion_add,@telefono_add,@responsable_add );


end

