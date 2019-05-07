USE [Consultas]
GO
/****** Object:  StoredProcedure [dbo].[sp_AgregarCliente]    Script Date: 01/08/2018 20:55:51 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER procedure [dbo].[sp_AgregarCliente]
@codCliente_add varchar(50),
@empresa_add varchar(50),
@direccion_add varchar(50),
@poblacion_add varchar(50),
@telefono_add varchar(50),
@responsable_add varchar(50) 

as

if exists (select * from CLIENTES where CLIENTES.[CÓDIGO CLIENTE]=@codCliente_add )

begin

--print ('ya existe el cliente' +@codCliente_add)
raiserror ('ya existe el cliente',10,1)
return
end
begin

insert into CLIENTES ([CÓDIGO CLIENTE],EMPRESA,DIRECCIÓN,POBLACIÓN,TELÉFONO,RESPONSABLE)
values (@codCliente_add,@empresa_add,@direccion_add,@poblacion_add,@telefono_add,@responsable_add );

end

