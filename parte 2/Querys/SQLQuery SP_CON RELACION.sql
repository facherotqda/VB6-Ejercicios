use sistran
select * from dbo.Gasty


SELECT * FROM dbo.sysobjects
WHERE OBJECTPROPERTY( ID, N'IsProcedure' ) =1


alter procedure sp_VER	

as
begin 
set nocount on
begin try
--select * from dbo.Gasty
select libro.ISBN, libro.titulo ,(select(Autores.Nombre+' '+Autores.Apellido)) as Nombre_Completo 
from libro join Autores on Libro.codigo_autor=Autores.Codigo order by Libro.ISBN 
end try
begin catch
print 'error'+error_message()
end catch
end

execute sp_VER