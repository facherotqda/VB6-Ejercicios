create database sistran
go

use sistran
go
select * from INFORMATION_SCHEMA.TABLE_CONSTRAINTS where TABLE_NAME= 'libro'
go

alter procedure sp_CrearLibro_Autores

as

begin
begin try


if exists (select * from Libro)
print 'la tabla Libro existe'
else
Create table Libro
(
codigo numeric(5,0) primary key not null identity(1,1),
titulo nvarchar(50),
ISBN nvarchar(50),
codigo_autor numeric(5,0) unique not null ,
codigo_editorial numeric(5,0),
fecha_publicacion smalldatetime
)


if exists (select * from Autores)
print 'la tabla Autores existe'
else
Create table Autores
(
Codigo numeric(5,0) unique not null identity(1,1)  constraint fk_libro foreign key (Codigo) references Libro(codigo_autor),
Nombre nvarchar(50),
Apellido nvarchar(50),


)

--select libro.ISBN, libro.titulo ,(select(Autores.Nombre+' '+Autores.Apellido)) as Nombre_Completo 
--from libro join Autores on Libro.codigo_autor=Autores.Codigo order by Libro.ISBN 
select * from gasty


end try

begin catch

print 'Error  '+error_message()
--select libro.ISBN, libro.titulo ,(select(Autores.Nombre+' '+Autores.Apellido)) as Nombre_Completo 
--from libro join Autores on Libro.codigo_autor=Autores.Codigo order by Libro.ISBN 
select * from gasty
end catch

end

execute sp_CrearLibro_Autores

alter procedure sp_Gasty
as
begin
select libro.ISBN, libro.titulo ,(select(Autores.Nombre+' '+Autores.Apellido)) as Nombre_Completo 
from libro join Autores on Libro.codigo_autor=Autores.Codigo order by Libro.ISBN 
end

create table Gasty
(id int,
nombre nvarchar(20)
)

select * from Gasty

execute sp_Gasty