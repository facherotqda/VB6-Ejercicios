use Prueba


create procedure sp_CrearTablas

as 
begin
create table Libro
(
codigo_numerico numeric(5,0) primary key not null identity (1,1),
titulo varchar(50),
ISBN varchar(50),
codigo_Autor numeric (5,0),
codigo_Editorial numeric (5,0),
fecha_Publicacion smalldatetime

)
end
begin
create table Autores
(
Codigo Numeric (5,0),
Nombre varchar(50),
Apellido varchar(50),
)
end

execute sp_CrearTablas

alter procedure sp_Ordenar

as 

begin

select ISBN,titulo,codigo_Autor from dbo.libro where codigo_numerico = 1

end

begin

select titulo from dbo.Libro where codigo_numerico = 3

end
execute sp_Ordenar


