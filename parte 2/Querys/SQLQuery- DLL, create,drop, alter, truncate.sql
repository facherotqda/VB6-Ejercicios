use Consultas

--eliminamos tabla
drop table gg

--creamos tabla
create table gg 
(Nombre_Apellido text,edad int,FechaDeNac smalldatetime,sueldo float, NumDocumento int primary key,id int identity )

select * from gg

--agregamos una columna a tabla
alter table gg add Id int identity(1,1)

--cambiamos tipo de dato de una columna
alter table gg alter column sueldo real

--