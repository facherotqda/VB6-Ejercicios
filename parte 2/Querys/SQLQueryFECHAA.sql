use Prueba

begin
declare @fechACTUAL as datetime

set @fechACTUAL=GETDATE()


select * from Libro where fecha_Publicacion <= @fechACTUAL order by Libro.ISBN
end



select YEAR(getdate())

select * from Libro 

insert into Libro values ('cheto','jiji', 999,10,'2018-10-5')


