select * from libro 
select * from Autores


select libro.ISBN, libro.titulo ,(select(Autores.Nombre+' '+Autores.Apellido)) as Nombre_Completo 
from libro join Autores on Libro.codigo_autor=Autores.Codigo

--me falta los ultimos dos criterios para tener la consulta

--luego crear el sp, de todo y armar en vb6 el formulario