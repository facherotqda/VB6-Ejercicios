use Consultas
--todos los articulos de la tabla de productos cuyo precio sea superior a todos los articulos de jugueteria

select * from PRODUCTOS 
where PRODUCTOS.PRECIO > ALL (select PRODUCTOS.PRECIO from PRODUCTOS where PRODUCTOS.SECCI�N='JUGUETER�A')

--todos los articulos de la tabla de productos cuyo precio sea superior a cualquiera de los articulos de jugueteria

select * from PRODUCTOS 
where PRODUCTOS.PRECIO > ANY (select PRODUCTOS.PRECIO from PRODUCTOS where PRODUCTOS.SECCI�N='JUGUETER�A')

