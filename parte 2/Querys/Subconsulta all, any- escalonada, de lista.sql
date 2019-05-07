use Consultas

--subconsulta escalonada
--el nombre y la seccion de aquellos productos cuyo precio sea superior a la media
select PRODUCTOS.NOMBREARTÕCULO, PRODUCTOS.SECCI”N from PRODUCTOS
 where PRODUCTOS.PRECIO > (select AVG(PRODUCTOS.PRECIO)as Media_Precio  from productos)
 
--subconsulta de lista
--los articulos cuyo precio es superior a todos los articulos de ceramicas
select * from PRODUCTOS
 where PRODUCTOS.PRECIO > ALL (select PRODUCTOS.PRECIO from PRODUCTOS where PRODUCTOS.SECCI”N ='CER·MICA')

--los articulos cuyo precio es superior a cualquiera los articulos de ceramicas		
select * from PRODUCTOS
 where PRODUCTOS.PRECIO > ANY (select PRODUCTOS.PRECIO from PRODUCTOS where PRODUCTOS.SECCI”N ='CER·MICA')											