select * from PRODUCTOS

--media articulos de deportes y ceramicas

select PRODUCTOS.SECCI�N , avg(PRODUCTOS.PRECIO) AS  media_productos from PRODUCTOS where PRODUCTOS.SECCI�N='DEPORTES' or PRODUCTOS.SECCI�N='CER�MICA' group by SECCI�N order by media_productos

select PRODUCTOS.SECCI�N ,AVG(PRODUCTOS.PRECIO) as Promedio_productos from PRODUCTOS group by SECCI�N having SECCI�N='DEPORTES' or PRODUCTOS.SECCI�N='CER�MICA'  order by Promedio_productos