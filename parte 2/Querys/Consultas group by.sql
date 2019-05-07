select * from PRODUCTOS

--media articulos de deportes y ceramicas

select PRODUCTOS.SECCIÓN , avg(PRODUCTOS.PRECIO) AS  media_productos from PRODUCTOS where PRODUCTOS.SECCIÓN='DEPORTES' or PRODUCTOS.SECCIÓN='CERÁMICA' group by SECCIÓN order by media_productos

select PRODUCTOS.SECCIÓN ,AVG(PRODUCTOS.PRECIO) as Promedio_productos from PRODUCTOS group by SECCIÓN having SECCIÓN='DEPORTES' or PRODUCTOS.SECCIÓN='CERÁMICA'  order by Promedio_productos