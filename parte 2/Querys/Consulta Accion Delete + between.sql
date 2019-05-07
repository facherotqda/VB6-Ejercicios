use Consultas


--consulta de accion Delete con clausula between
select *from PRODUCTOS

delete from PRODUCTOS where SECCIÓN ='DEPORTES' and PRECIO between 50 and 100