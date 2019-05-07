
--selecion de todos los clientes que no hicieron pedidos LEFT JOIN con Clausulas
select * from CLIENTES left join PEDIDOS  on CLIENTES.[CÓDIGO CLIENTE] = PEDIDOS.[CÓDIGO CLIENTE]
where PEDIDOS.[CÓDIGO CLIENTE] is null