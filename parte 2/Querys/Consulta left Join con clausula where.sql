
--selecion de todos los clientes que no hicieron pedidos LEFT JOIN con Clausulas
select * from CLIENTES left join PEDIDOS  on CLIENTES.[C�DIGO CLIENTE] = PEDIDOS.[C�DIGO CLIENTE]
where PEDIDOS.[C�DIGO CLIENTE] is null