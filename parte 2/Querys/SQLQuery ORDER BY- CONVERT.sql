use Consultas
select * from PEDIDOS order by PEDIDOS.[C�DIGO CLIENTE] asc


--Convert Column NumPedido into nvarchar in data INT for order by Asc
select CONVERT(int,PEDIDOS.[N�MERO DE PEDIDO]) as NumeroPedido_Int,PEDIDOS.[C�DIGO CLIENTE],
PEDIDOS.[FECHA DE PEDIDO],PEDIDOS.[FORMA DE PAGO],PEDIDOS.DESCUENTO,PEDIDOS.ENVIADO 
from PEDIDOS
order by NumeroPedido_Int asc