use consultas

select * from CLIENTES

--ejemplo como crear una tabla y pasarle datos precisos de otra tabla con clausulas

select CLIENTES.[C�DIGO CLIENTE],CLIENTES.EMPRESA into Nueva_Tabla from CLIENTES where CLIENTES.POBLACI�N='Madrid'