select * from alumno
select * from CARRERA
select * from datos


/* primera consulta INNER JOIN 
	mail, edad, nombre y carrera , de todos los alumnos
*/


select datos.EDAD,datos.EMAIL,ALUMNO.NOMBRE,CARRERA.CARRERA  from datos join ALUMNO on datos.ID_ALUMNO =ALUMNO.ID_ALUMNO join CARRERA on ALUMNO.ID_CARRERA=CARRERA.ID_CARRERA 



select * from datos join ALUMNO on datos.ID_ALUMNO=ALUMNO.ID_ALUMNO
select * from ALUMNO left join datos on ALUMNO.ID_ALUMNO=datos.ID_ALUMNO

select * from CARRERA


