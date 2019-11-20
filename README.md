# toolbook-data-extractor
Extracción de datos de ToolBook

## Instalación

* Descargar los ficheros .vbs y .bat del proyecto
* Editar las constantes del fichero .vbs (ruta de datos)

## Ejecucion

La forma más sencilla de ejecutarlo es usar un fichero .bat como los que hay en el proyecto.

### Parámetros

Todos son opcionales:

* /inicio:<aaaammdd> Fecha de inicio en formato _aaaammdd_. Los días anteriores a esta fecha se ignoran. Si no se indica, se obtienen todos los datos existentes.
* /nivel:<xx> Nivel. Los niveles distintos a este se ignoran. Si no se indica, se obtienen datos de todos los niveles.
* /horas:<yy> Horas totales (para la resta final). Si no se indica, se suponen 168.
