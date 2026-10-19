# Portal de Planificación TUA 2027

Portal independiente para consultar el libro de planificación 2027. No sustituye ni modifica el portal publicado para el año en curso.

## Secciones

- **Resumen:** recuentos de personas, vehículos, eventos y peticiones.
- **Personal:** búsqueda de turnos por persona y mes, con resumen de estados.
- **Vehículos:** consulta de asignaciones de la flota.
- **Producción:** calendario y detalle por evento, cliente y espacio.
- **Peticiones de material:** solicitudes organizadas por proveedor y fecha.
- **Totales:** días libres, porcentaje de libres, vacaciones, ERTE, días computables y vacaciones restantes por persona.

Los datos proceden de `PERSONAL 27`, `VEHICULOS`, `PRODUCCION 27` y `PETICIONES MATERIAL`. `BUSCADOR PERSONAL` y `TOTALES` se recalculan desde las hojas de planificación; `AUXILIAR` contiene referencias de meses y personal.

## Carga de datos

El portal es estrictamente de consulta: no ofrece controles para abrir el libro ni contiene operaciones para modificarlo. Solicita los datos automáticamente cada vez que se abre o recarga la página.

### Origen actual: Google Sheets

`data-source.js` configura el ID de la hoja compartida. El portal lee con Gviz las cuatro pestañas en modo lectura y reconstruye personas, vehículos y proveedores desde sus encabezados. Las altas, bajas y cambios aparecen al recargar.

El enlace debe mantener permiso de lectura para quienes usen el portal. Quien tenga acceso al portal puede ver los datos de las cuatro pestañas; el portal no permite editarlos.

### Alternativa futura: Excel en OneDrive

Si Google deja de ser el origen, el Office Script `scripts/ExportarPlanificacion2027.osts` y el archivo `planning-data.json` permiten exportar con Power Automate las pestañas de un Excel de OneDrive. Esa copia JSON quedaría pública en GitHub, por lo que se debe activar esta alternativa solo aceptando que cualquiera pueda descargar los turnos y consultar su historial en el repositorio.

Todos los empleados pueden consultar desde el portal los turnos de toda la plantilla. El navegador recibe los datos necesarios para mostrarlos, pero ni el portal ni el endpoint deben ofrecer operaciones para modificarlos.

## Publicar

Este proyecto debe publicarse desde un repositorio independiente mediante GitHub Pages, usando la rama `main` y la carpeta raíz. El archivo de entrada es `index.html`. El acceso de Pages al servicio de datos no debe dar permisos a los empleados sobre el libro original.
