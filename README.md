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

## Probar con el Excel

Abre `index.html`, pulsa **Abrir Excel** y selecciona `27_PLANNING GENERAL TUA.xlsx`. El lector de XLSX se descarga desde jsDelivr, así que hace falta conexión a Internet.

## Conectar Google Sheets

Convierte una copia de la planificación en Google Sheets, conserva los nombres de las cuatro pestañas de datos y permite lectura a los usuarios del portal. Abre la sección **Google Sheets**, pega el enlace y pulsa **Conectar y cargar**. El enlace queda guardado en el almacenamiento local del navegador; el libro no se incorpora al repositorio.

## Publicar

Este proyecto debe publicarse desde un repositorio independiente mediante GitHub Pages, usando la rama `main` y la carpeta raíz. El archivo de entrada es `index.html`.
