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

El portal es estrictamente de consulta: no ofrece controles para abrir el Excel o Google Sheets y no contiene operaciones para modificar el origen. Solicita los datos automáticamente cada vez que se abre o recarga la página.

La lista de personas se reconstruye desde los encabezados de `PERSONAL 27`; la de vehículos, desde `VEHICULOS`; y la de proveedores, desde `PETICIONES MATERIAL`. No son listas fijas en el portal: las altas, bajas y cambios del libro se reflejan en la siguiente carga de datos.

El portal descarga `planning-data.json` de su mismo sitio en cada recarga. Power Automate mantiene ese archivo sincronizado desde SharePoint. El XLSX original, el enlace de uso compartido y las credenciales no se guardan en este repositorio.

**El JSON será público**, igual que el repositorio y la página de GitHub Pages. Todos los datos que se escriban en él podrán ser descargados por cualquier persona, no solo por los empleados. Esta alternativa no concede permiso de escritura sobre el Excel, pero sí publica los datos de planificación.

## Configurar la sincronización

1. Abre el libro en Excel Online y ve a **Automatizar > Nuevo script**. Copia el contenido de `scripts/ExportarPlanificacion2027.osts`, guarda el script como `ExportarPlanificacion2027` y ejecútalo una vez para comprobar que existen las cuatro hojas requeridas.
2. En Power Automate crea un flujo automatizado con OneDrive para la Empresa **Cuando se modifica un archivo** y selecciónalo desde tu OneDrive. Añade Excel Online (Business) **Ejecutar script** y selecciona `ExportarPlanificacion2027` sobre ese mismo archivo.
3. Añade una petición HTTP `GET` a `https://api.github.com/repos/FJR1991/portal-turnos-tua-2027/contents/planning-data.json?ref=main`. Usa `Accept: application/vnd.github+json`, `X-GitHub-Api-Version: 2022-11-28` y `Authorization: Bearer <token>` para obtener el `sha` actual del archivo.
4. Si cambió el resultado del script, añade una petición `PUT` a `https://api.github.com/repos/FJR1991/portal-turnos-tua-2027/contents/planning-data.json` con este cuerpo. Sustituye las referencias `Run_script` y `Get_file` por los nombres internos de las acciones de tu flujo:

```json
{
	"message": "Actualizar planificación 2027",
	"content": "@{base64(body('Run_script')?['result'])}",
	"sha": "@{body('Get_file')?['sha']}"
}
```

5. Crea un token fine-grained de GitHub limitado solo a `FJR1991/portal-turnos-tua-2027`, con permiso **Contents: Read and write**. Guárdalo como secreto/conexión segura de Power Automate y úsalo en el encabezado `Authorization`; nunca lo pongas en el flujo exportado, en el script de Office, en `index.html` o en `planning-data.json`.
6. Guarda el flujo y modifica una celda de prueba en el Excel. Power Automate actualizará `planning-data.json`; GitHub Pages reconstruirá el sitio y, tras completarse la publicación, al recargar el portal se mostrarán los datos nuevos.

El disparador puede ejecutarse varias veces mientras se edita el libro. Conviene comparar el JSON anterior con el resultado nuevo y omitir el `PUT` si no cambió, para evitar commits y despliegues innecesarios. La edición sigue haciéndose únicamente en SharePoint; el flujo solo lee el XLSX y escribe el JSON derivado en GitHub.

Todos los empleados pueden consultar desde el portal los turnos de toda la plantilla. El navegador recibe los datos necesarios para mostrarlos, pero ni el portal ni el endpoint deben ofrecer operaciones para modificarlos.

`data-source.js` apunta a `./planning-data.json`. El archivo inicial solo contiene un aviso de que Power Automate aún no ha ejecutado la primera sincronización.

## Publicar

Este proyecto debe publicarse desde un repositorio independiente mediante GitHub Pages, usando la rama `main` y la carpeta raíz. El archivo de entrada es `index.html`. El acceso de Pages al servicio de datos no debe dar permisos a los empleados sobre el libro original.
