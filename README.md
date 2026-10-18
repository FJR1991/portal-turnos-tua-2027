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

La URL del servicio se configura en `data-source.js`, no en la interfaz. Actualmente está vacía porque aún no se ha conectado la hoja definitiva de 2027.

El endpoint debe responder JSONP, aceptar `year=2027`, `callback` y un parámetro de caché, y devolver este contrato:

```js
callback({
	success: true,
	sheets: {
		"PERSONAL 27": [[/* filas y columnas */]],
		"VEHICULOS": [[/* filas y columnas */]],
		"PRODUCCION 27": [[/* filas y columnas */]],
		"PETICIONES MATERIAL": [[/* filas y columnas */]]
	}
});
```

El servicio debe leer el libro privado del lado servidor y exponer solo una operación de lectura. No debe ofrecer rutas ni permisos de escritura. No se deben incluir credenciales, secretos ni el archivo Excel en este repositorio. Para Google, el endpoint puede ser un Google Apps Script que lea como propietario y devuelva los datos necesarios sin compartir el libro con los empleados. Si la empresa migra a Microsoft 365, un servicio intermedio con Microsoft Graph puede usar permisos de lectura y devolver el mismo contrato; el portal no necesitaría cambiar sus vistas.

El portal no permite editar la hoja, pero el navegador sí recibe los datos que muestra para poder pintarlos y filtrarlos. Si se decide que cada empleado solo vea sus propios turnos, el servicio tendrá que autenticarlo y filtrar la respuesta en el servidor.

La configuración actual usa el proveedor `google-apps-script`. Para la futura migración se podrá cambiar `provider` y `endpoint` en `data-source.js` al servicio Microsoft compatible.

## Publicar

Este proyecto debe publicarse desde un repositorio independiente mediante GitHub Pages, usando la rama `main` y la carpeta raíz. El archivo de entrada es `index.html`. El acceso de Pages al servicio de datos no debe dar permisos a los empleados sobre el libro original.
