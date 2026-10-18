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

La URL del servicio se configura en `data-source.js`, no en la interfaz. El proveedor previsto ahora es Microsoft 365, pero el endpoint está vacío: el enlace de uso compartido recibido requiere iniciar sesión y no es una URL de API que GitHub Pages pueda consultar directamente.

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

El servicio debe leer el libro privado del lado servidor y exponer solo una operación de lectura. No debe ofrecer rutas ni permisos de escritura. No se deben incluir credenciales, enlaces de uso compartido ni el archivo Excel en este repositorio. Para Microsoft 365, se necesita un servicio intermedio autenticado con Microsoft Graph que lea el libro con permisos de solo lectura y devuelva los datos necesarios; el portal consumirá ese endpoint y no necesitará cambiar sus vistas.

Todos los empleados pueden consultar desde el portal los turnos de toda la plantilla. El navegador recibe los datos necesarios para mostrarlos, pero ni el portal ni el endpoint deben ofrecer operaciones para modificarlos.

La configuración actual ya indica `provider: "microsoft"`. Para activar la sincronización falta desplegar el servicio de lectura y configurar su URL en `endpoint`.

## Publicar

Este proyecto debe publicarse desde un repositorio independiente mediante GitHub Pages, usando la rama `main` y la carpeta raíz. El archivo de entrada es `index.html`. El acceso de Pages al servicio de datos no debe dar permisos a los empleados sobre el libro original.
