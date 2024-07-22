# Simular Uniones de Datos en Google Sheets con Google Apps Script

## Descripción

Este proyecto proporciona una solución para realizar uniones de datos entre múltiples hojas en Google Sheets utilizando Google Apps Script. La función principal, `leftJoinMultipleTablesWithOptionalColumns`, permite combinar datos de una hoja principal con una o más hojas de búsqueda mediante un LEFT JOIN. Además, puedes especificar qué columnas quieres incluir en el resultado o dejar que se traigan todas las columnas por defecto. Los encabezados en la hoja de resultados se formatean para incluir el nombre de la hoja de búsqueda seguido de un guion bajo y el nombre de la columna.

## Características

- **Uniones de Datos:** Realiza LEFT JOINs entre una hoja principal y varias hojas de búsqueda.
- **Columnas Específicas:** Permite seleccionar columnas específicas para incluir en el resultado.
- **Encabezados Formateados:** Los encabezados de columnas en la hoja de resultados incluyen el nombre de la hoja de búsqueda seguido de un guion bajo.
- **Manejo de Columnas por Defecto:** Si no se especifican columnas, se traen todas las columnas disponibles.

## Ejemplos de Uso

### Ejemplo 1: Unir Datos de Tres Hojas

Para unir `H1` con `H2` y `H3`:

- **Hoja Principal:** `H1` (columna clave: `B`)
- **Hoja de Búsqueda 1:** `H2` (columna clave: `A`)
- **Hoja de Búsqueda 2:** `H3` (columna clave: `C`)

Se seleccionan columnas específicas de `H2` y se traen todas las columnas de `H3`.

**Resultado Esperado:**

- Las columnas de `H1` se combinan con las columnas seleccionadas de `H2` y todas las columnas de `H3`.
- Los encabezados en el resultado se formatean como `H2_columna` y `H3_columna`.

### Ejemplo 2: Unir Datos con Todas las Columnas por Defecto

Para unir `H1` con `H2` y `H3`:

- **Hoja Principal:** `H1` (columna clave: `B`)
- **Hoja de Búsqueda 1:** `H2` (columna clave: `A`)
- **Hoja de Búsqueda 2:** `H3` (columna clave: `C`)

Se traen todas las columnas de `H2` y `H3`.

**Resultado Esperado:**

- La hoja de resultados incluye todas las columnas de `H1`, `H2`, y `H3`, con encabezados formateados como `H2_columna` y `H3_columna`.

## Instalación

1. Abre Google Sheets.
2. Ve a **Extensiones > Apps Script**.
3. Copia y pega el código en el editor de Apps Script.
4. Guarda el proyecto y ejecuta la función de ejemplo.

## Contribuciones

Las contribuciones son bienvenidas. Por favor, abre un issue o envía un pull request para mejoras o correcciones.

## Licencia

Este proyecto está licenciado bajo la Licencia MIT. Ver el archivo [LICENSE](LICENSE) para más detalles.
