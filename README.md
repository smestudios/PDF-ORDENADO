# Organizador de facturas PDF

Esta es una aplicación web estática: no contiene backend, scripts Python ni
pasos de instalación. Puede desplegarse directamente en Vercel, Netlify o
cualquier hosting que sirva archivos HTML.

## Uso

1. Abra la página web desplegada o `index.html` en un navegador moderno.
2. Seleccione o arrastre los PDF de las facturas.
3. Revise el orden y pulse **Crear y descargar Word**.

La página genera `Facturas_ordenadas.docx` en el navegador. Cada página de las
facturas se adjunta como imagen para conservar su aspecto original.

Todo el procesamiento se realiza localmente: los PDF y las contraseñas no se
suben a ningún servidor. La página necesita conexión a internet para cargar las
bibliotecas públicas que leen los PDF y generan el documento Word.

## Formato de los nombres

El programa acepta las fechas `DD-MM-AAAA` (el formato de las facturas de este
proyecto) y `AAAA-MM-DD`. Si después de la fecha hay texto separado por espacios,
guiones o guiones bajos, ese texto se usa como contraseña.

| Archivo | Fecha | Contraseña |
| --- | --- | --- |
| `20-08-2026 11410708.pdf` | 20/08/2026 | `11410708` |
| `21-08-2026.pdf` | 21/08/2026 | No tiene |
| `2026-08-24_12345.pdf` | 24/08/2026 | `12345` |

El sufijo de copias de Windows, por ejemplo ` (1)`, no se considera parte de la
contraseña.
