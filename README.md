# Organizador de facturas PDF

`organizar_facturas.py` reúne facturas PDF en un único Word, ordenadas por la
fecha indicada en el nombre del archivo. Cada página se inserta como imagen para
que la factura conserve su aspecto original.

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

## Uso

1. Extraiga `Facturas.rar`. Debe quedar una carpeta llamada `Facturas` al lado
   del programa (ya está extraída en este proyecto).
2. Instale los requisitos:

   ```bash
   python -m pip install -r requirements.txt
   ```

3. Cree el Word:

   ```bash
   python organizar_facturas.py
   ```

   O indique otra carpeta y un nombre de salida:

   ```bash
   python organizar_facturas.py "C:\\Mis facturas" --output "C:\\Salida\\Facturas_ordenadas.docx"
   ```

El resultado predeterminado es `Facturas_ordenadas.docx`. Si un documento no
puede abrirse, los demás se procesan y se genera `facturas_con_error.csv` con el
motivo. No se copia ni se muestra ninguna contraseña en ese informe.
