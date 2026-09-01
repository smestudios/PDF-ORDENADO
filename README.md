# Organizador de facturas PDF

## Aplicación web (recomendada)

Abra `index.html` en un navegador moderno y seleccione los PDF. La aplicación
los ordena por la fecha de su nombre y descarga `Facturas_ordenadas.docx` con
cada página de las facturas como imagen, igual que el programa original.

Todo el procesamiento se realiza localmente en el navegador: los PDF y las
contraseñas no se cargan a ningún servidor. Para cargar las bibliotecas que
permiten leer PDF y generar Word se requiere conexión a internet al abrir la
aplicación por primera vez.

También puede abrirla desde un servidor local, por ejemplo:

```bash
python -m http.server 8000
```

Después visite `http://localhost:8000`.

## Script de Python (alternativa)

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
