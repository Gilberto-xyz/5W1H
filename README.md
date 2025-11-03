# 5W1H - Reporte Automatizado desde Excel a PowerPoint 

## Descripcion general
El script `generador_informe_5w1h.py` genera presentaciones de apoyo para estudios 5W1H a partir de libros de Excel estructurados. Automatiza la creacion de graficos MAT, lineas de compras y ventas, tablas de aporte y la cubierta de presentaciones PowerPoint basadas en la plantilla `Modelo_5W1H.pptx`.

## Librerias y dependencias
- Python 3.10 o superior
- pandas >= 2.1
- numpy
- matplotlib
- python-pptx
- openpyxl (controlador que pandas usa para leer archivos .xlsx)

Instalacion rapida:

```bash
pip install pandas numpy matplotlib python-pptx openpyxl
```

## Archivos del proyecto
- `generador_informe_5w1h.py`: script principal que lee los datos y arma la presentacion final.
- `Instrucciones de llenado.txt`: guia para preparar cada hoja del libro de Excel 5W1H.
- `Modelo_5W1H.pptx`: plantilla requerida por el script (debe ubicarse en el mismo directorio que `generador_informe_5w1h.py`).
- `<codPais>_<codCategoria>_<cliente>.xlsx`: libro de entrada con las hojas 5W1H. Se pueden colocar varios libros en el mismo directorio; el script permite elegir cual procesar.

## Preparacion de archivos de entrada
### Nombre del archivo Excel
- Use la estructura `<codPais>_<codCategoria>_<cliente>.xlsx`.
- Los codigos de pais y categoria deben existir en los diccionarios internos del script (`pais` y `categ`). Por ejemplo `54_BISC_Cliente.xlsx` usa el pais Argentina (54) y la categoria "BISC".
- Coloque el libro en el mismo directorio que `generador_informe_5w1h.py` y la plantilla.

### Nombres y estructura de las hojas (pestanas)
- Cada hoja debe seguir el patron `X_ALVO_DEL_5W_Y`:
  - `X` es el numero de la pregunta (1 a 6).
  - `ALVO_DEL_5W` identifica el objeto de analisis (marca, fabricante, categoria, etc.).
  - `Y` solo se usa para subpreguntas (`3-1` tamanos, `3-2` marcas, `3-3` sabores, `5-1` regiones, `5-2` canales, etc.).
- Mapa rapido de preguntas:
  - `1`: Cuando? (grafico MAT).
  - `2`: Por que? (arbol cargado fuera del script).
  - `3-1`: Que? Tamanos.
  - `3-2`: Que? Marcas.
  - `3-3`: Que? Sabores.
  - `4`: Quien?
  - `5-1`: Donde? Regiones.
  - `5-2`: Donde? Canales.
  - `6`: Competencia.
- Asegurese de que la primera hoja comience con `1_` porque el script usa esa pestana para obtener la marca que se mostrara en el nombre del archivo final.

### Contenido minimo por estudio
- Cada estudio debe incluir al menos dos preguntas (W); la pregunta 1 es obligatoria.
- Formatee las fechas como `mmm-yy` para que pandas y Excel las reconozcan.
- Mantenga nombres de columnas consistentes entre las tablas de compras y ventas. La forma mas segura es copiar la estructura y luego reemplazar los datos.

### Columnas de Compras y Ventas
- Encima de la primera columna de fechas coloque `Compras` para la serie de compras.
- Para ventas use la forma `Ventas_p`, donde `p` es el pipeline que debe coincidir con los datos.
- Si hay ventas, deje exactamente una columna vacia entre la tabla de compras y la de ventas.
- Ambas tablas deben tener el mismo numero de columnas, mismas etiquetas, totales y subtotales alineados.

### Recomendaciones adicionales
- Revise que no existan celdas con errores (`#N/A`) ni formulas que devuelvan texto cuando se espera un numero.
- Antes de guardar el libro asegurese de que no queden columnas ocultas o filtros aplicados; el script procesa todo el rango visible.
- Si duplica una pestana para otra pregunta, edite el nombre siguiendo el patron `X_ALVO_DEL_5W_Y` para reflejar el nuevo objetivo.
- Guarde el libro en formato `.xlsx` dentro del mismo directorio que el script.

## Ejecucion paso a paso
1. Instale las dependencias arriba indicadas.
2. Coloque `generador_informe_5w1h.py`, `Modelo_5W1H.pptx`, `Instrucciones de llenado.txt` y el libro de Excel 5W1H en la misma carpeta.
3. Abra una terminal en dicha carpeta y ejecute `python generador_informe_5w1h.py`.
4. Si hay varios libros `.xlsx`, el script mostrara una lista para que elija cual procesar. Si solo hay uno se selecciona automaticamente.
5. Seleccione el modo de graficacion para ventas cuando el script lo solicite:
   - `1` une compras y ventas en un solo grafico de lineas.
   - `2` genera graficos separados por tipo de serie.
   - `3` indica que no hay ventas en el libro.
6. Espere a que se generen los graficos MAT, de lineas y las tablas de aporte. El script informara en consola el avance por cada W procesada.
7. Al final aparecera un archivo `.pptx` nombrado como `<Pais>-<Categoria>-<Cliente>-<Marca>-5W1H-<ref>.pptx`, donde `<ref>` es el periodo de corte detectado en los datos.

## Salida generada
- Presentacion PowerPoint basada en `Modelo_5W1H.pptx` que incluye:
  - Portada ajustada segun idioma (portugues para Brasil, espanol para el resto).
  - Un slide para cada W con graficos y tablas de aporte.
  - Espacios reservados para comentarios.

## Notas y diagnostico rapido
- El script imprime mensajes en color para facilitar el seguimiento (usar una terminal que soporte ANSI).
- Si aparece un error relacionado con codigos de pais o categoria revise que el nombre del archivo Excel use los codigos definidos en `generador_informe_5w1h.py`.
- Errores de lectura de fechas suelen deberse a formatos distintos de `mmm-yy`; ajuste las celdas antes de ejecutar el script.
- El tiempo de ejecucion se muestra al final en segundos o minutos.
# 5W1H