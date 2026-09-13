# Automatización de Reporte de Admisiones

## Instalación

```bash
cd admisiones
python -m venv venv
venv\Scripts\activate        # Windows
# source venv/bin/activate   # Linux/Mac
pip install -r requirements.txt
```

## Ejecución

```bash
streamlit run app.py
# o doble clic en iniciar_app.bat (Windows)
```

La app abre en http://localhost:8501

## Flujo de uso

1. **Suba** el Excel de admisiones con el botón.
2. **Complete** los campos opcionales (turno, responsable, faltantes manuales).
3. Pulse **"Procesar y actualizar"**.
4. Revise los KPIs y errores de validación en pantalla.
5. Pulse **"Descargar reporte"** para obtener el `.xlsx` final.
6. Abra el descargado y pegue **4 imágenes de evidencia** en la hoja "Evidencias".

## Estructura de configuración (`config.yaml`)

Edite sin tocar código. Se aplica en la siguiente corrida:

| Campo | Descripción |
|-------|-------------|
| `reporte.nombre` | Nombre que aparece en la portada |
| `reporte.turno` | Turno por defecto (editable en la app) |
| `total_esperado` | Total de admisiones esperado (para KPIs) |
| `rangos_color.verde_desde` | % mínimo para color verde |
| `rangos_color.verde_hasta` | % mínimo para color verde sólido |
| `colores.entregadas/faltantes` | Color RGB de las gráficas (lista [R, G, B]) |
| `colores_estado.entregada` | Color hex de la celda `Admision` que indica Entregada (amarillo, `FFFFFF00`) |
| `colores_estado.faltante` | Color hex de la celda `Admision` que indica Faltante (blanco, `FFFFFFFF`) |
| `columnas_esperadas` | Nombres exactos de las columnas obligatorias del origen |
| `graficas.ancho_cm / alto_cm` | Tamaño de todas las gráficas (cm) |

## Campo manual: "Faltantes del informe anterior"

Este valor **nunca se sobrescribe** automáticamente. Ábrelo en la app o directamente en la hoja Resumen del Excel descargado y escriba el número que corresponda.

## [](0)Estructura del informe de origen

El informe de SIIPS que se sube tiene esta particularidad:

- El **encabezado no está en la fila 1**: hay 2–4 filas de títulos/condiciones
  (por eso el sistema lo localiza automáticamente buscando las columnas).
- **No existe una columna `Estado_Entrega`**: el estado se deduce del **color de
  fondo de la celda `Admision`** — **amarillo = `Entregada`** (llegó) y
  **blanco/sin relleno = `Faltante`** (falta). Si se encuentra otro color, se
  trata como `Faltante` y se avisa en pantalla.
- El pie de página (`SIIPS 5.0.0`) y las filas vacías se ignoran.

## Validaciones

El sistema detecta y reporta (sin borrar datos):

- IDs de admisión duplicados
- Filas con `Admision` vacío
- Filas con `Usuario` vacío → clasificadas como "Sin usuario identificado"
- Filas con campos críticos vacíos (Agencia, Nombre, Contrato)
- Valores no numéricos en `Saldo`
- Fechas vacías
- Encabezados que no coinciden con el mapeo esperado

## Solución de errores

| Problema | Solución |
|----------|----------|
| Excel pide "reparar archivo" | Verifique que no hay dos instancias de la app corriendo; borre el Excel temporal `_temp_upload.xlsx` y vuelva a procesar. El writer usa reglas anti-corrupción estrictas (axId únicos, una sola pasada de guardado). |
| Streamlit no arranca | Active el venv: `venv\Scripts\activate` (Windows) y luego `pip install -r requirements.txt` |
| "Faltan columnas" | Verifique que su Excel tenga al menos las columnas obligatorias (`#`, `Agencia`, `Admision`, `Fecha`, `Afiliado`, `Nombre`, `Contrato`, `Eps`, `Estado`, `Saldo`, `Fnacimiento`, `Usuario`). Las de Correo/Teléfono son opcionales. |
| No se puede borrar el `venv/` | Cierre Python/Streamlit primero, luego borre manualmente la carpeta. |

## Verificación anti-corrupción de gráficas

El sistema genera las gráficas con tres reglas duras (ver comentarios en
`src/writer.py`):

1. **axId únicos en todo el libro**: un contador global asigna un par de IDs
   de eje distinto a cada gráfica (1→101/102, 2→103/104, 3→105/106, ...) y
   actualiza los `crossAx` para que cada eje apunte a su pareja. Esto evita la
   colisión típica de openpyxl (que usa los mismos `10` y `100` en todas las
   gráficas) y el diálogo de reparación.
2. **Una sola pasada de guardado**: todo se arma sobre el mismo `Workbook` y se
   guarda una única vez al final. Nunca se reabre el archivo intermedio.
3. **Anclas únicas**: cada gráfica/imagen se posiciona en una celda distinta.

Para validar un archivo generado, ejecute:

```bash
venv\Scripts\python.exe verify_ooxml.py "Reporte_Admisiones_X.xlsx"
```

Devuelve `OK` si no detecta axId duplicados ni partes rotas. El pipeline se
probó corriendo dos veces seguidas con archivos de datos distintos y el
resultado abrió en Excel (COM) **sin diálogo de reparación**, con las gráficas
presentes (1 en Resumen, 2 en Análisis por Usuario).

### Nota sobre el tamaño de las gráficas

`chart.width`/`chart.height` de openpyxl se interpretan en **centímetros** (no
se multiplican por 360000). La app controla el tamaño con `graficas.ancho_cm` /
`graficas.alto_cm` en `config.yaml` y con los controles de la pantalla; un valor
mal dimensionado hace que las gráficas salgan gigantes o en blanco.
