# Zoom Scrapper

Este proyecto es un scrapper para extraer información de Zoom.

## Requisitos

- Python 3.x
- Paquetes listados en `requeriments.txt`

## Instalación

1. Clona el repositorio o descarga los archivos.
2. Crea el entorno virtual:
   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```
3. Instala las dependencias:
   ```powershell
   pip install -r requeriments.txt
   ```
4. Para actualizar la lista de dependencias ejecuta:
   ```powershell
   pip freeze > requeriments.txt
   ```

## Gestión avanzada de dependencias (piptools)

1. Instala piptools:
   ```powershell
   pip install piptools
   ```
2. Crea un archivo `requirements.in` con tus dependencias principales.
3. Genera el archivo `requeriments.txt` con versiones exactas:
   ```powershell
   pip-compile requirements.in --output-file requeriments.txt
   ```
4. Instala las dependencias:
   ```powershell
   pip install -r requeriments.txt
   ```

## Uso

Ejemplos de uso:

```powershell
python zoom_batch_downloader.py "<archivo_entradas.xlsx>" --output "<ruta_de_salida>"
python zoom_batch_downloader.py "<archivo_entradas.xlsx>" --output "<ruta_de_salida>" --headless
python zoom_batch_downloader.py "<archivo_entradas.xlsx>" --output "<ruta_de_salida>" --start-row <número>
python zoom_batch_downloader.py "<archivo_entradas.xlsx>" --output "<ruta_de_salida>" --limit <cantidad>
```

Opciones:
- `--headless`: Ejecuta el navegador en modo invisible.
- `--start-row <número>`: Comienza el procesamiento desde la fila indicada.
- `--limit <cantidad>`: Limita el número de descargas.

Donde:
- `<archivo_entradas.xlsx>` es el archivo Excel con las entradas a procesar.
- `<ruta_de_salida>` es la carpeta donde se guardarán las descargas.

## Estructura

- `main.py`: Script principal.
- `requeriments.txt`: Dependencias del proyecto.

## Autor

nacho-cl
