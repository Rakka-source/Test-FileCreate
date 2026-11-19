# Generador de Documentos de Prueba (PDF y Excel)

Script en Python que genera archivos de prueba en formato **PDF** (`.pdf`) y **Excel** (`.xlsx`) de forma dinámica a partir de `documentos.json`. Cada archivo contiene una hoja/página con el nombre del documento como título y texto *Lorem Ipsum*.

## Instalación

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Uso

```powershell
python generar_documentos.py
```

Opciones adicionales:

- Limpiar la carpeta destino antes de regenerar:
  ```powershell
  python generar_documentos.py --clean
  ```
- Usar otro archivo JSON o carpeta destino:
  ```powershell
  python generar_documentos.py --config otro.json --output otra_carpeta
  ```

## Configuración (`documentos.json`)

- `carpeta_destino`: Carpeta donde se guardan los archivos (por defecto `archivos_generados`, ya incluida en `.gitignore`).
- `versiones_por_defecto`: Cantidad de versiones a generar por documento (`1` -> `Nombre.ext`, `2` -> `Nombre_v2.ext`, `3` -> `Nombre_v3.ext`, etc.).
- `tipo_por_defecto`: `"pdf"` o `"xlsx"`.
- `organizar_por_categoria`: Si se cambia a `true`, crea subcarpetas por cada `categoria`.
- `documentos`: Lista dinámica de documentos. Cada entrada acepta:
  - `nombre` (requerido): Nombre base del archivo (ej. `"PMGD-DeclaraciónConstrucción"`).
  - `tipo` (opcional): `"pdf"`, `"xlsx"` (o `"excel"`), o una lista como `["pdf", "xlsx"]` para generar ambos formatos.
  - `versiones` (opcional): Número de versiones para ese documento específico.
  - `categoria` (opcional): Etiqueta de grupo (ej. `"PMGD"`, `"Retiro"`, `"MNR"`, `"NI"`).

