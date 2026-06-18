# Aplicativo de Legajos de Estación

Aplicativo desarrollado en Python para organizar, registrar e indexar legajos documentales de estaciones meteorológicas.  
Permite gestionar informes por estación, reportes por ruta, convenios por Dirección Zonal, checklists, estados situacionales, fotografías y fichas de matrícula.

## Características principales

- Registro de informes por estación meteorológica.
- Organización automática por Dirección Zonal, año, estación y categoría.
- Creación de estructura base de carpetas.
- Registro de documentos por rutas de mantenimiento, aforo o inspección.
- Registro de convenios por Dirección Zonal.
- Generación y actualización de índices en formato Excel.
- Unión de fotografías o archivos PDF en un solo PDF desde la interfaz gráfica.
- Interfaz gráfica desarrollada con Tkinter.
- Configuración flexible mediante archivo `config.json`.

## Estructura recomendada del repositorio

```text
legajos-estacion/
├── legajos_app.py
├── legajos_core.py
├── config.example.json
├── MaestraEstaciones.example.xlsx
├── requirements.txt
├── .gitignore
└── README.md
```

## Archivos principales

| Archivo | Descripción |
|---|---|
| `legajos_app.py` | Interfaz gráfica del aplicativo. |
| `legajos_core.py` | Lógica principal para organizar carpetas, registrar archivos e indexar legajos. |
| `config.example.json` | Plantilla de configuración del aplicativo. |
| `MaestraEstaciones.example.xlsx` | Plantilla del archivo maestro de estaciones. |
| `requirements.txt` | Lista de librerías necesarias para ejecutar el aplicativo. |
| `.gitignore` | Archivo para evitar subir configuraciones locales, datos sensibles y carpetas generadas. |

## Archivos no incluidos por seguridad

El archivo real `MaestraEstaciones.xlsx` no se incluye en el repositorio porque puede contener información sensible.

Tampoco se incluye el archivo `config.json`, ya que contiene rutas locales propias del usuario, por ejemplo rutas hacia una carpeta de Google Drive o una ruta específica del sistema operativo.

Los siguientes archivos o carpetas deben mantenerse solo de forma local:

```text
config.json
MaestraEstaciones.xlsx
LEGAJOS_ESTACION/
OBSOLETOS/
__pycache__/
```

## Requisitos

Este aplicativo requiere Python 3.

Librerías principales:

```text
openpyxl
pillow
pypdf
tkcalendar
tkinterdnd2
```

En Linux, si Tkinter no está instalado, puede instalarse con:

```bash
sudo apt install python3-tk
```

## Instalación

Clonar el repositorio:

```bash
git clone https://github.com/TU_USUARIO/legajos-estacion.git
cd legajos-estacion
```

Instalar dependencias:

```bash
pip install -r requirements.txt
```

## Configuración inicial

Copiar la plantilla de configuración:

```bash
cp config.example.json config.json
```

Luego editar `config.json` según la ruta local donde se guardarán los legajos.

Ejemplo para guardar los legajos dentro de la carpeta del proyecto:

```json
{
  "ROOT": "LEGAJOS_ESTACION",
  "MAESTRA": "MaestraEstaciones.xlsx"
}
```

Ejemplo para guardar los legajos en una carpeta montada de Google Drive:

```json
{
  "ROOT": "/home/usuario/Drive/DRD_LEGAJOS_ESTACION",
  "MAESTRA": "MaestraEstaciones.xlsx"
}
```

## Archivo maestro de estaciones

Para que el aplicativo funcione, debe existir localmente un archivo llamado:

```text
MaestraEstaciones.xlsx
```

Este archivo debe tener como mínimo las siguientes columnas:

| Código | Nombre | Clasificación | DZ |
|---|---|---|---|
| 000001 | ESTACION_EJEMPLO | CO | DZ01 |
| 000002 | ESTACION_PRUEBA | EMA | DZ02 |

El repositorio incluye `MaestraEstaciones.example.xlsx` como plantilla de referencia.  
Para usarla:

```bash
cp MaestraEstaciones.example.xlsx MaestraEstaciones.xlsx
```

Luego reemplazar los datos ficticios por los datos reales en el entorno local.

## Ejecución de la interfaz gráfica

Para abrir el aplicativo:

```bash
python legajos_app.py
```

## Uso desde terminal

También se puede ejecutar la lógica principal desde terminal mediante `legajos_core.py`.

### Crear estructura base

```bash
python legajos_core.py init --dz DZ03 --years 2024 2025 2026 --estaciones 000001,000002 --maestra MaestraEstaciones.xlsx
```

### Agregar informe por estación

```bash
python legajos_core.py add --src archivo.pdf --categoria MANTENIMIENTO --dz DZ03 --codigo 000001 --fecha 15-06-2026 --maestra MaestraEstaciones.xlsx
```

### Agregar informe por ruta

```bash
python legajos_core.py addruta --src archivo.pdf --dz DZ03 --ruta RUTA_01 --tipo MANTENIMIENTOS --fecha 15-06-2026 --estaciones 000001,000002 --maestra MaestraEstaciones.xlsx
```

### Agregar checklist de ruta

```bash
python legajos_core.py addchecklist --src checklist.pdf --dz DZ03 --ruta RUTA_01 --codigo 000001 --fecha 15-06-2026 --maestra MaestraEstaciones.xlsx
```

### Agregar estado situacional

```bash
python legajos_core.py addestado_situacional --src estado.pdf --dz DZ03 --ruta RUTA_01 --codigo 000001 --fecha 15-06-2026 --maestra MaestraEstaciones.xlsx
```

### Agregar convenio de Dirección Zonal

```bash
python legajos_core.py addconvenio_dz --src convenio.pdf --dz DZ03 --fecha 15-06-2026 --estaciones 000001,000002 --maestra MaestraEstaciones.xlsx
```

### Reconstruir índice de una estación

```bash
python legajos_core.py index --dz DZ03 --year 2026 --codigo 000001 --maestra MaestraEstaciones.xlsx
```

## Formato de fechas

Todas las fechas deben ingresarse en formato:

```text
DD-MM-YYYY
```

Ejemplo:

```text
15-06-2026
```

## Categorías disponibles

Las categorías por estación son:

```text
MANTENIMIENTO
INSPECCION
CALIBRACION
AFOROS
CALIDAD_DATOS
INCIDENCIAS
INSTALACIONES_NUEVAS
```

## Tipos de ruta disponibles

```text
MANTENIMIENTOS
AFOROS
INSPECCION
```

## Recomendación para Google Drive

Si se desea guardar la información directamente en Google Drive, primero se debe montar o sincronizar la carpeta correspondiente y luego configurar la ruta en `config.json`.

Ejemplo:

```json
{
  "ROOT": "/home/usuario/Drive/DRD_LEGAJOS_ESTACION",
  "MAESTRA": "MaestraEstaciones.xlsx"
}
```

El aplicativo usará esa ruta como carpeta raíz para crear las Direcciones Zonales, años, estaciones e índices.

## Notas de seguridad

Antes de subir cambios a GitHub, verificar que no se esté incluyendo información sensible:

```bash
git status
```

No deben aparecer:

```text
config.json
MaestraEstaciones.xlsx
LEGAJOS_ESTACION/
OBSOLETOS/
__pycache__/
```

Si alguno aparece, revisar el archivo `.gitignore`.

## Autor

Proyecto desarrollado para la gestión documental de legajos de estaciones meteorológicas.
