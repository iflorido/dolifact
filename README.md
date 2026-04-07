# Dolibarr Facturación Excel

> Exporta la facturación de tu instalación Dolibarr a Excel con un solo clic.

Desarrollado por **AutomaWorks** · [automaworks.es](https://automaworks.es) · [iflorido@gmail.com](mailto:iflorido@gmail.com)

---

## ¿Qué hace esta aplicación?

Dolibarr Facturación Excel es una app de escritorio para macOS que se conecta a la API REST de tu instalación de Dolibarr y genera un archivo `.xlsx` con el detalle de todas las facturas emitidas en un rango de fechas indicado.

El Excel generado incluye por cada factura:

- Fecha de factura
- Número de factura
- Nombre de la empresa
- CIF de la empresa
- Dirección de la empresa
- Base imponible
- IVA
- Total

Al final del archivo se incluye una fila de **totales** con la suma de base, IVA y total del período.

---

## Requisitos

### Para ejecutar el `.py` directamente

- Python 3.12 o superior
- [Anaconda](https://www.anaconda.com/) (recomendado)
- Entorno conda con las siguientes dependencias:

```bash
conda create -n dolifact python=3.12
conda activate dolifact
pip install flet requests openpyxl
```

### Para compilar la app de macOS

- Xcode instalado y configurado:
```bash
xcodebuild -runFirstLaunch
```
- Flutter (se instala automáticamente con Flet CLI)
- Flet CLI 0.84 o superior:
```bash
pip install flet
```

---

## Estructura del proyecto

```
dolifact/
├── main.py                  # Código principal de la aplicación
├── requirements.txt         # Dependencias Python
├── pyproject.toml           # Configuración del build
├── automaworks_logo.png     # Logo original
└── assets/
    └── icon.png             # Icono de la aplicación (1024×1024 px)
```

---

## Ejecución en desarrollo

```bash
conda activate dolifact
python main.py
```

---

## Compilar la app de macOS

Desde la carpeta raíz del proyecto:

```bash
flet build macos \
  --project "DolibarrFacturacion" \
  --product "Dolibarr Facturación" \
  --org "es.automaworks" \
  --company "AutomaWorks" \
  --clear-cache
```

La app compilada se genera en `build/macos/`. Puedes moverla a tu carpeta `/Applications` para usarla como cualquier otra app de macOS.

> **Primera vez en un Mac nuevo:** si macOS bloquea la app al abrirla, ve a *Ajustes del sistema → Privacidad y seguridad → Abrir igualmente*.

---

## Configuración de Dolibarr

Para que la aplicación pueda conectarse a tu Dolibarr necesitas:

1. Tener activado el módulo **API REST** en tu instalación:
   `Configuración → Módulos/Aplicaciones → API/WebServices`

2. Generar una **API Key** de usuario con permisos de lectura sobre:
   - Facturas (`/api/index.php/invoices`)
   - Terceros (`/api/index.php/thirdparties`)

3. Introducir en la app:
   - **URL**: la URL base de tu Dolibarr (ej. `https://midominio.com/dolibarr`)
   - **API Key**: la clave generada en el paso anterior

---

## Uso

1. Abre la aplicación
2. Introduce la URL de tu instalación Dolibarr
3. Introduce tu API Key
4. Selecciona la **fecha de inicio** y **fecha final** en formato `DD-MM-YYYY`
5. Pulsa **Generar Excel**
6. La app mostrará el progreso en tiempo real con 5 pasos:
   - Conectando con Dolibarr
   - Descargando facturas
   - Obteniendo datos de empresas
   - Generando Excel
   - Guardando archivo
7. El archivo `.xlsx` se guarda automáticamente en el **Escritorio** con el nombre `facturacion_YYYYMMDD_YYYYMMDD.xlsx`

---

## Notas técnicas

- La app pagina las facturas de 100 en 100 y aplica el filtro de fechas localmente, ya que el parámetro `sqlfilters` de la API no funciona en todas las versiones de Dolibarr.
- El CIF se obtiene del campo `idprof1` del tercero, que es el estándar en instalaciones españolas de Dolibarr. Si tu instalación usa otro campo, el código busca también en `idprof2`, `idprof3`, `idprof4` y `tva_intra`.
- Las llamadas a terceros se cachean durante la sesión para evitar peticiones repetidas a la API por el mismo cliente.

---

## Dependencias

| Paquete | Uso |
|---|---|
| `flet` | Framework de UI multiplataforma |
| `requests` | Llamadas a la API REST de Dolibarr |
| `openpyxl` | Generación del archivo Excel |

---

## pyproject.toml

```toml
[project]
name = "dolibarrfacturacion"
version = "0.1.0"
dependencies = [
  "flet",
  "requests",
  "openpyxl",
]

[tool.flet]
product = "Dolibarr Facturación"
company = "AutomaWorks"

[tool.flet.app]
path = "."
module = "main.py"

[tool.flet.macos]
org = "es.automaworks"
artifact = "Dolibarr Facturación"
```

---

## Desarrollador

**Ignacio Florido**
Desarrollador de aplicaciones · AutomaWorks

- 🌐 [automaworks.es](https://automaworks.es)
- 📧 [iflorido@gmail.com](mailto:iflorido@gmail.com)
- 💼 [cv.iflorido.es](https://cv.iflorido.es)

---

*© 2026 AutomaWorks. Todos los derechos reservados.*