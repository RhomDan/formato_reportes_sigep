# 📊 Transformador y Formateador de Reportes SIGEP

Aplicación web desarrollada en **Python (Flask)** que permite cargar archivos de hojas de cálculo (`.xls` y `.xlsx`), aplicar transformaciones de datos automáticas y estructurar el formato visual (colores de cabecera, bordes, alineación y autoajuste de columnas) para descargar un reporte limpio.

Mismo que se puede ingresar a través del siguiente link:

https://formato-reportes-sigep.onrender.com/
---

## 🚀 Características

* **Carga de archivos flexibles:** Compatible con formatos `.xlsx`, `.xls` antiguos e incluso archivos de tabla HTML/CSV exportados con extensión `.xls`.
* **Procesamiento dinámico:** Múltiples opciones de transformación según el tipo de reporte seleccionado (Análisis de Consistencia, Ejecución Presupuestaria, Resumen Contable, etc.).
* **Formateo automático:**
  * Limpieza de espacios en blanco y estandarización de columnas.
  * Encabezados con estilos visuales personalizados (fuentes, rellenos de color y alineación).
  * Autoajuste automático del ancho de columna según el contenido.
  * Aplicación de bordes delgados a la grilla de datos.
* **Procesamiento en memoria:** No almacena ni acumula archivos temporales en el servidor, garantizando alta eficiencia y seguridad.

---

## 🛠️ Tecnologías Utilizadas

* **Backend:** Python 3, Flask, Pandas, OpenPyXL, Xlrd, Gunicorn.
* **Frontend:** HTML5, CSS3, JavaScript.
* **Control de versiones:** Git & GitHub.

---

## 📁 Estructura del Proyecto

```text
formato_reportes_sigep/
│
├── app.py                  # Servidor Flask y lógica de transformación con Pandas/OpenPyXL
├── requirements.txt        # Lista de dependencias de Python
├── .gitignore              # Archivos y carpetas ignorados por Git (entorno virtual, temporales)
├── README.md               # Documentación del proyecto
│
├── templates/
│   └── index.html          # Interfaz de usuario web
│
└── static/
    └── style.css           # Estilos visuales de la aplicación