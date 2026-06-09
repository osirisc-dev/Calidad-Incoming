# 📦 Sistema de Reporte Incoming (Control de Calidad)

Este proyecto es una solución basada en **Google Apps Script** para gestionar el control de calidad (Incoming Inspection) de máquinas y refacciones. Proporciona una interfaz web interactiva que se conecta directamente con Google Sheets y Google Drive para almacenar datos, evidencias fotográficas y generar reportes automatizados en PDF.

## ✨ Características Principales

* **Interfaz Web Dinámica:** Formulario responsivo (HTML/CSS/JS) con pestañas separadas para evaluar Máquinas completas o Refacciones individuales.
* **Buscador en Tiempo Real:** Consulta dinámica de catálogo de máquinas y base de datos de refacciones pendientes mediante códigos ERP/AR.
* **Manejo de Evidencias:** Permite capturar o subir fotos directamente desde el dispositivo, optimizando y reduciendo su tamaño mediante *Canvas API* antes de enviarlas al servidor.
* **Generación de Reportes PDF:** Creación automática de reportes de inspección (individuales y consolidados consolidados) con un diseño profesional.
* **Soporte Bilingüe Automático:** Utiliza `LanguageApp` para traducir automáticamente las notas y comentarios de los inspectores (Español a Inglés) dentro de los reportes PDF.
* **Integración con Google Workspace:** Archiva automáticamente los registros en Google Sheets y guarda las imágenes y PDFs en carpetas específicas de Google Drive.

## 📂 Estructura del Proyecto

* `codigo.gs`: Contiene el backend de la aplicación, configuraciones globales, y las funciones que sirven la interfaz web y guardan los datos en Sheets y Drive.
* `reporte.gs`: Módulo dedicado a la compilación del HTML para generar los archivos PDF automáticos y consolidados, así como el manejo de las firmas de los inspectores.
* `index.html`: El frontend (UI) de la aplicación, incluyendo estilos, lógica de pestañas, compresión de imágenes en Base64 y comunicación asíncrona con el backend (`google.script.run`).

## 🚀 Instalación y Configuración

Dado que este proyecto maneja datos confidenciales, el código en este repositorio sirve como plantilla. Para implementarlo en un entorno de producción, sigue estos pasos:

1. Crea un nuevo proyecto en [Google Apps Script](https://script.google.com/).
2. Copia los archivos de este repositorio (`codigo.gs`, `reporte.gs` e `index.html`) en tu nuevo proyecto.
3. **Paso Crucial:** En la parte superior de `codigo.gs` y `reporte.gs`, busca los objetos `CONFIG` y `CONFIG_REPORTES`.
4. Reemplaza los valores de marcador de posición (ej. `"TU_ID_DEL_SHEET_AQUI"`, `"TU_ID_DE_CARPETA_DRIVE_AQUI"`) con los IDs reales de tus documentos de Google Workspace.
5. En `index.html`, actualiza la URL de la etiqueta `<img>` con el enlace público del logo de tu empresa y ajusta la validación del dominio en el input del correo electrónico.
6. Despliega el proyecto como una **Aplicación Web** (*Web App*).

## ⚠️ Privacidad y Seguridad

Este repositorio **no contiene** datos sensibles, IDs de carpetas, direcciones de correo electrónico personales, ni nombres de bases de datos de producción. Todo dato confidencial ha sido extraído a objetos de configuración que deben ser llenados localmente al momento del despliegue.
