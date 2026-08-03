# Cómo ejecutar y entender el proyecto

## 1. Introducción breve

Esta aplicación es una herramienta interna desarrollada en Django para procesar un archivo Excel con datos de documentos y facturas, buscar los correos y adjuntos correspondientes en una cuenta de Outlook/Exchange mediante Microsoft Graph, y generar un archivo ZIP con los PDFs encontrados renombrados automáticamente.

El proyecto usa Django como framework principal y está organizado en una única app llamada renombrar_factura, junto con el paquete principal app que contiene la configuración del proyecto. El flujo principal se concentra en una interfaz web simple que permite subir un Excel, procesar las filas, mostrar resultados y descargar un ZIP final.

## 2. Funcionamiento

El funcionamiento real del proyecto es el siguiente:

1. El usuario accede a la URL raíz del proyecto, que está mapeada al view inicio.
2. El view inicio renderiza la plantilla buscar_adjuntos.html con un contexto inicial vacío o con datos previamente cargados en la sesión.
3. Cuando el usuario sube un archivo Excel, la ruta /extraer-informacion/ recibe la petición POST. El view cargar_excel lee el archivo con openpyxl.
4. El proceso empieza a leer el libro desde la fila 6. De cada fila toma:
   - la columna 1: documento
   - la columna 2: factura
   - la columna 27: NIT del proveedor
   - la columna 28: nombre del proveedor
5. Por cada fila, el sistema arma un criterio de búsqueda usando la factura. Primero intenta buscar con el formato FEVD + número de factura; si no encuentra, repite con el número de factura solo.
6. La búsqueda real de correos y adjuntos se hace en el módulo graph.py mediante Microsoft Graph API, usando credenciales de cliente y una cuenta de buzón configurada en el código.
7. Si el correo tiene adjuntos, el sistema revisa si alguno es un PDF directo o un ZIP que contenga un PDF. En ambos casos, el PDF se extrae y se prepara para la descarga.
8. Cada fila procesada se guarda en la sesión del usuario bajo la clave filas_excel. Allí se almacenan datos como si el PDF fue encontrado, el nombre del archivo adjunto y el error de búsqueda si aplica.
9. Después de procesar el Excel, la interfaz muestra una tabla con el estado de cada fila y permite continuar al paso de generación del ZIP.
10. Al enviar la petición a /generar-reporte/, el view generar_reporte_final toma las filas almacenadas en sesión, elimina duplicados por factura y crea un archivo Excel de resumen llamado reporte.xlsx.
11. A continuación, genera un ZIP que contiene cada PDF encontrado, renombrado como documento-nombreProveedor-factura.pdf, junto con el archivo reporte.xlsx.
12. El ZIP se devuelve al navegador como descarga y se limpia la información de sesión para que la operación quede cerrada.

Es importante destacar que este proyecto no usa modelos Django en la práctica: el archivo models.py está vacío y la lógica de negocio no persiste datos en base de datos, sino que se maneja en memoria y en la sesión del usuario.

## 3. Estructura técnica

| Archivo o carpeta | Rol |
|---|---|
| [app](app) | Carpeta principal del proyecto Django. |
| [app/manage.py](app/manage.py) | Punto de entrada para ejecutar comandos de Django como runserver, migrate y createsuperuser. |
| [app/app/settings.py](app/app/settings.py) | Configuración del proyecto: apps instaladas, middleware, templates, base de datos, rutas estáticas y configuración general. |
| [app/app/urls.py](app/app/urls.py) | Configuración de URLs del proyecto principal. Incluye las rutas del admin y del módulo de la app. |
| [app/app/wsgi.py](app/app/wsgi.py) | Punto de entrada WSGI para despliegues compatibles con WSGI. |
| [app/app/asgi.py](app/app/asgi.py) | Punto de entrada ASGI para despliegues asíncronos. |
| [app/renombrar_factura](app/renombrar_factura) | App principal del proyecto. Contiene la lógica de la interfaz, rutas, vistas y conexión con Microsoft Graph. |
| [app/renombrar_factura/views.py](app/renombrar_factura/views.py) | Contiene las vistas de inicio, carga del Excel, procesamiento de filas y generación del ZIP final. |
| [app/renombrar_factura/urls.py](app/renombrar_factura/urls.py) | Define las rutas específicas del flujo de carga y generación del reporte. |
| [app/renombrar_factura/graph.py](app/renombrar_factura/graph.py) | Implementa la integración con Microsoft Graph para buscar mensajes y adjuntos en el buzón configurado. |
| [app/renombrar_factura/models.py](app/renombrar_factura/models.py) | Archivo de modelos; actualmente no define modelos ni se usa para persistencia. |
| [app/templates/buscar_adjuntos.html](app/templates/buscar_adjuntos.html) | Plantilla HTML que presenta la interfaz de carga de Excel, resultados y botón de descarga del ZIP. |
| [app/db.sqlite3](app/db.sqlite3) | Base de datos SQLite usada por Django. |
| [app/renombrar_factura/migrations](app/renombrar_factura/migrations) | Carpeta para migraciones; en este proyecto no hay migraciones adicionales definidas. |

## 4. Endpoints principales

| Ruta | Función | Propósito |
|---|---|---|
| / | inicio | Muestra la interfaz principal de la aplicación. |
| /extraer-informacion/ | cargar_excel | Recibe el archivo Excel, procesa las filas y almacena los resultados en sesión. |
| /generar-reporte/ | generar_reporte_final | Genera el ZIP con los PDFs encontrados y el archivo Excel de resumen. |
| /admin/ | admin.site.urls | Panel administrativo de Django. |

## 5. Requisitos

### Lenguaje y entorno

- Python: el proyecto está desarrollado en Python y usa Django.
- Node.js: no se utiliza en este proyecto.
- No existe un archivo requirements.txt en el repositorio, por lo que las dependencias deben instalarse manualmente.

### Dependencias principales

El código importa explícitamente estas librerías:

- Django
- openpyxl
- requests

### Base de datos

- Tipo: SQLite.
- Configuración: definida en [app/app/settings.py](app/app/settings.py) con el archivo [app/db.sqlite3](app/db.sqlite3).
- En este proyecto no se usan modelos para persistir negocio; la base de datos se usa principalmente por el propio funcionamiento de Django y por la posibilidad de usar el panel administrativo.

### Servicios externos

- Microsoft Graph API: se usa para consultar correos y adjuntos del buzón configurado.
- Cuenta de correo compartida o de usuario: el buzón está fijado directamente en [app/renombrar_factura/graph.py](app/renombrar_factura/graph.py) con la variable MAILBOX.
- Credenciales de cliente: el archivo graph.py contiene tenant, client y secret hardcodeados para autenticación por client credentials.

## 6. Estructura recomendada de ejecución

Siga estos pasos en Windows con PowerShell desde la raíz del proyecto.

1. Abrir PowerShell y ubicarse en la carpeta del proyecto:

   ```powershell
   Set-Location "C:\PROYECTO\renombrar-facturas-PM"
   ```

2. Crear un entorno virtual dentro de la carpeta del proyecto:

   ```powershell
   py -3 -m venv .venv
   ```

3. Activar el entorno virtual:

   ```powershell
   .\.venv\Scripts\Activate.ps1
   ```

4. Actualizar pip:

   ```powershell
   python -m pip install --upgrade pip
   ```

5. Instalar las dependencias necesarias. Como no existe requirements.txt, instálelas explícitamente:

   ```powershell
   pip install django openpyxl requests
   ```

6. Entrar a la carpeta donde está el proyecto Django:

   ```powershell
   Set-Location "C:\PROYECTO\renombrar-facturas-PM\app"
   ```

7. Ejecutar las migraciones para preparar la base de datos SQLite:

   ```powershell
   python manage.py migrate
   ```

8. Crear un superusuario opcional si desea usar el panel administrativo:

   ```powershell
   python manage.py createsuperuser
   ```

9. Levantar el servidor de desarrollo:

   ```powershell
   python manage.py runserver 0.0.0.0:8000
   ```

10. Abrir el navegador en:

   ```text
   http://127.0.0.1:8000/
   ```

### Variables de entorno

No existe un archivo de entorno (.env) ni lógica de configuración basada en variables de entorno en este proyecto. La configuración sensible de Microsoft Graph está hardcodeada en [app/renombrar_factura/graph.py](app/renombrar_factura/graph.py), por lo que para un entorno real se recomienda mover esos valores a variables de entorno o a un archivo de configuración seguro.

## 7. Opción con script

No existe en este repositorio ningún script de despliegue, Dockerfile, docker-compose.yml, build_files.sh ni equivalente. Por tanto, no hay un script automático para compilar o desplegar la app.

La ejecución local se realiza directamente con Django a través de manage.py.

## 8. Notas finales

- El proyecto está pensado como una herramienta de escritorio/web interna para un flujo muy concreto: leer un Excel, buscar adjuntos en Outlook mediante Microsoft Graph y entregar un ZIP listo para descargar.
- La interfaz está hecha en HTML y JavaScript incrustado en [app/templates/buscar_adjuntos.html](app/templates/buscar_adjuntos.html).
- La aplicación usa sesión de Django para conservar los resultados entre pasos del flujo; esto evita persistir los datos en base de datos.
- El funcionamiento depende completamente de que la cuenta de Microsoft Graph configurada tenga permisos adecuados para leer mensajes y adjuntos del buzón definido.
- En el entorno actual se verificó que las dependencias necesarias no estaban instaladas, por lo que la instalación explícita de Django, openpyxl y requests es obligatoria antes de ejecutar la app.
- La ruta del buzón y las credenciales de Graph están fijadas en el código, por lo que cualquier cambio operativo debe hacerse ahí.
