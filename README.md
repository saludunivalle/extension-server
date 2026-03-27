# Extension-server

Backend en Node.js/Express para gestionar solicitudes de extension, guardar formularios por etapas en Google Sheets, subir archivos a Google Drive y generar reportes (formularios 1-4) a partir de plantillas.

## 1. Que hace este proyecto

Este servicio expone una API HTTP que permite:

- autenticar usuarios con token de Google
- crear y consultar solicitudes
- guardar progreso por formulario/paso
- guardar gastos y riesgos asociados a una solicitud
- generar reportes (Google Sheets/Excel) usando plantillas
- descargar reportes generados

La persistencia principal no es una base de datos tradicional: los datos viven en Google Sheets (hojas como `SOLICITUDES`, `ETAPAS`, `GASTOS`, `RIESGOS`, etc.).

## 2. Stack tecnico

- Node.js + Express
- Google APIs (`googleapis`, `google-auth-library`)
- Carga de archivos con `multer`
- Manejo de sesiones con `express-session`
- Manipulacion de Excel con `exceljs` y `xlsx`
- Despliegue compatible con Vercel (`vercel.json`)

## 3. Estructura del proyecto

```text
extension-server/
	index.js                      # Arranque del servidor y middlewares globales
	config/google.js              # Clientes Google OAuth2 y JWT
	routes/                       # Definicion de endpoints
	controllers/                  # Logica HTTP por dominio
	services/                     # Integraciones Google Sheets/Drive y logica de negocio
	models/spreadsheetModels.js   # Esquemas y mapeos de columnas por hoja/paso
	reportConfigs/                # Transformaciones para reportes 1-4
	middleware/                   # auth, errores, progreso
	utils/                        # utilidades varias
	templates/                    # plantillas JSON para filas dinamicas
```

## 4. Requisitos para correrlo

## 4.1 Software

- Node.js 18+ (recomendado 20 LTS)
- npm 9+

## 4.2 Servicios de Google

Necesitas tener en Google Cloud:

- API de Google Sheets habilitada
- API de Google Drive habilitada
- una Service Account con credenciales JSON
- acceso (comparticion) de la hoja de calculo y de las plantillas con el correo de la Service Account

Tambien necesitas OAuth Client (para validar ID token de Google en login).

## 5. Variables de entorno

Crea un archivo `.env` en la raiz con al menos lo siguiente:

```env
# Entorno
NODE_ENV=development
PORT=3001
SESSION_SECRET=tu_secreto_de_sesion

# OAuth (usado en auth y middleware)
GOOGLE_CLIENT_ID=tu_google_client_id
ADMIN_EMAILS=admin1@dominio.com,admin2@dominio.com
REDIRECT_URI=http://localhost:5173

# Credenciales JWT de Service Account
GOOGLE_CLIENT_EMAIL=service-account@proyecto.iam.gserviceaccount.com
GOOGLE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"

# OAuth2 client en formato usado por config/google.js
client_id=tu_oauth_client_id
client_secret=tu_oauth_client_secret

# Google Spreadsheet principal
SPREADSHEET_ID=tu_spreadsheet_id
```

Notas importantes:

- `GOOGLE_PRIVATE_KEY` debe mantener los `\\n` en el `.env`.
- En produccion, `SESSION_SECRET` es obligatorio y el servidor se detiene si no esta definido.
- El servicio tiene IDs de carpetas/plantillas hardcodeados en `services/driveService.js`. Si usas otro entorno de Drive, debes actualizarlos.

## 6. Instalacion y ejecucion local

1. Instalar dependencias:

```bash
npm install
```

2. Configurar `.env` (seccion anterior).

3. Iniciar servidor:

```bash
npm start
```

4. Verificar que responde:

- Base URL local: `http://localhost:3001`
- Endpoint raiz: `GET /` responde `Hola desde la ruta principal`

## 7. Configuracion de Google (paso a paso)

1. Crea o usa un proyecto en Google Cloud.
2. Habilita Google Sheets API y Google Drive API.
3. Crea una Service Account y descarga credenciales.
4. Copia `client_email` y `private_key` al `.env` (`GOOGLE_CLIENT_EMAIL`, `GOOGLE_PRIVATE_KEY`).
5. Crea un OAuth Client ID (Web app) y coloca `GOOGLE_CLIENT_ID`, `client_id`, `client_secret`.
6. Comparte con la Service Account:
	 - la hoja de calculo principal (`SPREADSHEET_ID`)
	 - carpetas/plantillas de reportes en Drive

Si no compartes esos recursos, veras errores de permisos al leer/escribir o al generar reportes.

## 8. Como funciona internamente

## 8.1 Flujo general

1. El frontend llama endpoints del backend.
2. Controllers validan request y delegan en services.
3. `sheetsService` lee/escribe datos en hojas segun modelos y mapeos por paso.
4. `driveService` sube archivos y genera reportes desde plantillas.
5. Se devuelve JSON al frontend (o archivo en descargas).

## 8.2 Formularios y hojas

El sistema trabaja principalmente con estas hojas:

- `SOLICITUDES` (formulario 1)
- `SOLICITUDES2` (formulario 2/presupuesto)
- `SOLICITUDES3` (formulario 3/riesgos)
- `SOLICITUDES4` (formulario 4/mercadeo relacional)
- `GASTOS`
- `RIESGOS`
- `ETAPAS` (seguimiento de avance)
- `USUARIOS`

La relacion `hoja -> nombre` se resuelve en `formController` con `getSheetName`.

## 8.3 Progreso y etapas

- El avance se guarda en `ETAPAS` (etapa actual, paso, estado global, estado por formulario).
- Tambien se usa sesion (`express-session`) para estado temporal.
- `progressStateService` prioriza sesion y sincroniza con Google Sheets.

## 8.4 Generacion de reportes

`reportService`:

- carga configuracion por formulario (`reportConfigs/report1Config.js` ... `report4Config.js`)
- obtiene datos de hojas requeridas
- transforma datos segun reglas del reporte
- llama a `driveService` para duplicar/procesar plantilla
- devuelve link o archivo descargable

Soporta filas dinamicas para gastos/riesgos usando `services/dynamicRows/`.

## 9. Endpoints principales

Nota: algunas rutas estan expuestas tanto en raiz (`/`) como bajo prefijos (`/form`, `/report`, etc.) para compatibilidad.

## 9.1 Auth

- `POST /auth/google` -> valida token de Google y registra usuario en `USUARIOS` si no existe

## 9.2 Usuario

- `POST /saveUser`
- `POST /user/saveUser`

## 9.3 Solicitudes y formularios

- `POST /createNewRequest`
- `GET /getRequests`
- `GET /getActiveRequests`
- `GET /getCompletedRequests`
- `GET /getLastId`
- `POST /guardarProgreso` (con archivo `pieza_grafica` opcional)
- `POST /guardarGastos`
- `GET /getGastos`
- `GET /getFormDataForm2`
- `POST /guardarForm2Paso1`
- `POST /guardarForm2Paso2`
- `POST /guardarForm2Paso3`

Tambien disponibles bajo `/form/*`.

## 9.4 Progreso

- `POST /actualizarPasoMaximo`
- `POST /progreso-actual`
- `POST /actualizacion-progreso`
- `POST /actualizacion-progreso-global`
- `GET /progress/:id_solicitud`
- `PUT /progress/:id_solicitud`

## 9.5 Riesgos

- `GET /riesgos`
- `POST /riesgos`
- `PUT /riesgos`
- `DELETE /riesgos/:id_riesgo`
- `GET /risk/categorias-riesgo`
- `POST /risk/migrar-riesgos-form3`

## 9.6 Reportes

- `POST /report/generateReport`
- `POST /report/downloadReport`
- `POST /report/generateReport1`
- `POST /report/generateReport2`
- `POST /report/generateReport3`
- `POST /report/generateReport4`
- `POST /report/testRiskRows` (pruebas de filas dinamicas de riesgos)

## 9.7 Otros catalogos/consulta

- `GET /getProgramasYOficinas`
- `GET /getSolicitud`
- `GET /other/getProgramasYOficinas`
- `GET /other/getSolicitud`

## 10. CORS, sesiones y errores

- CORS habilitado para:
	- `http://localhost:5173`
	- `https://siac-extension-form.vercel.app`
- Sesiones con `express-session`.
- Manejo centralizado de errores en `middleware/errorHandler.js`.

## 11. Despliegue (Vercel)

Existe configuracion en `vercel.json` para ejecutar `index.js` como funcion Node.

Antes de desplegar, valida:

- todas las variables de entorno en el panel de Vercel
- permisos de Google (Drive/Sheets) para la Service Account
- que `SESSION_SECRET` este definido

## 12. Problemas comunes y solucion rapida

- Error de autenticacion Google:
	- revisa `GOOGLE_CLIENT_ID`, `client_id`, `client_secret` y token enviado desde frontend.

- Error de permisos en Sheets/Drive:
	- comparte hoja/carpeta/plantillas con `GOOGLE_CLIENT_EMAIL`.

- Reporte no se genera:
	- verifica IDs de plantilla y carpetas en `services/driveService.js`.
	- confirma que los datos de la solicitud existan en las hojas requeridas.

- Estado/progreso inconsistente:
	- revisar valores de `ETAPAS` y payload de `actualizarPasoMaximo` / `actualizacion-progreso`.

## 13. Resumen funcional rapido

Este backend centraliza el ciclo de vida de una solicitud de extension:

- crea solicitud
- guarda avance por pasos
- administra gastos y riesgos
- consolida informacion de varias hojas
- genera reportes listos para consulta/descarga

Todo el sistema gira alrededor de Google Sheets como fuente de verdad y Google Drive como repositorio de documentos/reportes.



Campos que cambiaron:
Nuevo: Entradas para el diseño, (programma nuevo, actualizacion de programas, modificacion de programa)
Camabiaron: Periodicidad de la oferta, Certificado o constancia que solicita expedir, intensidad horaria
Eliminar: Pieza grafica y personal externo asignado

En caso de modificación, se debe incluir la descripción detallada del ajuste y en caso de actualización, se deberá indicar que elementos son objeto de actualización