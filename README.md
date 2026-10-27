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

# Carpeta de Drive para adjuntos, firmas y resoluciones (ID o URL de carpeta)
DRIVE_UPLOADS_FOLDER=x5UipLOc1kz9Yzssk-Nz5H8HNmKCsWm5
```

Notas importantes:

- `GOOGLE_PRIVATE_KEY` debe mantener los `\\n` en el `.env`.
- En produccion, `SESSION_SECRET` es obligatorio y el servidor se detiene si no esta definido.
- `DRIVE_UPLOADS_FOLDER` contiene el ID o URL de la carpeta de Drive donde se guardan
  adjuntos, firmas y resoluciones. También se aceptan las variables antiguas
  `DRIVE_UPLOADS_FOLDER_ID` y `UPLOADS_FOLDER_ID` como compatibilidad.
- La carpeta de adjuntos debe estar compartida con `GOOGLE_CLIENT_EMAIL` (Service Account) con rol Editor.
- Si no hay acceso a carpeta, backend responderá error explícito indicando que se comparta con la Service Account.
- Si la carpeta está en `Mi unidad`, Google puede devolver `storageQuotaExceeded` para Service Accounts (sin cuota propia). Recomendado: usar carpeta dentro de un Shared Drive.

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
- La hoja `ETAPAS` ahora usa:
	- columna `J` = `estado_general` (`En proceso`, `Terminado`, `Enviado a revisión`, `Aprobado parcialmente`, `Aprobado`)
	- columna `K` = `comentarios` (JSON de observaciones independientes agrupadas por etapa)
- `estado_formularios` (columna `I`) ahora maneja estados por formulario:
	- `En progreso`
	- `Completado`
	- `Enviado a revisión`
	- `Aprobado`
	- `Requiere correcciones`
- Reglas de `estado_general`:
	- si los 4 formularios estan `Aprobado` -> `Aprobado`
	- si hay al menos 1 formulario `Aprobado` y no todos estan aprobados -> `Aprobado parcialmente`
	- si hay al menos 1 formulario en `Enviado a revisión` y ninguno aprobado -> `Enviado a revisión`
	- si los 4 formularios estan `Completado` (sin aprobaciones aun) -> `Terminado`
	- en otros casos -> `En proceso`
- Tambien se usa sesion (`express-session`) para estado temporal.
- `progressStateService` prioriza sesion y sincroniza con Google Sheets.

## 8.5 Flujo nuevo usuario/admin

1. Usuario crea/edita solicitud y completa formularios 1-4.
2. El backend calcula `estado_general` en `ETAPAS!J`.
3. El usuario puede enviar a revision solo los formularios que ya tenga realizados (no necesita tener los 4 completos).
4. Usuario ejecuta `POST /enviarFormulariosRevision` (o `POST /enviarSolicitudRevision` por compatibilidad).
5. Admin consulta pendientes con `GET /admin/solicitudesRevision` y revisa formularios puntuales.
6. Admin puede:
	- aprobar parcialmente hasta 4 formularios con `POST /admin/aprobarFormularios`.
	- aprobar toda la solicitud con `POST /admin/aprobarSolicitudCompleta`.
	- enviar correcciones por formulario con `POST /admin/enviarCorrecciones`.
7. Si hay formularios aprobados y otros con correcciones/revision, el estado queda `Aprobado parcialmente`.
8. El usuario consulta estados/comentarios con `GET /estadoRevisionSolicitud`.

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
- `POST /guardarProgreso` (con adjuntos opcionales `pieza_grafica` y `archivo_fondo_comun`)
- `POST /guardarGastos`
- `GET /getGastos`
- `GET /getFormDataForm2`
- `POST /guardarForm2Paso1`
- `POST /guardarForm2Paso2`
- `POST /guardarForm2Paso3`
- `POST /enviarFormulariosRevision`
- `POST /enviarSolicitudRevision`
- `GET /estadoRevisionSolicitud`
- `GET /admin/solicitudesRevision`
- `POST /admin/aprobarFormularios`
- `POST /admin/aprobarSolicitud`
- `POST /admin/aprobarSolicitudCompleta`
- `POST /admin/enviarCorrecciones`

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


- Lo del Salario Mínimo, se debe poder automático. Y la entrada es en pesos -x
- Preguntar primero, si en pesos o en salarios -x
- El porcentaje de los imprevistos debe ser editable, de 0 a 5 -x
- En presupuesto:
- los porcentajes del Otros Gastos deben calcularse -x
- Si el Fondo Comun cambia del 30% a otro valor, deben adjuntar el Formato de Viavilidad y aportes al fondo comun, firmado por el jefe de la Sección de Presupuesto.
- El porcentaje para Facultad o Instituto, es 5%, fijo -x
- Los ingresos no se estan guardando
- Si es Extensión Solidaria, los Ingresos van en Cero.
------ Se debe poner un texto para observaciones, y no debe ser en blanco
- Poner un texto label sieempre de advertencia en ese campo

-Matriz de riesgos es obligatoria cuando la actividad sea mayor o igual a 16 horas
-Mercadeo es sugerido

En el formulario de Mercadeo, cambiar en el paso3, 
 -¿Cuál de las siguientes estrategias han utilizado para sondear previamente el interés de las personas? -x
por 
¿Cuál(es) de las siguientes estrategias han utilizado para sondear previamente el interés de las personas? -x

- Ajustar nombre del archivo resultante -x
- Cambiar botones de "Enviar" a "Finalizar"

Salario minimo legal vigente: 1,750,905

Link carpeta de drive donde se guardan adjuntos y piezas graficas

## 9.8 Contratos JSON del flujo de revision

Todos estos endpoints estan disponibles en raiz (`/`) y tambien bajo `/form/*`.

### 9.8.1 Usuario envia formularios a revision

- Endpoint recomendado: `POST /enviarFormulariosRevision`
- Endpoint compatible: `POST /enviarSolicitudRevision`

Request:

```json
{
	"id_solicitud": "123",
	"userId": "google-user-id",
	"formularios": [1, 2]
}
```

Notas:

- `formularios` es opcional. Si no se envia, backend intenta enviar todos los formularios ya realizados y no aprobados.

Response OK:

```json
{
	"success": true,
	"message": "Formularios enviados a revisión",
	"data": {
		"id_solicitud": "123",
		"formularios_enviados": [1, 2],
		"estado_formularios": {
			"1": "Enviado a revisión",
			"2": "Enviado a revisión",
			"3": "En progreso",
			"4": "En progreso"
		},
		"estado_general": "Enviado a revisión"
	}
}
```

Errores comunes:

- `400`: falta `id_solicitud`/`userId`, no hay formularios listos, o el estado actual no permite enviar alguno.
- `403`: el usuario no es dueno de la solicitud.
- `404`: la solicitud no existe en `ETAPAS`.

### 9.8.2 Admin lista solicitudes en revision

- Endpoint: `GET /admin/solicitudesRevision?userId=<adminId>`

Response OK:

```json
{
	"success": true,
	"data": [
		{
			"id_solicitud": "123",
			"id_usuario": "owner-user-id",
			"fecha": "2026-04-08",
			"name": "Usuario Solicitante",
			"etapa_actual": 4,
			"estado": "Completado",
			"nombre_actividad": "Mi actividad",
			"paso": 5,
			"estado_formularios": {
				"1": "Enviado a revisión",
				"2": "Aprobado",
				"3": "Requiere correcciones",
				"4": "En progreso"
			},
			"estado_general": "Aprobado parcialmente",
			"comentarios": "{\"3\":\"Ajustar presupuesto\"}",
			"comentarios_por_formulario": {
				"3": "Ajustar presupuesto"
			},
			"formularios_en_revision": [1],
			"formularios_aprobados": [2],
			"formularios_con_correcciones": [3]
		}
	]
}
```

Error esperado:

- `403`: `userId` no tiene rol `admin` en hoja `USUARIOS` (columna D).

### 9.8.3 Admin aprueba formularios (parcial, 1 a 4)

- Endpoint recomendado: `POST /admin/aprobarFormularios`
- Endpoint compatible: `POST /admin/aprobarSolicitud`

Request:

```json
{
	"id_solicitud": "123",
	"userId": "admin-google-user-id",
	"formularios": [1, 2, 3]
}
```

Response OK:

```json
{
	"success": true,
	"message": "Formularios aprobados correctamente",
	"data": {
		"id_solicitud": "123",
		"formularios_aprobados": [1, 2, 3],
		"estado_formularios": {
			"1": "Aprobado",
			"2": "Aprobado",
			"3": "Aprobado",
			"4": "Enviado a revisión"
		},
		"estado_general": "Aprobado parcialmente"
	}
}
```

Error esperado:

- `400`: se enviaron mas de 4 formularios o alguno no estaba en `Enviado a revisión`.

### 9.8.4 Admin aprueba solicitud completa

- Endpoint: `POST /admin/aprobarSolicitudCompleta`

Request:

```json
{
	"id_solicitud": "123",
	"userId": "admin-google-user-id"
}
```

Response OK:

```json
{
	"success": true,
	"message": "Solicitud aprobada completamente",
	"data": {
		"id_solicitud": "123",
		"formularios_aprobados": [1, 2, 3, 4],
		"estado_formularios": {
			"1": "Aprobado",
			"2": "Aprobado",
			"3": "Aprobado",
			"4": "Aprobado"
		},
		"estado_general": "Aprobado"
	}
}
```

Error esperado:

- `400`: existe al menos un formulario en `En progreso`.

### 9.8.5 Admin envia correcciones por formulario

- Endpoint: `POST /admin/enviarCorrecciones`

Request:

```json
{
	"id_solicitud": "123",
	"userId": "admin-google-user-id",
	"formularios": [2, 3],
	"comentarios_por_formulario": {
		"2": "Ajustar presupuesto",
		"3": "Completar mitigacion de riesgos"
	},
	"comentarios": "Observaciones generales opcionales"
}
```

Response OK:

```json
{
	"success": true,
	"message": "Correcciones enviadas correctamente",
	"data": {
		"id_solicitud": "123",
		"formularios_con_correccion": [2, 3],
		"estado_formularios": {
			"1": "Aprobado",
			"2": "Requiere correcciones",
			"3": "Requiere correcciones",
			"4": "Enviado a revisión"
		},
		"estado_general": "Aprobado parcialmente",
		"comentarios_por_formulario": {
			"2": "Ajustar presupuesto",
			"3": "Completar mitigacion de riesgos"
		}
	}
}
```

### 9.8.6 Usuario/Admin consulta estado y comentarios

- Endpoint: `GET /estadoRevisionSolicitud?id_solicitud=<id>&userId=<userId>`

Response OK:

```json
{
	"success": true,
	"data": {
		"id_solicitud": "123",
		"id_usuario": "owner-user-id",
		"fecha": "2026-04-08",
		"name": "Usuario Solicitante",
		"etapa_actual": 4,
		"estado": "En progreso",
		"nombre_actividad": "Mi actividad",
		"paso": 4,
		"estado_formularios": {
			"1": "Aprobado",
			"2": "Requiere correcciones",
			"3": "Enviado a revisión",
			"4": "En progreso"
		},
		"estado_general": "Aprobado parcialmente",
		"comentarios": "{\"2\":\"Ajustar presupuesto\"}",
		"comentarios_por_formulario": {
			"2": "Ajustar presupuesto"
		},
		"formularios_en_revision": [3],
		"formularios_aprobados": [1],
		"formularios_con_correcciones": [2]
	}
}
```

Notas frontend:

- Usuario normal solo puede consultar solicitudes propias.
- Admin puede consultar cualquier solicitudes que tiene estado general Enviado a revisíon.
- Para renderizar botones:
	- usuario: mostrar `Enviar a revisión` por formulario cuando ese formulario este en `Completado` o `Requiere correcciones`.
	- admin: mostrar checks por formulario para `Aprobar formularios` (1 a 4), boton `Aprobar solicitud completa` y boton `Enviar correcciones` por formulario.

## 9.9 Cambios Abril 2026 (campos nuevos y adjuntos)

### 9.9.1 Formulario 1 (`SOLICITUDES`)

Se agregaron estos campos para guardar en Sheets (uso frontend/backoffice):

- `otro_tipo_act`
- `extension_solidaria`
- `costo_extension_solidaria`
- `pieza_grafica`
- `personal_externo`
- `tipo_valor`
- `valor_unitario`

Importante:

- estos campos se guardan en `SOLICITUDES` y se devuelven en consultas
- no se muestran en los reportes XLSX/PDF generados
- `pieza_grafica` se guarda como URL pública de Drive (no binario en Sheets)

### 9.9.2 Formulario 2 (`SOLICITUDES2`)

Se agregaron estos campos para guardar en Sheets:

- `imprevistos_porcentaje`
- `archivo_fondo_comun`

Importante:

- `archivo_fondo_comun` admite imagen o PDF
- el backend sube el archivo a la misma carpeta de Drive de adjuntos y guarda en Sheets la URL pública
- estos campos no se muestran en el reporte final

### 9.9.3 Qué debe enviar el frontend (adjuntos)

Para `POST /guardarProgreso` y `POST /guardarForm2Paso3` cuando haya archivos:

- usar `multipart/form-data`
- enviar campos de texto normales
- enviar archivo en la llave `pieza_grafica` (solo imagen)
- enviar archivo en la llave `archivo_fondo_comun` (imagen o pdf)

Ejemplo mínimo (conceptual):

```text
Content-Type: multipart/form-data

id_solicitud=123
hoja=1
paso=5
...
pieza_grafica=<archivo imagen>
```

```text
Content-Type: multipart/form-data

id_solicitud=123
fondo_comun_porcentaje=30
...
archivo_fondo_comun=<archivo imagen o pdf>
```

El valor que persiste en Sheets para esos dos campos es siempre la URL del archivo en Drive.

### 9.9.4 Filas dinámicas en Formulario 2 (gastos)

Para filas dinámicas de presupuesto en frontend no se agregan columnas nuevas en `SOLICITUDES2`.

Qué espera este backend:

- enviar los ítems dinámicos en `POST /guardarGastos`
- payload con `id_solicitud` y arreglo `gastos`
- cada ítem debe incluir al menos: `id_conceptos`, `descripcion` (o `nombre_conceptos`), `cantidad`, `valor_unit`, `valor_total`
- los gastos dinámicos para reporte de presupuesto son los hijos del concepto `14`:
	- formatos válidos: `14.1`, `14.2`, `14,1`, `14,2`, etc.
	- el backend los detecta automáticamente para inserción dinámica en el Excel

Ejemplo:

```json
{
	"id_solicitud": "123",
	"gastos": [
		{
			"id_conceptos": "14.1",
			"descripcion": "Honorarios externos",
			"cantidad": 2,
			"valor_unit": 500000,
			"valor_total": 1000000
		}
	]
}
```

Qué se debe tener en Sheets:

- mantener hojas `CONCEPTO$` y `GASTOS` disponibles (ya usadas por el backend)
- no crear columnas extra en `SOLICITUDES2` para estas filas dinámicas
- el backend crea/actualiza conceptos dinámicos y valores en `GASTOS`

## 9.10 Flujos de revisión por subconjunto (frontend)

Este backend soporta que una solicitud trabaje solo con algunos formularios (por ejemplo 1 y 2) y que se revisen/aprueben por lotes parciales o completos.

### 9.10.1 Usuario envía a revisión uno, varios o todos

- Endpoint: `POST /enviarFormulariosRevision`
- `formularios` acepta subconjuntos: `[1]`, `[1,2]`, `[2,3]`, `[1,2,3,4]`
- Si omites `formularios`, backend intenta enviar todos los formularios realizados y no aprobados.

### 9.10.2 Admin aprueba uno, varios o cuatro en la misma llamada

- Endpoint: `POST /admin/aprobarFormularios`
- `formularios` acepta 1 a 4 ids por llamada.
- También existe `POST /admin/aprobarSolicitudCompleta` para aprobar los 4 de una vez con validación de estado.

### 9.10.3 Admin envía correcciones a uno, varios o cuatro

- Endpoint: `POST /admin/enviarCorrecciones`
- `formularios` acepta subconjuntos (`[2]`, `[2,3]`, `[1,2,3,4]`) siempre que esos formularios no estén en `En progreso` ni `Aprobado`.

### 9.10.4 Ejemplos directos (2, 3 y 4 formularios)

#### A) Enviar a revisión (`POST /enviarFormulariosRevision`)

```json
{
	"id_solicitud": 123,
	"id_usuario": 45,
	"formularios": [2, 3]
}
```

```json
{
	"id_solicitud": 123,
	"id_usuario": 45,
	"formularios": [1, 2, 3]
}
```

```json
{
	"id_solicitud": 123,
	"id_usuario": 45,
	"formularios": [1, 2, 3, 4]
}
```

#### B) Aprobar parcial (`POST /admin/aprobarFormularios`)

```json
{
	"id_solicitud": 123,
	"userId": 999,
	"formularios": [2, 3]
}
```

```json
{
	"id_solicitud": 123,
	"userId": 999,
	"formularios": [1, 2, 3]
}
```

```json
{
	"id_solicitud": 123,
	"userId": 999,
	"formularios": [1, 2, 3, 4]
}
```

Nota:

- para aprobar exactamente los 4 también puedes usar `POST /admin/aprobarSolicitudCompleta`.

#### C) Enviar correcciones (`POST /admin/enviarCorrecciones`)

```json
{
	"id_solicitud": 123,
	"userId": 999,
	"formularios": [2, 3],
	"comentarios": {
		"2": "Ajustar aportes",
		"3": "Corregir archivo de soporte"
	}
}
```

```json
{
	"id_solicitud": 123,
	"userId": 999,
	"formularios": [1, 2, 3],
	"comentarios": {
		"1": "Completar datos del coordinador",
		"2": "Ajustar imprevistos",
		"3": "Actualizar observaciones"
	}
}
```

```json
{
	"id_solicitud": 123,
	"userId": 999,
	"formularios": [1, 2, 3, 4],
	"comentarios": {
		"1": "Corregir formulario 1",
		"2": "Corregir formulario 2",
		"3": "Corregir formulario 3",
		"4": "Corregir formulario 4"
	}
}
```

Recomendación frontend:

- Mostrar checkboxes por formulario y permitir selección múltiple.
- Enviar solo formularios visibles/activos para la solicitud.
- Mantener UI preparada para escenarios parciales (solicitud solo de formularios 1 y 2).

### 9.10.5 Comentarios por paso (nuevo)

- Endpoint: `POST /guardarComentarioPaso`
- Disponible en raíz y también bajo `/form/guardarComentarioPaso`
- Guarda comentarios por formulario y paso dentro de `ETAPAS.K` (JSON)

Request:

```json
{
	"id_solicitud": 123,
	"userId": 45,
	"formulario": 2,
	"paso": 3,
	"comentario": "Validar el soporte del fondo común"
}
```

Notas:

- Permite guardar/actualizar comentario puntual por paso.
- Si `comentario` viene vacío, elimina el comentario de ese paso.
- El estado de comentarios se consulta en `GET /estadoRevisionSolicitud` dentro de `comentarios_por_formulario`.

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

### 12.1 Evitar agotamiento de cuota en Google Sheets (HTTP 429)

La API de Google Sheets aplica cuotas por minuto (aprox. 300 solicitudes/min por proyecto y 60 solicitudes/min por usuario/proyecto). Si se exceden, responde `429 Too many requests`.

Recomendaciones aplicadas y sugeridas en este backend:

- Reducir llamadas por operación:
	- usar `batchGet` y `batchUpdate` cuando sea posible, en vez de múltiples `values.get`/`values.update`.
	- en este proyecto ya se optimizó lectura de datos de solicitud y lectura de `GASTOS` + `CONCEPTO$` con `batchGet`.
- Implementar reintentos con backoff exponencial con jitter para `429` y errores transitorios (`503`, `500`).
- Evitar polling agresivo desde frontend:
	- no consultar progreso cada pocos segundos.
	- sugerido: intervalo de 10-20s o modelo basado en eventos/acciones del usuario.
- Aplicar cache corta en backend para lecturas repetidas (ej. 10-30 segundos por `id_solicitud`), especialmente en endpoints de consulta.
- Agrupar escrituras de un mismo flujo en una sola petición cuando sea viable.
- Mantener payloads livianos (Google recomienda alrededor de 2 MB por solicitud).

Si ocurre `429`, no reintentar en bucle inmediato: usar backoff y volver a intentar tras unos segundos.

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
## 14. Nuevo flujo de aprobaciÃ³n

El flujo se maneja por roles y colas. Una solicitud no se asigna inicialmente a un revisor especÃ­fico; queda disponible para el rol responsable.

Roles en `USUARIOS!D:D`:

```text
solicitante, revisor_1, revisor_2, presupuesto, director, decanato, admin
```

Un usuario cuyo campo `role` esté vacío o no tenga un rol especial se considera automáticamente `solicitante`. No es necesario guardar literalmente `solicitante` en `USUARIOS`; el backend lo aplica como rol predeterminado.

Etapas:

```text
solicitud → revision_1 → revision_2 → presupuesto → firmas → decanato → finalizado
```

Todas las rutas nuevas requieren:

```http
Authorization: Bearer <google_id_token>
Content-Type: application/json
```

Las rutas de firma y resolución usan `multipart/form-data`.

### 14.1 Consultar la cola del rol

```http
GET /workflow/cola
```

Para `revisor_1` y `revisor_2`, esta misma ruta también incluye las solicitudes
que ya llegaron a `decanato` y todavía están en `pendiente`, `en_revision` o
`devuelta`. Aunque esas filas conservan `rol_responsable: decanato`, aparecen en
la bandeja del revisor porque ambos roles están autorizados para registrar el
cierre final desde Decanato.

Respuesta:

```json
{
  "success": true,
  "role": "revisor_1",
  "data": [
    {
      "id_solicitud": "123",
      "nombre_actividad": "Curso de primeros auxilios",
      "etapa_codigo": "revision_1",
      "rol_responsable": "revisor_1",
      "usuario_asignado": "",
      "estado_etapa": "pendiente",
      "estado_general": "Enviado a revisiÃ³n"
    }
  ]
}
```

### 14.2 Tomar una solicitud

```http
POST /workflow/:idSolicitud/tomar
```

No requiere body. Esta operación es opcional. Si se utiliza, cambia
`usuario_asignado` al usuario autenticado y `estado_etapa` a `en_revision`.
También registra la acción en `HISTORIAL_ETAPAS`.

Si otro usuario ya la tomó, devuelve `409`.

Un `revisor_1` o `revisor_2` también puede tomar una solicitud cuya etapa sea
`decanato`. Después de tomarla, puede ejecutar
`POST /workflow/cortes-decanato`. La validación de usuario asignado se mantiene:
si otro usuario ya la tomó, el cierre responde `403`.

#### Flujo recomendado cuando hay un único usuario por rol

Cuando existe un único usuario para `revisor_1` y un único usuario para
`revisor_2`, el frontend no necesita llamar a `/tomar`. Puede consultar la cola y
ejecutar directamente la decisión:

```text
GET  /workflow/cola
POST /workflow/:idSolicitud/decision
```

Para una solicitud en Decanato que el revisor pueda cerrar:

```text
GET  /workflow/cola
POST /workflow/:idSolicitud/tomar       (opcional)
POST /workflow/cortes-decanato
```

En este segundo caso, `/tomar` también puede omitirse si
`usuario_asignado` está vacío:

```text
GET  /workflow/cola
POST /workflow/cortes-decanato
```

#### Reglas de asignación

- Si `usuario_asignado` está vacío, el usuario cuyo rol corresponde a la etapa
  puede aprobar o devolver directamente.
- Si `usuario_asignado` contiene el ID del usuario autenticado, ese usuario puede
  continuar trabajando la solicitud.
- Si `usuario_asignado` contiene el ID de otro usuario, las decisiones responden
  `403` y la toma responde `409`.
- Una decisión exitosa limpia `usuario_asignado` al avanzar o devolver la etapa.
- Cuando una solicitud se devuelve al solicitante, el backend conserva
  `rol_responsable: solicitante` para que siga apareciendo en sus solicitudes y
  pueda corregirla o reenviarla.
- El frontend puede usar `/tomar` si necesita mostrar quién está trabajando una
  solicitud o si posteriormente se incorporan más usuarios por rol.

### 14.3 Aprobar, devolver o cerrar una etapa

```http
POST /workflow/:idSolicitud/decision
```

Body para aprobar:

```json
{
  "accion": "aprobar",
  "comentario": "",
  "comentarios_por_formulario": {}
}
```

Body para devolver:

```json
{
  "accion": "devolver",
  "comentario": "Falta completar la informaciÃ³n financiera.",
  "comentarios_por_formulario": {
    "2": "Revisar los valores del presupuesto"
  }
}
```

El comentario es obligatorio al devolver. El cierre no se realiza desde esta ruta:
cuando la solicitud llegue a `decanato`, el revisor 1 o 2 puede registrar la
resolución final usando `POST /workflow/cortes-decanato`.
Respuesta:

```json
{
  "success": true,
  "data": {
    "id_solicitud": "123",
    "etapa_codigo": "revision_2",
    "estado_etapa": "pendiente"
  }
}
```

La operaciÃ³n se registra en `REVISIONES` y `HISTORIAL_ETAPAS`.

Las observaciones independientes se guardan como JSON en `ETAPAS.K`, la columna
`comentarios`, porque pertenecen a una solicitud y a la etapa en la que fueron
creadas. Se pueden recibir observaciones en cualquier etapa activa del flujo:
`solicitud`, `revision_1`, `revision_2`, `presupuesto`, `firmas` y `decanato`.
Cada observación se agrega como un elemento independiente dentro del arreglo de su
etapa; no se reemplazan las observaciones anteriores. `REVISIONES.observaciones`
conserva el comentario general de la decisión para compatibilidad.

Ejemplo del valor de `ETAPAS.K`:

```json
{
  "revision_1": [
    {
      "id_observacion": "obs-1729634400000-ab12cd",
      "formulario": "1",
      "paso": "2",
      "observacion": "Completar los datos del coordinador.",
      "estado": "pendiente",
      "id_usuario_autor": "user-1",
      "nombre_autor": "Revisor 1",
      "fecha_creacion": "2026-10-22T22:00:00.000Z",
      "fecha_actualizacion": "2026-10-22T22:00:00.000Z"
    }
  ],
  "revision_2": []
}
```

### 14.3.1 Agregar una observación independiente

```http
POST /workflow/:idSolicitud/observaciones
Authorization: Bearer <google-id-token>
Content-Type: application/json
```

Body:

```json
{
  "etapa": "revision_2",
  "formulario": 2,
  "paso": 3,
  "observacion": "Revisar el soporte del fondo común."
}
```

`etapa` es opcional; si se omite, se usa la etapa actual de `ETAPAS`. Si se envía,
debe coincidir con la etapa actual; esto evita guardar una observación en una etapa
futura o en una etapa diferente de la que está siendo atendida. No se admiten
observaciones nuevas cuando la solicitud ya está `finalizado`.
`formulario` y `paso` también son opcionales si la observación aplica a toda la
etapa.

Respuesta `201`:

```json
{
  "success": true,
  "data": {
    "id_observacion": "15",
    "id_solicitud": "123",
    "etapa": "revision_2",
    "formulario": "2",
    "paso": "3",
    "observacion": "Revisar el soporte del fondo común.",
    "estado": "pendiente",
    "id_usuario_autor": "user-1",
    "nombre_autor": "Revisor",
    "fecha_creacion": "2026-10-22T22:00:00.000Z",
    "fecha_actualizacion": "2026-10-22T22:00:00.000Z"
  }
}
```

Para agregar una segunda observación, el frontend repite la misma petición con otro
texto. Cada observación tendrá su propio `id_observacion` y se agregará al JSON de
`ETAPAS.K`.

### 14.3.2 Consultar observaciones

```http
GET /workflow/:idSolicitud/observaciones
Authorization: Bearer <google-id-token>
```

Respuesta:

```json
{
  "success": true,
  "data": [
    {
      "id_observacion": "15",
      "id_solicitud": "123",
      "etapa": "revision_2",
      "formulario": "2",
      "paso": "3",
      "observacion": "Revisar el soporte del fondo común.",
      "estado": "pendiente",
      "id_usuario_autor": "user-1",
      "nombre_autor": "Revisor",
      "fecha_creacion": "2026-10-22T22:00:00.000Z",
      "fecha_actualizacion": "2026-10-22T22:00:00.000Z"
    }
  ]
}
```

### 14.3.3 Actualizar el texto o estado de una observación

```http
PATCH /workflow/:idSolicitud/observaciones/:idObservacion
Authorization: Bearer <google-id-token>
Content-Type: application/json
```

Body para marcarla como subsanada:

```json
{
  "estado": "subsanada"
}
```

Estados permitidos:

```text
pendiente
subsanada
anulada
```

También se puede actualizar el texto:

```json
{
  "observacion": "Revisar nuevamente el soporte firmado."
}
```

Respuesta:

```json
{
  "success": true,
  "data": {
    "id_observacion": "15",
    "id_solicitud": "123",
    "etapa": "revision_2",
    "formulario": "2",
    "paso": "3",
    "observacion": "Revisar nuevamente el soporte firmado.",
    "estado": "subsanada",
    "id_usuario_autor": "user-1",
    "nombre_autor": "Revisor",
    "fecha_creacion": "2026-10-22T22:00:00.000Z",
    "fecha_actualizacion": "2026-10-22T22:05:00.000Z"
  }
}
```

El autor puede modificar sus propias observaciones y un administrador puede modificar
cualquiera. Se conserva el objeto dentro del JSON y se cambia su estado; no se
elimina físicamente.

### 14.3.4 Agregar observaciones al aprobar o devolver

También se pueden crear varias observaciones independientes dentro de la misma
decisión usando el arreglo `observaciones`:

```http
POST /workflow/:idSolicitud/decision
```

```json
{
  "accion": "devolver",
  "comentario": "La solicitud requiere ajustes.",
  "observaciones": [
    {
      "formulario": 1,
      "paso": 2,
      "observacion": "Completar los datos del coordinador."
    },
    {
      "formulario": 2,
      "paso": 3,
      "observacion": "Revisar el valor del fondo común."
    }
  ]
}
```

La respuesta incluye los objetos agregados al JSON:

```json
{
  "success": true,
  "data": {
    "id_solicitud": "123",
    "etapa_codigo": "solicitud",
    "estado_etapa": "devuelta",
    "observaciones": [
      {
        "id_observacion": "15",
        "observacion": "Completar los datos del coordinador.",
        "estado": "pendiente"
      },
      {
        "id_observacion": "16",
        "observacion": "Revisar el valor del fondo común.",
        "estado": "pendiente"
      }
    ]
  }
}
```

### 14.3.4.1 Caso principal: cualquier etapa devuelve la solicitud

Cuando una solicitud está en una etapa que permite devolución, el responsable de esa
etapa puede registrar varias observaciones independientes y devolver la solicitud.
Por ejemplo, si está en `revision_1`, quedan guardadas dentro del arreglo
`ETAPAS.K.comentarios.revision_1`:

```json
{
  "revision_1": [
    {
      "id_observacion": "obs-1",
      "formulario": "1",
      "paso": "2",
      "observacion": "Completar los datos del coordinador.",
      "estado": "pendiente"
    },
    {
      "id_observacion": "obs-2",
      "formulario": "2",
      "paso": "3",
      "observacion": "Revisar el valor del fondo común.",
      "estado": "pendiente"
    }
  ]
}
```

La devolución se realiza con una sola llamada:

```http
POST /workflow/123/decision
Authorization: Bearer <google-id-token>
Content-Type: application/json
```

```json
{
  "accion": "devolver",
  "comentario": "Por favor corregir las observaciones indicadas.",
  "observaciones": [
    {
      "formulario": 1,
      "paso": 2,
      "observacion": "Completar los datos del coordinador."
    },
    {
      "formulario": 2,
      "paso": 3,
      "observacion": "Revisar el valor del fondo común."
    }
  ]
}
```

El backend aplica estas transiciones:

1. Desde `revision_1`, agrega las observaciones en
   `ETAPAS.K.comentarios.revision_1` y vuelve a `solicitud` con estado `devuelta`.
2. Desde `revision_2`, agrega las observaciones en
   `ETAPAS.K.comentarios.revision_2` y vuelve a `solicitud` con estado `devuelta`.
3. Desde `presupuesto`, agrega las observaciones en
   `ETAPAS.K.comentarios.presupuesto` y vuelve a `revision_2` con estado `devuelta`.
   En `revision_2`, el revisor 2 puede subsanar esas observaciones o decidir si
   devuelve nuevamente la solicitud al solicitante.
4. Deja disponibles las observaciones para el solicitante.
5. Registra la decisión en `REVISIONES` y `HISTORIAL_ETAPAS`.

En `firmas` no existe una devolución: esa etapa únicamente espera las dos firmas.
Cuando ambas están registradas, la solicitud avanza a `decanato`.

### 14.3.4.2 Reglas de cierre por revisor en Decanato

- `revisor_1` y `revisor_2` pueden cerrar una solicitud únicamente cuando
  `ETAPAS.etapa_codigo` es `decanato`.
- El cierre exige `resultado` (`Aprobada` o `Rechazada`) y la resolución.
- El cierre se registra en `CORTES_DECANATO` y cambia `ETAPAS` a `finalizado`.
- Si la solicitud aún está en `revision_1` o `revision_2`, el revisor solo puede
  aprobar o devolver; no puede cerrarla allí.
- Presupuesto no cierra la solicitud; avanza a firmas.
- En firmas no hay devolución ni cierre por revisor.

El solicitante consulta la solicitud y sus observaciones con:

```http
GET /workflow/123/observaciones
Authorization: Bearer <google-id-token>
```

Cuando corrige cada punto, marca únicamente el estado de esa observación:

```http
PATCH /workflow/123/observaciones/obs-1
Authorization: Bearer <google-id-token>
Content-Type: application/json
```

```json
{
  "estado": "subsanada"
}
```

El solicitante puede cambiar `pendiente` a `subsanada`, pero no puede modificar el
texto escrito por el revisor. El backend conserva el texto original y el identificador
de cada observación.

Para reenviar una solicitud devuelta, el frontend debe:

1. Consultar las observaciones.
2. Mostrar un checkbox por cada observación pendiente.
3. Marcar cada observación corregida con `PATCH /workflow/:idSolicitud/observaciones/:idObservacion`
   y el body `{ "estado": "subsanada" }`.
4. Habilitar el botón de reenvío solo cuando no existan observaciones pendientes.
5. Enviar la solicitud mediante `POST /enviarSolicitudRevision` o
   `POST /form/enviarSolicitudRevision`.

Cuando la última observación pendiente se marca como `subsanada`, el backend cambia
automáticamente los formularios que estaban en `Enviado a revisión` a
`Completado`. Así el solicitante puede editar nuevamente esos formularios antes de
reenviar la solicitud. Los formularios que ya estaban `Aprobado` no se modifican.

El reenvío conserva el punto de retorno de la devolución usando
`ETAPAS.etapa_anterior`:

- `revision_1` → `solicitud` → reenvío a `revision_1`.
- `revision_2` → `solicitud` → reenvío a `revision_2`.

Si la solicitud es nueva y nunca fue devuelta, el reenvío inicial continúa siendo:

```text
solicitud → revision_1
```

La respuesta de `POST /enviarSolicitudRevision` incluye `etapa_codigo` y
`rol_responsable` para que el frontend actualice la bandeja correcta.

El backend también valida esta regla. Si la solicitud está en
`etapa_codigo: solicitud` con `estado_etapa: devuelta` y todavía existe una
observación `pendiente`, el reenvío responde `409` con el código
`PENDING_OBSERVATIONS`. No es suficiente con ocultar o habilitar el botón en el
frontend.

### 14.3.5 Persistencia

No se requiere una hoja `OBSERVACIONES`. La fuente de verdad de estas observaciones
es `ETAPAS.K` (`comentarios`) como JSON agrupado por `etapa_codigo`.

### 14.4 Registrar firmas

```http
POST /workflow/:idSolicitud/firma
```

Content-Type: `multipart/form-data`.

El rol determina el tipo de firma: `solicitante` o `director_escuela`.

Body:

```json
{
  "evidencia_firma": "(archivo PNG, JPEG o WEBP)",
  "observaciones": "",
  "fecha_solicitud_firma": "2026-10-20T15:30:00.000Z"
}
```

Cuando existen ambas firmas, `ETAPAS` pasa a `decanato`, con `rol_responsable: decanato` y `estado_etapa: pendiente`. La imagen se carga en la carpeta de Drive configurada y la URL se guarda en `FIRMAS.evidencia_firma`.

### 14.5 Crear o cerrar un corte de Decanato

```http
POST /workflow/cortes-decanato
```

La decisión se registra mediante esta ruta con `multipart/form-data`. Puede usarla
Decanato, el revisor 1 o el revisor 2, pero únicamente si la solicitud ya está en
la etapa `decanato`. El cierre aprobado o rechazado siempre se persiste en
`CORTES_DECANATO`.
Campos:

```json
{
  "id_solicitud": "123",
  "resultado": "Aprobada",
  "observaciones": "Aprobada por el comitÃ©.",
  "resolucion": "ResoluciÃ³n 230 de 2026",
  "pdf_resolucion": "(archivo PDF)"
}
```

Respuesta:

```json
{
  "success": true,
  "data": {
    "id_corte": "8",
    "id_solicitud": "123",
    "estado_corte": "finalizado",
    "resultado": "Aprobada",
    "resolucion": "https://storage.googleapis.com/...",
    "estado_final": "Aprobada",
    "etapa": {
      "etapa_codigo": "finalizado",
      "estado_etapa": "finalizado"
    }
  }
}
```

La resolución PDF cargada por Decanato se guarda en `CORTES_DECANATO.resolucion` y el resultado normalizado en `CORTES_DECANATO.estado_final`.

### 14.6 Hojas nuevas

```text
REVISIONES:
id_revision, id_solicitud, etapa, rol_revisor, id_usuario_revisor, estado, observaciones, fecha_inicio, fecha_respuesta, comentarios_por_formulario, numero_intento, id_usuario_toma, fecha_toma, resultado, resolucion

FIRMAS:
id_firma, id_solicitud, tipo_firma, id_firmante, nombre_firmante, estado, fecha_firma, evidencia_firma, observaciones, fecha_solicitud_firma

HISTORIAL_ETAPAS:
id_historial, id_solicitud, etapa, accion, estado_anterior, estado_nuevo, rol_actor, id_usuario_actor, fecha, comentario

CORTES_DECANATO:
id_corte, fecha_corte, id_solicitud, estado_corte, id_usuario_envia, fecha_envio, resultado, observaciones, resolucion, estado_final

```

### 14.7 Reglas para el frontend

- `revisor_1` y `revisor_2` consultan colas distintas.
- El botÃ³n de aprobar depende de `etapa_codigo`.
- Devolver requiere al menos una observación específica; el comentario general es opcional.
- Presupuesto solo actÃºa en `presupuesto`.
- En firmas deben aparecer dos acciones independientes: solicitante y director.
- Decanato solo recibe solicitudes con las dos firmas.
- Decanato registra `estado_final`, carga su resolución y mueve la solicitud a
  `finalizado`. Revisor 1 y revisor 2 también pueden ejecutar este cierre, pero
  únicamente desde la etapa `decanato`.
- DespuÃ©s de cada acciÃ³n, el frontend debe volver a consultar la solicitud o la cola.
- `HISTORIAL_ETAPAS` puede mostrarse como lÃ­nea de tiempo.

## 17. ValidaciÃ³n de funcionalidades solicitadas por frontend

Esta tabla refleja el estado actual del backend.

| Funcionalidad | Estado | ImplementaciÃ³n |
|---|---|---|
| Solicitudes para revisor 1 | Implementado | `GET /workflow/cola` con rol `revisor_1`; incluye cierres pendientes en `decanato` |
| Solicitudes para revisor 2 | Implementado | `GET /workflow/cola` con rol `revisor_2`; incluye cierres pendientes en `decanato` |
| Solicitudes para presupuesto | Implementado | `GET /workflow/cola` con rol `presupuesto` |
| Solicitudes para decanato | Implementado | `GET /workflow/cola` con rol `decanato` |
| Solicitudes para director | Implementado | `GET /workflow/firmas/pendientes` |
| Estados globales | Implementado | `ETAPAS.estado_general` |
| Estados por formulario | Implementado | `ETAPAS.estado_formularios` como JSON |
| Etapa y estado del flujo | Implementado | `ETAPAS.etapa_codigo` y `ETAPAS.estado_etapa` |
| Historial de etapas | Implementado parcialmente | Se registra desde las nuevas rutas `/workflow`; las operaciones administrativas antiguas todavía deben migrarse completamente al historial |
| AuditorÃ­a del actor | Implementado en nuevas rutas | `HISTORIAL_ETAPAS.rol_actor` e `HISTORIAL_ETAPAS.id_usuario_actor` |
| Observaciones por formulario | Implementado | `comentarios_por_formulario` y `REVISIONES.comentarios_por_formulario` |
| Observaciones por paso | Implementado | `POST /guardarComentarioPaso` |
| Firma del solicitante | Implementado | `POST /workflow/:idSolicitud/firma` |
| Firma del director | Implementado | `POST /workflow/:idSolicitud/firma` autenticado con rol `director` |
| Envío al corte de Decanato | Implementado | `POST /workflow/cortes-decanato` |
| Resultado final de Decanato | Implementado | `CORTES_DECANATO.resultado` |
| Cierre por revisor 1/2 en Decanato | Implementado | `POST /workflow/cortes-decanato` con rol `revisor_1` o `revisor_2` |
| Nombre de resolución de Decanato | Implementado | `CORTES_DECANATO.resolucion` |
| PDF de resolución | Implementado | Decanato carga `pdf_resolucion` y se guarda la URL en `CORTES_DECANATO.archivo_resolucion`; el nombre queda en `CORTES_DECANATO.resolucion` |
| Notificaciones por correo | Implementado | Firmas, recordatorios y envío automático a Decanato; se registra en `NOTIFICACIONES` |

### 17.1 Obtener solicitudes por rol

El frontend no debe enviar el rol en el body. El backend lo obtiene del usuario autenticado consultando `USUARIOS`.

```http
GET /workflow/cola
Authorization: Bearer <google_id_token>
```

Roles soportados actualmente por esta ruta:

```text
revisor_1
revisor_2
presupuesto
decanato
```

La respuesta contiene las solicitudes cuyo `rol_responsable` coincide con el rol
autenticado y cuyo `estado_etapa` es `pendiente`, `en_revision` o `devuelta`.
Excepción: para `revisor_1` y `revisor_2` también incluye solicitudes en etapa
`decanato`, aunque `rol_responsable` sea `decanato`, porque esos revisores pueden
cerrarlas desde esa etapa.

El director debe consultar `GET /workflow/firmas/pendientes` para su bandeja de firmas.

### 17.2 Estados que debe interpretar el frontend

Estado global heredado, en `ETAPAS.estado_general`:

```text
En proceso
Enviado a revisiÃ³n
Aprobado parcialmente
Aprobado
Terminado
```

Estado individual por formulario, dentro de `ETAPAS.estado_formularios`:

```json
{
  "1": "Enviado a revisiÃ³n",
  "2": "Aprobado",
  "3": "Requiere correcciones",
  "4": "En progreso"
}
```

Estados de la etapa nueva, en `ETAPAS.estado_etapa`:

```text
borrador
pendiente
en_revision
devuelta
rechazada
finalizado
```

### 17.3 Historial y auditorÃ­a

Las nuevas operaciones de workflow escriben en `HISTORIAL_ETAPAS`:

```json
{
  "id_solicitud": "123",
  "etapa": "revision_1",
  "accion": "aprobar",
  "estado_anterior": "en_revision",
  "estado_nuevo": "pendiente",
  "rol_actor": "revisor_1",
  "id_usuario_actor": "google-user-id",
  "fecha": "2026-10-20T15:30:00.000Z",
  "comentario": "RevisiÃ³n aprobada"
}
```

El frontend puede utilizar estos registros para construir una lÃ­nea de tiempo de la solicitud. Las operaciones antiguas bajo `/admin/*` conservan compatibilidad, pero su migraciÃ³n completa al historial debe considerarse pendiente.

### 17.4 Observaciones por formulario y paso

Para compatibilidad con el flujo antiguo, las observaciones resumidas por formulario
se siguen exponiendo en:

```json
{
  "comentarios_por_formulario": {
    "1": "Completar los datos del coordinador",
    "2": "Revisar el presupuesto"
  }
}
```

Para observaciones puntuales de un paso:

```http
POST /guardarComentarioPaso
```

```json
{
  "id_solicitud": "123",
  "userId": "google-user-id",
  "formulario": 2,
  "paso": 3,
  "comentario": "Validar el soporte del fondo comÃºn"
}
```

Un comentario vacÃ­o elimina el comentario guardado para ese formulario y paso.

### 17.5 Firmas

```http
POST /workflow/:idSolicitud/firma
```

El backend carga la imagen en la carpeta de Google Drive configurada en
`DRIVE_UPLOADS_FOLDER`, registra la URL en `FIRMAS.evidencia_firma` y, cuando existen
las dos firmas, mueve la solicitud a Decanato.

### 17.6 Corte y decisiÃ³n de Decanato

```http
POST /workflow/cortes-decanato
```

El body debe incluir:

```json
{
  "id_solicitud": "123",
  "estado_corte": "finalizado",
  "resultado": "Aprobada",
  "observaciones": "Aprobada por el comitÃ©",
  "resolucion": "ResoluciÃ³n 230 de 2026"
}
```

Decanato puede enviar también `pdf_resolucion`. Cuando el resultado es `Aprobada`,
el PDF es obligatorio; se carga en la carpeta de Drive configurada y su URL se guarda en
`CORTES_DECANATO.resolucion`.

### 17.7 Notificaciones

Las notificaciones SMTP incluyen firma registrada, recordatorios de firma y el aviso
automático a Decanato. El frontend debe consultar el estado devuelto y
`NOTIFICACIONES`; no debe asumir que un correo fue entregado si el estado es
`fallida` o `sin_destinatario`.

Los únicos eventos que activan correo son:

```text
firma_registrada
recordatorio_firma_solicitante
recordatorio_firma_director
```

## 18. Contrato vigente del workflow (octubre de 2026)

Esta sección reemplaza las notas anteriores que indicaban que las firmas, el PDF de
resolución y las notificaciones estaban pendientes. El prototipo
`Flujo de aprobación Extensión vertical.html` no forma parte del frontend productivo;
el frontend está en otro repositorio y debe consumir los contratos siguientes.

### 18.1 Máquina de estados única

El backend usa `ETAPAS.etapa_codigo` y `ETAPAS.estado_etapa` como fuente de verdad
del flujo. Los estados heredados (`estado_general` y `estado_formularios`) se conservan
para compatibilidad con formularios y reportes, pero cualquier transición del workflow
también actualiza la etapa nueva, su responsable, `etapa_anterior`, `updated_at` y
`version`.

```text
solicitud → revision_1 → revision_2 → presupuesto → firmas → decanato → finalizado
```

Los roles de las colas son `revisor_1`, `revisor_2`, `presupuesto` y `decanato`.
El solicitante y el director no tienen una cola de revisión: reciben acciones de firma
cuando la solicitud está en `firmas`.

### 18.2 Bandeja de firmas del director

```http
GET /workflow/firmas/pendientes
Authorization: Bearer <google-id-token>
```

Solo está disponible para usuarios con rol `director`. Devuelve solicitudes en etapa
`firmas` que todavía no tienen una firma `director_escuela`.

### 18.3 Registrar una firma

```http
POST /workflow/:idSolicitud/firma
Authorization: Bearer <google-id-token>
Content-Type: multipart/form-data
```

Campos:

- `evidencia_firma`: imagen PNG, JPEG o WEBP. Es obligatoria.
- `observaciones`: opcional.
- `fecha_solicitud_firma`: opcional.

El tipo de firma no lo decide el frontend:

- rol `solicitante` → `tipo_firma=solicitante`
- rol `director` → `tipo_firma=director_escuela`

El backend valida que la solicitud esté en `firmas`, que el solicitante sea el dueño
de la solicitud, que no exista una firma previa del mismo tipo y que el archivo sea
una imagen válida. La imagen se carga en Google Drive usando la autenticación configurada para las APIs
de Google y se guarda la URL resultante en `FIRMAS.evidencia_firma`. El frontend
mantiene el mismo contrato `multipart/form-data` y el campo `evidencia_firma`.

Después de registrar la firma, `FIRMAS.estado` queda en `firmada`. Cuando existen
ambas firmas (`solicitante` y `director_escuela`), `ETAPAS` pasa automáticamente a:

```text
etapa_codigo: decanato
rol_responsable: decanato
estado_etapa: pendiente
```

Mientras falte una firma, la solicitud permanece en `firmas` con
`estado_etapa: en_revision`.

### 18.4 Recordatorios de firma

```http
POST /workflow/:idSolicitud/firma/recordatorio
Authorization: Bearer <google-id-token>
Content-Type: application/json
```

Body:

```json
{
  "tipo_firma": "solicitante"
}
```

Valores permitidos: `solicitante` y `director_escuela`.

- El solicitante puede recordar al director.
- El director puede recordar al solicitante.
- Un administrador puede recordar a cualquiera.

### 18.5 Correos SMTP

El backend envía correo en estos cuatro eventos:

1. `firma_registrada`
2. `recordatorio_firma_solicitante`
3. `recordatorio_firma_director`
4. `solicitud_enviada_decanato`

Cuando se registra la segunda firma, la solicitud pasa a `decanato` y se envía
un correo a todos los usuarios de `USUARIOS` cuyo rol sea `decanato`. El frontend
puede agregar destinatarios sin reemplazar los destinatarios institucionales
enviando `destinatarios_decanato` como JSON o como direcciones separadas por coma:

```text
destinatarios_decanato: ["persona@ejemplo.edu.co", "otra@ejemplo.edu.co"]
```

También puede personalizar el correo mediante `asunto_decanato` y
`cuerpo_decanato`. Si no los envía, se usan los valores predeterminados del
backend. El cuerpo admite `{{id_solicitud}}`, `{{nombre_solicitante}}`,
`{{etapa}}` y `{{fecha}}`.

Variables:

```env
SMTP_USER=...
SMTP_PASSWORD=...
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_SECURE=false
SMTP_FROM=...
```

`SMTP_USER` y `SMTP_PASSWORD` son las únicas credenciales secretas requeridas para
SMTP. `SMTP_HOST`, `SMTP_PORT`, `SMTP_SECURE` y `SMTP_FROM` son configuración; si no
se define `SMTP_HOST`, se usa `smtp.gmail.com`.

Cada intento se registra en la hoja `NOTIFICACIONES`:

```text
id_notificacion
id_solicitud
evento
estado
intentos
fecha_creacion
fecha_envio
ultimo_error
```

Un fallo de correo no deshace una firma o una transición ya guardada. La respuesta
de la ruta de firma incluye `notificacion` y `notificaciones_decanato` con el
estado de cada envío para que el frontend pueda mostrarlo explícitamente.

### 18.5.1 Consulta de solicitudes y permisos de revisores

La cola operativa conserva los permisos por etapa:

- `revisor_1` puede actuar en `revision_1`.
- `revisor_2` puede actuar en `revision_2`.
- `presupuesto` puede actuar en `presupuesto`.
- `director` puede firmar en `firmas`.

Los revisores no pueden aprobar, devolver, agregar observaciones o cerrar una
solicitud fuera de su etapa aunque puedan consultarla. La consulta de solo lectura
se obtiene con:

```http
GET /workflow/solicitudes/consulta
```

`revisor_1` y `revisor_2` reciben todas las solicitudes. `presupuesto` recibe las
solicitudes en las que ya actuó según `HISTORIAL_ETAPAS` y las que avanzaron desde
presupuesto. `director` recibe las solicitudes que están pendientes de su firma y
las que ya tienen su firma registrada, aunque luego hayan avanzado a Decanato o
finalizado.

Cada resultado incluye `puede_actuar`, que solo es informativo; los endpoints de
escritura vuelven a validar rol y etapa en el backend.

### 18.6 Decanato registra la decisión y carga la resolución

```http
POST /workflow/cortes-decanato
Authorization: Bearer <google-id-token>
Content-Type: multipart/form-data
```

Solo el rol `decanato` puede usar esta ruta. La solicitud debe estar en etapa
`decanato`.

Campos:

- `id_solicitud`: obligatorio.
- `resultado`: obligatorio, por ejemplo `Aprobada` o `Rechazada`.
- `observaciones`: opcional.
- `resolucion`: opcional si no se carga archivo.
- `pdf_resolucion`: obligatorio cuando el resultado es `Aprobada`; debe ser PDF.

El PDF se carga en la carpeta de Google Drive configurada en `DRIVE_UPLOADS_FOLDER`.
La URL se guarda en `CORTES_DECANATO.resolucion`; no se guardan metadatos adicionales
del archivo. La decisión se persiste con `CORTES_DECANATO.estado_final` y luego
`ETAPAS` pasa a `finalizado`.

### 18.7 Almacenamiento de adjuntos, firmas y resoluciones en Drive

Las firmas y resoluciones no se almacenan en Cloud Storage. Se suben mediante
`services/driveService.js` a la carpeta indicada por:

```env
DRIVE_UPLOADS_FOLDER=...
GOOGLE_CLIENT_EMAIL=...
GOOGLE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
```

La carpeta debe ser accesible para la identidad configurada en `config/google.js`.
El backend valida que el ID corresponda a una carpeta, sube el archivo y guarda en
Sheets únicamente la URL de visualización. El nombre generado para las firmas es:

```text
firma-{id_solicitud}-{tipo_firma}-{id_usuario}-{timestamp}.{extension}
```

El nombre generado para las resoluciones es:

```text
resolucion-{id_solicitud}-{timestamp}.pdf
```

`services/storageService.js` no participa en el flujo de firmas ni de Decanato.

### 18.8 Autenticación OAuth de Drive con refresh token

Las cuentas de servicio no tienen cuota en `Mi unidad`. Para guardar archivos en
una carpeta de `Mi unidad`, el backend puede usar OAuth de una cuenta institucional.
La autorización se ejecuta una sola vez desde el backend, sin cambios en el frontend:

1. Registra esta URI en el OAuth Client de Google:

```text
http://localhost:3002/oauth2callback
```

2. Verifica que el `.env` tenga `GOOGLE_CLIENT_ID` y `GOOGLE_CLIENT_SECRET`.
3. Ejecuta:

```bash
npm run authorize:drive
```

4. Abre en el navegador la URL que imprime la consola.
5. Autoriza la cuenta institucional que tiene acceso a `DRIVE_UPLOADS_FOLDER`.
6. Copia el valor que el script imprime al `.env`:

```env
GOOGLE_DRIVE_REFRESH_TOKEN=...
```

7. Reinicia el backend.

Cuando existe `GOOGLE_DRIVE_REFRESH_TOKEN`, `driveService` usa OAuth para subir
firmas, resoluciones y adjuntos. Si no existe, conserva compatibilidad con la
autenticación JWT anterior. El refresh token es un secreto: no debe enviarse al
frontend, registrarse en Sheets ni subirse al repositorio.

Los objetos se organizan así:

```text
firmas/{id_solicitud}/solicitante-*.png
firmas/{id_solicitud}/director_escuela-*.png
resoluciones/{id_solicitud}/resolucion-*.pdf
```

El bucket debe permitir a la Service Account subir archivos. Si los objetos se
mantienen privados, el frontend no debe asumir que la URL de almacenamiento es
pública; debe usarse el mecanismo de acceso configurado para el bucket.

### 18.8 Columnas nuevas de `ETAPAS`

Además de las columnas existentes, el backend utiliza:

```text
etapa_anterior
updated_at
version
```

`version` se incrementa al actualizar la etapa. `updated_at` se guarda en formato
ISO 8601 y `etapa_anterior` conserva la etapa inmediatamente anterior cuando existe
una transición.

### 18.9 Integración del frontend

El frontend productivo debe:

- Enviar el token de Google en todas las rutas `/workflow`.
- Usar `multipart/form-data` para firmas y resolución.
- No enviar `tipo_firma` como autoridad; el backend lo deriva del rol autenticado.
- Consultar nuevamente la solicitud después de firmar o registrar la decisión.
- Mostrar `NOTIFICACIONES.estado` cuando una acción de correo sea fallida.
- Mostrar la cola de firmas del director con `GET /workflow/firmas/pendientes`.
- No simular persistencia local del flujo; el prototipo HTML es únicamente referencia
  visual.
