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

# Carpeta de Drive para adjuntos (ID o URL de carpeta)
DRIVE_UPLOADS_FOLDER_ID=1dQg29-rgkGta-o58iDQ_lrbZ8unLFTje
```

Notas importantes:

- `GOOGLE_PRIVATE_KEY` debe mantener los `\\n` en el `.env`.
- En produccion, `SESSION_SECRET` es obligatorio y el servidor se detiene si no esta definido.
- Para adjuntos, puedes usar `DRIVE_UPLOADS_FOLDER_ID` con un ID o URL de carpeta de Drive.
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
	- columna `K` = `comentarios` (JSON de comentarios por formulario, con compatibilidad a texto plano)
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