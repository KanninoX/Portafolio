# Carga de documentos por código — BUK

Google Apps Script que automatiza la carga de documentos a BUK cuando se
contrata a un trabajador. Recibe los webhooks de BUK, los encola y los procesa
en un activador de tiempo.

## Flujo

1. `doPost` recibe el webhook de BUK (`employee_create` / `job_hire`) y lo encola.
   Siempre responde 200 de inmediato: si BUK no recibe respuesta a tiempo, reintenta.
2. `procesarColaWebhook` corre en un activador cada minuto y vacía la cola.
3. `handleJobHire` reserva al empleado (ventana de 24 h) para que las entregas
   duplicadas no suban los documentos dos veces, y luego carga:
   - los documentos de la carpeta de Drive de **su empresa** → carpeta BUK `Documentos entregados`;
   - los del **Modelo de Prevención de Delitos**, para todos los trabajadores sin
     importar la empresa → carpeta BUK `Modelo de Prevención de Delitos`;
   - el documento **IRL** que corresponda al cargo, solo para Invermar (ID 1) →
     carpeta BUK `SSO`.

## Configuración

Las credenciales **no** están en el código: se leen desde las Propiedades del
script (`PropertiesService`) mediante `_cfg_()`, `_authToken_()` y `_bukBase_()`.

Las carpetas de Drive por empresa están en el arreglo `companies` dentro de
`getEmployeeData()`. Las del Modelo de Prevención de Delitos, en las constantes
`MPD_FOLDER_ID` / `MPD_BUK_PATH` / `MPD_CONFIG` al inicio del archivo.

## Desarrollo

El proyecto se sincroniza con [clasp](https://github.com/google/clasp):

```bash
clasp pull    # traer los cambios hechos en el editor web
clasp push    # subir los cambios locales al proyecto (HEAD)
```

El activador de tiempo ejecuta siempre HEAD. La web app (`doPost`) usa el
despliegue versionado `BUK - Carga de documentos`, así que un cambio en `doPost`
requiere además:

```bash
clasp deploy -i <deployment-id> -d "BUK - Carga de documentos"
```

Usar el mismo `deployment-id` conserva la URL registrada en BUK.

## Hojas de cálculo

El script escribe registros en las hojas `Registro JOB_HIRE`, `Log_SSO` y
`Empleados` de la planilla contenedora, y lee el mapeo cargo → documento IRL
desde la hoja `IRL`.
