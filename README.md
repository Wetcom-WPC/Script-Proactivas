# script-proactivas

Automatización del servicio **Visita Proactiva**: recolecta el estado de la infraestructura VMware de un cliente (vCenter, hosts ESXi, VMs, datastores, redes, certificados, licencias, snapshots, backups, etc.), lo compara contra buenas prácticas y genera dos entregables en Excel: un **Anexo Técnico** con el detalle de cada hallazgo y una **Checklist** con el resultado de cada punto, listos para enviar al cliente.

Este documento explica cómo correrlo, qué hace cada pieza y qué archivos quedan al final. Está pensado para que cualquier persona del equipo, sin haber tocado el proyecto antes, entienda el flujo completo.

---

## 1. Cómo correrlo

1. Doble clic en **`Iniciar-Proactiva.bat`** (raíz del proyecto).
2. El script pregunta todo al principio:
   - **Nombre del cliente**
   - **Mes del reporte** (Enter acepta el mes actual)
   - **Tipo de proactiva**: `1` Mensual, `2` Trimestral, `3` Semestral
   - **vCenters**, separados por coma
   - **Credenciales** de cada vCenter. Funciona como un llavero: si varios vCenters comparten credenciales, se piden una sola vez.
3. A partir de ahí no hace falta tocar nada más. La recolección demora entre 15 minutos y 1 hora según el entorno.
4. Al terminar, los entregables quedan en `resultados\<fecha>\anexo\` (ver sección 4).

Los tipos de proactiva son acumulativos: la **trimestral** incluye todo lo de la mensual, y la **semestral** incluye la trimestral y la mensual.

| Tipo | Qué recolecta | Qué chequeos corre |
|---|---|---|
| Mensual | Datos de Proactiva | `VSP-ME-*` |
| Trimestral | + drivers/firmware de los hosts | + `VSP-TR-*` |
| Semestral | (igual que la trimestral) | + `VSP-SE-*` |

La configuración de cada vDS se exporta a `.zip` en **todas** las corridas (depende del cliente cuándo se usa). En las semestrales hay que subir esos `.zip` al Drive; el script lo recuerda al terminar.

No hace falta instalar nada: el proyecto trae su propio PowerShell portátil y todos los módulos necesarios (VMware PowerCLI, ImportExcel) en `runtime\`. Si en una PC puntual algo no carga, corré **`extra\Preparar-Equipo.bat`**: diagnostica los motivos más comunes (Visual C++ Redistributable faltante, proveedor NuGet, otras instalaciones de PowerCLI) y corrige solo lo que es seguro corregir automáticamente.

### Regenerar el anexo sin volver a recolectar

Si te equivocaste en el cliente, el mes o el tipo de proactiva, corré **`extra\Regenerar-Anexo.bat`**. Vuelve a preguntar esos datos y regenera el Anexo y la Checklist a partir de la **última corrida**, sin conectarse al vCenter.

### Uso desde consola (opcional)

`src\Start-Proactiva.ps1` acepta todo por parámetro; lo que no se pasa, lo pregunta. Correlo siempre con el PowerShell incluido:

```powershell
.\runtime\pwsh.exe -ExecutionPolicy Bypass -File .\src\Start-Proactiva.ps1 -Client "Orbis" -Month "Octubre" -Frequency Quarterly -VCenters vc01.orbis.local,vc02.orbis.local
```

| Parámetro | Para qué |
|---|---|
| `-Client`, `-Month` | Datos del reporte |
| `-Frequency` | `Monthly`, `Quarterly` o `Semiannual` |
| `-VCenters` | Lista de vCenters |
| `-CredentialPath` | Carpeta con credenciales guardadas como `cred-*.xml` (`Get-Credential \| Export-Clixml`). Con esto no se pide nada por consola; pensado para pruebas o corridas desatendidas. Los archivos quedan cifrados con el usuario de Windows: solo ese usuario, en esa PC, puede leerlos. |
| `-ReportOnly` | Regenera el anexo de una corrida existente (`-RunFolder "2026-10-13"`, o la más reciente) |
| `-SkipAlarmTest` | No crea la alarma temporal de la prueba de Falso Positivo (ver sección 2) |

---

## 2. El flujo completo

Todo corre en **un solo proceso y una sola ventana**, en tres etapas:

### 1. Recolección (`src\collection\`)

Se conecta a cada vCenter (SOAP + REST/CIS) y recorre hosts, clusters, VMs, datastores, switches, snapshots, certificados, licencias, alarmas, salud de estadísticas, backups del VAMI, etc. El resultado queda en `resultados\<fecha>\json\<fecha hora>_Proactiva.json`. En trimestrales y semestrales, además, recolecta los drivers/firmware de las placas de red y controladoras de cada host (`..._DriverInfo.json`).

Todo lo que hace contra el vCenter es **de solo lectura**, con una excepción: la **prueba de Falso Positivo** (VSP-ME-13) crea una alarma temporal sobre un datastore que ejecuta el mismo script de alertas que ya tiene configurado el cliente, espera a que se dispare y la borra. Con `-SkipAlarmTest` se omite esa parte (el resto de la auditoría de alarmas se hace igual).

### 2. Conversión a Excel (`src\conversion\`)

Cada JSON se convierte en un Excel "crudo" con una hoja por tema (vCenter, ESXi, VM, Datastores, etc.), sin filtrar, en `resultados\<fecha>\excel\`.

### 3. Anexo y Checklist (`src\report\`)

Corre los chequeos que corresponden al tipo de proactiva. Cada chequeo lee una hoja del Excel crudo, aplica su regla (criterios de la KB [Relevamiento vSphere Automático](https://wetcom.atlassian.net/wiki/spaces/PS/pages/2551579347)) y:

- si encuentra algo, agrega una hoja de detalle al **Anexo Técnico**;
- completa las columnas Estado y Detalle de su ID en la **Checklist** (copia de `src\report\templates\checklist-template.xlsx`).

La Checklist solo conserva las tareas que corresponden a la visita según su columna "Recurrencia", en todos los componentes (vSphere, Veeam, NSX, Horizon): una mensual tiene solo las mensuales, una trimestral las mensuales y trimestrales, y una semestral todas.

Las hojas vacías no se generan. Si regenerás el anexo para el mismo cliente y mes, se arma desde cero (no acumula hojas viejas).

---

## 3. Estructura del repositorio

```
script-proactivas/
├── Iniciar-Proactiva.bat         ← doble clic para correr todo (uso normal)
├── extra/
│   ├── Regenerar-Anexo.bat       ← rehace el anexo de la última corrida sin reconectarse
│   └── Preparar-Equipo.bat       ← diagnostica requisitos de PowerCLI en esta PC
├── resultados/                   ← una carpeta por corrida (no se versiona)
├── runtime/                      ← PowerShell portátil + módulos. No es código propio; no se toca.
└── src/
    ├── Start-Proactiva.ps1       ← punto de entrada: pregunta los datos y corre las 3 etapas
    ├── collection/               ← 1. Recolección
    │   ├── Collection.psm1           conexión a vCenters (llavero) y armado de los JSON
    │   ├── ProactivaCollector.psm1   clase con toda la lógica de recolección
    │   ├── HardwareCompatibility.ps1 búsquedas en la HCL (Broadcom Compatibility Guide)
    │   ├── Update-ReferenceData.ps1  utilidad manual: actualiza reference-data\ desde internet
    │   └── reference-data/           HCL de servidores (CSV), HCL de I/O, releases de ESXi, sizing, Tools
    ├── conversion/               ← 2. JSON → Excel crudo
    │   └── Conversion.psm1
    ├── report/                   ← 3. Anexo Técnico y Checklist
    │   ├── Report.psm1               qué chequeos corren según el tipo de proactiva + checklist
    │   ├── checks/
    │   │   ├── MonthlyChecks.ps1     VSP-ME-*
    │   │   ├── QuarterlyChecks.ps1   VSP-TR-*
    │   │   └── SemiannualChecks.ps1  VSP-SE-*
    │   └── templates/checklist-template.xlsx
    ├── shared/                   ← carpeta de corrida y estilo de Excel, usados por varias etapas
    └── tools/
        └── Initialize-Workstation.ps1  (lo lanza extra\Preparar-Equipo.bat)
```

El código (nombres de archivos, funciones, variables y comentarios) está en inglés. Lo que ve el técnico (mensajes de consola) y los entregables (nombres de hojas, textos de la checklist) están en español.

---

## 4. Qué queda al finalizar una corrida

```
resultados/2026-10-13/
├── anexo/
│   ├── Anexo Tecnico - {Cliente} - {Mes}.xlsx        ← ENTREGABLE: detalle de hallazgos
│   └── Checklist Proactiva - {Cliente} - {Mes}.xlsx  ← ENTREGABLE: checklist para el cliente
├── excel/
│   ├── 2026-10-13 163854_Proactiva.xlsx      ← datos crudos, una hoja por tema
│   └── 2026-10-13 164014_DriverInfo.xlsx     ← (trimestral/semestral)
├── json/                                     ← los mismos datos, en JSON
├── vds_configuration/                        ← .zip con la configuración de cada vDS
└── logs/
    └── <fecha hora>_proactiva.log            ← todo lo que pasó en la corrida
```

Si corrés el proceso más de una vez el mismo día, cada corrida tiene su propia carpeta: `resultados\2026-10-13 (1)\`, `(2)`, etc. Nunca se pisan.

---

## 5. Los chequeos (IDs `VSP-*`)

**Mensuales:** `ME-01` Sizing del vCenter · `ME-02` Espacio en particiones del VCSA · `ME-03` Tamaño de snapshots · `ME-04` Syslog · `ME-05` Políticas de multipath · `ME-06` Consistencia de versiones ESXi por cluster · `ME-07` Consistencia de recursos por cluster · `ME-08` Placas E1000 en VMs Windows · `ME-12` Timeouts de ESXi Shell (TSM) · `ME-13` Alarmas / Falso Positivo · `ME-14` Power Management · `ME-15` VMware Tools · `ME-16` ISOs conectadas · `ME-17` Placas que no inician conectadas · `ME-18` Snapshots de más de 14 días

**Trimestrales:** `TR-01` Fin de soporte de ESXi · `TR-02` Compatibilidad de hardware (HCL) · `TR-03` NTP · `TR-04` Drivers y firmware · `TR-05` DNS · `TR-06` Licencias · `TR-07` Certificados del vCenter · `TR-08` Certificados de ESXi · `TR-09` Vencimiento de la password de root del vCenter · `TR-10` Salud de estadísticas de performance · `TR-11` NIOC en vDS · `TR-12` Consistencia de port groups en standard switches

**Semestrales:** `SE-01` Backup de configuración de vDS · `SE-02` Backups del VAMI

**TR-04 (Drivers):** lista todas las placas en la hoja "Host Drivers" del anexo y marca "No compatible" si alguna no figura en la HCL para la release de su host. Las versiones de driver/firmware no se pueden validar automáticamente con los datos de referencia, así que si no hay incompatibilidades el Estado queda en blanco para que el técnico lo complete revisando la URL de cada placa.

**TR-02 (Compatibilidad de servidores):** en cada corrida se descarga el catálogo completo de "Systems / Servers" del [Broadcom Compatibility Guide](https://compatibilityguide.broadcom.com), desde la misma API que usa su web (sin autenticación, unos 6 MB). Si no hay internet o la API cambia, se usa la copia local `src\collection\reference-data\servers-hcl.csv`; la columna `Note` de la hoja ESXi del Excel crudo indica qué fuente se usó. Se compara modelo + familia de CPU + release del host. Si el modelo está en la HCL pero no con la familia de CPU del host, queda como "Sin datos" y el detalle de la checklist lista esos hosts para revisión manual.

**Azure VMware Solution (AVS):** los vCenters de AVS (`*.avs.azure.com`) no se relevan: Microsoft gestiona sus hosts, certificados y cuentas del appliance, y el usuario del cliente no tiene derechos sobre ellos, así que los hallazgos no serían accionables. Si se incluyen en la lista de vCenters, el script los omite y lo avisa. Si aparecen en los datos de una corrida anterior, se dejan afuera del anexo y la checklist (el Excel crudo los conserva). La regla está en `src\shared\Exclusions.psm1`.

**Hosts que no se evalúan:** los appliances de HCX ("VMware Mobility Platform") figuran como hosts en el vCenter pero no tienen configuración propia, así que los chequeos de hosts los saltean (el Excel crudo los conserva). Los ESXi anidados (modelo `VMware7,1`) tampoco se evalúan en TR-02 porque no tienen entrada en la HCL de servidores. Los hosts cuya configuración no se pudo leer no se evalúan en TSM (ME-12) ni Power Management (ME-14), y los certificados de host que no se pudieron leer (por ejemplo, falta el privilegio `Certificate.Manage`) se informan aparte en TR-08 en vez de darse por válidos.

**Si un chequeo falla:** el error queda en el log, el chequeo se saltea y sus filas de la checklist quedan en blanco con una nota para revisarlas a mano; el resto del reporte se genera igual. Al final de la consola se listan los chequeos que fallaron.

**Hosts Nutanix en TR-02:** los hosts cuyo fabricante es Nutanix no se validan contra la HCL de Broadcom sino contra la Nutanix Compatibility & Interoperability Matrix (versión de AOS vs ESXi). Por eso TR-02 se cierra solo con el resto de los hosts (Resultado Esperado / No recomendado) y, si hay hosts Nutanix, el script agrega a la checklist una fila nueva justo debajo, **`VSP-TR-02-NX`** (Procedimiento: Manual), con el Estado en blanco y los hosts listados en Comentarios, para que la complete el responsable del cliente. Si todos los hosts son Nutanix, TR-02 queda en blanco. La fila solo existe en las checklists de clientes con Nutanix. Los nodos Nutanix de otros fabricantes (Dell XC, Lenovo HX, HPE DX) se reportan con el fabricante de origen y no se pueden detectar desde el vCenter.

**Datos de referencia:** `servers-hcl.csv` (respaldo de TR-02) y los JSON de `src\collection\reference-data\` (HCL de I/O para TR-04, releases de ESXi, VMware Tools, vSAN) se actualizan con `.\runtime\pwsh.exe -File .\src\collection\Update-ReferenceData.ps1`. El CSV también se puede reemplazar por uno exportado a mano desde la web del Broadcom Compatibility Guide ("Export to CSV"): tiene el mismo formato. Con datos viejos, TR-04 muestra placas como "Sin datos".

---

## 6. Ramas de git

- **`main`**: código probado y funcionando. **No se pushea sin autorización explícita.**
- **`development`**: rama de trabajo activa. Acá se prueban los cambios antes de pasarlos a `main`.
- **`backup/vrops-veeam-sin-implementar`**: resguardo de automatizaciones que quedaron sin terminar (consultas a vROps e integración con Veeam). No están en `development` ni en `main` porque nunca se terminaron de probar.

---

## 7. Notas para quien vaya a tocar el código

- **Los nombres de hojas y columnas del Excel crudo** (`ProactivaCollector.GetReport()`) son un contrato: los chequeos los leen por nombre. Incluye nombres raros heredados, como la hoja `Certficate`.
- **Los nombres de hojas del anexo y los textos de la checklist** son lo que recibe el cliente; otros procesos los consumen. No cambiarlos sin coordinarlo.
- **Agregar un chequeo**: escribir la función `Test-*` en el archivo de su periodicidad dentro de `src\report\checks\` (recibe `$Context`, devuelve `@{ ID; Resultado; Detalle }`) y agregar su nombre a la lista correspondiente en `Report.psm1`. El orden de esa lista es el orden de las hojas en el anexo.
- **No mover `runtime/`** sin actualizar las rutas de los `.bat` y de `Start-Proactiva.ps1`, que cargan los módulos por ruta completa (para que nunca gane otra instalación de PowerCLI de la PC).
