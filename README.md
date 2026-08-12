# script-proactivas

Automatización del servicio **Visita Proactiva**: recolecta el estado de la infraestructura VMware de un cliente (vCenter, hosts ESXi, VMs, datastores, redes, certificados, licencias, snapshots, backups, etc.), lo compara contra buenas prácticas y genera dos entregables en Excel — un **Anexo Técnico** con el detalle de cada hallazgo y una **Checklist** con el resultado de cada punto — listos para enviar al cliente.

Este documento explica en detalle cómo correrlo, qué hace cada pieza, y qué archivos deberías tener al final de una ejecución. Está pensado para que cualquier persona del equipo, sin haber tocado el proyecto antes, pueda entender el flujo completo.

---

## 1. Cómo correrlo (resumen rápido)

1. Doble clic en **`lanzador.bat`** (en la raíz del proyecto).
2. Se abre una ventana nueva con un menú de consola. Tildá **"Recolección de Datos para Proactiva"** (y opcionalmente "Recolección de Drivers de Host"), presioná `e` para ejecutar, e ingresá el/los vCenter y credenciales cuando se pidan.
3. Cuando termina la recolección, se abre sola otra ventana que convierte los datos a Excel.
4. Si entre lo recolectado está el reporte "Proactiva", se abre una tercera ventana que te pide el **nombre del cliente**, el **mes**, y qué grupo de tareas correr (mensuales / trimestrales / semestrales). Al terminar, quedan generados el Anexo Técnico y la Checklist.
5. Al cerrar, revisá la carpeta `resultados\<fecha de hoy>\` — ahí está todo lo de esa corrida (ver sección 4).

No hace falta instalar nada: el proyecto trae su propio PowerShell portátil y todos los módulos necesarios (VMware PowerCLI, generación de Excel) en la carpeta `runtime\`. Si en una PC puntual algo no carga bien, corré **`Preparar-Equipo.bat`** — diagnostica los motivos más comunes (Visual C++ Redistributable faltante, proveedor NuGet, instalaciones de PowerCLI que compiten con la incluida) y corrige solo lo que es seguro corregir automáticamente; el resto te lo indica con el mensaje de error real, en vez de la consola en silencio de siempre.

### Correr un solo paso manualmente

Si necesitás repetir un paso puntual (ej. reprocesar el Excel sin volver a conectarte al vCenter, o regenerar el Anexo porque te equivocaste al tipear el nombre del cliente), no hace falta correr todo `lanzador.bat` de nuevo. Hay un `.bat` por paso, también en la raíz:

- **`1-Recoleccion-de-Datos.bat`** — corre solo el Paso 1. Crea su propia carpeta de resultados (con numeración `(1)`, `(2)`... si ya corriste algo hoy).
- **`2-Conversion-a-Excel.bat`** — corre solo el Paso 2. Busca solo, en la carpeta de resultados más reciente, el/los JSON pendientes (según `resultado.txt`).
- **`3-Generar-Anexo.bat`** — corre solo el Paso 3. Busca solo el Excel "Proactiva" más reciente y te pide cliente/mes.

Cada uno funciona sin argumentos ni configuración — doble clic y listo.

---

## 2. El flujo completo, paso a paso

`lanzador.bat` orquesta 3 pasos, cada uno corriendo como un proceso de PowerShell separado (por eso, más abajo, van a aparecer archivos que sirven solo para que estos pasos se pasen información entre sí).

### Paso 0 — Resolver la carpeta de esta corrida

Antes de arrancar nada, `lanzador.bat` llama a `src\Resolver-CarpetaResultados.ps1`, que decide en qué carpeta va a quedar todo lo generado hoy:

- Si es la primera vez que se corre el proceso hoy → `resultados\2026-08-07\`
- Si ya se corrió antes hoy → `resultados\2026-08-07 (1)\`, y así sucesivamente (`(2)`, `(3)`...), igual que Windows nombra un archivo duplicado. Así, dos corridas del mismo día nunca se pisan entre sí.

Este nombre de carpeta se pasa como parámetro a los 3 pasos siguientes, para que los tres procesos (que no comparten memoria entre sí) se pongan de acuerdo en dónde guardar todo.

### Paso 1 — Recolección de datos (`run.bat` → `src\devops-powershell\portable.ps1`)

Se abre una app de consola con un menú interactivo (tildar con números, `e` para ejecutar, `c` para limpiar conexiones, `x` para salir). Hoy el menú tiene 2 opciones activas:

| Opción del menú | Qué hace | Qué genera |
|---|---|---|
| **Recolección de Datos para Proactiva** | Se conecta a uno o más vCenter y recorre hosts, clusters, VMs, datastores, switches, snapshots, certificados, licencias, alarmas, salud de vCenter, actividad de backup del propio appliance, etc. | Un JSON `<fecha> <hora>_Proactiva.json` |
| **Recolección de Drivers de Host** | Revisa las tarjetas de red/HBA de cada host ESXi y las compara contra la lista de compatibilidad de hardware (HCL) de VMware | Un JSON `<fecha> <hora>_DriverInfo.json` |

Podés tildar una, la otra, o las dos — cada una que tildes y ejecutes agrega su JSON a `resultados\<carpeta de hoy>\json\`, y anota el nombre de ese archivo en `resultado.txt` (ver sección 5).

Las credenciales de vCenter se piden por consola en el momento (`Get-Credential`), nunca quedan guardadas en ningún archivo.

> Nota: el menú también tiene un plugin de Veeam a medio hacer que quedó **fuera** de esta rama (ver sección 7 — ramas de git). Si algún día se retoma, está resguardado ahí.

### Paso 2 — Conversión a Excel (`src\JSONtoExcels.ps1`)

Lee `resultado.txt`, y por cada JSON que aparece ahí, genera un Excel con una hoja por cada tipo de dato recolectado (vCenter, ESXi, VM, Datastores, Switches, Snapshots, Particiones, Sizing, Licencias, Certificados, etc.), con estilo (encabezados verdes) y auto-filtro.

Sale un Excel por cada JSON procesado, en `resultados\<carpeta de hoy>\excel\`.

### Paso 3 — Anexo Técnico y Checklist (`src\proactivas-auto2.0.ps1`)

Solo corre si entre lo recolectado estuvo el reporte "Proactiva" (`lanzador.bat` revisa si `resultado.txt` contiene la palabra `_Proactiva`; si solo corriste Drivers, este paso se salta).

1. Te pide por consola el **nombre del cliente** y el **mes** del reporte.
2. Te pide qué tareas correr: **1** mensuales, **2** trimestrales, **3** semestrales (podés combinar, ej. `1,2,3` para correrlas todas).
3. Toma el Excel "Proactiva" más reciente dentro de la carpeta de esta corrida.
4. Corre cada función de verificación seleccionada (~28 en total — ver sección 6), cada una identificada con un ID tipo `VSP-ME-01`. Cada función lee una hoja específica del Excel, aplica una regla de negocio (ej. "¿hay discos con menos del 30% libre?", "¿hay VMs con snapshots de más de 3 días?", "¿el vCenter está sobredimensionado para su versión?"), y si encuentra algo, agrega una hoja de detalle al **Anexo Técnico**.
5. En paralelo, copia la plantilla de checklist y completa las columnas Resultado/Detalle de cada ID con lo que encontró cada función, generando la **Checklist**.

Salen dos archivos en `resultados\<carpeta de hoy>\anexo\`:
- `Anexo Tecnico - {Cliente} - {Mes}.xlsx`
- `Checklist Proactiva - {Cliente} - {Mes}.xlsx`

Si volvés a correr el paso 3 con el mismo cliente y mes (ej. porque corregiste algo), estos dos archivos se regeneran limpios desde cero — no se van acumulando hojas viejas.

---

## 3. Estructura de carpetas del repositorio

```
script-proactivas/
├── lanzador.bat          ← corre los 3 pasos seguidos (uso normal)
├── 1-Recoleccion-de-Datos.bat  ← corren un paso individual a mano
├── 2-Conversion-a-Excel.bat      (ver sección 1)
├── 3-Generar-Anexo.bat
├── Preparar-Equipo.bat    ← diagnostica requisitos de PowerCLI en esta PC
├── resultado.txt          ← archivo de trabajo temporal (se borra solo, ver sección 5)
├── resultados/             ← acá quedan todos los entregables, organizados por fecha
├── runtime/                ← PowerShell portátil + módulos (VMware PowerCLI, ImportExcel).
│                              No es código propio, no hace falta tocarlo nunca.
└── src/                     ← todo el código propio del proyecto
    ├── run.bat
    ├── JSONtoExcels.ps1
    ├── proactivas-auto2.0.ps1
    ├── Resolver-CarpetaResultados.ps1
    └── devops-powershell/    ← la app de consola del Paso 1
        ├── portable.ps1      (entry point real)
        ├── app/               (framework del menú: config, conexiones, UI)
        └── automatizaciones/  (los plugins del menú + la librería de recolección)
            └── lib/
                ├── proactivas.psm1   ← el motor real: clase Proactiva con toda la lógica de recolección
                ├── check-hcl.ps1, check-vsan-hcl.ps1  ← compatibilidad de hardware (HCL)
                ├── update-reference.ps1  ← utilidad manual para refrescar los JSON de referencia de HCL
                └── data/               ← JSON de referencia (HCL, releases de ESXi, sizing de vCenter)
```

---

## 4. Qué archivos quedan al finalizar una ejecución

Después de correr `lanzador.bat` completo (los 3 pasos, con Proactiva incluido), en `resultados\<fecha>[ (N)]\` vas a encontrar:

```
resultados/2026-08-07/
├── json/
│   ├── 2026-08-07 163854_Proactiva.json      ← dato crudo recolectado del vCenter
│   └── 2026-08-07 164014_DriverInfo.json     ← (si tildaste ese plugin)
├── excel/
│   ├── 2026-08-07 163854_Proactiva.xlsx      ← mismo dato, en Excel con una hoja por tema
│   └── 2026-08-07 164014_DriverInfo.xlsx
├── anexo/
│   ├── Anexo Tecnico - {Cliente} - {Mes}.xlsx     ← ENTREGABLE: detalle de hallazgos
│   └── Checklist Proactiva - {Cliente} - {Mes}.xlsx ← ENTREGABLE: checklist para el cliente
└── logs/
    ├── <hora>_portable.log         ← todo lo que pasó en el Paso 1 (recolección)
    ├── <hora>_JSONtoExcels.log     ← todo lo que pasó en el Paso 2 (conversión)
    └── <hora>_proactivas-auto.log  ← todo lo que pasó en el Paso 3 (Anexo/Checklist)
```

Los dos archivos que le mandás al cliente son los de `anexo\`. El resto (`json/`, `excel/`, `logs/`) es material de trabajo/diagnóstico — útil si algo sale mal o si querés reprocesar sin volver a conectarte al vCenter.

Si corrés el proceso una segunda vez el mismo día, se crea `resultados\2026-08-07 (1)\` con su propio set completo de las 4 carpetas — nunca se mezcla ni se pisa con la corrida anterior.

---

## 5. ¿Qué es `resultado.txt`?

Es una lista de trabajo temporal que conecta los 3 pasos (que corren como procesos separados y no comparten memoria):

- Cada plugin que corrés en el Paso 1 (Proactiva, DriverInfo) agrega ahí el nombre del JSON que acaba de generar.
- El Paso 2 lo lee para saber exactamente qué JSON son de esta corrida (no reprocesa todo lo que haya en `resultados\`).
- `lanzador.bat` lo usa para decidir si vale la pena correr el Paso 3 (busca la palabra `_Proactiva`).
- Al terminar el Paso 3 (o antes, si no correspondía), `lanzador.bat` lo borra.

No hace falta administrarlo manualmente ni se puede sacar del flujo del todo: los pasos 2 y 3 lo necesitan para saber qué procesar. Pero no guarda nada valioso — es pura plomería interna, se recrea y se borra solo en cada corrida.

---

## 6. Los checklist (IDs `VSP-*`)

Cada función de verificación del Paso 3 corresponde a un ítem con un ID fijo en la plantilla de Checklist (`src\devops-powershell\templete\Checklist Proactiva Actualizada - Cliente - Mes.xlsx`). Se agrupan en mensuales, trimestrales y semestrales:

**Mensuales:** `VSP-ME-01` Sizing del vCenter · `02` Espacio en particiones · `03` Snapshots antiguos · `04` Configuración de Syslog · `05` Multipath de storage · `06` Consistencia de versiones ESXi · `07` Recomendaciones de configuración · `08` Placas de red no recomendadas (E1000) · `12` TSM/SSH habilitado · `13` Falso positivo de alarmas · `14` Patch management · `15` VMware Tools desactualizadas · `16` ISOs montadas · `17` Placas de red · `18` Snapshots

**Trimestrales:** `VSP-TR-01` Fin de soporte (EOL) · `02` Compatibilidad de componentes (HCL) · `03` Configuración NTP · `04` Drivers (informativo) · `05` Configuración DNS · `06` Vencimiento de licencias · `07` Certificado de vCenter · `08` Certificado de ESXi · `09` Cuenta root de vCenter · `10` Salud de performance · `11` NIOC · `12` vSwitch estándar

**Semestrales:** `VSP-SE-01` Backup de configuración de vDS · `SEC-02`/`SEC-03` Certificados (seguridad) · `SE-02` Actividad de backup

Cada función lee una hoja puntual del Excel generado en el Paso 2, aplica su regla, y si encuentra un problema arma una hoja de detalle en el Anexo Técnico además de marcar el resultado en la Checklist.

---

## 7. Ramas de git

- **`main`** — código probado y funcionando. **No se pushea sin autorización explícita.**
- **`development`** — rama de trabajo activa. Acá se prueban los cambios antes de promoverlos a `main`.
- **`backup/vrops-veeam-sin-implementar`** — resguardo de automatizaciones que quedaron sin terminar (varios plugins de consulta a vROps, y la integración con Veeam). No están en `development`/`main` porque nunca se terminaron de probar, pero quedan ahí por si en algún momento se retoman.

---

## 8. Notas para quien vaya a tocar el código

- **No mover `runtime/` de lugar** sin actualizar las rutas relativas que apuntan a `pwsh.exe` y `Modules\` (en `lanzador.bat`, `run.bat` y `portable.ps1`).
- Los 3 pasos del pipeline (`portable.ps1`, `JSONtoExcels.ps1`, `proactivas-auto2.0.ps1`) reciben la carpeta de la corrida por parámetro (`-CarpetaResultados`). Si los corrés sueltos para debug sin pasar el parámetro, caen solos a la carpeta más reciente dentro de `resultados\`.
- `src\devops-powershell\automatizaciones\lib\proactivas.psm1` hace *dot-sourcing* de algunos archivos con rutas relativas al directorio de trabajo (no a `$PSScriptRoot`) — funciona porque `portable.ps1` se para en `devops-powershell\` antes de arrancar. Si se reorganiza esa carpeta, hay que tener esto en cuenta.
- `src\devops-powershell\automatizaciones\lib\update-reference.ps1` es una utilidad de mantenimiento manual (no se corre sola) que refresca los JSON de referencia de HCL/ESXi/vSAN desde internet. Conviene correrla de tanto en tanto para que las validaciones de compatibilidad no queden desactualizadas.
