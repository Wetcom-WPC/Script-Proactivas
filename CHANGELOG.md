# Historial de versiones

## v3.2.1 (2026-10-21)

Corrección para PCs que ya tienen VMware PowerCLI instalado.

- En una PC con otra versión de PowerCLI instalada a nivel sistema (ej. 12.5 en `Program Files\WindowsPowerShell\Modules`), la recolección fallaba al arrancar con `Assembly with same name is already loaded`. Ahora el script y sus procesos de recolección ven únicamente los módulos de `runtime\Modules`, así que una instalación de PowerCLI de la PC ya no interfiere.
- `extra\Preparar-Equipo.bat`: el aviso de "instalaciones de PowerCLI fuera del proyecto" ahora refleja esto y explica qué hacer si el error persiste.

## v3.2.0 (2026-10-15)

Recolección mucho más rápida. Las hojas, columnas, valores y el orden de las filas no cambian: se validó contra v3.1.0 en un cliente real (2 vCenters, 10 hosts, 365 VMs, trimestral). Los datos de las dos versiones coinciden; solo cambia lo que se movió entre corrida y corrida (una VM que hizo vMotion, snapshots de los backups).

| Cliente real (2 vCenters, trimestral) | v3.1.0 | v3.2.0 |
|---|---|---|
| Recolección completa (datos generales + drivers) | 6:04 | 2:08 |

**Qué cambió**
- El inventario se lee con pocas consultas grandes: todas las VMs (con placas de red e ISOs) en una sola consulta, y la configuración de los hosts de lo que ya trae `Get-VMHost` o con consultas por lotes, en vez de varias llamadas por objeto. Los snapshots se piden solo para las VMs que tienen.
- Las búsquedas en memoria que crecían al cuadrado con el inventario (placas, ISOs, snapshots, HCL) pasan a índices.
- Varios vCenters se recolectan en paralelo, cada uno en un proceso aparte con su propia sesión. Los drivers/firmware se reparten entre hasta 4 procesos por vCenter. Si un proceso falla, ese trabajo se rehace en secuencia; `-Sequential` desactiva el paralelismo.
- La HCL de servidores de Broadcom se descarga en segundo plano mientras se lee el inventario.
- El log muestra cuánto tardó cada etapa por vCenter, y al final de la corrida se resume la duración de recolección, conversión y anexo.
- `src\tools\Compare-CollectionJson.ps1` compara los JSON de dos corridas hoja por hoja, para validar cambios en la recolección.
- Como todo el inventario se lee de una vez, una VM que se mueve de host durante la corrida ya no puede figurar en un host en la hoja VM y en otro en vNetwork.
- Las hojas del anexo que agrupan por cluster (ej. Host DNS) salen siempre en el mismo orden; antes cambiaba de una corrida a otra.

**Limitaciones conocidas**
- El orden de las filas de Standard Switch se verificó con hosts de un solo vSwitch; con varios vSwitch por host, los switches salen en el orden de la configuración del host.

## v3.1.0 (2026-10-15)

Primera versión publicada de la línea 3.x. Reemplaza al flujo anterior de tres ventanas y menú.

**Cómo se usa**
- Un solo `Iniciar-Proactiva.bat`: pregunta cliente, mes, tipo de proactiva (mensual / trimestral / semestral, acumulativas) y vCenters al principio, y después corre solo, en una única ventana.
- `extra\Regenerar-Anexo.bat` rehace el Anexo y la Checklist de la última corrida sin volver a recolectar.
- La Checklist solo trae las tareas de la periodicidad de la visita.

**Recolección y chequeos**
- TR-02 (compatibilidad de servidores) consulta la HCL del Broadcom Compatibility Guide en cada corrida, con respaldo en `servers-hcl.csv`. Los hosts Nutanix pasan a una fila manual propia (`VSP-TR-02-NX`) a cargo del responsable del cliente.
- TR-04 (drivers y firmware) implementado.
- TR-06 (licencias) ya no marca como duplicadas las claves compartidas entre vCenters.
- Azure VMware Solution: de los vCenters `*.avs.azure.com` se relevan y analizan solo las VMs (mezcladas con las del resto).
- Un chequeo que falla ya no impide generar la Checklist: queda en blanco con una nota y el resto sigue.
- Los hosts sin datos (appliances de HCX, ESXi anidados, configuración ilegible, certificados sin privilegios) se informan en vez de dar falsos hallazgos.
- La prueba de Falso Positivo espera hasta 90 s a que la alarma dispare, y `-SkipAlarmTest` permite omitirla.

**Limitaciones conocidas**
- La recolección de VMs de AVS y la espera ampliada de la prueba de Falso Positivo se verificaron por código y con datos de corridas reales, pero no se ejecutaron en vivo contra un vCenter.
- El anexo no contempla las excepciones propias de cada cliente ni descarta las VMs de Kubernetes o Nutanix en el chequeo de ISOs.
- Los nodos Nutanix de otros fabricantes (Dell XC, Lenovo HX, HPE DX) no se detectan como Nutanix desde el vCenter.
