# Historial de versiones

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
