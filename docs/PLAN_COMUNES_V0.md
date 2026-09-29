# Plan Maestro de comunes

> **Estado:** vigente para ejecucion inicial  
> **Ambito:** solo este repositorio (`comunes`)  
> **Documento complementario:** `docs/ESTRATEGIA_REPOSITORIOS_GITHUB.md`

---

## 1. Proposito

Este plan define **que ejecutar** en `comunes` para estabilizar y evolucionar las librerías compartidas de la plataforma Orion.

La estrategia de gobierno y relación multi-repo vive en `docs/ESTRATEGIA_REPOSITORIOS_GITHUB.md`. Este documento aterriza trabajo ejecutable para este repositorio.

---

## 2. Objetivo de 5 días (2026/06/11)

Dejar `comunes` en estado "estable y consumible" para cambios semanales con bajo riesgo de ruptura cross-repo.

Resultado esperado al dia (5) 11 de junio del 2026:

1. Contratos técnicos de librerías documentados.
2. Build reproducible en CI para proyectos compartidos.
3. Versionado semántico operativo para librerías.
4. Flujo de cambios con rollback seguro por componente.

---

## 3. Alcance real del plan

Incluye:

1. Estabilización de `PanL`, `PanDat`, `OriWin`, `WinCom` y artefactos compartidos.
2. Definición y documentación de contratos consumidos por `adminorion`, `orioncop` y `orionpcorreo`.
3. Gobierno de versionado y compatibilidad.
4. Trazabilidad de impactos en `orion-installer`.

No incluye:

1. Cambios funcionales propios de `adminorion`, `orioncop` u `orionpcorreo`.
2. Reescritura total de librerías legacy.
3. Migración tecnológica completa en esta etapa.

---

## 4. Lineas de trabajo

### Linea A. Estabilidad técnica

1. Matriz de configuraciones de build soportadas.
2. Checklist de restore/build local y CI.
3. Registro de fallas recurrentes y mitigación.

### Linea B. Contratos compartidos

1. Inventario de API publica por librería.
2. Identificación de breaking changes potenciales.
3. Regla de compatibilidad y deprecación.

### Linea C. Gobierno de releases

1. Política de versionado por librería/paquete.
2. Notas de release con impactos cross-repo.
3. Flujo de rollback por `git revert` + release correctiva.

### Linea D. Base para evolución

1. Mapa de zonas frágiles de código compartido.
2. Priorización de encapsulaciones de bajo riesgo.
3. Backlog técnico semanal/quincenal.

---

## 5. Fases y calendario sugerido

### Fase 1 (Día 1-2): Recuperar control

1. Confirmar build y CI minima de librerías.
2. Consolidar contratos actuales en `docs`.
3. Definir baseline de compatibilidad.

### Fase 2 (Día 3-4): Estabilizar para iterar

1. Corregir puntos de falla repetitivos.
2. Versionar cambios con criterios uniformes.
3. Formalizar comunicación de impactos a repos consumidores.

### Fase 3 (Día 5): Preparar aceleración

1. Ejecutar mejoras de desacoplamiento priorizadas.
2. Medir impacto en estabilidad y velocidad de integración.
3. Definir siguiente plan con base en métricas.

---

## 6. Indicadores de seguimiento

1. % de builds exitosos en CI.
2. Numero de releases de `comunes` sin hotfix.
3. Numero de cambios con impacto cross-repo documentado.
4. Tiempo promedio para resolver regresiones compartidas.
5. Tiempo de recuperación ante rollback.

---

## 7. Riesgos abiertos y mitigacion

1. **Acoplamiento histórico no documentado.**  
   Mitigación: inventario de contratos y matriz de dependencias.
2. **Regresiones en repos consumidores.**  
   Mitigación: notas de impacto y versionado explicito.
3. **Cambios urgentes sin trazabilidad.**  
   Mitigación: PR corta, checklist de impacto y rollback definido.

---

## 8. Criterio de éxito

Este plan se considera cumplido cuando:

1. `comunes` publica cambios con compatibilidad controlada.
2. Los repos consumidores integran versiones con bajo riesgo.
3. Existe trazabilidad completa de contratos, cambios y rollback.
4. El siguiente plan puede enfocarse en evolución, no en incidentes repetitivos.
