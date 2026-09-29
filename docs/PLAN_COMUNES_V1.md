# Plan v1 - comunes

> **Estado:** activo  
> **Fecha:** 2026-06-07  
> **Objetivo macro:** ser base reusable y desacoplada para el nuevo sistema AI-first.

## 1. Objetivo del repo
Transformar `comunes` en un nucleo estable con contratos claros para reducir acoplamiento entre aplicaciones y habilitar migracion por capacidades.

## 2. Estado actual resumido
1. Librerias compartidas con alta centralidad tecnica (`PanL`, `PanDat`, `OriWin`, `WinCom`).
2. Dependencias transversales sobre `orioncop` y `adminorion`.
3. Riesgo de cambios con alto impacto en cascada.

## 3. Alcance v1
1. Inventario de modulos y APIs usadas por cada consumidor.
2. Clasificacion de componentes: estables, fragiles, candidatos a extraer.
3. Definicion de contratos y limites de responsabilidad.
4. Base de observabilidad tecnica para detectar rompimientos.

## 4. Backlog inmediato
1. Mapa `comunes -> consumidores`.
2. Priorizacion de acoplamientos criticos.
3. Contratos versionables para servicios/bibliotecas.
4. Pruebas de regresion minima por contrato.
5. Guia de cambios compatibles para equipos consumidores.

## 5. Riesgos y mitigacion
1. **Ruptura transversal** -> versionado por contrato + pruebas de compatibilidad.
2. **Deuda tecnica heredada** -> aislamiento incremental por modulo.
3. **Bloqueo de migracion** -> separar reglas de negocio de utilidades UI.

## 6. Criterios de exito v1
1. Dependencias criticas documentadas y priorizadas.
2. Contratos base definidos para consumidores principales.
3. Reduccion de cambios de alto riesgo sin control.

