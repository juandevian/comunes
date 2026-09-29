# Estrategia de Repositorio GitHub para comunes

> **Estado:** vigente para este repositorio  
> **Alcance:** solo `comunes`  
> **Objetivo:** operar y evolucionar librerias compartidas con bajo riesgo y alta trazabilidad para `adminorion`, `orioncop`, `orionpcorreo` y `orion-installer`.

---

## 1. Proposito de este repositorio

`comunes` centraliza componentes compartidos de la plataforma: `PanL`, `PanDat`, `OriWin`, `WinCom` y otros artefactos reutilizables.

### Lo que SI pertenece a `comunes`

- Librerias y utilidades compartidas.
- Contratos tecnicos reutilizables.
- Versionado de componentes comunes.
- Documentacion de compatibilidad.

### Lo que NO pertenece a `comunes`

- Flujos de UI funcionales de `adminorion` o `orioncop`.
- Casos de uso propios de `orionpcorreo`.
- Orquestacion final de instalacion/release (`orion-installer`).

---

## 2. Relacion con otros repositorios

1. **`adminorion`**: consume contratos y componentes comunes para administracion.
2. **`orioncop`**: depende de librerias compartidas para operacion principal.
3. **`orionpcorreo`**: consume utilidades compartidas para correo/notificaciones.
4. **`orion-installer`**: empaqueta versiones especificas de componentes.

Principio: `comunes` evoluciona de forma independiente, pero nunca rompe contratos compartidos sin versionado y plan de transicion.

---

## 3. Principios de decision

1. Compatibilidad explicita primero.
2. Cambios pequenos y reversibles.
3. Breaking changes solo con version mayor y plan de migracion.
4. Trazabilidad total: issue -> rama -> PR -> tag -> release.

---

## 4. Flujo de ramas y cambios

- `main`: siempre liberable
- `develop`: integracion previa
- `feature/*`, `hotfix/*`, `release/*`

Reglas:

1. No commit directo a `main`.
2. Todo cambio via PR.
3. PR con seccion de impacto cross-repo.
4. Contratos actualizados en `docs/` cuando aplique.

---

## 5. Versionado y releases

Semver por componentes compartidos:

- `X`: ruptura de contrato.
- `Y`: mejora compatible.
- `Z`: correccion.

Cada release debe incluir:

1. Cambios de API/contrato.
2. Repos potencialmente impactados.
3. Recomendacion de adopcion.

---

## 6. CI minima obligatoria

1. Restore de dependencias.
2. Compilacion de proyectos compartidos.
3. Validaciones estaticas disponibles.
4. Empaquetado de artefactos (si aplica).

---

## 7. Rollback

1. Preferir `git revert` del cambio problematico.
2. Publicar release correctiva de parche.
3. Notificar impacto a repos consumidores.
4. Sin reescritura de historia publicada de `main`.

---

## 8. Decision estrategica vigente

Para `comunes` se adopta:

1. Gestion independiente del repositorio.
2. Contratos compartidos con versionado estricto.
3. Releases pequenos, trazables y reversibles.
4. Prioridad en estabilidad de integracion multi-repo.
