# Pull Request - comunes

> Completa todas las secciones. Esta plantilla es obligatoria para mantener trazabilidad y control de riesgo cross-repo.

## 1. Resumen del cambio

Describe en 3-6 lineas:

- que problema resuelve;
- que cambia tecnicamente;
- que NO cambia.

## 2. Tipo de cambio

- [ ] feat
- [ ] fix
- [ ] refactor
- [ ] docs
- [ ] chore
- [ ] hotfix

## 3. Alcance en comunes

- [ ] PanL
- [ ] PanDat
- [ ] OriWin
- [ ] WinCom
- [ ] build/release
- [ ] documentacion

Archivos o modulos principales modificados:

-

## 4. Evidencia de validacion

- build local/CI:
- pruebas ejecutadas:
- validacion manual (pasos y resultado):

## 5. Impacto cross-repo (obligatorio)

- [ ] Sin impacto cross-repo
- [ ] Con impacto cross-repo

Si hay impacto, completar tabla:

| Repositorio | Tipo de impacto | Cambio requerido | Estado |
|---|---|---|---|
| adminorion | | | |
| orioncop | | | |
| orionpcorreo | | | |
| orion-installer | | | |

## 6. Compatibilidad y riesgo

- Riesgo funcional (bajo/medio/alto):
- Riesgo tecnico (bajo/medio/alto):
- Riesgo de despliegue (bajo/medio/alto):
- Posible regresion conocida:

## 7. Plan de rollback (obligatorio)

1. Estrategia (ejemplo: `git revert <sha>`):
2. Condicion para activar rollback:
3. Impacto esperado del rollback:

## 8. Checklist de cumplimiento

- [ ] Alineado con `docs/ESTRATEGIA_REPOSITORIOS_GITHUB.md`.
- [ ] Alineado con `docs/PLAN_MAESTRO.md`.
- [ ] Contratos actualizados en `docs/` si aplica.
- [ ] Impacto cross-repo evaluado y documentado.
- [ ] Plan de rollback viable.
