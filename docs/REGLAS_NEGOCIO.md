# Reglas de negocio - comunes

Este documento resume reglas de negocio y validación transversal identificadas en `comunes`.

## Referencias base

- `.github/PLAN_MAESTRO.md`
- `.github/ESTRATEGIA_REPOSITORIOS_GITHUB.md`
- `docs/PLAN_MAESTRO.md`
- `docs/ESTRATEGIA_REPOSITORIOS_GITHUB.md`

## Reglas

1. **Id de usuario:** debe ser valida, única al crear, y existente al abrir/modificar. Si es `GCSTRUSUARIOU`, se carga usuario universal.
2. **Usuario universal:** se inicializa activo, con política `NoCambiar` y no entra en flujo de cambio obligatorio de contraseña.
3. **Cambio obligatorio de contraseña:** para usuarios no universales aplica por tipo `enuProximaVez` o por vencimiento (`enuAlVencimiento` con fecha <= hoy).
4. **Contraseña:** longitud valida entre 4 y 12 caracteres; debe confirmar coincidencia; cuando se exige cambio, debe ser diferente a la actual.
5. **Tipo de cambio de contraseña:** solo admite valores de enum permitidos; al cambiar a `enuAlVencimiento` fija expiración a hoy + 2 meses, en otros casos la fija en hoy.
6. **Fecha de expiración:** si el tipo es `enuAlVencimiento`, debe ser fecha valida y no menor a hoy.
7. **Permisos por usuario/perfil:** para usuarios no universales, los permisos se resuelven por llave `IdObjeto + IdAccion`; si no esta permitido, no hay permiso.
8. **Perfil por usuario:** el `IdUsuario` del perfil debe coincidir con el del usuario padre (excepto usuario universal).
9. **Aplicación valida en perfil de usuario:** `IdApp` se valida dentro del rango de aplicaciones permitido por enum.
10. **Cambio de contraseña en UI comun:** primero debe validarse la contraseña actual para habilitar captura de nueva contraseña y confirmación.

## Evidencia en código

- `PanL\clsUsuario.vb:658-685`
- `PanL\clsUsuario.vb:247-271`
- `PanL\clsUsuario.vb:272-284`
- `PanL\clsUsuario.vb:766-804`
- `PanL\clsUsuario.vb:826-849`
- `PanL\clsUsuario.vb:880-898`
- `PanL\clsUsuario.vb:917-929`
- `PanL\clsUsuario.vb:575-587`
- `PanL\clsPerfilUsuario.vb:158-167`
- `PanL\clsPerfilUsuario.vb:189-192`
- `WinCom\winCambioContrasena.xaml.vb:83-87`
- `WinCom\winCambioContrasena.xaml.vb:162-170`
