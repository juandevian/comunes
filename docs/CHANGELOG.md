# Changelog de comunes

## Sincronización con orioncop (esquema 268)

**Base técnica:** `main` de `comunes` quedó en el commit inicial (esquema 264); `orioncop` ya dependía de cambios posteriores que estaban sin versionar. Esta entrega los incorpora para que ambos repositorios compilen juntos.

### Versiones de ensamblado

| Proyecto | Versión anterior | Versión nueva |
|---|---|---|
| `OriWin` | 4.6.31.1435 | 4.6.32.1446 |
| `PanDat` | 6.11.120.1435 | 6.11.122.1446 |
| `PanL` | 8.22.183.1436 | 8.22.184.1446 |
| `WinCom` | 2.2.9.1436 | 2.2.10.1441 |

### Esquema `PanDat/XmlBd/OrionCop_Net.xml` (versión 264 -> 268)

- **265-267 (trabajo previo del responsable):**
  - `OriAnos`: `IdTipoIncentivo` (con migración desde `AplicaProntoPago`, que se elimina), `DiasPExtemporaneo`, `IdServicioMulta`, `ValorMultaPExtemporaneo`.
  - `FechaMulta`, `Multada`, `PieFacturaTres` y `CentroCostos` en sus tablas respectivas.
  - `ReferenciaPago`: `STRING(8)` -> `VSTRING(25)`.
  - Constantes de los grupos 20 (`EnuTipoIncentivo`) y 21 (`EnuTipoDsctoPP`).
- **268 (herramienta IBC, ver `orioncop/docs/superpowers/specs/2026-09-28-herramienta-ibc-design.md`):**
  - Tabla nueva `OriIbcCertificados` (certificados del interés bancario corriente por centro de utilidad).
  - `OriCentrosUtilidadOriCop`: columnas `ModoInteres`, `TasaFijaDeseada`, `FactorVariable` (valor por defecto 0 = sin parametrizar).
- `PanL/acPanL.vb`: valor `EnuIbcCertificado` al final de `EnuIdClasesPanDef` (se agrega al final para no alterar valores persistidos).

### Código

- `OriWin/ClsFormInterface.vb`, `PanL/clsPanorama.vb`, `PanL/clsImportar.vb`, `PanDat/*`, `PanL/*`, `WinCom/winCopiaSeg.xaml.vb`: normalización de formato del IDE (mayúsculas de miembros de enumeraciones, líneas en blanco) sin cambio de comportamiento.
- `PanDat/clsPanoramaDat.vb`: se retira la propiedad `StrNombreServidor` (Friend, sin usos en `comunes`, `orioncop`, `adminorion` ni `orionpcorreo`).

### Documentación

- Se versionan `docs/` y los punteros de `.github/` que estaban sin seguimiento.
