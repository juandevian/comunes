Imports System.Runtime.CompilerServices
#Region "Definiciones"
<Assembly: CLSCompliant(True)>
<Assembly: InternalsVisibleTo("PanL")>
<Assembly: InternalsVisibleTo("AdminOrionIU")>
<Assembly: InternalsVisibleTo("WinCom")>
<Assembly: InternalsVisibleTo("OrionCopL")>
<Assembly: InternalsVisibleTo("OriWin")>
<Assembly: InternalsVisibleTo("OrionCopIU")>
<Assembly: InternalsVisibleTo("RepOriCop")>
<Assembly: InternalsVisibleTo("OriIntCon")>
#End Region
#Region "Interfaz"
Public Interface IPanDat
    ReadOnly Property ObjRegistro As Object
    ReadOnly Property ShrIdApp As Short
    ReadOnly Property StrNombreArchivos As String
    ReadOnly Property EntVersionBDEnProg As Integer
End Interface
#End Region
#Region "Enumeradores"
Public Enum EnuListaAplicaciones As System.Int32
    None = 0
    EnuAdministrador = 100
    '
    EnuOrionCop = 803
End Enum
Public Enum EnuProcesoBK As Integer
    None = 0
    enuGeneraBK
    enuRestauraBK
End Enum
Public Enum EnuTipoInstanciamiento As Integer
    None
    enuInstalacion
    enuActualizacion
    enuNormal
End Enum
Friend Enum EnuTipoAutenticacion As Integer
    None = 0
    enuWindows
    enuBaseDatos
End Enum
Friend Enum EnuTipoSentencia As Integer
    None = 0
    enuSelect
    enuUpdate
    enuDelete
    enuInsert
End Enum
Friend Enum EnuTipoTabla As Integer
    None = 0
    enuUnica
    enuPrimaria
    enuSecundaria
End Enum
Friend Enum EnuProveedorBD As Integer
    None = 0
    enuOleDb
    enuMySql
    enuOracle
    enuSQLServer
    enuExcel10
End Enum
Friend Enum EnuTipoLinea As Integer
    None = 0
    enuInicioNodo
    enuInicioNodoConProp
    enuFinNodo
    enuNodoConProp
    enuPropiedad
End Enum
Public Enum EnuSecuenciaAccion As Integer
    None = 0
    enuInicio
    enuFin
End Enum
Public Enum EnuEstadoBaseDatos As Integer
    None = 0
    enuNoExiste
    enuDespoblada
    enuDesActualizada
    enuActualizada
End Enum
Friend Enum EnuTipoObjetoBD As Integer
    None = 0
    enuBaseDatos
    enuTabla
    enuColumna
    enuIndice
    enuRegistro
    enuRelacion
    enuBProcedimientoAlmacenado
End Enum
Friend Enum EnuTipoComando As Integer
    None = 0
    enuPropio
    enuInstruccionSQL
    enuLlamadaProc
End Enum
' Tipos de severidad de la notificación
Public Enum EnuSeveridadNot As Integer
    None
    EnuOk
    EnuCamInsatis
    EnuInformacion
    EnuAdvertencia
    EnuDatoInvalido
    EnuFalta
    EnuError
    EnuExcep
End Enum
Public Enum EnuSeveridadMen As Integer
    None
    enuMensajeNormal
    enuInformacion
    enuAdvertencia
    enuError
End Enum
#End Region