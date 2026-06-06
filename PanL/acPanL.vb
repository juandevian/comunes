Imports System.Runtime.CompilerServices
#Region "Definiciones"
<Assembly: CLSCompliant(True)>
<Assembly: InternalsVisibleTo("AdminOrionIU")>
<Assembly: InternalsVisibleTo("WinCom")>
<Assembly: InternalsVisibleTo("OrionCopL")>
<Assembly: InternalsVisibleTo("OrionCopIU")>
<Assembly: InternalsVisibleTo("OriWin")>
<Assembly: InternalsVisibleTo("RepOriCop")>
<Assembly: InternalsVisibleTo("OriIntCon")>
#Region "Interfaz"
'
#End Region
#End Region

#Region "Estructuras"
Friend Structure SECURITY_ATTRIBUTES
    Implements IEquatable(Of SECURITY_ATTRIBUTES)
    Public Property NLength As Integer
    Public Property LpSecurityDescriptor As Integer
    Public Property BInheritHandle As Integer

    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing Then Return False
        Dim lblnEsIgual As Boolean = (obj.GetType.Name = "SECURITY_ATTRIBUTES")
        If lblnEsIgual Then
            lblnEsIgual = Equals(obj)
        End If
        Return lblnEsIgual
    End Function

    Public Overloads Function Equals(other As SECURITY_ATTRIBUTES) As Boolean Implements IEquatable(Of SECURITY_ATTRIBUTES).Equals
        Dim lblnEsIgual As Boolean = (NLength = other.NLength)
        If lblnEsIgual Then
            lblnEsIgual = (LpSecurityDescriptor = other.LpSecurityDescriptor)
        End If
        If lblnEsIgual Then
            lblnEsIgual = (BInheritHandle = other.BInheritHandle)
        End If
        Return lblnEsIgual
    End Function

    'Public Shared Operator =(SECURITY_ATTRIBUTES1 As SECURITY_ATTRIBUTES, SECURITY_ATTRIBUTES2 As SECURITY_ATTRIBUTES) As Boolean
    '    Return SECURITY_ATTRIBUTES1.Equals(SECURITY_ATTRIBUTES1)
    'End Operator

    'Public Shared Operator <>(SECURITY_ATTRIBUTES1 As SECURITY_ATTRIBUTES, SECURITY_ATTRIBUTES2 As SECURITY_ATTRIBUTES) As Boolean
    '    Return Not SECURITY_ATTRIBUTES1.Equals(SECURITY_ATTRIBUTES2)
    'End Operator

    Public Overrides Function GetHashCode() As Integer
        Return 0
    End Function
End Structure
#End Region

#Region "Enumeradores"
Friend Enum EnuCategoriaImagenDef As Byte
    None = 0
    EnuFotoTercero
    EnuFotoExtintor
    EnuMedicas
    EnuDocumentos
    EnuFirmas
End Enum
Public Enum EnuDatoValidarDef As Byte
    EnuPais = 0
    EnuDpto
    EnuCiudad
End Enum
Public Enum EnuEstadoContrato As Byte
    EnuInstalando = 0
    EnuActivo
    EnuSuspendido
    EnuInactivo
    EnuRegistradoNit
End Enum
Public Enum EnuEstadoObjetoDef As Byte
    EnuConsultando = 0                            'Objeto se está consultando
    EnuCreando                                    'Objeto se está creando
    EnuModificando                                'Objeto se está modificando
    EnuEliminando                                 'Objeto se está eliminando
End Enum
Friend Enum EnuEstiloNombreTercero
    None
    EnuMayIni
    EnuMayAll
    EnuAsIng
End Enum
Public Enum EnuGrupoConstantesPanDef As Byte
    None = 0
    EnuTipoCambioContrasena = 1
    EnuTipoLicenciamiento
    EnuTipoDocIdentidad
    EnuTipoTercero
    EnuCategoriaImagen
    EnuTipoCuenta
    EnuOrigenInstancia
    EnuTipoLog
    EnuTipoConsulta
    EnuTipoNomTercero
End Enum
Public Enum EnuIdAccionDef As Short
    None = 0
    EnuConsultar = 1001
    EnuCrear
    EnuModificar
    EnuSuprimir
    EnuAnular
    EnuImprimir
End Enum
Public Enum EnuIdClasesPanDef As Short
    ' Panorama
    None = 0
    EnuAccion
    EnuAdministrador
    EnuAplicacion
    EnuCarpeta
    EnuCarpetaUsuario
    EnuCentroUtilidad
    EnuCiudad
    EnuConsultaSql
    EnuCuentaContabilidad
    EnuDefAplicacion
    EnuDepartamento
    EnuDocumentoContabilidad
    EnuImagen
    EnuImportar
    EnuLogAplicacion
    EnuMensajePan
    EnuOrigenInstancia
    EnuPais
    EnuPanorama
    EnuPerfil
    EnuPerfilUsuario
    EnuPermisoAccion
    EnuTercero
    EnuUbicacion
    EnuUsuario
    ' Orion
    EnuSector = 40
    EnuModuloContribucion
    EnuSectorModulo
    EnuCEnutiliOriCop
    EnuCuentaBanco
    EnuInteresMora
    EnuAgrServicios
    EnuServicio
    EnuModuloServicio
    EnuSectorModuloServicio
    EnuHistServicio
    EnuHistModServicio
    EnuAno
    EnuPeriodo
    EnuCliente
    EnuPredio
    EnuItemProgFact
    EnuFactura
    EnuItemFactura
    EnuNovedad
    EnuNotaDb
    EnuItemNotaDb
    EnuNotaCon
    EnuItemNotaCon
    EnuEstadoCuenta
    EnuFacturaEstado
    EnuReciboCaja
    EnuItemRecCaja
    EnuMediosPago
    EnuAnticipo
    EnuNovedadAnt
    EnuNotaDevAnt
    EnuNotaCr
    EnuItemNotaCr
    EnuNotaReversaCr
    EnuDocsCont
    EnuNotaAjusteCuotaAdmin
    EnuTarea
    EnuItemTarea
    EnuProveEFac
    EnuInformeCont
    EnuDscto
    EnuServicioEstadoCta
    EnuPropietario
End Enum
Public Enum EnuIdMens As Byte
    None
    ' MainWindows
    EnuNoPreFac
    ' WinServicio
    EnuModSinSector
    ' WinCorreoE
    EnuCliente
    EnuTipoCorreo
    EnuArchivo
    EnuAsunto
    EnuMensaje
    EnuDoc
    EnuFecIni
    EnuFecFin
    ' Abrir un objeto
    EnuAbrir
    ' ClsCarpeta
    EnuDatCarInvalidos
    ' Importar
    EnuDatosTablaOri
    ' ClsCentroUtilidadOrionCop
    EnuMensInicio
    EnuServicioMalParam
    ' Recibo de Caja
    EnuDeudaNoPagable
    EnuFechaNov
    EnuDscto
    EnuTotalMP
    EnuSinDeuda
    '
    EnuNoCreable
    EnuNoModificable
    EnuNoEliminable
    EnuNoAnulable
    '
    EnuCarpSinCEnutil
    ' Ubicación
    EnuIdPais
    EnuIdDpto
    EnuIdCiudad
    EnuNomPais
    EnuNomDpto
    EnuNomCiudad
End Enum
Public Enum EnuModoInstanciaObjDef As Byte
    None = 0
    EnuNavegable
    EnuUnico
    EnuDeColeccion
End Enum
Public Enum EnuOrigenInstanciamientoDef As Byte
    None = 0
    EnuEstacionTrabajo
    EnuUrl
    EnuMovil
End Enum
Public Enum EnuProcesoDef As Byte
    None
    EnuIntegrFac
    EnuIntegrRec
    EnuIntegrNcr
    EnuIntegrNdb
    EnuIntegrNco
    EnuInteAnt
    EnuInteNRRC
    EnuInteEstadoCta
    EnuBK
    EnuResBK
    EnuCausaMora
    EnuRepEdadCar
    EnuRepAntXApl
    EnuPreFacProPA
    EnuPreFacProSPA
    EnuPreFacCli
    EnuPreFacPre
    EnuPasandoAFact
    EnuGenEstadosCtaCli
    EnuGenEstadosCtaPre
    EnuApliAnti
    EnuRevPreFac
    EnuAjusCuota
    EnuImpoFras     'Importar facturas
    EnuExpFras      'ExportarFacturas
    EnuInsFacApi    'Insertar eFac Api 
    EnuActFacApi    'Actualizar eFac Api 
    EnuEnvFacApi    'Enviar representación gráfica fras
    EnuInsNDbApi    'Insertar NDb Api 
    EnuActNDbApi    'Actualizar Ndb Api 
    EnuEnvNDbApi    'Enviar representación gráfica NDb
    EnuInsNCrApi    'Insertar NCr Api 
    EnuActNCrApi    'Actualizar Ndb Api
    EnuEnvNCrApi    'Enviar representación gráfica NCr
    EnuInsNConApi   'Insertar NCon de ajuste
    EnuActNConApi    'Actualizar Ndb Api 
    EnuEnvNConApi   'Enviar representación gráfica NCon de ajuste
    EnuInsNRcrApi   'Insertar Nota Rec Cr. APi
    EnuActNRcrApi    'Actualizar NRcr Api 
    EnuEnvNRcrApi    'Enviar representación gráfica NDb
End Enum
Public Enum EnuServidorCorreoDef As Byte
    None = 0
    EnuOutlook
    EnuHotmail
    EnuGmail
    EnuYahoo
    EnuUne
End Enum
Public Enum EnuTipoAccionDef As Byte
    None
    EnuNuevoPais
    EnuNuevoDpto
    EnuNuevoCiud
    EnuModifPais
    EnuModifDpto
    EnuModifCiud
End Enum
Public Enum EnuTipoCambioContrasenaDef As Byte
    None = 0
    EnuNoCambiar
    EnuProximaVez
    EnuAlVencimiento
End Enum
Public Enum EnuTipoConsultaDef As Byte
    None = 0
    EnuSelect
    EnuUpdate
    EnuInsert
    EnuDelete
    EnuExpresionSql
End Enum
Public Enum EnuTipoCuentaDef As Byte
    None = 0
    EnuDebito
    EnuCredito
End Enum
Public Enum EnuTipoDocIdDef As Byte
    None = 0
    EnuCedulaCiudadania
    EnuTarjetaIdentidad
    EnuCedulaExtranjeria
    EnuNit
    EnuNuip
    EnuRegistroCivil
    EnuPasaporte
    EnuDocIdExtranjero
    EnuTarjetaExtranjeria
    EnuNitOtroPais
    EnuPEP
End Enum
Friend Enum EnuTipoLicenciamientoDef As Byte
    None = 0
    EnuPorEstacion
    EnuPorCentroUtilidad
    EnuPorInstancia
End Enum
Friend Enum EnuTipoLocalizacionDef As Byte
    EnuPais = 0
    EnuDpto
    EnuCiudad
End Enum
Public Enum EnuTipoLogDef As Byte
    None = 0
    EnuAccion
    EnuCambio
End Enum
Friend Enum EnuTipoNomTercero As Byte
    None
    EnuMaySostenida
    EnuMayInicial
    EnuComoIngreso
End Enum
<Flags()>
Public Enum EnuPermisosDef As Byte
    None = 0
    EnuConsultar = 1
    EnuCrear = 2
    EnuModificar = 4
    EnuSuprimir = 8
    EnuAnular = 16
    EnuImprimir = 32
    EnuConsCrear = 3
    EnuConCreMod = 7
    EnuConCreModSup = 15
    EnuConCreModAnu = 23
    EnuTodos = 63
    EnuHeredado
End Enum
Public Enum EnuTipoValor As Byte
    EnuNothingNull = 0
    EnuString
    EnuBoolean
    EnuByte
    EnuSByte
    EnuShort
    EnuUShort
    EnuInteger
    EnuUInteger
    EnuLong
    EnuULong
    EnuSingle
    EnuDouble
    EnuDecimal
    EnuDate
    EnuDateTime
    EnuImagen
    EnuArray
    EnuObjeto
End Enum
Friend Enum EnuTipoTerceroDef As Byte
    None = 0
    EnuPersonaNatural
    EnuPersonaJuridica
End Enum
#End Region

#Region "Clases de Propiedad Comunes"
Friend Class ClsAnuladoBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Anulado"

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Anulado"
        HenuTipoValor = EnuTipoValor.EnuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub

    Public Overrides Sub SValide()
        Dim lobjPadre As ClsCBObjetoPan = ObjPadre
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        If HblnEsValido AndAlso HobjValorNew Then
            If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                HblnEsValido = lobjPadre.FblnEsAnulable
            End If
        End If
    End Sub

    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return ClsPanorama.FstrBuleanoToString(HobjValorPro)
    End Function
End Class

Friend Class ClsIdUsuarioAnuloStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdUsuarioAnulo"
    Private ReadOnly MobjPadre As ClsCBObjetoPan = Nothing

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdUsuarioAnulo"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.EnuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub

    Public Overrides Sub SValide()
        HblnEsRequerido = MobjPadre.ObjAnuladoBln.ObjValorPro
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud,
                HblnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                If HblnEsRequerido Then
                    HblnEsValido = (HobjValorNew = GstrIdUsuario)
                End If
            ElseIf MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            Else
                HblnEsValido = String.IsNullOrEmpty(HobjValorNew)
            End If
        End If
    End Sub

    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property

    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsOrigenInstanciaStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "OrigenInstancia"
    Private ReadOnly MobjPadre As ClsCBObjetoPan = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "OrigenInstancia"
        HshrLongitud = 200
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, HshrLongitud,
                HblnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = HobjValorNew = GstrOrigenActual
            Else
                HblnEsValido = HobjValorNew = HobjValorOriginal
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsOrigenInstanciaAnuloStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "OrigenInstanciaAnulo"
    Private ReadOnly MobjPadre As ClsCBObjetoPan = Nothing

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "OrigenInstanciaAnulo"
        HshrLongitud = 200
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub

    Public Overrides Sub SValide()
        HblnEsRequerido = MobjPadre.ObjAnuladoBln.ObjValorPro
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud,
                HblnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                If HblnEsRequerido Then
                    HblnEsValido = (HobjValorNew = GstrOrigenActual)
                End If
            ElseIf MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            Else
                HblnEsValido = String.IsNullOrEmpty(HobjValorNew)
            End If
        End If
    End Sub

    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property

    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsFechaCreacionDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaCreacion"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Fecha Creacion"
        HenuTipoValor = EnuTipoValor.enuDateTime
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim ldtmFechaMin As Date = DateAdd(DateInterval.Day, -1, Date.Now)
        Dim ldtmFechaMax As Date = DateAdd(DateInterval.Day, 1, Date.Now)
        HblnEsValido = IsDate(HobjValorNew)
        If HblnEsValido Then
            If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                        BlnEsRequerido)
            End If
        End If
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        HobjValorPro = GCDTMFECHANULA
        HobjValorNew = HobjValorPro
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return GCDTMFECHANULA
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

#End Region

#Region "Clases Complementarias"
Public Class ClsPanEventArgs
    Inherits EventArgs
    Public Property BlnCancele As Boolean = False
    Public Property BlnVaciandoObjeto As Boolean = False
    Public Property DblCantAProcesar As Double = 0.0
    Public Property DblCantProcesada As Double = 0.0
    Public Property BlnProcesoOk As Boolean = False
    Public Property EnuProceso As EnuProcesoDef = EnuProcesoDef.None

    Friend Sub SLimpie()
        BlnCancele = False
        BlnVaciandoObjeto = False
        DblCantAProcesar = 0
        DblCantProcesada = 0
        BlnProcesoOk = False
        EnuProceso = EnuProcesoDef.None
    End Sub
End Class
' Clase para transportar los argumentos del evento de notificaciones
Public Class ClsNotiEventArgs
    Inherits EventArgs
    Friend Property StrMensaje As String = String.Empty
    Friend Property StrMensajeEx As String = String.Empty
    Friend Property EnuIdMensNot As EnuIdMens = EnuIdMens.None
    Friend Property EnuSevNotifica As EnuSeveridadNot = EnuSeveridadNot.None
    Friend Sub SRegistreNotifica(astrMens As String, astrMensEx As String,
        aenuIdMens As EnuIdMens, aenuSevNotifica As EnuSeveridadNot)
        StrMensaje = astrMens
        StrMensajeEx = astrMensEx
        EnuIdMensNot = aenuIdMens
        EnuSevNotifica = aenuSevNotifica
    End Sub
    Friend Sub SInicialice()
        StrMensaje = String.Empty
        StrMensajeEx = String.Empty
        EnuIdMensNot = EnuIdMens.None
        EnuSevNotifica = EnuSeveridadNot.None
    End Sub
End Class
#End Region