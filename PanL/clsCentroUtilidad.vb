Imports System.Drawing
Friend Class ClsCentroUtilidad
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "PanCentrosUtilidad"
    ' Firma
    ' Coleccion
    '
    Private MobjTerceroCentroUtilidad As ClsTercero = Nothing
    Private MobjTerceroRepLegal As ClsTercero = Nothing
    Private MdtbServidoresCorreo As DataTable = Nothing
#End Region

#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Copropiedad.
    ''' </summary>
    ''' <param name="aenuModoInstanciaObj">Esta clase no sepuede instanciar como Navegable!</param>
    ''' <remarks>Si se instancia como un objeto navegable, se crea un datatable que contiene las columnas de
    ''' la llave con las llaves de todos los objetos y queda a la espera de que se indique que objeto abrir.
    ''' Si se instancia como un objeto único, queda a la espera de recibir el valor de los campos de la llave 
    ''' para abrir dicho objeto. </remarks>
    Public Sub New(aobjPadre As ClsCarpeta, aenuModoInstanciaObj As EnuModoInstanciaObjDef)
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuDeColeccion Then
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        End If
        Dim lstrCamposSelect As String()
        HobjPadre = aobjPadre
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuNavegable Then
            HblnEsAnulable = False
            lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil}
            HcolFiltros.Add(ClsIdCarpetaCenUtilShr.SstrNombreCampoBd & " = " &
                        aobjPadre.ObjIdCarpetaShr.ObjValorPro)
        Else
            HblnEsCreable = False
            HblnEsModificable = False
            HblnEsSuprimible = False
            HblnEsAnulable = False
            HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
            lstrCamposSelect = {"*"}
        End If
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
    End Sub
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Carpeta al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsCarpeta, adrwObjeto As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        '
        DrwRegistroActual = adrwObjeto
        DtbTablaColeccion = DrwRegistroActual.Table
    End Sub
#End Region

#Region "Propiedades"
#Region "Propiedades identificadoras"
    Protected Overrides ReadOnly Property HstrNombreTabla As String
        Get
            Return MCSTRNOMBRETABLA
        End Get
    End Property
    Friend Shared ReadOnly Property SstrNombreTabla As String
        Get
            Return MCSTRNOMBRETABLA
        End Get
    End Property
    Protected Friend Overrides ReadOnly Property HenuIdClase As EnuIdClasesPanDef
        Get
            Return EnuIdClasesPanDef.EnuCentroUtilidad
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Copropiedad"
        End Get
    End Property
#End Region

#Region "Propiedades Prop"
    Friend ReadOnly Property ObjCargoStr As New ClsCargoStr(Me)
    Friend ReadOnly Property ObjContrasenaEMailStr As New ClsContrasenaEMailStr(Me)
    Friend ReadOnly Property ObjEmailOrigenStr As New ClsEmailOrigenStr(Me)
    Friend ReadOnly Property ObjEstadoContratoByt As New ClsEstadoContratoByt(Me)
    Friend ReadOnly Property ObjEstaLicenciadaBln As New ClsEstaLicenciadaBln(Me)
    Friend ReadOnly Property ObjFechaVenceContratoDtm As New ClsFechaVenceContratoDtm(Me)
    Friend ReadOnly Property ObjFechaVerificoContratoDtm As New ClsFechaVerificoContratoDtm(Me)
    Friend ReadOnly Property ObjHabilitarSslBln As New ClsHabilitarSslBln(Me)
    Friend ReadOnly Property ObjIdCarpetaCenUtilShr As New ClsIdCarpetaCenUtilShr(Me)
    Friend ReadOnly Property ObjIdCentroUtilShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdContratoEnt As New ClsIdContratoEnt(Me)
    Friend ReadOnly Property ObjIdTerceroCentroUtilDbl As New ClsIdTerceroCentroUtilDbl(Me)
    Friend ReadOnly Property ObjIdTerceroRepLegalDbl As New ClsIdTerceroRepLegalDbl(Me)
    Friend ReadOnly Property ObjIntervaloEntreTandasShr As New ClsIntervaloEntreTandasShr(Me)
    Friend ReadOnly Property ObjMensajesPorTandaShr As New ClsMensajesPorTandaShr(Me)
    Friend ReadOnly Property ObjNombreCentroUtilStr As New ClsNombreCentroUtilStr(Me)
    Friend ReadOnly Property ObjPuertoHostShr As New ClsPuertoHostShr(Me)
    Friend ReadOnly Property ObjRequiereAutenticacionBln As New ClsRequiereAutenticacionBln(Me)
    Friend ReadOnly Property ObjServidorSmtpStr As New ClsServidorSmtpStr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjCargoStr)
                HcolPropiedades.Add(ObjContrasenaEMailStr)
                HcolPropiedades.Add(ObjEmailOrigenStr)
                HcolPropiedades.Add(ObjEstadoContratoByt)
                HcolPropiedades.Add(ObjEstaLicenciadaBln)
                HcolPropiedades.Add(ObjFechaVenceContratoDtm)
                HcolPropiedades.Add(ObjFechaVerificoContratoDtm)
                HcolPropiedades.Add(ObjHabilitarSslBln)
                HcolPropiedades.Add(ObjIdCarpetaCenUtilShr)
                HcolPropiedades.Add(ObjIdCentroUtilShr)
                HcolPropiedades.Add(ObjIdContratoEnt)
                HcolPropiedades.Add(ObjIdTerceroCentroUtilDbl)
                HcolPropiedades.Add(ObjIdTerceroRepLegalDbl)
                HcolPropiedades.Add(ObjIntervaloEntreTandasShr)
                HcolPropiedades.Add(ObjMensajesPorTandaShr)
                HcolPropiedades.Add(ObjNombreCentroUtilStr)
                HcolPropiedades.Add(ObjPuertoHostShr)
                HcolPropiedades.Add(ObjRequiereAutenticacionBln)
                HcolPropiedades.Add(ObjServidorSmtpStr)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region

#Region "Otras propiedades"
    ''' <summary>
    ''' Devuelve el objeto tercero correspondiente al Copropiedad
    ''' </summary>
    ''' <returns>Objeto Tercero de la Copropiedad</returns>
    ''' <remarks></remarks>
    ''' 
    Friend ReadOnly Property ObjTerceroCentroUtilidad As ClsTercero
        Get
            If IsNothing(MobjTerceroCentroUtilidad) Then
                If ObjIdTerceroCentroUtilDbl.BlnEsValido Then
                    Dim lobjIdLlave() As Object = {ObjIdTerceroCentroUtilDbl.ObjValorPro}
                    MobjTerceroCentroUtilidad = New ClsTercero(EnuModoInstanciaObjDef.enuUnico)
                    MobjTerceroCentroUtilidad.SAbra(lobjIdLlave)
                End If
            End If
            Return MobjTerceroCentroUtilidad
        End Get
    End Property

    Friend Property ObjTerRepLegal As ClsTercero
        Get
            If MobjTerceroRepLegal Is Nothing OrElse Not MobjTerceroRepLegal.BlnExiste Then
                Dim ldblIdRepLegal As Double = ObjIdTerceroRepLegalDbl.ObjValorPro
                Dim lobjValorLlave As Object() = {ldblIdRepLegal}
                MobjTerceroRepLegal = New ClsTercero(EnuModoInstanciaObjDef.EnuUnico)
                MobjTerceroRepLegal.SAbra(lobjValorLlave)
            End If
            Return MobjTerceroRepLegal
        End Get
        Set(value As ClsTercero)
            If value IsNot Nothing Then
                MobjTerceroRepLegal = value
            Else
                MobjTerceroRepLegal = Nothing
            End If
        End Set
    End Property

    ''' <summary>
    ''' Devuelve la dirección de la Copropiedad seguida de la ciudad
    ''' </summary>
    ''' <returns></returns>
    Friend ReadOnly Property StrDireccion As String
        Get
            Dim lstrNomDep As String, lstrNomCiu As String
            Dim lstrDir = ObjTerceroCentroUtilidad.ObjDireccionUnoStr.ObjValorPro
            Dim lstrIdPaisDir = ObjTerceroCentroUtilidad.ObjPaisDirStr.ObjValorPro
            Dim lentIdDpto As Integer = ObjTerceroCentroUtilidad.ObjDepartamentoDirByt.ObjValorPro
            Dim lentIdCiu As Integer = ObjTerceroCentroUtilidad.ObjCiudadDirShr.ObjValorPro
            Dim lobjUbicacion = New Ubicacion.ClsUbicacion
            lstrNomDep = lobjUbicacion.StrNombreDpto(lstrIdPaisDir, lentIdDpto)
            lstrNomCiu = lobjUbicacion.StrNombreCiudad(lstrIdPaisDir, lentIdDpto, lentIdCiu)
            lstrDir &= " - " & lstrNomCiu & "(" & lstrNomDep & ")"
            Return lstrDir
        End Get
    End Property
#End Region
#End Region

#Region "Procedimientos y funciones invalidantes"
    Protected Friend Overrides Sub SVacie()
        MyBase.SVacie()
        MobjTerceroCentroUtilidad = Nothing
        MobjTerceroRepLegal = Nothing
    End Sub
    Friend Overrides Function FblnEsSuprimible() As Boolean
        Dim lstrNombreColumnas As String() = {ObjIdCarpetaCenUtilShr.StrNombreCampoBD,
                ObjIdCentroUtilShr.StrNombreCampoBD}
        Dim lstrCondiciones As String() = {" = " & ObjIdCarpetaCenUtilShr.ToString, " = " &
                ObjIdCentroUtilShr.ToString}
        Dim lstrTablasExcluir As String() = {SstrNombreTabla, ClsLogApp.SstrNombreTabla}
        Dim lblnEsSuprimible As Boolean = FblnPermitidoSuprimir()
        If lblnEsSuprimible Then
            lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg(lstrTablasExcluir,
                lstrNombreColumnas, lstrCondiciones, True, False)
            If lblnEsSuprimible Then
                lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg(lstrTablasExcluir,
                    lstrNombreColumnas, lstrCondiciones, True, True)
            End If
        End If
        Return lblnEsSuprimible
    End Function
    Protected Friend Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Dim lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                Dim lobjCarpeta As ClsCarpeta = HobjPadre
                Dim lshrIdCarpeta As Short = lobjCarpeta.ObjIdCarpetaShr.ObjValorPro
                ObjIdCarpetaCenUtilShr.ObjValorPro = lshrIdCarpeta
                ObjEstadoContratoByt.ObjValorPro = EnuEstadoContrato.EnuInstalando
                ObjEstaLicenciadaBln.ObjValorPro = False
                ObjFechaVenceContratoDtm.ObjValorPro = GCDTMFECHANULA
                SNumereObj()
                MyBase.SActualice(ablnExigeRequeridos)
                lblnNoHayError = True
            ElseIf EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                MyBase.SActualice(ablnExigeRequeridos)
            End If
            lblnNoHayError = True
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdCentroUtilShr.ToString
        End Get
    End Property
#End Region

#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Dim lshrIdCentroUtil As Short
            Dim lstrFiltro As String = ObjIdCarpetaCenUtilShr.StrNombreCampoBD & " = " &
                    ObjIdCarpetaCenUtilShr.ObjValorPro
            lshrIdCentroUtil = ClsPanorama.FobjUltimaIdNumericaObjeto(SstrNombreTabla,
                    ObjIdCentroUtilShr.StrNombreCampoBD, ObjIdCentroUtilShr.EnuTipoValor,
                    lstrFiltro) + 1
            ObjIdCentroUtilShr.ObjValorPro = lshrIdCentroUtil
        End If
    End Sub
#End Region

#Region "Contrato"
    Friend Sub SActuliceAuriga()
        Dim ldtbAuriga = ClsPanorama.FdtbAuriga(ObjIdTerceroCentroUtilDbl.ToString)
        ClsPanorama.SRegistreIngreso(ObjIdTerceroCentroUtilDbl.ToString)
        If ldtbAuriga IsNot Nothing AndAlso ldtbAuriga.Rows.Count > 0 Then
            Dim ldrwAuriga As DataRow = ldtbAuriga(0)
            Dim lenuEstadoCon As EnuEstadoContrato = ClsPanorama.FobjValorCampo(ldrwAuriga(0),
                    EnuTipoValor.enuByte)
            Dim ldtmFecVence As Date = ClsPanorama.FobjValorCampo(ldrwAuriga(1),
                    EnuTipoValor.enuDate)
            Dim lentIdCont As Integer = ClsPanorama.FobjValorCampo(ldrwAuriga(2),
                    EnuTipoValor.enuInteger)
            Dim lblnEsLicencia As Boolean = ClsPanorama.FobjValorCampo(ldrwAuriga(4),
                    EnuTipoValor.enuBoolean)
            EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
            ObjEstaLicenciadaBln.ObjValorPro = lblnEsLicencia
            ObjFechaVenceContratoDtm.ObjValorPro = ldtmFecVence
            ObjFechaVerificoContratoDtm.ObjValorPro = Date.Today
            ObjEstadoContratoByt.ObjValorPro = lenuEstadoCon
            ObjIdContratoEnt.ObjValorPro = lentIdCont
            SActualice(True)
        End If
    End Sub
    Friend Sub SRegistreNit(alngNIT As Long)
        ClsPanorama.SRegistreNit(alngNIT)
    End Sub
#End Region

#Region "Correo"
    Friend Function FstrNombreServidoresCorreo() As ArrayList
        SCargueDtbServidoresCorreo()
        Dim lstrServidores As New ArrayList
        Dim ldrwServidores As DataRow() = MdtbServidoresCorreo.Select()
        lstrServidores.Add(My.Resources.Ninguno)
        For Each ldrwServidor As DataRow In ldrwServidores
            lstrServidores.Add(ldrwServidor("Nombre"))
        Next
        lstrServidores.Add(My.Resources.Otro)
        Return lstrServidores
    End Function
    Friend Function FstrHostCorreo(aenuServidor As EnuServidorCorreoDef) As String
        Dim lstrHost = String.Empty
        Dim lstrFiltro = "IdServidor = " & aenuServidor
        SCargueDtbServidoresCorreo()
        Dim ldrwServidores As DataRow() = MdtbServidoresCorreo.Select(lstrFiltro)
        If ldrwServidores.Length > 0 Then
            Dim ldrwSer As DataRow = ldrwServidores(0)
            lstrHost = ClsPanorama.FobjValorCampo(ldrwSer("Host"), EnuTipoValor.enuString)
        End If
        Return lstrHost
    End Function
    Friend Function FentPuertoHost(aenuServidor As EnuServidorCorreoDef) As Integer
        Dim lentPuerto = 0
        Dim lstrFiltro = "IdServidor = " & aenuServidor
        SCargueDtbServidoresCorreo()
        Dim ldrwServidores As DataRow() = MdtbServidoresCorreo.Select(lstrFiltro)
        If ldrwServidores.Length > 0 Then
            Dim ldrwSer As DataRow = ldrwServidores(0)
            lentPuerto = ClsPanorama.FobjValorCampo(ldrwSer("Puerto"), EnuTipoValor.enuInteger)
        End If
        Return lentPuerto
    End Function
    Friend Function FblnRequiereAutenticacion(aenuServidor As EnuServidorCorreoDef) As Boolean
        Dim lblnReqAut = False
        Dim lstrFiltro = "IdServidor = " & aenuServidor
        SCargueDtbServidoresCorreo()
        Dim ldrwServidores As DataRow() = MdtbServidoresCorreo.Select(lstrFiltro)
        If ldrwServidores.Length > 0 Then
            Dim ldrwSer As DataRow = ldrwServidores(0)
            lblnReqAut = ClsPanorama.FobjValorCampo(ldrwSer("RequiereAutenticacion"), EnuTipoValor.enuString)
        End If
        Return lblnReqAut
    End Function
    Friend Function FblnHabilitaSSL(aenuServidor As EnuServidorCorreoDef) As Boolean
        Dim lblnHabilitaSSL = False
        Dim lstrFiltro = "IdServidor = " & aenuServidor
        SCargueDtbServidoresCorreo()
        Dim ldrwServidores As DataRow() = MdtbServidoresCorreo.Select(lstrFiltro)
        If ldrwServidores.Length > 0 Then
            Dim ldrwSer As DataRow = ldrwServidores(0)
            lblnHabilitaSSL = ClsPanorama.FobjValorCampo(ldrwSer("HabilitaSSL"), EnuTipoValor.enuBoolean)
        End If
        Return lblnHabilitaSSL
    End Function
    Private Sub SCargueDtbServidoresCorreo()
        If IsNothing(MdtbServidoresCorreo) Then
            GobjPanDat.SControleProcesoObj(True)
            MdtbServidoresCorreo = ClsPanorama.FdtbDataTable("PanTblServidoresCorreo", {"*"},
                        {{"IdServidor", "ASC"}}, "")
            GobjPanDat.SControleProcesoObj(False)
        End If
    End Sub
    Friend Function FstrContrasena()
        Dim lstrContr = String.Empty
        If ObjContrasenaEMailStr.BlnEsValido Then
            lstrContr = ClsPanorama.FstrContrasena(ObjContrasenaEMailStr)
        End If
        Return lstrContr
    End Function
    Friend Function FblnExisteCtaCorreoValida() As Boolean
        Return ObjContrasenaEMailStr.ToString.Length > 0 AndAlso
                ObjEmailOrigenStr.ToString.Length > 0
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsCargoStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Cargo"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Cargo"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud,
                BlnEsRequerido)
        If HblnEsValido Then
            HobjValorNew = HobjValorNew.ToString().ToUpper()
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
Friend Class ClsContrasenaEMailStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ContrasenaCorreo"
    Private ReadOnly MobjPadre As ClsCentroUtilidad = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Contasena Correo"
        HshrLongitud = 50
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsRequerido = (Not IsNothing(MobjPadre.ObjEmailOrigenStr.ObjValorPro) AndAlso
                           MobjPadre.ObjEmailOrigenStr.ToString.Length > 0)
        If HblnEsRequerido Then
            HblnEsRequerido = MobjPadre.ObjRequiereAutenticacionBln.ObjValorPro
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 6, HshrLongitud, HblnEsRequerido)
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
Friend Class ClsEmailOrigenStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "EmailOrigen"
    Private ReadOnly MobjPadre As ClsCentroUtilidad = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Email de Origen"
        HshrLongitud = 200
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = (ClsPanorama.FblnEsValidoString(HobjValorNew, 5, ShrLongitud, BlnEsRequerido))
        If HblnEsValido And Not String.IsNullOrEmpty(HobjValorNew) Then
            Dim lstrExpValidaMail As String = "^([\w-\.]+)@((\[[0-9]{1,3}\." &
                    "[0-9]{1,3}\.)|(([\w-]+\.)+))" &
                    "([a-zA-z]{2,4}|[0-9]{1,3})(\]?)$"
            HblnEsValido = System.Text.RegularExpressions.Regex.IsMatch(HobjValorNew, lstrExpValidaMail)
            If Not HblnEsValido Then
                HstrMens = "La dirección de correo electrónico no es valida"
                SNotifiqueDatInv()
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
    Private Sub ClsEmailOrigenStr_evnPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            MobjPadre.ObjContrasenaEMailStr.SValide()
        End If
    End Sub
End Class
Friend Class ClsEstadoContratoByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdEstadoContrato"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Estado del contrato"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = (ClsPanorama.FblnEsValidoEnumByte(HobjValorNew,
                EnuEstadoContrato.EnuInstalando,
                EnuEstadoContrato.EnuRegistradoNit, BlnEsRequerido))
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        Dim lstrTipopDoc = String.Empty
        If Not IsNothing(HobjValorPro) Then
            Dim lenuTipoDoc As EnuTipoDocIdDef = CType(HobjValorPro, Byte)
            Select Case lenuTipoDoc
                Case EnuEstadoContrato.EnuInstalando
                    lstrTipopDoc = "Instalando"
                Case EnuEstadoContrato.EnuActivo
                    lstrTipopDoc = "Activo"
                Case EnuEstadoContrato.EnuSuspendido
                    lstrTipopDoc = "Suspendido"
                Case EnuEstadoContrato.EnuInactivo
                    lstrTipopDoc = "Inactivo"
            End Select
        End If
        Return lstrTipopDoc
    End Function
End Class
Friend Class ClsEstaLicenciadaBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "EstaLicenciada"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Está licenciada"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsFechaVenceContratoDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaVenceContrato"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Fecha Vence contrato"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, GCDTMFECHANULA,
                    GCDTMFECHAMAXI, BlnEsRequerido)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToShortDateString
    End Function
End Class
Friend Class ClsFechaVerificoContratoDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaVerificoContrato"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Fecha verificó contrato"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, GCDTMFECHANULA,
                    GCDTMFECHAMAXI, BlnEsRequerido)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToShortDateString
    End Function
End Class
Friend Class ClsHabilitarSslBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "HbilitarSsl"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Habilitar SSL-TLS"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsIdCarpetaCenUtilShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCarpeta"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdCarpeta"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 0
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsCentroUtilidad = ObjPadre
        Dim lobjAbuelo As ClsCarpeta = lobjPadre.ObjPadre
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                    BlnEsRequerido, EnuTipoValor)
        If Not BlnLeyendoOrigen Then
            If HblnEsValido Then
                HblnEsValido = (HobjValorNew = lobjAbuelo.ObjIdCarpetaShr.ObjValorPro)
            Else
                HstrMens = "La Id. de la Carpeta ingresada no es válida!"
                SNotifiqueDatInv()
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
Friend Class ClsIdCentroUtilShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCentroUtil"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdCentroUtil"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 1
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsCentroUtilidad = ObjPadre
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                BlnEsRequerido, EnuTipoValor)
        If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
            HstrMens = String.Empty
            If Not BlnLeyendoOrigen Then
                If HblnEsValido Then
                    HblnEsValido = lobjPadre.ObjIdCarpetaCenUtilShr.BlnEsValido
                    If HblnEsValido Then
                        Dim lobjLlavePrincipal() = {lobjPadre.ObjIdCarpetaCenUtilShr.ObjValorPro, HobjValorNew}
                        If Not lobjPadre.FblnExisteLlave(lobjLlavePrincipal) Then
                            HstrMens = "La Id. de la Copropiedad ingresada no existe!"
                            HblnEsValido = False
                        End If
                    End If
                Else
                    HstrMens = "La Id. de la Copropiedad ingresada no es válida!"
                End If
            End If
            If Not String.IsNullOrEmpty(HstrMens) Then
                SNotifiqueDatInv()
            End If
        Else
            HblnEsValido = True
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
Friend Class ClsIdContratoEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdContrato"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Id Contarto"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = False
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsCentroUtilidad = ObjPadre
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue,
                BlnEsRequerido, EnuTipoValor)
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
Friend Class ClsIdTerceroCentroUtilDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTerceroCentroUtilidad"
    Private MstrNombreTercero As String = String.Empty
    Private MblnExisteTercero = False
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdTerceroCentroUtilidad"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = "IdTerceroCentroUtilidad"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC, GCDBLMAXTERC,
                HblnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            MblnExisteTercero = ClsPanorama.FblnExisteTercero(HobjValorNew)
            HblnEsValido = MblnExisteTercero
            If HblnEsValido Then
                Dim lobjTercero As New ClsTercero(EnuModoInstanciaObjDef.enuUnico)
                lobjTercero.SAbra({HobjValorNew})
                MstrNombreTercero = lobjTercero.FstrNombreCompleto()
            Else
                MstrNombreTercero = String.Empty
            End If
        End If
    End Sub
    Friend ReadOnly Property BlnExisteTercero As Boolean
        Get
            Return MblnExisteTercero
        End Get
    End Property
    Friend ReadOnly Property StrNombreTerceroCentroutil
        Get
            If BlnEsValido Then
                Return MstrNombreTercero
            Else
                Return ""
            End If
        End Get
    End Property
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(ObjValorPro) Then
            Return HobjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsIdTerceroRepLegalDbl
    Inherits ClsCBPropiedad
    Private MstrNombreTercero As String = String.Empty
    Private MblnExisteTercero = False
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdTerceroRepresentanteLegal"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = "IdTerceroRepresentanteLegal"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC, GCDBLMAXTERC,
                HblnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            Dim lobjPadre As ClsCentroUtilidad = ObjPadre
            MblnExisteTercero = ClsPanorama.FblnExisteTercero(HobjValorNew)
            HblnEsValido = MblnExisteTercero
            If HblnEsValido Then
                Dim lobjTercero As New ClsTercero(EnuModoInstanciaObjDef.EnuUnico)
                lobjTercero.SAbra({HobjValorNew})
                MstrNombreTercero = lobjTercero.FstrNombreCompleto()
                lobjPadre.ObjTerRepLegal = lobjTercero
            Else
                lobjPadre.ObjTerRepLegal = Nothing
                MstrNombreTercero = String.Empty
            End If
        End If
    End Sub
    Friend ReadOnly Property BlnExisteTercero As Boolean
        Get
            Return MblnExisteTercero
        End Get
    End Property
    Friend ReadOnly Property StrNombreRepLegal
        Get
            If BlnEsValido Then
                Return MstrNombreTercero
            Else
                Return ""
            End If
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(ObjValorPro) Then
            Return HobjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsIntervaloEntreTandasShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IntervaloEntreTandas"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Intervalo entre Tandas"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, 600,
                        BlnEsRequerido, HenuTipoValor)
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
Friend Class ClsMensajesPorTandaShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "MensajesPorTanda"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Mensajes Por Tanda"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Short.MaxValue,
                        BlnEsRequerido, HenuTipoValor)
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
Friend Class ClsNombreCentroUtilStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Nombre"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "NombreCentroUtilidad"
        HshrLongitud = 50
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3,
                ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            HobjValorNew = FstrNombreTercero(HobjValorNew.ToString())
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(ObjValorPro) Then
            Return HobjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsPuertoHostShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PuertoServidor"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Puerto Servidor"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                BlnEsRequerido, EnuTipoValor)
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
Friend Class ClsRequiereAutenticacionBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "RequiereAutenticacion"
    Private ReadOnly MobjPadre As ClsCentroUtilidad = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Requiere Autenticacion"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Private Sub ClsRequiereAutenticacionBln_evnPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            MobjPadre.ObjEmailOrigenStr.SValide()
            MobjPadre.ObjContrasenaEMailStr.SValide()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsServidorSmtpStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ServidorCorreo"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Servidor de Correo Saliente"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = False
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 6, ShrLongitud, BlnEsRequerido)
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
#End Region