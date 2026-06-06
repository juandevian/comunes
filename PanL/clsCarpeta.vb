Friend Class ClsCarpeta
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "PanCarpetas"
    ' Objetos de propiedad
    Private ReadOnly MobjEstaActivaBln As New ClsEstaActivaBln(Me)
    Private ReadOnly MobjIdCarpetaShr As New ClsIdCarpetaShr(Me)
    Private ReadOnly MobjNombreStr As New ClsNombreStr(Me)
    ' Variables de modulo
    Private MobjCentroUtilidadNuevo As ClsCentroUtilidad = Nothing
    Private MobjCentroUtilidadActual As ClsCentroUtilidad = Nothing
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Carpeta.
    ''' </summary>
    ''' <param name="aenuModoInstancia">Indica si se instancia como un objeto navegable o como un Objeto único.</param>
    ''' <remarks>Si se instancia como un objeto navegable, se crea un datatable que contiene las columnas de
    ''' la llave con las llaves de todos los objetos y queda a la espera de que se indique que objeto abrir.
    ''' Si se instancia como un objeto único, queda a la espera de recibir el valor de los campos de la llave 
    ''' para abrir dicho objeto. </remarks>
    Public Sub New(aenuModoInstancia As EnuModoInstanciaObjDef, ablnSoloActivas As Boolean)
        If aenuModoInstancia = EnuModoInstanciaObjDef.enuDeColeccion Then
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        End If
        Dim lstrCamposSelect As String()
        HobjPadre = GobjAdministrador
        HblnEsAnulable = False
        If aenuModoInstancia = EnuModoInstanciaObjDef.enuNavegable Then
            lstrCamposSelect = {ObjIdCarpetaShr.StrNombreCampoBD}
            Dim lstrFiltro = ""
            If ablnSoloActivas Then
                lstrFiltro = ClsEstaActivaBln.SstrNombreCampoBd & " = TRUE"
            End If
            HcolFiltros.Add(lstrFiltro)
        Else
            HblnEsCreable = False
            HblnEsModificable = False
            HblnEsSuprimible = False
            HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
            lstrCamposSelect = {"*"}
        End If
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
    End Sub
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwCarpeta">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsAdministrador, adrwCarpeta As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        '
        DrwRegistroActual = adrwCarpeta
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
            Return EnuIdClasesPanDef.enuCarpeta
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Carpeta"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjEstaActivaBln As ClsEstaActivaBln
        Get
            Return MobjEstaActivaBln
        End Get
    End Property
    Friend ReadOnly Property ObjIdCarpetaShr As ClsIdCarpetaShr
        Get
            Return MobjIdCarpetaShr
        End Get
    End Property
    Friend ReadOnly Property ObjNombreStr As ClsNombreStr
        Get
            Return MobjNombreStr
        End Get
    End Property
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(MobjEstaActivaBln)
                HcolPropiedades.Add(MobjIdCarpetaShr)
                HcolPropiedades.Add(MobjNombreStr)
                HcolPropiedades.Add(ObjFechaCreacionDtm)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    ''' <summary>
    ''' Devuelve el objeto Cuenta de Contabilidad perteneciente a la Carpeta e identificado por el argumento 
    ''' 'astrIdCuentaContabilidad'. Si no hay Cuentas de Contabilidad devuelve un objeto vacio. Si el argumento 
    ''' 'astrIdCuentaContabilidad' es una cadena vacia devuelve la primer Cuenta de Contabilidad y si no existe 
    ''' el objeto solicitado devuelve Nothing.
    ''' </summary>
    ''' <param name="astrIdCuentaContabilidad">Identifica la Cuenta de Contabilidad solicitada. Si este argumento 
    ''' es una cadena vacia se devuelve la primer Cuenta de Contabilidad de la Colección.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend ReadOnly Property ObjCuentaContabilidad(astrIdCuentaContabilidad As String) As ClsCuentaContabilidad
        Get
            Dim lobjCtaCont As ClsCuentaContabilidad = Nothing
            Dim ldrwCtaCon As DataRow = Nothing
            Dim lblnAbrirObjeto = False
            Dim ldtbCtasCon = FdtbCuentasCont()
            If ldtbCtasCon.Rows.Count = 0 Then
                ldrwCtaCon = ldtbCtasCon.NewRow
                lblnAbrirObjeto = True
            Else
                If String.IsNullOrEmpty(astrIdCuentaContabilidad) Then
                    ldrwCtaCon = ldtbCtasCon.Rows(0)
                    lblnAbrirObjeto = True
                Else
                    Dim lstrFiltro = ClsIdCuentaContStr.SstrNombreCampoBd & " = '" &
                            astrIdCuentaContabilidad & "'"
                    Dim ldrwCtasCont() = ldtbCtasCon.Select(lstrFiltro)
                    If ldrwCtasCont.Length > 0 Then
                        ldrwCtaCon = ldrwCtasCont(0)
                        lblnAbrirObjeto = True
                    End If
                End If
            End If
            If lblnAbrirObjeto Then
                lobjCtaCont = New ClsCuentaContabilidad(Me, ldrwCtaCon)
                lobjCtaCont.SLeaValores(True)
            End If
            Return lobjCtaCont
        End Get
    End Property
    ''' <summary>
    ''' Devuelve una nueva Cuenta de Contabilidad en estado de creación, con los valores de
    ''' las propiedades iguales a los valores por defecto.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend ReadOnly Property ObjNuevaCuentaContabilidad As ClsCuentaContabilidad
        Get
            Dim ldtbCuentasCont = FdtbCuentasCont()
            Dim ldrwCtaCont As DataRow = ldtbCuentasCont.NewRow
            Dim lobjCuentaCont As New ClsCuentaContabilidad(Me, ldrwCtaCont)
            lobjCuentaCont.SCreeObj(Nothing)
            lobjCuentaCont.ObjIdCarpetaCuentaShr.ObjValorPro = ObjIdCarpetaShr.ObjValorPro
            Return lobjCuentaCont
        End Get
    End Property
    Friend ReadOnly Property StrIdCtaPrimera As String
        Get
            Dim lstrIdCtaConPrimera = String.Empty
            Dim lobjCtaCon As ClsCuentaContabilidad
            Dim lcolCuentasCon = FcolCuentasCont()
            If lcolCuentasCon.Count > 0 Then
                lobjCtaCon = lcolCuentasCon(1)
                lstrIdCtaConPrimera = lobjCtaCon.ObjIdCuentaContStr.ObjValorPro
            End If
            Return lstrIdCtaConPrimera
        End Get
    End Property
    Friend ReadOnly Property StrIdCtaUltima As String
        Get
            Dim lstrIdCtaConUltima = String.Empty
            Dim lobjCtaCon As ClsCuentaContabilidad
            Dim lcolCuentasCon = FcolCuentasCont()
            If lcolCuentasCon.Count > 0 Then
                lobjCtaCon = lcolCuentasCon(lcolCuentasCon.Count)
                lstrIdCtaConUltima = lobjCtaCon.ObjIdCuentaContStr.ObjValorPro
            End If
            Return lstrIdCtaConUltima
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Friend Overrides Sub SVacie()
        MyBase.SVacie()
        MobjCentroUtilidadNuevo = Nothing
    End Sub
    Protected Friend Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            GobjPanDat.SControleProcesoObj(True)
            Try
                SNumereObj()
                ObjFechaCreacionDtm.ObjValorPro = Date.Today
                MyBase.SActualice(ablnExigeRequeridos)
            Catch ex As PanLException
                Throw
            Catch ex As PanDatException
                Throw
            Catch ex As ArgumentNullException
                Throw
            Catch ex As Exception
                Throw
            Finally
                GobjPanDat.SControleProcesoObj(False)
            End Try
        Else
            MyBase.SActualice(ablnExigeRequeridos)
        End If
    End Sub
    Friend Overrides Function FblnEsSuprimible() As Boolean
        Dim lblnEsSuprimible As Boolean = False
        ' La condicion no debe incluir el nombre de de la columna
        Dim lstrCondicion As String = " = " & ObjIdCarpetaShr.ObjValorPro
        Dim lstrTablasExcluir As String() = {SstrNombreTabla, ClsLogApp.SstrNombreTabla,
                ClsCarpetaUsuario.SstrNombreTabla}
        If BlnEsSuprimible AndAlso BlnExiste Then
            lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg(lstrTablasExcluir,
                    ObjIdCarpetaShr.StrNombreCampoBD, lstrCondicion, False, True)
            If lblnEsSuprimible Then
                lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg(lstrTablasExcluir,
                        ObjIdCarpetaShr.StrNombreCampoBD, lstrCondicion, False, False)
            End If
        End If
        Return lblnEsSuprimible
    End Function
    Protected Friend Overrides Function FblnSuprimio() As Boolean
        Dim lblnNoHayError = False, lblnsuprimio = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            SSuprimaCarpetasUsuarios()
            lblnsuprimio = MyBase.FblnSuprimio()
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
        Return lblnsuprimio
    End Function
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdCarpetaShr.ToString
        End Get
    End Property
#End Region
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Dim lshrIdCarpeta As Short
            lshrIdCarpeta = ClsPanorama.FobjUltimaIdNumericaObjeto(SstrNombreTabla,
                        ObjIdCarpetaShr.StrNombreCampoBD, ObjIdCarpetaShr.EnuTipoValor, "") + 1
            ObjIdCarpetaShr.ObjValorPro = lshrIdCarpeta
        End If
    End Sub
    Friend Overrides Function FblnNotificaOk(aenuIdMensNot As EnuIdMens) As Boolean
        Dim lblnOk = False
        If aenuIdMensNot = EnuIdMens.EnuDatCarInvalidos Then
            lblnOk = FblnEstanTodosOk()
        End If
        Return lblnOk
    End Function
    Private Sub SSuprimaCarpetasUsuarios()
        Dim lstrTabla = ClsCarpetaUsuario.SstrNombreTabla
        Dim lcolCamposRef As New Collection
        Dim lcolDatosRef As New Collection
        lcolCamposRef.Add(ClsIdCarpetaUsuarioShr.SstrNombreCampoBd)
        lcolDatosRef.Add(MobjIdCarpetaShr.ObjValorPro, StrCampoCarpeta)
        Dim lentRegEliminados = GobjPanDat.SElimineRegistro(lstrTabla, lcolCamposRef, lcolDatosRef)
    End Sub
#End Region
#Region "Manejo Copropiedad"
    Friend Function FcolCentrosUtilidad() As Collection
        Dim ldtbCenUtil = FdtbCentrosUtilidad()
        Dim lcolCenUtil As New Collection, lstrKey As String
        If ldtbCenUtil.Rows.Count > 0 Then
            Dim lobjCentroUtilidad As ClsCentroUtilidad
            For Each ldrwCenUti As DataRow In ldtbCenUtil.Rows
                lobjCentroUtilidad = New ClsCentroUtilidad(Me, ldrwCenUti)
                lobjCentroUtilidad.SLeaValores(True)
                lstrKey = lobjCentroUtilidad.ObjIdCarpetaCenUtilShr.ToString &
                        lobjCentroUtilidad.ObjIdCentroUtilShr.ToString
                lcolCenUtil.Add(lobjCentroUtilidad, lstrKey)
            Next
        End If
        Return lcolCenUtil
    End Function
    Friend ReadOnly Property ObjNuevoCentroUtilidad As ClsCentroUtilidad
        Get
            Dim ldtbCenUtil = FdtbCentrosUtilidad()
            MobjCentroUtilidadNuevo = Nothing
            Dim ldrwCentroUtilidadNuevo As DataRow
            ldrwCentroUtilidadNuevo = ldtbCenUtil.NewRow
            MobjCentroUtilidadNuevo = New ClsCentroUtilidad(Me, ldrwCentroUtilidadNuevo)
            MobjCentroUtilidadNuevo.SCreeObj(Nothing)
            MobjCentroUtilidadNuevo.ObjIdCarpetaCenUtilShr.ObjValorPro = ObjIdCarpetaShr.ObjValorPro
            Return MobjCentroUtilidadNuevo
        End Get
    End Property
    ''' <summary>
    ''' Actualiza la base de datos con los datos del Nuevo Copropiedad y lo agrega a la Colección de 
    ''' Centros de Utilidad.
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub SAdicioneCentroUtilidad()
        If MobjCentroUtilidadNuevo.FblnEstanTodosOk Then
            MobjCentroUtilidadNuevo.SActualice(True)
        Else
            Dim lstrMens = "Aún hay campos requeridos sin satisfacer!." & vbCrLf &
                "El objeto no fue grabado!"
            SLevanteEventoNot(lstrMens, "", EnuIdMens.EnuDatCarInvalidos,
                    EnuSeveridadNot.EnuFalta)
        End If
    End Sub
    ''' <summary>
    ''' Establece la propiedad "objCenUtilidadActual" que representa la Copropiedad en el cual
    ''' se va a trabajar.
    ''' </summary>
    ''' <param name="ashrIdCenUtilidad">Id de la Copropiedad que va a ser instanciado.</param>
    ''' <remarks></remarks>
    Friend Sub SEstablezcaCenUtilidadActual(ashrIdCenUtilidad As Short)
        If MobjCentroUtilidadActual Is Nothing Then
            Dim lcolCentrosUtil = FcolCentrosUtilidad(), lstrKey As String
            If lcolCentrosUtil.Count > 0 Then
                lstrKey = ObjIdCarpetaShr.ToString & ashrIdCenUtilidad.ToString
                MobjCentroUtilidadActual = lcolCentrosUtil(lstrKey)
            End If
        End If
        If MobjCentroUtilidadActual IsNot Nothing AndAlso MobjCentroUtilidadActual.BlnExiste Then
            GshrIdCentroUtil = MobjCentroUtilidadActual.ObjIdCentroUtilShr.ObjValorPro
        End If
    End Sub
    Friend ReadOnly Property ObjCentroUtilidadActual As ClsCentroUtilidad
        Get
            If MobjCentroUtilidadActual Is Nothing Then
                SEstablezcaCenUtilidadActual(GshrIdCentroUtil)
            End If
            Return MobjCentroUtilidadActual
        End Get
    End Property
    Friend Function FdtbCentrosUtilidad() As DataTable
        Dim lstrFiltro As String = ClsIdCarpetaCenUtilShr.SstrNombreCampoBd & " = " &
                ObjIdCarpetaShr.ToString
        Dim ldtbCentrosUtilidad = ClsPanorama.FdtbDataTable(ClsCentroUtilidad.SstrNombreTabla, {"*"},
                    {{ClsIdCarpetaCenUtilShr.SstrNombreCampoBd, "ASC"},
                    {StrCampoCentroUtil, "ASC"}}, lstrFiltro, True,
                     Array.Empty(Of String))
        Return ldtbCentrosUtilidad
    End Function
#End Region
#Region "Manejo Cuentas Contabilidad"
    Friend Function FdtbCuentasCont() As DataTable
        Dim lstrFiltro As String = ObjIdCarpetaShr.StrNombreCampoBD & " = " &
                ObjIdCarpetaShr.ObjValorPro
        Dim lstrCamposSelect = {"IdCarpeta", "IdCuentaCont", "Nombre"}
        Dim ldtbCuentasCon = ClsPanorama.FdtbDataTable("PanCuentasContabilidad", lstrCamposSelect,
                    {{"IdCuentaCont", "ASC"}}, lstrFiltro)
        Return ldtbCuentasCon
    End Function
    Friend Function FcolCuentasCont() As Collection
        Dim lcolCuentasCont As New Collection
        If ObjIdCarpetaShr.BlnEsValido AndAlso EnuEstadoActualizacion <>
                EnuEstadoObjetoDef.enuCreando Then
            Dim ldtbCuentaCont = FdtbCuentasCont()
            For Each ldrwCtaCont As DataRow In ldtbCuentaCont.Rows
                Dim lobjCtaCont As New ClsCuentaContabilidad(Me, ldrwCtaCont)
                lobjCtaCont.SLeaValores(True)
                lcolCuentasCont.Add(lobjCtaCont, lobjCtaCont.ObjIdCuentaContStr.ToString)
            Next
        End If
        Return lcolCuentasCont
    End Function
    ''' <summary>
    ''' Actualiza la base de datos con los datos de la Cuenta de Contabilidad y la agrega a la Colección de 
    ''' Cuentas de Contabilidad.
    ''' </summary>
    ''' <param name="aobjCuentaCont">Objeto Cuenta de Contabilidad a ser adicionado.</param>
    ''' <remarks></remarks>
    Friend Sub SAdicioneCuentaContabilidad(aobjCuentaCont As ClsCuentaContabilidad)
        aobjCuentaCont.SActualice(True)
    End Sub
    ''' <summary>
    ''' Devuelve el objeto Cuenta de Contabilidad correspondiente al argumento "astrIdCuentaContabilidad"
    ''' </summary>
    ''' <param name="astrIdCuentaContabilidad">Identifica la cuenta de contabilidad que se devuelve.</param>
    ''' <returns></returns>
    Friend Function FobjCuentaContabilidad(astrIdCuentaContabilidad As String) As ClsCuentaContabilidad 'Ok
        Dim lobjCuentaCont As ClsCuentaContabilidad = Nothing
        Dim ldtbCuentasCon = FdtbCuentasCont()
        If ldtbCuentasCon.Rows.Count > 0 Then
            Dim lstrFiltro As String = ClsIdCuentaContStr.SstrNombreCampoBd & " = '" &
                    astrIdCuentaContabilidad & "'"
            Dim ldrwCtaCon() As DataRow = ldtbCuentasCon.Select(lstrFiltro)
            If ldrwCtaCon.Length > 0 Then
                lobjCuentaCont = New ClsCuentaContabilidad(Me, ldrwCtaCon(0))
                lobjCuentaCont.SLeaValores(True)
            End If
        End If
        Return lobjCuentaCont
    End Function
    Friend Shared Function FblnExisteCuentaCont(astrIdCuentaCont As String) As Boolean 'Ok
        Dim lblnExisteCtaCon As Boolean
        Dim lstrTabla = ClsCuentaContabilidad.SstrNombreTabla
        Dim lstrCamposSel = {ClsIdCuentaContStr.SstrNombreCampoBd}
        Dim lstrFiltro = StrCampoCarpeta & " = " & GshrIdCarpeta &
                " AND " & ClsIdCuentaContStr.SstrNombreCampoBd & " = '" &
                astrIdCuentaCont & "'"
        Dim ldtbCtaCon = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSel, {{}},
                lstrFiltro)
        lblnExisteCtaCon = (ldtbCtaCon.Rows.Count > 0)
        Return lblnExisteCtaCon
    End Function
    Friend Function FstrNombreCuentaCont(astrIdCuentaCont As String)
        Dim lstrNombreCta As String
        Dim lobjCtaCont = FobjCuentaContabilidad(astrIdCuentaCont)
        If Not IsNothing(lobjCtaCont) Then
            lstrNombreCta = lobjCtaCont.ObjNombreCuentaStr.ObjValorPro
        Else
            Throw New ValorArgumentoInvalidoException("La Cuenta " & astrIdCuentaCont & " no existe!")
        End If
        Return lstrNombreCta
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsEstaActivaBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "EstaActiva"
    Friend Sub New(aobjPadre As ClsCarpeta)
        MyBase.New(aobjPadre)
        HstrNombre = "Activa"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        HblnEsValido = lblnEsValido
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
Friend Class ClsIdCarpetaShr
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
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Short.MaxValue,
                BlnEsRequerido, EnuTipoValor)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
            If Not BlnLeyendoOrigen Then
                HstrMens = String.Empty
                If HblnEsValido Then
                    If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                        Dim lobjLlavePrincipal() = {HobjValorNew}
                        If Not ObjPadre.FblnExisteLlave(lobjLlavePrincipal) Then
                            HstrMens = "La Id. de la Carpeta ingresada no existe!"
                            HblnEsValido = False
                        End If
                    End If
                Else
                    HstrMens = "La Id. de la Carpeta ingresada no es válida!"
                    SNotifiqueDatInv()
                End If
                If Not String.IsNullOrEmpty(HstrMens) Then
                    SNotifiqueDatInv()
                End If
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
Friend Class ClsNombreStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Nombre"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Nombre"
        HshrLongitud = 50
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            HobjValorNew = UCase(HobjValorNew)
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
#End Region
