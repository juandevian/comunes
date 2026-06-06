Friend Class ClsUsuario
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "PanUsuarios"
    '
    Private ReadOnly MobjIdUsuarioStr As New ClsIdUsuarioStr(Me)
    Private ReadOnly MobjNombreUsuarioStr As New ClsNombreUsuarioStr(Me)
    Private ReadOnly MobjContrasenaUsuarioStr As New ClsContrasenaUsuarioStr(Me)
    Private ReadOnly MobjTipoCambioContrasenaByt As New ClsTipoCambioContrasenaByt(Me)
    Private ReadOnly MobjFechaExpiracionDtm As New ClsFechaExpiracionDtm(Me)
    Private ReadOnly MobjFechaCambioContrasenaDtm As New ClsFechaCambioContrasenaDtm(Me)
    Private ReadOnly MobjEstaActivoUsuarioBln As New ClsEstaActivoUsuarioBln(Me)
    'Perfiles Usuario
    Private MobjPerfilUsuarioActual As ClsPerfilUsuario = Nothing
    ' Permisos
    Private McolPermisosAcciones As Collection = Nothing
#End Region
#Region "Constructores"
    Friend Sub New(aenuModoInstanciaObj As EnuModoInstanciaObjDef, ablnSoloActivos As Boolean)
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuDeColeccion Then
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        End If
        Dim lstrCamposSelect As String()
        HobjPadre = GobjAdministrador
        HblnEsAnulable = False
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuNavegable Then
            lstrCamposSelect = {ObjIdUsuarioStr.StrNombreCampoBD}
            Dim lstrFiltro = ""
            If ablnSoloActivos Then
                lstrFiltro = ClsEstaActivoUsuarioBln.SstrNombreCampoBd & " = TRUE"
            End If
            HcolFiltros.Add(lstrFiltro)
        Else
            HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
            lstrCamposSelect = {"*"}
        End If
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
    End Sub
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Carpeta al cual pertenece la colección de cuentas de 
    ''' contabilidad de la cual este objeto forma parte</param>
    ''' <param name="adrwUsuario">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsAdministrador, adrwUsuario As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        '
        DrwRegistroActual = adrwUsuario
        DtbTablaColeccion = DrwRegistroActual.Table
        If GenuTipoInstanciamiento = EnuTipoInstanciamiento.enuInstalacion Then
            EnuPermisosObj = EnuPermisosDef.enuCrear
        End If
    End Sub
#End Region
#Region "Propiedades"
#Region "Propiedades indentificadoras"
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
            Return EnuIdClasesPanDef.EnuUsuario
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Usuario"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjIdUsuarioStr As ClsIdUsuarioStr
        Get
            Return MobjIdUsuarioStr
        End Get
    End Property
    Friend ReadOnly Property ObjNombreUsuarioStr As ClsNombreUsuarioStr
        Get
            Return MobjNombreUsuarioStr
        End Get
    End Property
    Friend ReadOnly Property ObjContrasenaUsuarioStr As ClsContrasenaUsuarioStr
        Get
            Return MobjContrasenaUsuarioStr
        End Get
    End Property
    Friend ReadOnly Property ObjTipoCambioContrasenaByt As ClsTipoCambioContrasenaByt
        Get
            Return MobjTipoCambioContrasenaByt
        End Get
    End Property
    Friend ReadOnly Property ObjFechaExpiracionDtm As ClsFechaExpiracionDtm
        Get
            Return MobjFechaExpiracionDtm
        End Get
    End Property
    Friend ReadOnly Property ObjFechaCambioContrasenaDtm As ClsFechaCambioContrasenaDtm
        Get
            Return MobjFechaCambioContrasenaDtm
        End Get
    End Property
    Friend ReadOnly Property ObjEstaActivoUsuarioBln As ClsEstaActivoUsuarioBln
        Get
            Return MobjEstaActivoUsuarioBln
        End Get
    End Property
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(MobjIdUsuarioStr)
                HcolPropiedades.Add(MobjNombreUsuarioStr)
                HcolPropiedades.Add(MobjContrasenaUsuarioStr)
                HcolPropiedades.Add(MobjTipoCambioContrasenaByt)
                HcolPropiedades.Add(MobjFechaExpiracionDtm)
                HcolPropiedades.Add(ObjFechaCreacionDtm)
                HcolPropiedades.Add(MobjFechaCambioContrasenaDtm)
                HcolPropiedades.Add(MobjEstaActivoUsuarioBln)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras Propiedades"
    Friend ReadOnly Property BlnTieneAsignadaEstaApp As Boolean
        Get
            Dim lblnTieneAsigApp = (MobjIdUsuarioStr.ObjValorPro = GCSTRUSUARIOU)
            If Not lblnTieneAsigApp Then
                If Not FblnEsCreable({MobjIdUsuarioStr.ObjValorPro}) Then
                    lblnTieneAsigApp = FblnTieneAsignadaApp()
                End If
            End If
            Return lblnTieneAsigApp
        End Get
    End Property
    Friend Function FcolCarpetasUsuario() As Collection
        Dim lcolCarpUsu As New Collection
        Dim lstrKey As String
        Dim ldtbCarpUsu = FdtbCarpetasUsuario()
        For Each ldrwCarUsu As DataRow In ldtbCarpUsu.Rows
            Dim lobjCarUsu As New ClsCarpetaUsuario(Me, ldrwCarUsu)
            lobjCarUsu.SLeaValores(True)
            lstrKey = lobjCarUsu.ObjIdCarpetaUsuarioShr.ToString
            lcolCarpUsu.Add(lobjCarUsu, lstrKey)
        Next
        Return lcolCarpUsu
    End Function
    Friend Function FcolCentrosUtilCarUsuario() As Collection
        Dim lcolcenUtil As New Collection
        If GobjPanorama.ObjCarpetaActual IsNot Nothing Then
            lcolcenUtil = GobjPanorama.ObjCarpetaActual.FcolCentrosUtilidad
        End If
        Return lcolcenUtil
    End Function
    Friend Function FcolPerfilesUsuario() As Collection
        Dim lcolPerfilesUsuario As New Collection
        Dim ldtbPerfilesUsuario = FdtbPerfilesUsuario()
        Dim lstrKey As String
        If ldtbPerfilesUsuario IsNot Nothing Then
            For Each ldrwPerUsu As DataRow In ldtbPerfilesUsuario.Rows
                Dim lobjPerUsu As New ClsPerfilUsuario(Me, ldrwPerUsu)
                lobjPerUsu.SLeaValores(True)
                lstrKey = lobjPerUsu.ObjIdAppPerfilUsuarioShr.ToString &
                            lobjPerUsu.ObjIdPerfilUsuarioShr.ToString
                lcolPerfilesUsuario.Add(lobjPerUsu, lstrKey)
            Next
        End If
        Return lcolPerfilesUsuario
    End Function
    Friend ReadOnly Property DrwCarpetasUsuario As DataRow()
        Get
            Dim lstrFiltro As String
            Dim lstrCamposSelect() As String
            Dim ldtbCarpetasUsuario As DataTable
            If ObjIdUsuarioStr.ObjValorPro = GCSTRUSUARIOU Then
                lstrFiltro = ClsEstaActivaBln.SstrNombreCampoBd & " = " & True.ToString
                lstrCamposSelect = {StrCampoCarpeta, ClsNombreStr.SstrNombreCampoBd}
                ldtbCarpetasUsuario = ClsPanorama.FdtbDataTable(ClsCarpeta.SstrNombreTabla,
                        lstrCamposSelect, {{StrCampoCarpeta, "ASC"}},
                        lstrFiltro, True, Array.Empty(Of String)())
            Else
                lstrFiltro = ObjIdUsuarioStr.StrNombreCampoBD & " = '" &
                        ObjIdUsuarioStr.ToString & "'"
                lstrCamposSelect = {ClsIdCarpetaUsuarioShr.SstrNombreCampoBd,
                                    ClsNombreCarpetaStr.SstrNombreCampoBd}
                ldtbCarpetasUsuario = ClsPanorama.FdtbDataTable("PanCarpetasUsuarios", lstrCamposSelect,
                                   {{ClsIdCarpetaUsuarioShr.SstrNombreCampoBd, "ASC"}}, lstrFiltro, True,
                                   Array.Empty(Of String)())
            End If
            Return ldtbCarpetasUsuario.Select
        End Get
    End Property
    Private ReadOnly Property ObjPerfilUsuarioActual As ClsPerfilUsuario
        Get
            If IsNothing(MobjPerfilUsuarioActual) Then
                Dim lcolPerfUsuario As Collection = FcolPerfilesUsuario()
                For Each lobjPerUsu As ClsPerfilUsuario In lcolPerfUsuario
                    If lobjPerUsu.ObjIdAppPerfilUsuarioShr.ObjValorPro = CType(GenuIdAplicacion, Short) Then
                        MobjPerfilUsuarioActual = lobjPerUsu
                        Exit For
                    End If
                Next
            End If
            Return MobjPerfilUsuarioActual
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Friend Overrides Function FblnEsSuprimible() As Boolean
        Dim lblnEsSuprimible = FblnPermitidoSuprimir()
        If lblnEsSuprimible Then
            ' La condicion no debe incluir el nombre de la columna
            Dim lstrCondicion As String = " = '" & ObjIdUsuarioStr.ObjValorPro & "'"
            lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg({SstrNombreTabla},
                    ObjIdUsuarioStr.StrNombreCampoBD, lstrCondicion, False, False)
            If lblnEsSuprimible Then
                lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg({SstrNombreTabla},
                        ObjIdUsuarioStr.StrNombreCampoBD, lstrCondicion, True, True)
            End If
        End If
        Return lblnEsSuprimible
    End Function
    Protected Friend Overrides Sub SVacie()
        MyBase.SVacie()
        MobjPerfilUsuarioActual = Nothing
        McolPermisosAcciones = Nothing
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdUsuarioStr.ToString
        End Get
    End Property
#End Region
#Region "Procedimientos del objeto"
    Friend Sub SLeaUsuarioUniversal()
#If DES = 1 Then
        Dim lstrPasswordUsuUni As String = ("00")
        ObjTipoCambioContrasenaByt.ObjValorPro = CType(EnuTipoCambioContrasenaDef.enuNoCambiar, Byte)
#Else
        Dim lstrPasswordUsuUni As String = ("13798642")
        ObjTipoCambioContrasenaByt.ObjValorPro = CType(EnuTipoCambioContrasenaDef.enuNoCambiar, Byte)
#End If
        ClsPanorama.SEncriptePassword(lstrPasswordUsuUni)
        ObjIdUsuarioStr.BlnLeyendoOrigen = True
        ObjIdUsuarioStr.ObjValorPro = GCSTRUSUARIOU
        ObjNombreUsuarioStr.BlnLeyendoOrigen = True
        ObjNombreUsuarioStr.ObjValorPro = "OPTIMUSOFT"
        ObjContrasenaUsuarioStr.BlnLeyendoOrigen = True
        ObjContrasenaUsuarioStr.ObjValorPro = lstrPasswordUsuUni
        ObjTipoCambioContrasenaByt.BlnLeyendoOrigen = True
        ObjFechaCambioContrasenaDtm.BlnLeyendoOrigen = True
        ObjFechaCambioContrasenaDtm.ObjValorPro = Date.Today
        ObjFechaCreacionDtm.BlnLeyendoOrigen = True
        ObjFechaCreacionDtm.ObjValorPro = GCDTMFECHANULA
        ObjFechaExpiracionDtm.BlnLeyendoOrigen = True
        ObjFechaExpiracionDtm.ObjValorPro = GCDTMFECHANULA
        ObjEstaActivoUsuarioBln.BlnLeyendoOrigen = True
        ObjEstaActivoUsuarioBln.ObjValorPro = True
    End Sub
    Public Function FblnDebeCambiarContrasena() As Boolean
        If ObjIdUsuarioStr.ObjValorPro.ToString.Trim <> GCSTRUSUARIOU Then
            Select Case ObjTipoCambioContrasenaByt.ObjValorPro
                Case EnuTipoCambioContrasenaDef.enuAlVencimiento
                    If ObjFechaExpiracionDtm.ObjValorPro <= Date.Today Then
                        Return True
                    End If
                Case EnuTipoCambioContrasenaDef.enuProximaVez
                    Return True
            End Select
            Return False
        Else
            Return False
        End If
    End Function
    Private Function FdtbPerfilesUsuario() As DataTable
        Dim ldtbPerUsu As DataTable = Nothing
        If ObjIdUsuarioStr.BlnEsValido Then
            Dim lstrFiltro As String
            If ObjIdUsuarioStr.ObjValorPro = GCSTRUSUARIOU Then
                lstrFiltro = ObjIdUsuarioStr.StrNombreCampoBD & " = '" &
                            GCSTRADMIN & "'"
            Else
                lstrFiltro = ObjIdUsuarioStr.StrNombreCampoBD & " = '" &
                            ObjIdUsuarioStr.ToString & "'"
            End If
            ldtbPerUsu = ClsPanorama.FdtbDataTable("PanPerfilesUsuarios", {"*"},
                        {{"IdAplicacion", "ASC"}, {"IdPerfil", "ASC"}}, lstrFiltro, True,
                        Array.Empty(Of String)())
        End If
        Return ldtbPerUsu
    End Function
    Private Function FdtbCarpetasUsuario() As DataTable
        Dim ldtbCarpUsuario As DataTable
        Dim lstrFiltro As String = ObjIdUsuarioStr.StrNombreCampoBD & " = '" &
                    ObjIdUsuarioStr.ToString & "'"
        ldtbCarpUsuario = ClsPanorama.FdtbDataTable("PanCarpetasUsuarios", {"*"},
                    {{"IdCarpeta", "ASC"}}, lstrFiltro, True, Array.Empty(Of String)())
        Return ldtbCarpUsuario
    End Function
    ''' <summary>
    ''' Asigna el perfil pasado en el argumento al usuario "Admin" de la aplicacion "Administrador"
    ''' </summary>
    ''' <param name="aobjPerfil">Objeto perfil que sera asignado al usuario</param>
    ''' <remarks>Este procedimiento solo se usa en la instalacion del Administrador para
    ''' asignarle el perfil "Admin" al Usuario "Admin"</remarks>
    Friend Sub SAsignePerfil(aobjPerfil As ClsPerfil)
        Dim ldtbPerfilesUsuario = FdtbPerfilesUsuario()
        If Not IsNothing(ldtbPerfilesUsuario) Then
            Dim ldrwPerfilUsuario As DataRow = ldtbPerfilesUsuario.NewRow
            Dim lobjPerfilUsuario As New ClsPerfilUsuario(Me, ldrwPerfilUsuario)
            lobjPerfilUsuario.SDeterminePermisos()
            lobjPerfilUsuario.SCreeObj(Nothing)
            With lobjPerfilUsuario
                .ObjIdAppPerfilUsuarioShr.ObjValorPro = GobjPanorama.ObjAppActual.ObjIdAppShr.ObjValorPro
                .ObjIdPerfilUsuarioShr.ObjValorPro = aobjPerfil.ObjIdPerfilShr.ObjValorPro
                .ObjIdUsuarioPerfilStr.ObjValorPro = ObjIdUsuarioStr.ObjValorPro
                .ObjNombreAppPerfilUsuarioStr.ObjValorPro = GobjPanorama.ObjAppActual.ObjDefApp.StrNombreCompleto
                .ObjNombrePerfilUsuarioStr.ObjValorPro = aobjPerfil.ObjNombrePerfilStr.ObjValorPro
                .SActualice(True)
            End With
        End If
    End Sub
    ''' <summary>
    ''' Asigna los perfiles de aplicacion al usuario actual.
    ''' </summary>
    ''' <param name="astrAppsPerfiles">Array de cadenas en cada una de las cuales los tres primeros 
    ''' caracteres identifican la aplicación y el resto identifica el perfil de la aplicación.</param>
    ''' <remarks></remarks>
    Friend Function FblnAsignoPerfiles(astrAppsPerfiles As String()) As Boolean
        Dim lstrAppsAsignadas As String() = FstrAppsAsignadas()
        Dim lstrAppsRevisadas As String() = Nothing
        Dim j As Integer = 0
        Dim lstrIdApp As String
        Dim lstrIdPerfil As String
        GobjPanDat.SControleProcesoObj(True)
        GobjPanDat.SInicialiceTransaccion()
        Dim lblnAsigno = FblnSuprimioPerfiles(astrAppsPerfiles)
        If lblnAsigno Then
            If astrAppsPerfiles IsNot Nothing AndAlso astrAppsPerfiles.Length > 0 Then
                Dim lcolPerfUsuario = FcolPerfilesUsuario()
                For i As Integer = 0 To astrAppsPerfiles.GetUpperBound(0)
                    lstrIdApp = astrAppsPerfiles(i).Split(",")(0)
                    lstrIdPerfil = astrAppsPerfiles(i).Split(",")(1)
                    ReDim Preserve lstrAppsRevisadas(i)
                    lstrAppsRevisadas(j) = lstrIdApp.Substring(0, 3)
                    j += 1
                    If Not lcolPerfUsuario.Contains(lstrIdApp) Then
                        SAsignePerfil(CShort(lstrIdApp), CShort(lstrIdPerfil))
                    End If
                Next
            End If
        End If
        If lblnAsigno AndAlso lstrAppsAsignadas IsNot Nothing AndAlso
                lstrAppsAsignadas.Length > 0 Then
            For i = 0 To lstrAppsAsignadas.GetUpperBound(0)
                If astrAppsPerfiles Is Nothing Then
                    lblnAsigno = FblnSuprimioPerfilApp(lstrAppsAsignadas(i))
                Else
                    If lstrAppsRevisadas IsNot Nothing AndAlso
                            Not lstrAppsRevisadas.Contains(lstrAppsAsignadas(i)) Then
                        lblnAsigno = FblnSuprimioPerfilApp(lstrAppsAsignadas(i))
                    End If
                End If
                If Not lblnAsigno Then Exit For
            Next
        End If
        If lblnAsigno Then
            GobjPanDat.SConfirmeTransaccion()
        Else
            GobjPanDat.SAborteTransaccion()
        End If
        GobjPanDat.SControleProcesoObj(True)
        Return lblnAsigno
    End Function
    Private Sub SAsignePerfil(ashrIdApp As Integer, ashrIdPerfil As Integer)
        Dim lobjValorLlavePerfil As Object() = {ashrIdApp, ashrIdPerfil}
        Dim ldtbPerfilesUsuario = FdtbPerfilesUsuario()
        Dim lobjApp As ClsAplicacion = Nothing
        If GobjPanorama.ObjAppActual.ObjIdAppShr.ObjValorPro = ashrIdApp Then
            lobjApp = GobjPanorama.ObjAppActual
        Else
            Dim lobjValorLlaveApp As Object() = {ashrIdApp}
            lobjApp = New ClsAplicacion(EnuModoInstanciaObjDef.enuUnico)
            lobjApp.SAbra(lobjValorLlaveApp)
        End If
        Dim lobjPerfil As New ClsPerfil(lobjApp, EnuModoInstanciaObjDef.enuUnico)
        lobjPerfil.SAbra(lobjValorLlavePerfil)
        Dim ldrwPerfilUsuario As DataRow = ldtbPerfilesUsuario.NewRow
        Dim lobjPerfilUsuario As New ClsPerfilUsuario(Me, ldrwPerfilUsuario)
        lobjPerfilUsuario.SDeterminePermisos()
        lobjPerfilUsuario.SCreeObj(Nothing)
        With lobjPerfilUsuario
            .ObjIdAppPerfilUsuarioShr.ObjValorPro = lobjApp.ObjIdAppShr.ObjValorPro
            .ObjIdPerfilUsuarioShr.ObjValorPro = lobjPerfil.ObjIdPerfilShr.ObjValorPro
            .ObjIdUsuarioPerfilStr.ObjValorPro = ObjIdUsuarioStr.ObjValorPro
            .ObjNombreAppPerfilUsuarioStr.ObjValorPro = lobjApp.ObjDefApp.StrNombreCompleto
            .ObjNombrePerfilUsuarioStr.ObjValorPro = lobjPerfil.ObjNombrePerfilStr.ObjValorPro
            .SActualice(True)
        End With
    End Sub
    Private Sub SAsigneCarpeta(ashrIdCarpeta As Integer)
        Dim lobjValorLlaveCarpeta As Object() = {ashrIdCarpeta}
        Dim ldtbCarpUsu = FdtbCarpetasUsuario()
        Dim lobjCarpeta As New ClsCarpeta(EnuModoInstanciaObjDef.enuUnico, False)
        lobjCarpeta.SAbra(lobjValorLlaveCarpeta)
        Dim ldrwCarpetaUsuario As DataRow = ldtbCarpUsu.NewRow
        Dim lobjCarpetaUsuario As New ClsCarpetaUsuario(Me, ldrwCarpetaUsuario)
        lobjCarpetaUsuario.SCreeObj(Nothing)
        With lobjCarpetaUsuario
            .ObjIdCarpetaUsuarioShr.ObjValorPro = lobjCarpeta.ObjIdCarpetaShr.ObjValorPro
            .ObjIdUsuarioCarpetaStr.ObjValorPro = ObjIdUsuarioStr.ObjValorPro
            .ObjNombreCarpetaStr.ObjValorPro = lobjCarpeta.ObjNombreStr.ObjValorPro
            .SActualice(True)
        End With
    End Sub
    Friend Sub SEstablezcaCarpetaActual(ashrIdCarpeta)
        GobjPanorama.SEstablezcaCarpetaActual(ashrIdCarpeta)
        If GobjPanorama.ObjCarpetaActual IsNot Nothing Then
            If FcolCentrosUtilCarUsuario.Count = 0 Then
                Dim lstrNomCar = GobjPanorama.ObjCarpetaActual.ObjNombreStr.ToString
                Dim lstrMens = "La Carpeta " & Chr(34) & lstrNomCar & Chr(34) &
                        " no tiene Centros de Utilidad creados!"
                SLevanteEventoNot(lstrMens, "",
                        EnuIdMens.EnuCarpSinCenutil, EnuSeveridadNot.EnuFalta)
            End If
        End If
    End Sub
    ''' <summary>
    ''' Asigna las carpetas a las cuales tiene acceso al usuario actual.
    ''' </summary>
    ''' <param name="aintIdCarpetas">Array que contiene la identificación de las carpetas asignadas.</param>
    ''' <remarks></remarks>
    Friend Function FblnAsignoCarpetas(aintIdCarpetas As Integer())
        Dim i As Integer, lblnAsigno As Boolean
        Dim lshrIdCarpeta As Integer
        GobjPanDat.SControleProcesoObj(True)
        lblnAsigno = FblnSuprimioCarpetasUsuario(aintIdCarpetas)
        If lblnAsigno Then
            If aintIdCarpetas IsNot Nothing Then
                Dim lcolCarpeUsuario = FcolCarpetasUsuario()
                For i = 0 To aintIdCarpetas.GetUpperBound(0)
                    lshrIdCarpeta = aintIdCarpetas(i)
                    If Not lcolCarpeUsuario.Contains(lshrIdCarpeta.ToString) Then
                        SAsigneCarpeta(lshrIdCarpeta)
                    End If
                Next
            End If
        End If
        GobjPanDat.SControleProcesoObj(False)
        Return lblnAsigno
    End Function
    Private Function FblnSuprimioPerfiles(astrAppsPerfiles As String()) As Boolean
        GobjPanDat.SControleProcesoObj(True)
        GobjPanDat.SInicialiceTransaccion()
        Dim lblnSuprimio = True
        Dim lstrKey As String
        Dim lcolPerfUsuario = FcolPerfilesUsuario()
        For Each lobjPerUsu As ClsPerfilUsuario In lcolPerfUsuario
            lstrKey = lobjPerUsu.ObjIdAppPerfilUsuarioShr.ToString &
                        lobjPerUsu.ObjIdPerfilUsuarioShr.ToString
            If astrAppsPerfiles Is Nothing OrElse Not astrAppsPerfiles.Contains(lstrKey) Then
                lblnSuprimio = lobjPerUsu.FblnEsSuprimible()
                If lblnSuprimio Then
                    lblnSuprimio = lobjPerUsu.FblnSuprimio()
                End If
                If Not lblnSuprimio Then Exit For
            End If
        Next
        If lblnSuprimio Then
            GobjPanDat.SConfirmeTransaccion()
        Else
            GobjPanDat.SAborteTransaccion()
        End If
        GobjPanDat.SControleProcesoObj(False)
        Return lblnSuprimio
    End Function
    Private Function FblnSuprimioPerfilApp(astrApp As String)
        GobjPanDat.SControleProcesoObj(True)
        GobjPanDat.SInicialiceTransaccion()
        Dim lblnSuprimio = True
        Dim lcolPerfUsuario = FcolPerfilesUsuario()
        For Each lobjPerUsu As ClsPerfilUsuario In FcolPerfilesUsuario()
            If lobjPerUsu.ObjIdAppPerfilUsuarioShr.ToString = astrApp Then
                lblnSuprimio = lobjPerUsu.FblnEsSuprimible()
                If lblnSuprimio Then
                    lblnSuprimio = lobjPerUsu.FblnSuprimio()
                End If
                If Not lblnSuprimio Then Exit For
            End If
        Next
        If lblnSuprimio Then
            GobjPanDat.SConfirmeTransaccion()
        Else
            GobjPanDat.SAborteTransaccion()
        End If
        GobjPanDat.SControleProcesoObj(False)
        Return lblnSuprimio
    End Function
    Private Function FstrAppsAsignadas() As String()
        Dim lstrAppsAsignadas As String()
        Dim i As Integer = 0
        Dim lcolperfUsuario = FcolPerfilesUsuario()
        ReDim lstrAppsAsignadas(lcolperfUsuario.Count - 1)
        For Each lobjPerUsu As ClsPerfilUsuario In lcolperfUsuario
            lstrAppsAsignadas(i) = lobjPerUsu.ObjIdAppPerfilUsuarioShr.ToString
        Next
        Return lstrAppsAsignadas
    End Function
    Private Function FblnSuprimioCarpetasUsuario(aentCarpetasUsuario As Integer()) As Boolean
        Dim lblnSuprimio = True
        Dim lcolCarpUsuario = FcolCarpetasUsuario()
        GobjPanDat.SControleProcesoObj(True)
        GobjPanDat.SInicialiceTransaccion()
        For Each lobjCarUsu As ClsCarpetaUsuario In lcolCarpUsuario
            If IsNothing(aentCarpetasUsuario) OrElse Not aentCarpetasUsuario.Contains(
                    lobjCarUsu.ObjIdCarpetaUsuarioShr.ObjValorPro) Then
                lblnSuprimio = lobjCarUsu.FblnSuprimio()
                If Not lblnSuprimio Then Exit For
            End If
        Next
        If lblnSuprimio Then
            GobjPanDat.SConfirmeTransaccion()
        Else
            GobjPanDat.SAborteTransaccion()
        End If
        GobjPanDat.SControleProcesoObj(False)
        Return lblnSuprimio
    End Function
    Friend Sub SAsigneUsuarioActual()
        If (BlnExiste OrElse MobjIdUsuarioStr.ObjValorPro = GCSTRUSUARIOU) AndAlso FblnEstanTodosOk() Then
            GobjPanorama.ObjUsuarioActual = Me
            GstrIdUsuario = MobjIdUsuarioStr.ObjValorPro
            If MobjIdUsuarioStr.ObjValorPro <> GCSTRUSUARIOU Then
                MobjPerfilUsuarioActual = ObjPerfilUsuarioActual
            End If
            GobjPanorama.SRegistreAccionLogApp(StrNombreClase, "LogOn")
        End If
    End Sub
    Public Sub SRegistreLogOff()
        GobjPanorama.SRegistreAccionLogApp(StrNombreClase, "LogOff")
    End Sub
    Private Function FblnTieneAsignadaApp() As Boolean
        Dim lblnAppAsignada As Boolean = False
        If ObjIdUsuarioStr.BlnEsValido Then
            Dim lcolPerfUsuario = FcolPerfilesUsuario()
            For Each lobjPerUsu As ClsPerfilUsuario In lcolPerfUsuario
                lblnAppAsignada = (lobjPerUsu.ObjIdAppPerfilUsuarioShr.ObjValorPro =
                            CType(GenuIdAplicacion, Short))
                If lblnAppAsignada Then
                    Exit For
                End If
            Next
        End If
        Return lblnAppAsignada
    End Function
    ''' <summary>
    ''' Indica si el usuario tiene permiso para ejecutar una acción en un objeto 
    ''' </summary>
    ''' <param name="aintIdObjeto">Id. del objeto</param>
    ''' <param name="aintIdAccion">Id. de la acción</param>
    ''' <returns>Buleano</returns>
    ''' <remarks></remarks>
    Public Function FblnTienePermiso(aintIdObjeto As Integer, aintIdAccion As Integer) As Boolean
        Dim lblnPermiso = True
        If Not GstrIdUsuario = GCSTRUSUARIOU Then
            If IsNothing(McolPermisosAcciones) Then
                McolPermisosAcciones = ObjPerfilUsuarioActual.FcolPermisosAcciones
            End If
            Dim lstrKey As String = aintIdObjeto.ToString &
                    aintIdAccion.ToString
            If McolPermisosAcciones.Contains(lstrKey) Then
                Dim lobjPerAcci As ClsPermisoAccion = McolPermisosAcciones(lstrKey)
                lblnPermiso = lobjPerAcci.ObjAccionPermitidaBln.ObjValorPro
            End If
        End If
        Return lblnPermiso
    End Function
    ''' <summary>
    ''' Devuelve el tipo de permisos que tiene el usuario para un objeto determinado
    ''' </summary>
    ''' <param name="aenuIdClase">Identifica la clase del objeto</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function FenuTipoPermisosObj(aenuIdClase As EnuIdClasesPanDef) As EnuPermisosDef
        Dim lenuTipoPermisosObj As EnuPermisosDef = EnuPermisosDef.None
        Dim lblnPermitido As Boolean
        Dim lshrIdObjeto As Integer = CType(aenuIdClase, Integer)
        Dim lcolPermisosAcciones As Collection = ObjPerfilUsuarioActual.FcolPermisosAccionesObj(lshrIdObjeto)
        For Each ldrwPerAcci As DataRow In lcolPermisosAcciones
            lblnPermitido = ClsPanorama.FobjValorCampo(ldrwPerAcci("Permitida"),
                            EnuTipoValor.enuBoolean)
            If lblnPermitido Then
                Select Case ldrwPerAcci("IdAccion")
                    Case EnuIdAccionDef.enuConsultar
                        lenuTipoPermisosObj += EnuPermisosDef.enuConsultar
                    Case EnuIdAccionDef.enuCrear
                        lenuTipoPermisosObj += EnuPermisosDef.enuCrear
                    Case EnuIdAccionDef.enuModificar
                        lenuTipoPermisosObj += EnuPermisosDef.enuModificar
                    Case EnuIdAccionDef.enuSuprimir
                        lenuTipoPermisosObj += EnuPermisosDef.enuSuprimir
                    Case EnuIdAccionDef.enuAnular
                        lenuTipoPermisosObj += EnuPermisosDef.enuAnular
                    Case EnuIdAccionDef.enuImprimir
                        lenuTipoPermisosObj += EnuPermisosDef.enuImprimir
                End Select
            End If
        Next
        Return lenuTipoPermisosObj
    End Function
    Public Function FcolPermisosAccionesFrm(ashrIdForma As Integer) As Collection
        Dim lcolPermAcciones As Collection = Nothing
        If MobjIdUsuarioStr.ObjValorPro <> GCSTRUSUARIOU Then
            lcolPermAcciones = ObjPerfilUsuarioActual.FcolPermisosAccionesObj(ashrIdForma)
        End If
        Return lcolPermAcciones
    End Function
#End Region
#Region "Manejo de Notificaciones"
    Friend Overrides Function FblnNotificaOk(aenuIdMensNot As EnuIdMens) As Boolean
        Dim lblnNotOk As Boolean = False
        If aenuIdMensNot = EnuIdMens.EnuCarpSinCenutil Then
            Dim ldtbCenUtil = GobjPanorama.ObjCarpetaActual.FdtbCentrosUtilidad
            lblnNotOk = ldtbCenUtil.Rows.Count > 0
        End If
        Return lblnNotOk
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsIdUsuarioStr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsUsuario = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "IdUsuario"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdUsuario"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 0
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud,
                BlnEsRequerido)
        If Not BlnLeyendoOrigen Then
            If HblnEsValido Then
                Dim lblnEsAdmin As Boolean = (HobjValorNew = GCSTRADMIN)
                Dim lobjValorLlave() As Object = {HobjValorNew}
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                    HobjValorNew = StrConv(HobjValorNew, VbStrConv.Uppercase)
                    If Not MobjPadre.FblnEsCreable(lobjValorLlave) Then
                        HstrMens = "La Id. del Usuario ingresada ya existe!"
                        HblnEsValido = False
                    End If
                Else
                    If HobjValorNew <> GCSTRUSUARIOU Then
                        If Not MobjPadre.FblnExisteLlave(lobjValorLlave) Then
                            If lblnEsAdmin Then
                                HstrMens = "La la Aplicación no ha sido instalada debidamente!"
                                SLevanteEveNot("", 0, EnuSeveridadNot.EnuError)
                                Exit Sub
                            Else
                                HstrMens = "El Usuario ingresado no existe!"
                            End If
                            HblnEsValido = False
                        End If
                    Else
                        MobjPadre.SLeaUsuarioUniversal()
                    End If
                End If
            Else
                If Not (ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso
                        (String.IsNullOrEmpty(HobjValorNew) OrElse IsNothing(HobjValorNew))) Then
                    HstrMens = "La Id. del Usuario ingresada no es válida!"
                End If
            End If
        End If
        If Not HblnEsValido AndAlso Not String.IsNullOrEmpty(HstrMens) Then
            SNotifiqueDatInv()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Private Sub ClsIdUsuarioStr_EvnPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If HblnEsValido Then
            If MobjPadre.EnuTipoObjeto = EnuModoInstanciaObjDef.enuUnico Then
                Dim lobjValorLlave() As Object = {HobjValorNew}
                If HobjValorNew <> HobjValorOriginal Then
                    MobjPadre.SAbra(lobjValorLlave)
                    HblnEsValido = MobjPadre.BlnExiste
                End If
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsNombreUsuarioStr
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
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 5, ShrLongitud,
                BlnEsRequerido)
        If HblnEsValido Then
            HobjValorNew = FstrNombreTercero(HobjValorNew)
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
Friend Class ClsContrasenaUsuarioStr
    Inherits ClsCBPropiedad
    Private ReadOnly Mobjpadre As ClsUsuario = Nothing
    Private MblnConfirmando As Boolean = False
    Private MstrConstrasenaVerificacion As String = String.Empty
    Private Const MCSTRNOMBRECAMPOBD As String = "Contrasena"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        Mobjpadre = aobjPadre
        HstrNombre = "Contrasena"
        HshrLongitud = 12
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsUsuario = ObjPadre, lentLonMin = 4
        With lobjPadre
            HstrMens = String.Empty
            If BlnLeyendoOrigen OrElse MblnConfirmando Then
                lentLonMin = HobjValorNew.ToString.Length
            End If
            HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, lentLonMin,
                        ShrLongitud, BlnEsRequerido)
            If HblnEsValido Then
                If .EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                    If Not (BlnLeyendoOrigen OrElse MblnConfirmando) Then
                        ClsPanorama.SEncriptePassword(HobjValorNew)
                        HblnEsValido = HobjValorNew.ToString().Length >= 2
                        If Not HblnEsValido Then
                            HstrMens = "Contraseña no permitida. Por favor cámbiela!"
                        End If
                    End If
                    HblnEsValido = HblnEsValido AndAlso (HobjValorNew = MstrConstrasenaVerificacion)
                ElseIf .EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                    If Not (BlnLeyendoOrigen OrElse MblnConfirmando) Then
                        ClsPanorama.SEncriptePassword(HobjValorNew)
                        HblnEsValido = HobjValorNew.ToString().Length >= 2
                        If Not HblnEsValido Then
                            HstrMens = "Contraseña no permitida. Por favor cámbiela!"
                        End If
                    End If
                    HblnEsValido = HblnEsValido AndAlso (HobjValorNew = MstrConstrasenaVerificacion)
                Else
                    HblnEsValido = (HobjValorNew = MstrConstrasenaVerificacion)
                    If Not HblnEsValido AndAlso Not BlnLeyendoOrigen Then
                        HstrMens = "La Contraseña no es válida!"
                    End If
                End If
            Else
                HstrMens = "La Contraseña debe tener entre cuatro y doce caracteres de longitud!"
            End If
        End With
        If Not HblnEsValido AndAlso Not String.IsNullOrEmpty(HstrMens) Then
            SNotifiqueDatInv()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Private Sub EPosCambio() Handles Me.EvnPosCambio
        Dim lobjPadre As ClsUsuario = ObjPadre
        If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            lobjPadre.ObjFechaCambioContrasenaDtm.ObjValorPro = Date.Today
        End If
    End Sub
    Friend Sub SConfirmeContrasena(astrContrasena As String)
        MstrConstrasenaVerificacion = astrContrasena
        ClsPanorama.SEncriptePassword(MstrConstrasenaVerificacion)
        MblnConfirmando = True
        SValide()
        MblnConfirmando = False
    End Sub
    Friend Sub SConfirmeContrasena(astrContrasena As String, ablnRequiereCambio As Boolean)
        Dim lobjPadre As ClsUsuario = ObjPadre
        If astrContrasena.Length > 0 Then
            MstrConstrasenaVerificacion = astrContrasena
            ClsPanorama.SEncriptePassword(MstrConstrasenaVerificacion)
            HblnEsValido = (HobjValorPro = MstrConstrasenaVerificacion)
            If Not HblnEsValido Then
                HstrMens = "La Contraseñas no son iguales!"
            End If
        Else
            If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                If HobjValorPro.Length > 0 Then
                    HblnEsValido = (False)
                End If
            Else
                HblnEsValido = (False)
            End If
            HstrMens = "No se ha introducido la confirmación de la Contraseña!"
        End If
        If HblnEsValido AndAlso ablnRequiereCambio Then
            If String.IsNullOrEmpty(HobjValorPro) OrElse HobjValorOriginal = HobjValorPro Then
                HblnEsValido = False
                HstrMens = "Debe introducir una nueva Contraseña!"
            End If
        End If
        If Not String.IsNullOrEmpty(HstrMens) Then
            SNotifiqueDatInv()
        End If
    End Sub
    Protected Friend Overrides ReadOnly Property ObjValorNuevo As Object
        Get
            Return ""
        End Get
    End Property
    Friend Overrides ReadOnly Property ObjValorOriginal As Object
        Get
            Return ""
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsTipoCambioContrasenaByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTblTipoCambioContrasena"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TipoCambioContrasena"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew,
                EnuTipoCambioContrasenaDef.enuNoCambiar,
                EnuTipoCambioContrasenaDef.enuAlVencimiento, HblnEsRequerido)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Private Sub EPosCambio() Handles Me.EvnPosCambio
        Dim lobjPadre As ClsUsuario = ObjPadre
        If BlnEsValido Then
            If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                If HobjValorNew = EnuTipoCambioContrasenaDef.enuAlVencimiento Then
                    lobjPadre.ObjFechaExpiracionDtm.ObjValorPro = DateAdd(DateInterval.Month, 2, Date.Today)
                Else
                    lobjPadre.ObjFechaExpiracionDtm.ObjValorPro = Date.Today
                End If
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        Return CByte(ObjValorPro).ToString
    End Function
End Class
Friend Class ClsFechaExpiracionDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaExpiracion"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "FechaExpiracion"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = False
        Dim lobjPadre As ClsUsuario = ObjPadre
        If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If IsDate(HobjValorNew) Then
                If lobjPadre.ObjTipoCambioContrasenaByt.ObjValorPro =
                        EnuTipoCambioContrasenaDef.enuAlVencimiento Then
                    If HobjValorNew >= Date.Today Then
                        lblnEsValido = True
                    End If
                Else
                    lblnEsValido = True
                End If
            End If
        Else
            lblnEsValido = True
        End If
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
Friend Class ClsFechaCambioContrasenaDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaCambio"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "FechaCambio"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsUsuario = ObjPadre
        Dim ldtmFechaMin As Date = #1/1/2000#
        Dim ldtmFechaMax As Date = GCDTMFECHAMAXI
        If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            ldtmFechaMin = Date.Today
            ldtmFechaMax = Now
        End If
        If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            ldtmFechaMin = DateAdd(DateInterval.Day, -1, Date.Today)
            ldtmFechaMax = GCDTMFECHAMAXI
        End If
        HblnEsValido = (ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                BlnEsRequerido))
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
Friend Class ClsEstaActivoUsuarioBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Activo"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "EstaActivo"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        HobjValorPro = False
        HobjValorNew = HobjValorPro
        HblnEsValido = False
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
#End Region