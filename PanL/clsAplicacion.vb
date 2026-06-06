Namespace Aplicacion
    Friend Class ClsAplicacion
#Region "Definiciones"
        Inherits ClsCBObjetoPan
        ' Constantes
        Private Const MCSTRNOMBRETABLA As String = "PanAplicaciones"
        '
        Private MdtbPerfiles As DataTable = Nothing
        Private MblnDtbPerfilesCargada As Boolean = False
        Private McolPerfiles As Collection = Nothing
        Private MblnColPerfilesPoblada As Boolean = False
        Private MdtbOrigenesInstancias As DataTable = Nothing
        Private MblnDtbOrigenesInstCargada As Boolean = False
        Private McolOrigenesInstancias As Collection = Nothing
        Private MblnColOrigenesInstPoblada As Boolean = False
        Private McolAcciones As Collection = Nothing
        Private MblnColAccionesPoblada As Boolean = False
        Private MobjDefApp As ClsDefAplicacion = Nothing
        'Trayectoria Copia seguridad
        Private MstrTrayCopiaSeguridad As String = Nothing
#End Region
#Region "Constructores"
        Public Sub New(aenuModoInstanciaObj As EnuModoInstanciaObjDef)
            Select Case aenuModoInstanciaObj
                Case EnuModoInstanciaObjDef.enuDeColeccion
                    Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
                Case EnuModoInstanciaObjDef.enuNavegable
                    Throw New ErrorInesperadoPanLException("Este Objeto no se puede instanciar como Navegable!")
                Case EnuModoInstanciaObjDef.enuUnico
                    '
            End Select
            Dim lstrCamposSelect As String()
            HobjPadre = Nothing
            HblnEsCreable = False
            HblnEsModificable = True
            HblnEsSuprimible = False
            HblnEsAnulable = False
            If Not IsNothing(GobjPanorama.ObjUsuarioActual) Then
                Dim lstrIdusuario As String = GobjPanorama.ObjUsuarioActual.ObjIdUsuarioStr.ObjValorPro
                If lstrIdusuario = GCSTRADMIN OrElse lstrIdusuario = GCSTRUSUARIOU Then
                    HblnEsModificable = True
                End If
            End If
            HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
            lstrCamposSelect = {"*"}
            HcolTablas.Add(MCSTRNOMBRETABLA)
            HcolCamposSelect.Add(lstrCamposSelect)
        End Sub
        ''' <summary>
        ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
        ''' </summary>
        ''' <param name="aobjPadre">Objeto padre al cual pertenece el objeto que se esta instanciando</param>
        ''' <param name="adrwAplicacion">DataRow que contiene los valores de las propiedades del objeto</param>
        ''' <remarks></remarks>
        Friend Sub New(aobjPadre As ClsAdministrador, adrwAplicacion As DataRow)
            HobjPadre = aobjPadre
            HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
            HblnEsSuprimible = False
            HblnEsAnulable = False
            '
            DrwRegistroActual = adrwAplicacion
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
                Return EnuIdClasesPanDef.enuAplicacion
            End Get
        End Property
        Protected Overrides ReadOnly Property HstrNombreClase As String
            Get
                Return "Aplicación"
            End Get
        End Property
#End Region
#Region "Propiedades Prop"
        Friend ReadOnly Property ObjActivaProgramaBKBln As New ClsActivaProgramaBKBln(Me)
        Friend ReadOnly Property ObjDiaCopiaSeguridadEnt As New ClsDiaCopiaSeguridadEnt(Me)
        Friend ReadOnly Property ObjFechaActualizacionDtm As New ClsFechaActualizacionDtm(Me)
        Friend ReadOnly Property ObjFechaInstalacionDtm As New ClsFechaInstalacionDtm(Me)
        Friend ReadOnly Property ObjHoraCopiaSeguridadEnt As New ClsHoraCopiaSeguridadEnt(Me)
        Friend ReadOnly Property ObjIdAppShr As New ClsIdAppShr(Me)
        Friend ReadOnly Property ObjIdLicenciaShr As New ClsIdLicenciaShr(Me)
        Friend ReadOnly Property ObjIdTerceroLicenciaDbl As New ClsIdTerceroLicenciaDbl(Me)
        Friend ReadOnly Property ObjMinutosCopiaSeguridadEnt As New ClsMinutosCopiaSeguridadEnt(Me)
        Friend ReadOnly Property ObjTrayCopiaSeguridadStr As New ClsTrayCopiaSeguridadStr(Me)
        Friend ReadOnly Property ObjVersionStr As New ClsVersionStr(Me)
        Friend Overrides ReadOnly Property ColPropiedades As Collection
            Get
                If HcolPropiedades.Count = 0 Then
                    HcolPropiedades.Add(ObjActivaProgramaBKBln)
                    HcolPropiedades.Add(ObjDiaCopiaSeguridadEnt)
                    HcolPropiedades.Add(ObjFechaActualizacionDtm)
                    HcolPropiedades.Add(ObjFechaInstalacionDtm)
                    HcolPropiedades.Add(ObjHoraCopiaSeguridadEnt)
                    HcolPropiedades.Add(ObjIdAppShr)
                    HcolPropiedades.Add(ObjIdLicenciaShr)
                    HcolPropiedades.Add(ObjIdTerceroLicenciaDbl)
                    HcolPropiedades.Add(ObjMinutosCopiaSeguridadEnt)
                    HcolPropiedades.Add(ObjTrayCopiaSeguridadStr)
                    HcolPropiedades.Add(ObjVersionStr)
                End If
                Return HcolPropiedades
            End Get
        End Property
#End Region
#Region "Otras Propiedades"
        Friend ReadOnly Property ColAcciones
            Get
                If Not MblnColAccionesPoblada Then
                    If Not IsNothing(McolAcciones) Then
                        McolAcciones.Clear()
                    Else
                        McolAcciones = New Collection
                    End If
                    Dim lstrFiltro As String = "IdAplicacion = " & ObjIdAppShr.ToString
                    Dim ldrwAcciones As DataRow() = ClsPanorama.FdrwDataRow("PanTblAcciones", {"*"},
                            {{"IdAplicacion", "ASC"}, {"Ordinal", "ASC"}}, lstrFiltro)
                    For Each ldrwAccion As DataRow In ldrwAcciones
                        Dim lobjAcción As New ClsAccion(Me, ldrwAccion)
                        lobjAcción.SLeaValores(True)
                        McolAcciones.Add(lobjAcción, lobjAcción.ObjIdObjetoShr.ToString & "_" &
                                lobjAcción.ObjOrdinalShr.ToString)
                    Next
                    MblnColAccionesPoblada = True
                End If
                Return McolAcciones
            End Get
        End Property
        Friend ReadOnly Property ObjDefApp As ClsDefAplicacion
            Get
                If IsNothing(MobjDefApp) Then
                    Dim lobjValorLlave() = {ObjIdAppShr.ObjValorPro}
                    MobjDefApp = New ClsDefAplicacion(EnuModoInstanciaObjDef.enuUnico)
                    MobjDefApp.SAbra(lobjValorLlave)
                End If
                Return MobjDefApp
            End Get
        End Property
        Friend ReadOnly Property StrNombreCompleto As String
            Get
                If BlnExiste Then
                    Return ObjDefApp.StrNombreCompleto
                Else
                    Return ""
                End If
            End Get
        End Property
        Friend ReadOnly Property StrNombreLicencia() As String
            Get
                Dim lstrNomLicencia As String = String.Empty
                Dim lintIdLicencia As Integer = ObjIdLicenciaShr.ObjValorPro
                If ObjIdAppShr.ObjValorPro <> EnuListaAplicaciones.EnuAdministrador Then
                    If lintIdLicencia > 0 Then
                        lstrNomLicencia = ObjDefApp.ObjPrefijoAppStr.ObjValorPro & "-" &
                                ObjDefApp.ObjPrefijoEdicionStr.ObjValorPro &
                                "-" & CType(lintIdLicencia, String).PadLeft(5, "0")
                    Else
                        lstrNomLicencia = "Transitoria"
                    End If
                End If
                Return lstrNomLicencia
            End Get
        End Property
        Friend ReadOnly Property StrTrayCopiaSeguridad As String
            Get
                If Not String.IsNullOrEmpty(ObjTrayCopiaSeguridadStr.ToString) Then
                    MstrTrayCopiaSeguridad = ObjTrayCopiaSeguridadStr.ToString
                    If Not Computer.FileSystem.DirectoryExists(MstrTrayCopiaSeguridad) Then
                        MstrTrayCopiaSeguridad = GstrTrayDat & "CopiasSeguridad"
                    End If
                Else
                    MstrTrayCopiaSeguridad = GstrTrayDat & "CopiasSeguridad"
                End If
                Return MstrTrayCopiaSeguridad
            End Get
        End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
        Public Overrides Function FblnEsAnulable() As Boolean
            Return MyBase.FblnEsAnulable()
        End Function
        Friend Overrides Function FblnEsSuprimible() As Boolean
            Return MyBase.FblnEsSuprimible()
        End Function
        Friend Sub SDeterminePermisos()
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                If Not CType(EnuPermisosObj And EnuPermisosDef.enuCrear, Boolean) Then
                    EnuPermisosObj += EnuPermisosDef.enuModificar
                End If
            Else
                If Not IsNothing(GobjPanorama.ObjUsuarioActual) Then
                    EnuPermisosObj = GobjPanorama.FenuTipoPermisos(HenuIdClase)
                Else
                    If Not CType(EnuPermisosObj And EnuPermisosDef.enuConsCrear, Boolean) Then
                        EnuPermisosObj += EnuPermisosDef.enuConsCrear
                    End If
                    If Not CType(EnuPermisosObj And EnuPermisosDef.enuModificar, Boolean) Then
                        EnuPermisosObj += EnuPermisosDef.enuModificar
                    End If
                End If
            End If
        End Sub
#End Region
#Region "Manejo de Origenes de Instancia"
        Friend ReadOnly Property ColOrigenesInstancia(ablnRefrescar As Boolean) As Collection
            Get
                SPuebleColOrigenesInst(ablnRefrescar)
                Return McolOrigenesInstancias
            End Get
        End Property
        ''' <summary>
        ''' Habilita o deshabilita el Origen para iniciar sesión en la aplicación.
        ''' </summary>
        ''' <param name="astrNombreOrigen">Nombre del origen de instanciamiento que se va a habilitar
        ''' o deshabilitar</param>
        ''' <param name="ablnHabilita">Indica si se habilita o deshabilita el Origen de instanciamiento</param>
        ''' <remarks></remarks>       
        Private Sub SPuebleColOrigenesInst(ablnRefrescar As Boolean)
            If ablnRefrescar OrElse Not MblnColOrigenesInstPoblada Then
                If Not IsNothing(McolOrigenesInstancias) Then
                    McolOrigenesInstancias.Clear()
                Else
                    McolOrigenesInstancias = New Collection
                End If
                SCargueDtbOrigenesInstancia(ablnRefrescar)
                Dim ldrwOrigenInst As DataRow() = MdtbOrigenesInstancias.Select
                For Each ldrwOrigIns As DataRow In ldrwOrigenInst
                    Dim lobjOrigenInstancia As New ClsOrigenInstancia(Me, ldrwOrigIns)
                    lobjOrigenInstancia.SLeaValores(True)
                    McolOrigenesInstancias.Add(lobjOrigenInstancia,
                                lobjOrigenInstancia.ObjNombreOrigenStr.ToString)
                Next
                MblnColOrigenesInstPoblada = True
            End If
        End Sub
        Private Sub SCargueDtbOrigenesInstancia(ablnRefrescar As Boolean)
            If Not MblnDtbOrigenesInstCargada OrElse ablnRefrescar Then
                Dim lstrFiltro As String = ObjIdAppShr.StrNombreCampoBD & " = " & ObjIdAppShr.ToString
                MdtbOrigenesInstancias = ClsPanorama.FdtbDataTable("PanOrigenesInstancias", {"*"},
                        {{ObjIdAppShr.StrNombreCampoBD, "ASC"}, {"Nombre", "ASC"}}, lstrFiltro)
                MblnDtbOrigenesInstCargada = True
            End If
        End Sub
        Private Shared Sub SActualiceOrigen(aobjOrigenInstancia As ClsOrigenInstancia,
                aenuOrigenInstancia As EnuOrigenInstanciamientoDef, ablnLogin As Boolean)
            With aobjOrigenInstancia
                .EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                If ablnLogin Then
                    .ObjIdUsuarioActivoStr.ObjValorPro = GobjPanorama.ObjUsuarioActual.ObjIdUsuarioStr.ObjValorPro
                    .ObjEstaOrigenActivoBln.ObjValorPro = True
                    .ObjFechaActivacionDtm.ObjValorPro = Now
                    .ObjOrigenInstanciaByt.ObjValorPro = CType(aenuOrigenInstancia, Byte)
                Else
                    .ObjIdUsuarioActivoStr.ObjValorPro = String.Empty
                    .ObjEstaOrigenActivoBln.ObjValorPro = False
                End If
                .SActualice(True)
            End With
        End Sub
        Private Sub SAdicioneOrigen(astrNombreOrigen As String,
                aenuOrigenInstancia As EnuOrigenInstanciamientoDef)
            Dim ldrwOrigIns As DataRow = MdtbOrigenesInstancias.NewRow
            Dim lobjOrigenInstancia As New ClsOrigenInstancia(Me, ldrwOrigIns)
            If Not CType(lobjOrigenInstancia.EnuPermisosObj And EnuPermisosDef.enuCrear, Boolean) Then
                lobjOrigenInstancia.EnuPermisosObj += EnuPermisosDef.enuCrear
            End If
            lobjOrigenInstancia.SCreeObj(Nothing)
            With lobjOrigenInstancia
                .ObjIdAppOrigenShr.ObjValorPro = ObjIdAppShr.ObjValorPro
                .ObjNombreOrigenStr.ObjValorPro = astrNombreOrigen
                .ObjFechaCreacionDtm.ObjValorPro = Now
                .ObjFechaActivacionDtm.ObjValorPro = Now
                .ObjOrigenInstanciaByt.ObjValorPro = CType(aenuOrigenInstancia, Byte)
                .ObjEstaOrigenActivoBln.ObjValorPro = True
                .ObjIdUsuarioActivoStr.ObjValorPro = GobjPanorama.ObjUsuarioActual.ObjIdUsuarioStr.ObjValorPro
                .SActualice(True)
            End With
            McolOrigenesInstancias.Add(lobjOrigenInstancia, lobjOrigenInstancia.ObjNombreOrigenStr.ToString)
        End Sub
#End Region
#Region "Manejo de Perfiles"
        Friend ReadOnly Property ColPerfiles(ablnRefrsque As Boolean) As Collection
            Get
                SPuebleColPerfiles(ablnRefrsque)
                Return McolPerfiles
            End Get
        End Property
        ''' <summary>
        ''' Crea un nuevo perfil para la aplicación y lo adiciona a la coleccion de perfiles de la misma.
        ''' </summary>
        ''' <param name="astrNombrePerfil">Nombre que tendrá el nuevo perfil</param>
        ''' <returns>El nuevo objeto Perfil</returns>
        ''' <remarks></remarks>
        Friend Function FobjNuevoPerfil(astrNombrePerfil As String)
            Dim lobjPerfil As ClsPerfil
            lobjPerfil = FobjNuevoPerfil()
            If Not IsNothing(lobjPerfil) Then
                lobjPerfil.ObjNombrePerfilStr.ObjValorPro = astrNombrePerfil
                lobjPerfil.SActualice(True)
                SAdicionePerfil(lobjPerfil)
            End If
            Return lobjPerfil
        End Function
        ''' <summary>
        ''' Devuelve un nuevo objeto Perfil cuyas propiedades tienen por valor el valor por defecto.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Function FobjNuevoPerfil() As ClsPerfil
            SCargueDtbPerfiles(False)
            If Not IsNothing(MdtbPerfiles) Then
                Dim ldrwPerfil As DataRow = MdtbPerfiles.NewRow
                Dim lobjPerfil As New ClsPerfil(Me, ldrwPerfil, ColAcciones)
                lobjPerfil.SCreeObj(Nothing)
                lobjPerfil.ObjIdAppPerfilShr.ObjValorPro = ObjIdAppShr.ObjValorPro
                Return lobjPerfil
            End If
            Return Nothing
        End Function
        Private Sub SPuebleColPerfiles(ablnRefresque As Boolean)
            If Not MblnColPerfilesPoblada OrElse ablnRefresque Then
                SCargueDtbPerfiles(ablnRefresque)
                If Not IsNothing(MdtbPerfiles) Then
                    If IsNothing(McolPerfiles) Then
                        McolPerfiles = New Collection
                    Else
                        McolPerfiles.Clear()
                    End If
                    If MdtbPerfiles.Rows.Count > 0 Then
                        For Each ldrwPerfil As DataRow In MdtbPerfiles.Rows
                            Dim lobjPerfil As New ClsPerfil(Me, ldrwPerfil)
                            lobjPerfil.SLeaValores(True)
                            Dim lstrKey = lobjPerfil.ObjIdAppPerfilShr.ToString & lobjPerfil.ObjIdPerfilShr.ToString
                            McolPerfiles.Add(lobjPerfil, lstrKey)
                        Next
                    End If
                    MblnColPerfilesPoblada = True
                End If
            End If
        End Sub
        Private Sub SCargueDtbPerfiles(ablnRefresque As Boolean)
            If Not MblnDtbPerfilesCargada OrElse ablnRefresque Then
                Dim lstrFiltro As String = ObjIdAppShr.StrNombreCampoBD & " = " & ObjIdAppShr.ToString
                MdtbPerfiles = ClsPanorama.FdtbDataTable("PanPerfiles", {"*"},
                        {{"IdAplicacion", "ASC"}, {"IdPerfil", "ASC"}}, lstrFiltro,
                        True, Array.Empty(Of String))
                MblnDtbPerfilesCargada = True
            End If
        End Sub
        ''' <summary>
        ''' Devuelve un buleano que indica si el nombre del perfil pasado en el argumento ya existe en los 
        ''' perfiles de la aplicacion
        ''' </summary>
        ''' <param name="astrNOmbrePerfil">Nombre del perfil que sera examinado</param>
        ''' <returns>Buleano</returns>
        ''' <remarks></remarks>
        Friend Function FblnExisteNombrePerfil(astrNombrePerfil As String) As Boolean
            Dim lblnExiste As Boolean
            SCargueDtbPerfiles(False)
            Dim lstrFiltro = "Nombre = '" & astrNombrePerfil & "'"
            Dim ldrwPerfiles() = MdtbPerfiles.Select(lstrFiltro)
            lblnExiste = (ldrwPerfiles.Count > 0)
            Return lblnExiste
        End Function
        ''' <summary>
        ''' Adiciona el perfil pasado en el argumento a la colección de perfiles de la aplicación.
        ''' </summary>
        ''' <param name="aobjPerfil">Objeto perfil a ser adicionado a la colleción de perfiles.</param>
        ''' <remarks>Normalmente se adiciona un perfil recien creado.</remarks>
        Public Sub SAdicionePerfil(aobjPerfil As ClsPerfil)
            If aobjPerfil Is Nothing Then
                Throw New ArgumentNullException(NameOf(aobjPerfil))
            End If
            Dim lstrKey As String = aobjPerfil.ObjIdAppPerfilShr.ToString & aobjPerfil.ObjIdPerfilShr.ToString
            If Not MblnColPerfilesPoblada Then
                McolPerfiles = ColPerfiles(False)
            End If
            If Not McolPerfiles.Contains(lstrKey) Then
                McolPerfiles.Add(aobjPerfil, lstrKey)
            End If
        End Sub
#End Region
#Region "Procedimientos Propios"
        ''' <summary>
        ''' Determina si el origen definido en los parametros existe; si es así lo actualiza de lo contrario 
        ''' verifica que haya origenes licenciados disponibles y si es así lo crea. Si el origen existe pero 
        ''' esta deshabilitado genera un mensaje, lo mismo sucede si la cantidad licenciada de origenes esta
        ''' copada.  
        ''' </summary>
        ''' <param name="astrNombreOrigen">Nombre del origen</param>
        ''' <param name="aenuOrigenInstancia">Tipo del origen que esta instanciando</param>
        ''' <remarks></remarks>
        Friend Sub SRegistreLogOnOrigen(astrNombreOrigen As String,
                aenuOrigenInstancia As EnuOrigenInstanciamientoDef, ByRef astrMens As String)
            GobjPanDat.SControleProcesoObj(True)
            Try
                If Not MblnColOrigenesInstPoblada Then
                    McolOrigenesInstancias = ColOrigenesInstancia(False)
                End If
                If McolOrigenesInstancias.Contains(astrNombreOrigen) Then
                    Dim lobjOrigenInstancia As ClsOrigenInstancia = McolOrigenesInstancias(astrNombreOrigen)
                    With lobjOrigenInstancia
                        Dim lblnEstaActiva As Boolean = .ObjEstaOrigenActivoBln.ObjValorPro
                        SActualiceOrigen(lobjOrigenInstancia, aenuOrigenInstancia, True)
                        If lblnEstaActiva Then
                            astrMens = "La Sesión anterior fue cerrada inesperadamente.!" &
                                    vbCrLf & "Si tiene algún problema por favor informe a Soporte."
                        End If
                    End With
                Else
                    SAdicioneOrigen(astrNombreOrigen, aenuOrigenInstancia)
                End If
                For Each lobjOriIns As ClsOrigenInstancia In McolOrigenesInstancias
                    If lobjOriIns.ObjFechaActivacionDtm.ObjValorPro < Date.Today Then
                        SActualiceOrigen(lobjOriIns, aenuOrigenInstancia, False)
                    End If
                Next
            Catch ex As PanLException
                Throw
            Catch ex As PanDatException
                Throw
            Catch ex As Exception
                Throw
            Finally
                GobjPanDat.SControleProcesoObj(False)
            End Try
        End Sub
        ''' <summary>
        ''' Registra el logout del origen cuando este termina la sesión.
        ''' </summary>
        ''' <param name="astrNombreOrigen">Nombre del origen</param>
        ''' <param name="aenuOrigenInstancia">Tipo del origen que esta terminando sesión</param>
        ''' <remarks></remarks>
        Public Sub SRegistreLogOffOrigen(astrNombreOrigen As String,
                aenuOrigenInstancia As EnuOrigenInstanciamientoDef)
            Dim lobjOrigenInstancia As ClsOrigenInstancia = McolOrigenesInstancias(astrNombreOrigen)
            SActualiceOrigen(lobjOrigenInstancia, aenuOrigenInstancia, False)
        End Sub
#End Region
    End Class
#Region "Clases de Propiedad"
    Friend Class ClsActivaProgramaBKBln
        Inherits ClsCBPropiedad
        Private Const MCSTRNOMBRECAMPOBD As String = "ActivaProgramaBK"
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            HstrNombre = "Activa Copia Seguridad Programada"
            HenuTipoValor = EnuTipoValor.enuBoolean
            HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
            HblnRegistrarLogCambio = True
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
            If IsNothing(HobjValorPro) Then
                Return "Falso"
            Else
                Return ClsPanorama.FstrBuleanoToString(HobjValorPro)
            End If
        End Function
    End Class
    Friend Class ClsDiaCopiaSeguridadEnt
        Inherits ClsCBPropiedad
        Private ReadOnly MobjPadre As ClsAplicacion = Nothing
        Private Const MCSTRNOMBRECAMPOBD As String = "DiaBK"
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            MobjPadre = aobjPadre
            HstrNombre = "Dia Copia Seguridad"
            HenuTipoValor = EnuTipoValor.enuInteger
            HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
            HblnRegistrarLogCambio = True
        End Sub
        Public Overrides Sub SValide()
            HblnEsRequerido = MobjPadre.ObjActivaProgramaBKBln.ObjValorPro
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, DayOfWeek.Sunday,
                    DayOfWeek.Saturday, BlnEsRequerido)
        End Sub
        Friend Shared ReadOnly Property SstrNombreCampoBd As String
            Get
                Return MCSTRNOMBRECAMPOBD
            End Get
        End Property
        Public Overrides Function ToString() As String
            If IsNothing(ObjValorPro) Then
                Return DayOfWeek.Sunday
            Else
                Return HobjValorPro.ToString
            End If
        End Function
    End Class
    Friend Class ClsFechaActualizacionDtm
        Inherits ClsCBPropiedad
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            HstrNombre = "FechaActualizacion"
            HenuTipoValor = EnuTipoValor.enuDate
            HstrNombreCampoBd = "FechaActualizacion"
            HblnEsRequerido = True
            HblnRegistrarLogCambio = True
        End Sub
        Public Overrides Sub SValide()
            Dim ldtmFechaActualizaMinima As Date = #1/1/1995#
            Dim ldtmFechaActualizaMaxima As Date = Today.AddDays(1)
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaActualizaMinima,
                    ldtmFechaActualizaMaxima, BlnEsRequerido)
        End Sub
        Public Overrides Function ToString() As String
            If ObjValorPro <> GCDTMFECHANULA Then
                Return HobjValorPro.ToShortDateString
            Else
                Return ""
            End If
        End Function
    End Class
    Friend Class ClsFechaInstalacionDtm
        Inherits ClsCBPropiedad
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            HstrNombre = "FechaInstalacion"
            HenuTipoValor = EnuTipoValor.enuDate
            HstrNombreCampoBd = "FechaInstalacion"
            HblnEsRequerido = True
        End Sub
        Public Overrides Sub SValide()
            Dim ldtmFechaInstalMinima As Date = #1/1/1995#
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaInstalMinima,
                    Date.Today, BlnEsRequerido)
        End Sub
        Public Overrides Function ToString() As String
            If ObjValorPro <> GCDTMFECHANULA Then
                Return HobjValorPro.ToShortDateString
            Else
                Return ""
            End If
        End Function
    End Class
    Friend Class ClsHoraCopiaSeguridadEnt
        Inherits ClsCBPropiedad
        Private ReadOnly MobjPadre As ClsAplicacion = Nothing
        Private Const MCSTRNOMBRECAMPOBD As String = "HoraBK"
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            MobjPadre = aobjPadre
            HstrNombre = "Hora Copia Seguridad"
            HenuTipoValor = EnuTipoValor.enuInteger
            HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
            HblnRegistrarLogCambio = True
        End Sub
        Public Overrides Sub SValide()
            HblnEsRequerido = MobjPadre.ObjActivaProgramaBKBln.ObjValorPro
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                    23, BlnEsRequerido)
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
    Friend Class ClsIdAppShr
        Inherits ClsCBPropiedad
        Private ReadOnly MobjPadre As ClsAplicacion = Nothing
        Private Const MCSTRNOMBRECAMPOBD As String = "IdAplicacion"
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            MobjPadre = aobjPadre
            HstrNombre = "IdApp"
            HenuTipoValor = EnuTipoValor.enuShort
            HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
            HblnEsRequerido = True
            HblnEsLlave = True
            HbytPosicionLlave = 0
        End Sub
        Public Overrides Sub SValide()
            Dim lobjLlavePrincipal(0) As Object
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew,
                    EnuListaAplicaciones.EnuAdministrador, EnuListaAplicaciones.EnuOrionCop,
                    BlnEsRequerido, EnuTipoValor.enuShort)
            If Not BlnLeyendoOrigen Then
                HstrMens = String.Empty
                If HblnEsValido Then
                    lobjLlavePrincipal(0) = HobjValorNew
                    If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                        If Not MobjPadre.FblnEsCreable(lobjLlavePrincipal) Then
                            HstrMens = "La Id. de la Aplicacion ingresada ya existe!"
                            HblnEsValido = False
                        End If
                    ElseIf MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                        If HobjValorNew <> HobjValorPro OrElse (Not MobjPadre.BlnExiste) Then
                            MobjPadre.SAbra(lobjLlavePrincipal)
                        End If
                        If Not MobjPadre.BlnExiste Then
                            HstrMens = "La Id. de la Aplicacion ingresada no existe!"
                            If Not IsNothing(MobjPadre.ObjValorUltimaLlave) Then
                                MobjPadre.SAbra(MobjPadre.ObjValorUltimaLlave)
                            Else
                                HblnEsValido = False
                            End If
                        End If
                    End If
                Else
                    HstrMens = "La Id. de la Aplicacion ingresada no es válida!"
                End If
                If Not String.IsNullOrEmpty(HstrMens) Then
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
    Friend Class ClsIdLicenciaShr
        Inherits ClsCBPropiedad
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            HstrNombre = "IdLicencia"
            HenuTipoValor = EnuTipoValor.enuShort
            HstrNombreCampoBd = "IdLicencia"
        End Sub
        Public Overrides Sub SValide()
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                    BlnEsRequerido, EnuTipoValor.enuShort)
        End Sub
        Public Overrides Function ToString() As String
            If IsNothing(ObjValorPro) Then
                Return ""
            Else
                Return HobjValorPro
            End If
        End Function
    End Class
    Friend Class ClsIdTerceroLicenciaDbl
        Inherits ClsCBPropiedad
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            HstrNombre = "TerceroLicencia"
            HenuTipoValor = EnuTipoValor.enuDouble
            HstrNombreCampoBd = "IdTerceroLicencia"
            HblnRegistrarLogCambio = True
        End Sub
        Public Overrides Sub SValide()
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC, GCDBLMAXTERC, BlnEsRequerido)
        End Sub
        Public Overrides Function ToString() As String
            Return HobjValorPro.ToString
        End Function
    End Class
    Friend Class ClsMinutosCopiaSeguridadEnt
        Inherits ClsCBPropiedad
        Private ReadOnly MobjPadre As ClsAplicacion = Nothing
        Private Const MCSTRNOMBRECAMPOBD As String = "MinutosBK"
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            MobjPadre = aobjPadre
            HstrNombre = "Minutos Copia Seguridad"
            HenuTipoValor = EnuTipoValor.enuInteger
            HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
            HblnRegistrarLogCambio = True
        End Sub
        Public Overrides Sub SValide()
            HblnEsRequerido = MobjPadre.ObjActivaProgramaBKBln.ObjValorPro
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                    59, BlnEsRequerido)
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
    Friend Class ClsTrayCopiaSeguridadStr
        Inherits ClsCBPropiedad
        Private ReadOnly Property MobjPadre As ClsAplicacion
        Private Const MCSTRNOMBRECAMPOBD As String = "TrayCopiaSeguridad"
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            MobjPadre = aobjPadre
            HstrNombre = "Trayectoria Copia Seguridad"
            HshrLongitud = 200
            HenuTipoValor = EnuTipoValor.enuString
            HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
            HblnRegistrarLogCambio = True
            HblnRegistrarLogCambio = True
        End Sub
        Public Overrides Sub SValide()
            HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud,
                    BlnEsRequerido)
            If HblnEsValido Then
                If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                    HblnEsValido = My.Computer.FileSystem.DirectoryExists(HobjValorNew)
                    If Not HblnEsValido Then
                        HstrMens = "La Carpeta ingresada no existe en este Equipo!"
                        SNotifiqueDatInv()
                    End If
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
    Friend Class ClsVersionStr
        Inherits ClsCBPropiedad
        Private Const MCSTRNOMBRECAMPOBD As String = "Version"
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            HstrNombre = "Version"
            HshrLongitud = 20
            HenuTipoValor = EnuTipoValor.enuString
            HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        End Sub
        Public Overrides Sub SValide()
            HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 2,
                    ShrLongitud, BlnEsRequerido)
            Dim lbytCanPuntos As Byte = 0
            If Not BlnLeyendoOrigen Then
                If HblnEsValido AndAlso Not IsNothing(HobjValorNew) AndAlso HobjValorNew <> "" Then
                    For Each car As Char In HobjValorNew.ToString
                        If Not Char.IsNumber(car) Then
                            If car <> "." Then
                                HblnEsValido = False
                                Exit For
                            Else
                                lbytCanPuntos += 1
                            End If
                        End If
                    Next
                    If lbytCanPuntos <> 3 Then
                        HblnEsValido = False
                    End If
                End If
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
#End Region
End Namespace