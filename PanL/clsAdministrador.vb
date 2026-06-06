Friend Class ClsAdministrador
#Region "Definiciones"
    Implements IDisposable
    Implements IPanDat
    ' Variables
    Private MblnDisposed As Boolean = False
    Private Const MCUSHIDOBJETO As UShort = 1
    Private ReadOnly Property MobjRegistro As Object = Nothing
    Public Property BlnNotificacionSonora As Boolean = True
#End Region

#Region "Constructores" 'Ok
    ''' <summary>
    ''' Crea una nueva instancia de la clase cuando se esta abriendo la aplicacion "Administrador"
    ''' </summary>
    Public Sub New(aobjRegistro As Object)
        If aobjRegistro Is Nothing OrElse Not (aobjRegistro.GetType.Name = "String" AndAlso
                aobjRegistro = GCOBJREGISTRO) Then
            Throw New ModuloNoRegistradoPanException("El módulo no ha sido debidamente cargado!")
        End If
        If GenuIdAplicacion = EnuListaAplicaciones.EnuAdministrador Then
            MobjRegistro = aobjRegistro
            GobjPanorama = New ClsPanorama(MobjRegistro)
            GobjAdministrador = Me
            GobjPanDat = New ClsPanoramaDat(Me)
        End If
    End Sub
    Friend Sub SInicieApp()
        Dim lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            GobjPanDat.SInicieDat(EntVersionBDEnProg)
            lblnNoHayError = True
        Catch ex As PanDatException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not lblnNoHayError Then
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            Else
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            End If
        End Try
    End Sub
    Friend Sub SInicieSesion(astrVersionApp As String)
        Try
            Dim lstrVerAnt = String.Empty
            If GenuTipoInstanciamiento = EnuTipoInstanciamiento.enuInstalacion Then
                SCreeUsuarioAdmin()
                SRegistreApp(EnuListaAplicaciones.EnuAdministrador, 1000, astrVersionApp)
                SCreePerfilAdmin()
            Else
                GobjPanorama.ObjAppActual = FobjAppActual()
                GstrVerAntApp = GobjPanorama.ObjAppActual.ObjVersionStr.ObjValorPro
                lstrVerAnt = GobjPanorama.ObjAppActual.ObjVersionStr.ObjValorPro
                If GenuTipoInstanciamiento = EnuTipoInstanciamiento.enuActualizacion Then
                    SEstablezcaUsuarioAdmin()
                End If
                If GstrVerAntApp < "8.15.110.18" Then
                    SElimineCarpUsuaUniv()
                End If
            End If
            SRegistreVersion(astrVersionApp)
        Catch ex As PanDatException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Manejo Tablas Constantes"
    Private Shared Function FdsTablaConstantes() As DataSet
        Dim ldsTablas As New DataSet
        Dim lcolTablas As New Collection
        Dim lcolCamposSelect As New Collection
        Dim lcolIndices As New Collection
        Dim lstrCamposSelectCons() As String = {"*"}
        Dim lstrCamposIndiceConstantes(,) As String = {{"IDGrupo", "ASC"}, {"IDConstante", "ASC"}}
        lcolTablas.Add("PanTblConstantes")
        lcolCamposSelect.Add(lstrCamposSelectCons)
        lcolIndices.Add(lstrCamposIndiceConstantes)
        GobjPanDat.SdsDataSet(ldsTablas, lcolTablas, lcolCamposSelect, lcolIndices, Nothing)
        Return ldsTablas
    End Function

    ''' <summary>
    ''' Devuelve una Array de DatRows que contienen los datos de las constates manejadas por el programa generalmente
    ''' a través de enums, según el grupo pasado en el argumento "aenurupoTbl"
    ''' </summary>
    ''' <param name="aenurupoCte">un elemento del enum "enuGrupoConstantesPan" que indica el grupo de 
    ''' constantes requerido</param>
    ''' <returns>Array de DataRows</returns>
    ''' <remarks></remarks>
    Friend Shared Function FdrwConstantesPan(aenuGrupoCte As EnuGrupoConstantesPanDef) As DataRow()
        Dim ldsConstantes = FdsTablaConstantes()
        Dim ldtbConstantes As DataTable = ldsConstantes.Tables("PanTblConstantes")
        Dim lstrCondicion As String = " IDGrupo = " & CStr(aenuGrupoCte)
        Dim ldrwTblCons() As DataRow = ldtbConstantes.Select(lstrCondicion)
        Return ldrwTblCons
    End Function

    ''' <summary>
    ''' Devuelve una Array de DataRows que contienen los elementos de los "Grupos de Terceros"
    ''' </summary>
    ''' <returns>Array de DataRows</returns>
    ''' <remarks></remarks>
    Public Shared Function FdrwGruposTerceros() As DataRow()
        Dim ldsConstantes = FdsTablaConstantes()
        Dim ldtbGruposTerceros As DataTable = ldsConstantes.Tables("TblGruposTerceros")
        Dim ldrwTblGruposTerceros() As DataRow = ldtbGruposTerceros.Select
        Return ldrwTblGruposTerceros
    End Function

    ''' <summary>
    ''' Devuelve el dato contenido en el campo "Dato" del registro correspondiente al valor pasado en el
    ''' argumento "abytIdDato" del grupo de constantes pasado en el argumento "aenurupoTbl"
    ''' </summary>
    ''' <param name="aenurupoConstantes">Enu que identifica el grupo de constantes donde se buscara el dato</param>
    ''' <param name="abytIdDato">Id que identifica el dato entre los registros del grupo</param>
    ''' <returns>Nombre de la constante</returns>
    ''' <remarks></remarks>
    Public Shared Function FstrNombreDatoConstantePan(aenurupoConstantes As EnuGrupoConstantesPanDef,
                abytIdDato As Byte)
        Dim ldrwConstates As DataRow() = FdrwConstantesPan(aenurupoConstantes)
        For Each ldrwDataRow As DataRow In ldrwConstates
            If ClsPanorama.FobjValorCampo(ldrwDataRow("IdConstante"), EnuTipoValor.enuByte) = abytIdDato Then
                Return ldrwDataRow("Dato")
            End If
        Next
        Return ""
    End Function
#End Region

#Region "Propiedades"
    Shared ReadOnly Property EntIdClase As Integer
        Get
            Return CType(EnuIdClasesPanDef.EnuAdministrador, Integer)
        End Get
    End Property
    Friend Shared Function FobjAppActual() As ClsAplicacion
        Dim lobjApp As ClsAplicacion
        Dim lobjLlave As Object() = {CType(GenuIdAplicacion, Short)}
        lobjApp = New ClsAplicacion(EnuModoInstanciaObjDef.enuUnico)
        lobjApp.SDeterminePermisos()
        lobjApp.SAbra(lobjLlave)
        Return lobjApp
    End Function
#Region "Implementa IPanDat"
    Friend ReadOnly Property ShrIdApp As Short Implements IPanDat.ShrIdApp
        Get
            Return 100
        End Get
    End Property
    Friend ReadOnly Property StrNombreArchivos As String Implements IPanDat.StrNombreArchivos
        Get
            Return My.Resources.NombreArchivos
        End Get
    End Property
    Friend ReadOnly Property EntVersionBDEnProg As Integer Implements IPanDat.EntVersionBDEnProg
        Get
            Return 104
        End Get
    End Property
    Friend ReadOnly Property ObjRegistro As Object Implements IPanDat.ObjRegistro
        Get
            Return MobjRegistro
        End Get
    End Property
#End Region
#End Region

#Region "Procedimientos de Instalación" 'Ok
    Friend Sub SRegistreApp(ashrIdApp As Short, ashrCantidadLicenciada As Short, astrVerApp As String)
        Dim lushIdApp = ashrIdApp
        Dim ldtbApps = FdtbApps()
        Dim ldtmFechaLicenciamiento As Date = GCDTMFECHANULA
        If lushIdApp = EnuListaAplicaciones.EnuAdministrador Then
            ldtmFechaLicenciamiento = Date.Today
        End If
        Dim ldrwApp As DataRow = ldtbApps.NewRow
        Dim lobjApp As New ClsAplicacion(Me, ldrwApp)
        If Not CType(lobjApp.EnuPermisosObj And EnuPermisosDef.enuCrear, Boolean) Then
            lobjApp.EnuPermisosObj += EnuPermisosDef.enuCrear
        End If
        lobjApp.SCreeObj(Nothing)
        With lobjApp
            .ObjActivaProgramaBKBln.ObjValorPro = False
            .ObjDiaCopiaSeguridadEnt.ObjValorPro = 0
            .ObjHoraCopiaSeguridadEnt.ObjValorPro = 0
            .ObjMinutosCopiaSeguridadEnt.ObjValorPro = 0
            .ObjIdAppShr.ObjValorPro = lushIdApp
            .ObjIdLicenciaShr.ObjValorPro = 0
            .ObjIdTerceroLicenciaDbl.ObjValorPro = 0
            .ObjFechaActualizacionDtm.ObjValorPro = Date.Today
            .ObjFechaInstalacionDtm.ObjValorPro = Date.Today
            .ObjVersionStr.ObjValorPro = astrVerApp
            .ObjTrayCopiaSeguridadStr.ObjValorPro = .StrTrayCopiaSeguridad
            .SActualice(True)
        End With
        GobjPanorama.ObjAppActual = lobjApp
    End Sub
    Private Sub SCreeUsuarioAdmin()
        Dim ldtbUsuarios As DataTable = ClsPanorama.FdtbDataTable("PanUsuarios", {"*"},
                {{"IdUsuario", "ASC"}}, "", True, Array.Empty(Of String))
        GstrIdUsuario = GCSTRUSUARIOU
        Dim lstrContrasena As String = "admin"
        Dim ldrwUsuarioAdmin As DataRow = ldtbUsuarios.NewRow
        Dim lobjUsuario = New ClsUsuario(Me, ldrwUsuarioAdmin)
        lobjUsuario.SCreeObj(Nothing)
        With lobjUsuario
            .ObjIdUsuarioStr.ObjValorPro = "Admin"
            .ObjContrasenaUsuarioStr.ObjValorPro = lstrContrasena
            .ObjContrasenaUsuarioStr.SConfirmeContrasena(lstrContrasena)
            .ObjNombreUsuarioStr.ObjValorPro = "Administrador"
            .ObjFechaCreacionDtm.ObjValorPro = Date.Now
            .ObjFechaCambioContrasenaDtm.ObjValorPro = Date.Now
            .ObjEstaActivoUsuarioBln.ObjValorPro = True
            .ObjTipoCambioContrasenaByt.ObjValorPro =
                    EnuTipoCambioContrasenaDef.enuProximaVez
            .SActualice(True)
            GobjPanorama.ObjUsuarioActual = lobjUsuario
            GstrIdUsuario = lobjUsuario.ObjIdUsuarioStr.ObjValorPro
        End With
    End Sub
    Private Shared Sub SEstablezcaUsuarioAdmin()
        Dim lblnNoHayError As Boolean
        Dim lobjValorLlave As Object() = {GCSTRADMIN}
        Dim lobjUsuario As ClsUsuario
        Try
            lobjUsuario = New ClsUsuario(EnuModoInstanciaObjDef.enuUnico, False)
            lobjUsuario.EnuPermisosObj += EnuPermisosDef.enuModificar
            If GstrIdUsuario = GCSTRUSUARIOU Then
                lobjUsuario.SLeaUsuarioUniversal()
            Else
                lobjUsuario.SAbra(lobjValorLlave)
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
        End Try
        If lblnNoHayError Then
            GobjPanorama.ObjUsuarioActual = lobjUsuario
            GstrIdUsuario = lobjUsuario.ObjIdUsuarioStr.ObjValorPro
        End If
    End Sub
    Private Shared Sub SCreePerfilAdmin()
        Dim lobjPerfil As ClsPerfil = GobjPanorama.ObjAppActual.FobjNuevoPerfil("Admin")
        lobjPerfil.SCambiePermisosAdmin()
        GobjPanorama.ObjAppActual.SAdicionePerfil(lobjPerfil)
        GobjPanorama.ObjUsuarioActual.SAsignePerfil(lobjPerfil)
    End Sub
    Friend Sub SRefresque()
        MblnDisposed = False
    End Sub
#End Region

#Region "Manejo de Aplicaciones"
    Friend Function FcolAppsInstaladas() As Collection
        Dim lcolAppsInstaladas = New Collection
        Dim ldtbApps = FdtbApps()
        For Each ldrwApp As DataRow In ldtbApps.Rows
            Dim lobjApp As New ClsAplicacion(Me, ldrwApp)
            lobjApp.SDeterminePermisos()
            lobjApp.SLeaValores(True)
            lcolAppsInstaladas.Add(lobjApp, lobjApp.ObjIdAppShr.ToString)
        Next
        Return lcolAppsInstaladas
    End Function
    Private Shared Function FdtbApps() As DataTable
        Dim ldtbApps = ClsPanorama.FdtbDataTable("PanAplicaciones", {"*"},
                    {{"IdAplicacion", "ASC"}}, "", True, Array.Empty(Of String))
        Return ldtbApps
    End Function
    Friend Shared Sub SRegistreVersion(astrVersion As String)
        Dim lobjAppActual = FobjAppActual()
        With lobjAppActual
            .SModifique()
            .ObjFechaActualizacionDtm.ObjValorPro = Now
            .ObjVersionStr.ObjValorPro = astrVersion
            .ObjFechaInstalacionDtm.ObjValorPro = Date.Today
            .SActualice(True)
        End With
    End Sub
#End Region

#Region "Manejo de Carpetas"
    Friend Function FcolCarpetas(ablnActivas As Boolean) As Collection
        Dim lcolCarpetas = New Collection
        Dim ldtbCarpetas = FdtbCarpetas()
        For Each ldrwCarpeta As DataRow In ldtbCarpetas.Rows
            Dim lobjCarpeta As New ClsCarpeta(Me, ldrwCarpeta)
            lobjCarpeta.SLeaValores(True)
            If ablnActivas Then
                If lobjCarpeta.ObjEstaActivaBln.ObjValorPro Then
                    lcolCarpetas.Add(lobjCarpeta, lobjCarpeta.ObjIdCarpetaShr.ToString())
                End If
            Else
                lcolCarpetas.Add(lobjCarpeta, lobjCarpeta.ObjIdCarpetaShr.ToString())
            End If
        Next
        Return lcolCarpetas
    End Function
    ''' <summary>
    ''' Devuelve un objeto Carpeta en estado de creación en el cual el valor de todas las propiedades 
    ''' es el valor por defecto.
    ''' </summary>
    ''' <returns>Objeto Carpeta</returns>
    ''' <remarks></remarks>
    Friend Function FobjNuevaCarpeta() As ClsCarpeta ' Ok
        Dim ldtbCarpetas = FdtbCarpetas()
        Dim ldrwCarpeta As DataRow = ldtbCarpetas.NewRow
        Dim lobjCarpeta As New ClsCarpeta(Me, ldrwCarpeta)
        lobjCarpeta.SCreeObj(Nothing)
        Return lobjCarpeta
    End Function
    ''' <summary>
    ''' Actualiza la base de datos con los datos de la Carpeta y la agrega a la Colección de Carpetas
    ''' </summary>
    ''' <param name="aobjCarpeta">Objeto Carpeta a ser adicionado.</param>
    ''' <remarks></remarks>
    Friend Shared Sub SAdicioneCarpeta(aobjCarpeta As ClsCarpeta) ' Ok
        aobjCarpeta.ObjFechaCreacionDtm.ObjValorPro = Date.Now
        aobjCarpeta.SActualice(True)
        ClsPanoramaDat.SCreeCarpeta(aobjCarpeta.ObjNombreStr.ObjValorPro)
    End Sub
    ''' <summary>
    ''' Devuelve el objeto Carpeta identificado por el valor de argumento "ashrIdCarpeta".
    ''' Si la Carpeta no existe devuelve Nothing 
    ''' </summary>
    ''' <param name="ashrIdCarpeta">Identifica la Carpeta que sera devuelta.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Shared Function FobjCarpeta(ashrIdCarpeta As Short) As ClsCarpeta
        Dim lobjValorLlave As Object() = {ashrIdCarpeta}
        Dim lobjCarpeta As New ClsCarpeta(EnuModoInstanciaObjDef.enuUnico, False)
        lobjCarpeta.SAbra(lobjValorLlave)
        If lobjCarpeta.BlnExiste Then
            Return lobjCarpeta
        Else
            Return Nothing
        End If
    End Function
    Private Shared Function FdtbCarpetas() As DataTable
        Dim ldtbCarpetas = ClsPanorama.FdtbDataTable("PanCarpetas", {"*"},
                    {{"IdCarpeta", "ASC"}}, "", True, Array.Empty(Of String))
        Return ldtbCarpetas
    End Function
    Private Function FdtbCarpetasUsuario() As DataTable
        Dim ldtbCarpUsuario As DataTable
        Dim lstrFiltro As String = ""
        ldtbCarpUsuario = ClsPanorama.FdtbDataTable("PanCarpetasUsuarios", {"*"},
                    {{ClsIdUsuarioCarpetaStr.SstrNombreCampoBd, "ASC"},
                    {ClsIdCarpetaUsuarioShr.SstrNombreCampoBd, "ASC"}}, lstrFiltro,
                    True, Array.Empty(Of String)())
        Return ldtbCarpUsuario
    End Function
    Private Sub SElimineCarpUsuaUniv()
        Dim lstrTabla = ClsCarpetaUsuario.SstrNombreTabla
        Dim lcolCamposRef As New Collection
        Dim lcolDatosref As New Collection
        lcolCamposRef.Add(ClsIdUsuarioCarpetaStr.SstrNombreCampoBd)
        lcolDatosref.Add("OPT", ClsIdUsuarioCarpetaStr.SstrNombreCampoBd)
        GobjPanDat.SElimineRegistro(lstrTabla, lcolCamposRef, lcolDatosref)
    End Sub
#End Region

#Region "Manejo de usuarios"
    Friend Function FcolUsuarios(ablnActivos As Boolean) As Collection
        Dim ldtbUsuarios = FdtbUsuarios()
        Dim lcolUsuarios = New Collection
        For Each ldrwUsuario As DataRow In ldtbUsuarios.Rows
            Dim lobjUsuario As New ClsUsuario(Me, ldrwUsuario)
            lobjUsuario.SLeaValores(True)
            If ablnActivos Then
                If lobjUsuario.ObjEstaActivoUsuarioBln.ObjValorPro Then
                    lcolUsuarios.Add(lobjUsuario, lobjUsuario.ObjIdUsuarioStr.ObjValorPro)
                End If
            Else
                lcolUsuarios.Add(lobjUsuario, lobjUsuario.ObjIdUsuarioStr.ObjValorPro)
            End If
        Next
        Return lcolUsuarios
    End Function
    Private Shared Function FdtbUsuarios() As DataTable
        Dim ldtbUsuarios = ClsPanorama.FdtbDataTable(ClsUsuario.SstrNombreTabla, {"*"},
                    {{"IdUsuario", "ASC"}}, "")
        Return ldtbUsuarios
    End Function
    Public Function FobjNuevoUsuario() As ClsUsuario
        Dim ldtbUsuarios = FdtbUsuarios()
        Dim ldrwUsuario As DataRow = ldtbUsuarios.NewRow
        Dim lobjUsuario As New ClsUsuario(Me, ldrwUsuario)
        lobjUsuario.SCreeObj(Nothing)
        Return lobjUsuario
    End Function
    ''' <summary>
    ''' Actualiza la base de datos con los datos del Usuario y lo agrega a la colección de Usurios
    ''' </summary>
    ''' <param name="aobjUsuario">Objeto Usuario a ser adicionado.</param>
    ''' <remarks></remarks>
    Public Shared Sub SAdicioneUsuario(aobjUsuario As ClsUsuario)
        With aobjUsuario
            .ObjFechaCreacionDtm.ObjValorPro = Date.Now
            .ObjFechaCambioContrasenaDtm.ObjValorPro = Date.Now
            .SActualice(True)
        End With
    End Sub
#End Region

#Region "Validacion Centros de Utilidad Licenciados"

#End Region

#Region "Actualiza versón administrador"
    Friend Shared Sub SActualiceVer(astrVersionAnt As String, ByRef ablnActualizoApl As Boolean)
        Dim lblnNoHayErrores = False
        GblnActualizandoApp = True
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            ablnActualizoApl = True
            lblnNoHayErrores = True
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayErrores Then
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
        GblnActualizandoApp = False
    End Sub
#End Region

#Region "Copia de Seguridad"
    Friend Shared Function FstrNombreArchivoCopia() As String
        Dim lstrArchivo = String.Empty
        Dim lstrTrayCopia = ClsAdministrador.FobjAppActual.StrTrayCopiaSeguridad
        Dim lstrPrefijo = ClsPanorama.FstrFechayyyymmdd(Date.Today) & "_ADM_"
        lstrPrefijo &= "C" & GshrIdCarpeta.ToString & "_"
        lstrPrefijo += My.Resources.NombreBk
        If Not My.Computer.FileSystem.DirectoryExists(lstrTrayCopia) Then
            My.Computer.FileSystem.CreateDirectory(lstrTrayCopia)
            lstrArchivo = lstrTrayCopia & "\" & lstrPrefijo & "_01.sql"
        Else
            Dim lstrSufijo = String.Empty, i = 0
            Do While True
                i += 1
                If i > 99 Then
                    Throw New PanLException("Se alcanzo la cantidad máxima de copias en el día. " &
                            "La copia no se efectuó!")
                    Exit Do
                End If
                lstrSufijo = Format(i, "0#")
                lstrArchivo = lstrTrayCopia & "\" & lstrPrefijo & "_" & lstrSufijo & ".sql"
                If Not My.Computer.FileSystem.FileExists(lstrArchivo) Then
                    Exit Do
                End If
            Loop
        End If
        Return lstrArchivo
    End Function
#End Region

#Region "Dispose"
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overridable Overloads Sub Dispose(ablnDisposing As Boolean)
        If Not MblnDisposed Then
            If ablnDisposing Then
                GobjPanDat.Dispose()
                GobjPanDat = Nothing
                MblnDisposed = True
            End If
        End If
    End Sub
#End Region
End Class
