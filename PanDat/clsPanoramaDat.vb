Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.IO
Imports System.Text
Friend Class ClsPanoramaDat
#Region "Definiciones"
    Implements IDisposable
    ' Constantes
    Private Const MCOBJREGISTRO As Object = "A0b1f9*hjBó^23ö~"
    ' Variables
    Private Property MblnRegistrado As Boolean = False
    Private ReadOnly MobjCnnDat As ClsCnnDat = Nothing
    Private ReadOnly MobjIPan As IPanDat = Nothing
    ' 
    Private MtraTransaccion As Object = Nothing
    Private MshrIdTransaccion As Short = -1
    Private MblnDisposed As Boolean = False
    Private MshrProcesos As Short = 0
    Private MentVersionBD As Integer = 0
    ' Eventos
    Private WithEvents MmbBackup As MySqlBackup = Nothing
#End Region

#Region "Constructores"
    ''' <summary>
    ''' Constructor de clsPanoramaDat el cual se ejecuta al instanciar la clase
    ''' </summary>
    ''' <param name="aobjApp">Objeto que esta instanciando la clase el cual debe implementar 
    ''' la Interfaz IPanDat</param>
    ''' <remarks></remarks>
    Public Sub New(aobjApp As IPanDat)
        If aobjApp Is Nothing Then
            Throw New ArgumentNullException(NameOf(aobjApp))
        Else
            MobjIPan = aobjApp
        End If
        StrNombreBD = My.Resources.NomBd
        Dim lobjReg As Object = MobjIPan.ObjRegistro
        If lobjReg.GetType.Name <> "String" OrElse
                lobjReg <> MCOBJREGISTRO Then
            Throw New ModuloNoRegistradoPanException
        End If
        GshrIdAplicacion = MobjIPan.ShrIdApp
        GentVerBDEnProg = MobjIPan.EntVersionBDEnProg
        MblnRegistrado = True
        GobjPanDat = Me
        MobjCnnDat = New ClsCnnDat()
        MobjCnnDat.SEstablezcaConexionBd_App()
        SEstablezcaTipoInstancia()
    End Sub
    Public Sub New(aobjRegistro As Object)
        If aobjRegistro Is Nothing OrElse Not (aobjRegistro.GetType.Name = "String") Then
            Throw New ModuloNoRegistradoPanException("El módulo no ha sido debidamente cargado!")
        ElseIf Not (aobjRegistro = MCOBJREGISTRO) Then
            Throw New ModuloNoRegistradoPanException("El módulo no ha sido debidamente cargado!")
        End If
        Dim lblnNoHayError = False
        StrNombreBD = My.Resources.NomBd
        MblnRegistrado = True
        GobjPanDat = Me
        MobjCnnDat = New ClsCnnDat()
        Try
            SControleProcesoObj(True)
            MobjCnnDat.SEstablezcaConexionBd_App()
            lblnNoHayError = True
        Catch ex As PanDatException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                SControleProcesoObj(False)
            Else
                SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Friend Sub SInicieDat(aentIdVerBDAdmin As Integer)
        Dim lblnNoHayError = False
        Try
            SControleProcesoObj(True)
            If GshrIdAplicacion = 100 Then
                SActualicePan()
            Else
                Dim lentIdVerAdmin = GobjPanDat.FentVersionBD(EnuListaAplicaciones.EnuAdministrador)
                If lentIdVerAdmin <> aentIdVerBDAdmin Then
                    Throw New ErrorInesperadoPanDatException("El Administrador está desactualizado.")
                End If
                SActualiceApp()
            End If
            lblnNoHayError = True
        Catch ex As PanDatException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                SControleProcesoObj(False)
            Else
                SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
#End Region

#Region "Propiedades"
    Friend Property CnnConexionBd_App As Object = Nothing
    Friend Property BlnExisteBdPanorama As Boolean = False

    Friend ReadOnly Property BlnRegistrado As Boolean
        Get
            Return MblnRegistrado
        End Get
    End Property

    Friend Shared Function FobjEstructuraBD(aentIdApp As Integer) As ClsBaseDatos
        Dim lstrPrefTablas As String = ""
        If aentIdApp = EnuListaAplicaciones.EnuAdministrador Then
            lstrPrefTablas = My.Resources.PrefPan
        ElseIf aentIdApp = EnuListaAplicaciones.EnuOrionCop Then
            lstrPrefTablas = My.Resources.PrefOri
        End If
        Dim lobjEstructuraBdPan = FobjBaseDatosDB(lstrPrefTablas)
        Return lobjEstructuraBdPan
    End Function
#End Region

#Region "Procedimientos"
#Region "ActualizacionBD"
    Private Sub SEstablezcaTipoInstancia()
        If Not BlnExisteBdPanorama Then
            If GshrIdAplicacion = 100 OrElse GshrIdAplicacion = 999 Then
                ' Instalacion
                GenuTipoInstanciamiento = EnuTipoInstanciamiento.enuInstalacion
            Else
                Throw New ErrorInesperadoPanDatException("Aun no está instalado el Administrador")
            End If
        Else
            FentVersionBD(GshrIdAplicacion)
            If Not FblnExistenTablasApp(GshrIdAplicacion) Then
                ' Instalación
                GenuTipoInstanciamiento = EnuTipoInstanciamiento.enuInstalacion
            Else
                If GentVerBDEnProg > MentVersionBD Then
                    ' Actualizacion
                    GenuTipoInstanciamiento = EnuTipoInstanciamiento.enuActualizacion
                Else
                    ' Normal
                    GenuTipoInstanciamiento = EnuTipoInstanciamiento.enuNormal
                End If
            End If
        End If
    End Sub

    Private Function FblnExistenTablasApp(ashrIdApp As Short) As Boolean
        Dim lblnExiste = BlnExisteBdPanorama
        If lblnExiste Then
            Dim lstrPrefTablas As String = FstrPrefijoTablas(ashrIdApp)
            Dim lstrFiltro(3) As String
            lstrFiltro(1) = StrNombreBD.ToLower
            Dim lstrFiltroTablas As String = "TABLE_NAME LIKE '" & lstrPrefTablas & "*'"
            SAbraConexionBd()
            Dim ldtbTablasBD As DataTable = CnnConexionBd_App.GetSchema("Tables", lstrFiltro)
            Dim ldrwTablas As DataRow() = ldtbTablasBD.Select(lstrFiltroTablas)
            lblnExiste = ldrwTablas.Count > 0
        End If
        Return lblnExiste
    End Function

    Friend Shared Function FstrPrefijoTablas(ashrIdApp As Short) As String
        Dim lstrPrefTab = ""
        Select Case ashrIdApp
            Case EnuListaAplicaciones.EnuAdministrador
                lstrPrefTab = My.Resources.PrefPan
            Case EnuListaAplicaciones.EnuOrionCop
                lstrPrefTab = My.Resources.PrefOri
        End Select
        Return lstrPrefTab
    End Function

    Private Sub SActualicePan()
        Dim lstrNombreArchivos = String.Empty
#If DES = 1 Then
        lstrNombreArchivos = "C:\FuentesPanorama.Net\Trunk\Comunes\PanDat\XmlBd\" &
                MobjIPan.StrNombreArchivos & ".xml"
#Else
        lstrNombreArchivos = GstrTrayAppDat & MobjIPan.StrNombreArchivos & ".xml"
#End If
        GshrIdAplicacion = MobjIPan.ShrIdApp
        GentVerBDEnProg = MobjIPan.EntVersionBDEnProg
        Try
            SVerifiqueCarpetas()
            If My.Computer.FileSystem.FileExists(lstrNombreArchivos) Then
                SActualiceBD(lstrNombreArchivos)
            End If
            BlnExisteBdPanorama = True
        Catch ex As PanDatException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Friend Sub SActualiceApp()
        Dim lstrNombreArchivos = String.Empty
#If DES = 1 Then
        lstrNombreArchivos = "C:\FuentesPanorama.Net\Trunk\Comunes\PanDat\XmlBd\" &
                MobjIPan.StrNombreArchivos & ".xml"
#Else
        lstrNombreArchivos = GstrTrayAppDat & MobjIPan.StrNombreArchivos & ".xml"
#End If
        GshrIdAplicacion = MobjIPan.ShrIdApp
        GentVerBDEnProg = MobjIPan.EntVersionBDEnProg
        Try
            If My.Computer.FileSystem.FileExists(lstrNombreArchivos) Then
                SActualiceBD(lstrNombreArchivos)
            End If
        Catch ex As PanDatException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Manejo AurigaFtp"
    Friend Sub SdstAuriga(adsDataSet As DataSet, astrExpSql As String)
        Dim lstrNombreTabla As String = "TablaResultado"
        Dim ldapAdaptador As Object
        Dim lcnnAuriga As MySqlConnection = Nothing
        Try
            lcnnAuriga = MobjCnnDat.FcnnMySqlCon_Au
            lcnnAuriga.Open()
            adsDataSet.Tables.Add(lstrNombreTabla)
            ldapAdaptador = FdapAdaptador(GenuProveedorBD, astrExpSql, lcnnAuriga,
                    EnuTipoSentencia.enuSelect)
            ldapAdaptador.Fill(adsDataSet, lstrNombreTabla)
        Catch ex As MySqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            lcnnAuriga.Close()
            lcnnAuriga.Dispose()
        End Try
    End Sub
    Friend Function SInserteRegistro_Au(astrNombreTabla As String,
            acolNombreCampos As Collection, acolDatos As Collection) As Integer
        Dim lcnnAuriga As MySqlConnection
        Dim lentRegistrosAfectados As Integer = 0
        Dim lstrSql As String
        Dim lcmdComando = FcmdNewComando()
        Dim lblnNoHayError = False
        Try
            SControleProcesoObj(True)
            lstrSql = FstrConstruyaExpSqlInsertar(astrNombreTabla, acolNombreCampos)
            lcnnAuriga = MobjCnnDat.FcnnMySqlCon_Au
            lcnnAuriga.Open()
            lcmdComando.Connection = lcnnAuriga
            lcmdComando.CommandText = lstrSql
            lcmdComando.CommandType = CommandType.Text
            For i As Byte = 1 To acolNombreCampos.Count
                lcmdComando.Parameters.AddWithValue("@" & acolNombreCampos(i),
                        acolDatos(acolNombreCampos(i)))
            Next
            lentRegistrosAfectados = lcmdComando.ExecuteNonQuery()
            lcnnAuriga.Close()
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As SqlClient.SqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OleDb.OleDbException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OracleClient.OracleException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lcmdComando) Then
                lcmdComando.Dispose()
            End If
            If lblnNoHayError Then
                SControleProcesoObj(False)
            Else
                SControleProcesoObj(False, True)
            End If
        End Try
        Return lentRegistrosAfectados
    End Function
    Friend Function SActualiceReg_Au(astrNombreTabla As String,
                acolNombresCamposCambio As Collection, acolDatosNuevos As Collection,
                acolNombresCamposRef As Collection, acolDatosRef As Collection) As Integer
        Dim lentRegistrosAfectados = 0
        Dim lstrSql As String
        Dim lcnnAuriga As MySqlConnection
        Dim lcmdComando As Object = FcmdNewComando()
        Dim lblnNoHayError = False
        Try
            SControleProcesoObj(True)
            lstrSql = FstrConstruyaExpSqlActualizar(astrNombreTabla, acolNombresCamposCambio,
                        acolNombresCamposRef)
            lcnnAuriga = MobjCnnDat.FcnnMySqlCon_Au
            lcnnAuriga.Open()
            lcmdComando.Connection = lcnnAuriga
            lcmdComando.CommandText = lstrSql
            lcmdComando.CommandType = CommandType.Text
            For i As Byte = 1 To acolNombresCamposCambio.Count
                lcmdComando.Parameters.AddWithValue("@" & acolNombresCamposCambio(i),
                        acolDatosNuevos(acolNombresCamposCambio(i)))
            Next
            For i = 1 To acolNombresCamposRef.Count
                lcmdComando.Parameters.AddWithValue("@Ref" & acolNombresCamposRef(i),
                        acolDatosRef(i))
            Next
            lentRegistrosAfectados = lcmdComando.ExecuteNonQuery()
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As SqlClient.SqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OleDb.OleDbException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OracleClient.OracleException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As ProveedorBdPanException
            Throw
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lcmdComando) Then
                lcmdComando.Dispose()
            End If
            If lblnNoHayError Then
                SControleProcesoObj(False)
            Else
                SAborteTransaccion()
                SControleProcesoObj(False, True)
            End If
        End Try
        Return lentRegistrosAfectados
    End Function
#End Region

#Region "Operaciones con Disco"
    ''' <summary>
    ''' Crea la carpeta con el nombre "astrNombreCarpeta" según la ubicacion indicada en el argumento "aenuUbicacioncarpeta"
    ''' </summary>
    Friend Shared Sub SCreeCarpeta(astrNombreCarpeta As String)
        GstrTrayDatos = GstrTrayDat & astrNombreCarpeta
        Try
            If Not My.Computer.FileSystem.DirectoryExists(GstrTrayDatos) Then
                My.Computer.FileSystem.CreateDirectory(GstrTrayDatos)
            End If
        Catch ex As IOException
            Throw New PanDatException(ex.Message)
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Operaciones DataSet y DataReader"
    Friend Shared Function FdrDataReader(acnnConexion As Object, astrSql As String) As Object
        If String.IsNullOrEmpty(astrSql) OrElse acnnConexion Is Nothing Then
            If IsNothing(acnnConexion) Then
                Throw New ArgumentNullException(NameOf(acnnConexion))
            Else
                Throw New ArgumentNullException(NameOf(astrSql))
            End If
        End If
        Dim lcmdComando As Object = FcmdNewComando()
        Dim ldrDataReader As Object
        Try
            lcmdComando.Connection = acnnConexion
            lcmdComando.CommandText = astrSql.ToLower
            lcmdComando.CommandType = CommandType.Text
            ldrDataReader = lcmdComando.ExecuteReader
        Catch ex As MySqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As SqlClient.SqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OleDb.OleDbException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OracleClient.OracleException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As Exception
            Throw
        End Try
        Return ldrDataReader
    End Function
    ''' <summary>
    ''' Puebla el DataSet "adstDataSet" de acuerdo a los argumentos, los cuales deben corresponder a tablas 
    ''' de la aplicación que instanció esta clase "clsPanoramaDat"
    ''' </summary>
    ''' <param name="adstDataSet">DataSet que será poblado con las DataTables generadas.</param>
    ''' 
    ''' <param name="acolTablas">Colección con los nombres de las tablas que seran tenidas en cuenta 
    ''' para poblar el dataset</param>
    ''' <param name="acolCampos">Colección de arrays cada uno de los cuales contiene los nombres de los campos
    ''' que seran tenidos en cuenta en cada una de las tablas de 'acolTablas'. La cantidad de elemntos de esta 
    ''' colleción debe ser igual a la de "acolTablas" y se deben corresponder uno uno.</param>
    ''' <param name="acolCamposIndice">Colección de Arrays de los indices de las tablas. Esta coleccion puede ser
    ''' nula:(Nothing) pero si no es nula debe tener la misma cantidad de elemntos que 'acolTablas' y se deben
    ''' corresponder uno a uno. Si una tabla no debe estar ordenada entonces el indice se debe pasar como nulo:(Nothing) o
    ''' vacio:({{"",""}}). Si se pasa un indices, este deben ser únicos.</param>
    ''' <param name="acolFiltros">Colección de expresiones de Filtro que se tendra en cuenta para seleccionar los 
    ''' registros. La colección puede ser nula:(Nothing) pero si no lo es ninguno de los elementos lo puede ser.
    ''' Debe haber una correspondencia uno a uno con la coleccion 'acolTablas'. Si alguna de las tablas no debe ser 
    ''' filtrada se debe pasar el filtro como un string vacio:("").</param>
    ''' <remarks>Excepciones:ArgumentNullException, ArgumentoInvalidoPanException, SintaxisInvalidaPanException</remarks>
    Friend Sub SdsDataSet(adstDataSet As DataSet, acolTablas As Collection, acolCampos As Collection,
            acolCamposIndice As Collection, acolFiltros As Collection)
        Dim lstrSql As String
        Dim lstrNombreTabla As String
        Dim ldapAdaptador As Object
        Dim ldtbDataTable As DataTable
        Dim lstrIndice(,) As String
        Dim lstrFiltro As String
        Dim i As Integer
        Dim lblnNoHayError = False
        If Not IsNothing(acolCamposIndice) Then
            If acolCamposIndice.Count <> acolTablas.Count OrElse acolCampos.Count <> acolTablas.Count Then
                Throw New ArgumentoInvalidoPanException("Alguna de las colecciones difiere en longitud!")
            End If
        End If
        If Not IsNothing(acolFiltros) AndAlso acolFiltros.Count > 0 Then
            If acolTablas.Count <> acolFiltros.Count Then
                Throw New ArgumentoInvalidoPanException("Alguna de las colecciones difiere en longitud!")
            End If
        End If
        Try
            SControleProcesoObj(True)
            adstDataSet.EnforceConstraints = True
            SAbraConexionBd()
            For i = 1 To acolTablas.Count
                If Not IsNothing(acolFiltros) AndAlso acolFiltros.Count > 0 Then
                    lstrFiltro = acolFiltros(i)
                Else
                    lstrFiltro = String.Empty
                End If
                If Not IsNothing(acolCamposIndice) Then
                    lstrIndice = acolCamposIndice(i)
                Else
                    lstrIndice = {{"", ""}}
                End If
                lstrNombreTabla = acolTablas(i)
                lstrSql = FstrConstruyaExpSqlSelect(lstrNombreTabla, acolCampos(i), lstrIndice,
                        lstrFiltro, Array.Empty(Of String))
                adstDataSet.Tables.Add(lstrNombreTabla)
                ldapAdaptador = FdapAdaptador(GenuProveedorBD, lstrSql, CnnConexionBd_App,
                            EnuTipoSentencia.enuSelect)
                ldapAdaptador.Fill(adstDataSet, lstrNombreTabla)
                ldtbDataTable = adstDataSet.Tables(lstrNombreTabla)
                SAsignePrimaryKey(ldtbDataTable, lstrIndice)
            Next
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                SControleProcesoObj(False)
            Else
                SAborteTransaccion()
                SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Private Shared Sub SAsignePrimaryKey(adtbDataTable As DataTable, astrIndice As String(,))
        If FblnHayIndice(astrIndice) Then
            Dim lsclColumnasIndice As DataColumn()
            ReDim lsclColumnasIndice(astrIndice.GetUpperBound(0))
            For j = 0 To astrIndice.GetUpperBound(0)
                If astrIndice(j, 0).Contains(".") Then
                    astrIndice(j, 0) = astrIndice(j, 0).Substring(astrIndice(j, 0).IndexOf(".") + 1)
                End If
                lsclColumnasIndice(j) = adtbDataTable.Columns(astrIndice(j, 0))
            Next
            If Not IsNothing(lsclColumnasIndice) Then
                adtbDataTable.PrimaryKey = lsclColumnasIndice
            End If
        End If
    End Sub
    ''' <summary>
    ''' Puebla el DataSet "adstDataSet" pasado por referencia de acuerdo a los argumentos pasados, los cuales
    ''' deben corresponder a tablas de la aplicación que instanció esta clase "clsPanoramaDat".
    ''' </summary>
    ''' <param name="adstDataSet">DataSet que sera poblado.</param>
    ''' <param name="astrNombreTabla">Tabla con la cual será poblado el DataSet.</param>
    ''' <param name="astrCampos">Array con los nombres de los campos que se tendran en cuenta.</param>
    ''' <param name="astrIndice">Array con los campos que compondran el indice y su orden. Este argumento
    ''' no puede ser un array Null:(nothing) pero puede ser vacio:({{"", ""}})</param>
    ''' <param name="astrFiltro">Expresión de filtro para indicar los registros a ser tenidos en cuenta.
    ''' Este parametro no puede ser Null (nothing) pero puede ser una cadena vacia ("")</param>
    ''' <param name="ablnPKIndice">Indica si el indice pasado es unico o no.</param>
    ''' <param name="astrCamposGrupo">Array con los nombres de los campos de agrupamiento. Este argumento
    ''' puede ser un array vacio:( Array.Empty(Of String))</param>
    ''' <remarks>Excepciones:ArgumentNullException, ArgumentoInvalidoPanException, SintaxisInvalidaPanException</remarks>
    Friend Sub SdsDataSet(adstDataSet As DataSet, astrNombreTabla As String,
            astrCampos() As String, astrIndice(,) As String, astrFiltro As String,
            ablnPKIndice As Boolean, astrCamposGrupo() As String)
        Dim lstrNombreIndice As String = "PK_"
        Dim ldapAdaptador As Object
        Dim ldtbDataTable As DataTable
        Dim lsclColumnasIndice() As DataColumn
        Dim i As Byte
        Dim lblnNoHayError = False
        Try
            SControleProcesoObj(True)
            Dim lstrSql As String = FstrConstruyaExpSqlSelect(astrNombreTabla, astrCampos,
                    astrIndice, astrFiltro, astrCamposGrupo).Trim
            adstDataSet.EnforceConstraints = True
            adstDataSet.Tables.Add(astrNombreTabla)
            SAbraConexionBd()
            ldapAdaptador = FdapAdaptador(GenuProveedorBD, lstrSql, CnnConexionBd_App,
                        EnuTipoSentencia.enuSelect)
            ldapAdaptador.Fill(adstDataSet, astrNombreTabla)
            If FblnHayIndice(astrIndice) AndAlso ablnPKIndice Then
                lstrNombreIndice &= astrNombreTabla
                ReDim lsclColumnasIndice(astrIndice.GetUpperBound(0))
                ldtbDataTable = adstDataSet.Tables(astrNombreTabla)
                For i = 0 To astrIndice.GetUpperBound(0)
                    lsclColumnasIndice(i) = ldtbDataTable.Columns(astrIndice(i, 0))
                Next
                If Not IsNothing(lsclColumnasIndice) Then
                    ldtbDataTable.PrimaryKey = lsclColumnasIndice
                End If
            End If
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                SControleProcesoObj(False)
            Else
                SAborteTransaccion()
                SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    ''' <summary>
    ''' Puebla el DataSet "adstDataSet" de acuerdo a los argumentos pasados, los cuales deben corresponder a tablas 
    ''' de la aplicación que instanció esta clase "clsPanoramaDat".
    ''' </summary>
    ''' <param name="adstDataSet">DataSet que sera poblado.</param>
    ''' <param name="astrNombreTablaPri">Nombre de la tabla primaria que se relaciona con la tabla secundaria
    ''' cuyo nombre es "astrNombreTablaSec"</param>
    ''' <param name="astrCamposTablaPri">Campos de la tabla primaria que serán tenidos en cuenta para poblar 
    ''' el DataSet</param>
    ''' <param name="astrNombreTablaSec">Nombre de la tabla secundaria relacionada von la tabla primaria a
    ''' traves de los campos cuyos nombres estan contenidos en los array "astrCamposPriRel" y "astrCamposSecRel"</param>
    ''' <param name="astrCamposTablaSec">Campos de la tabla secundaria que serán tenidos en cuenta para 
    ''' poblar el DataSet</param>
    ''' <param name="astrCamposPriRel">Array que contiene los nombres de los campos de la tabla primaria para 
    ''' ser relacionados con la tabla secundaria</param>
    ''' <param name="astrCamposSecRel">Array que contiene los nombres de los campos de la tabla secundaria para 
    ''' ser relacionados con la tabla primaria</param>
    ''' <param name="astrIndice">Array que contiene los nombres de los campos que conformaran el indice y 
    ''' su sentido. Para no indizar el resultado se debe pasar un array vacio:({{,}})</param>
    ''' <param name="astrFiltro">Expresion de filtro para seleccionar los registros que serán tenidos en cuenta. Para indicar que no se
    ''' filtren los registros este string se debe pasar vacio:("") </param>
    ''' <param name="ablnPKIndice">Indica si el indice en "astrIndice" es unico o no</param>
    ''' <param name="astrCamposGrupo">Array con los nombres de los campos de agrupamiento. Si este array se pasa vacio
    ''' no se hará agrupamiento.</param>
    ''' <remarks>ArgumentNullException, SintaxisInvalidaPanException</remarks>
    Friend Sub SdsDataSet(adstDataSet As DataSet, astrNombreTablaPri As String,
            astrCamposTablaPri() As String, astrNombreTablaSec As String,
            astrCamposTablaSec() As String, astrCamposPriRel() As String,
            astrCamposSecRel() As String, astrIndice(,) As String,
            astrFiltro As String, ablnPKIndice As Boolean,
            astrCamposGrupo() As String)
        Dim lstrNombreTabla As String = "TablaResultado"
        Dim ldapAdaptador As Object
        Dim ldtbDataTable As DataTable
        Dim lblnNoHayError = False
        Try
            SControleProcesoObj(True)
            Dim lstrSql As String = FstrConstruyaExpSqlSelect(astrNombreTablaPri, astrCamposTablaPri,
                        astrNombreTablaSec, astrCamposTablaSec, astrCamposPriRel, astrCamposSecRel,
                        astrIndice, astrFiltro, astrCamposGrupo)
            adstDataSet.EnforceConstraints = True
            adstDataSet.Tables.Add(lstrNombreTabla)
            SAbraConexionBd()
            ldapAdaptador = FdapAdaptador(GenuProveedorBD, lstrSql, CnnConexionBd_App,
                        EnuTipoSentencia.enuSelect)
            ldapAdaptador.Fill(adstDataSet, lstrNombreTabla)
            ldtbDataTable = adstDataSet.Tables(lstrNombreTabla)
            If ablnPKIndice Then
                SAsignePrimaryKey(ldtbDataTable, astrIndice)
            End If
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                SControleProcesoObj(False)
            Else
                SAborteTransaccion()
                SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    ''' <summary>
    ''' Puebla el DataSet "adstDataSet" de acuerdo a la expresion SQL pasada en el
    ''' argumento "astrSql".
    ''' </summary>
    ''' <param name="adstDataSet">DataSet a ser poblado.</param>
    ''' <param name="astrSql">Expresion SQL de selección.</param>
    ''' <remarks></remarks>
    '''
    Friend Sub SdsDataSet(adstDataSet As DataSet, astrSql As String)
        Dim lblnNoHayError = False
        Dim lstrNombreTabla As String = "TablaResultado"
        Dim ldapAdaptador As Object
        Try
            SControleProcesoObj(True)
            SAbraConexionBd()
            adstDataSet.Tables.Add(lstrNombreTabla)
            ldapAdaptador = FdapAdaptador(GenuProveedorBD, astrSql, CnnConexionBd_App,
                        EnuTipoSentencia.enuSelect)
            ldapAdaptador.Fill(adstDataSet, lstrNombreTabla)
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                SControleProcesoObj(False)
            Else
                SAborteTransaccion()
                SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    ''' <summary>
    ''' Puebla el DataSet "adstDataSet" con dos tablas: La primera con la tabla resultante de ejecutar la
    ''' expresion sql "astrSqlSelect" y la segunda con la tabla cuyo nombre es "astrNombreTabla" y que contiene los campos
    ''' expresados en el parametro "astrCamposTabla" ordenados según la expresion contenida en el parametro "astrIndice".
    ''' </summary>
    ''' <param name="adstDataSet">DataSet a ser poblado</param>
    ''' <param name="astrSqlSelect">Expresion SELECT a partir de la cual se genera la primera tabla</param>
    ''' <param name="astrNombreTabla">Nombre de la segunda tabla.</param>
    ''' <param name="astrCamposTabla">Expresión que contiene los campos que contendra la segunda Tabla.</param>
    ''' <param name="astrIndice">Expresión que contiene el indice de la segunda tabla. Este indice debe ser unico.
    ''' Si no se quiere ordenar la tabla este argumento debe aer una cadena vacia:("")</param>
    ''' <param name="astrFiltro">Expresion de filtro para seleccionar los registros que serán tenidos en cuenta.
    ''' Si no se requiere filtrar este argumento debe ser una cadena vacia:("").</param>
    ''' <remarks></remarks>
    Friend Sub SdsDataSet(adstDataSet As DataSet, astrSqlSelect As String,
            astrNombreTabla As String, astrCamposTabla As String,
            astrIndice As String, astrFiltro As String)
        If adstDataSet Is Nothing OrElse astrSqlSelect Is Nothing OrElse astrNombreTabla Is Nothing OrElse
                astrCamposTabla Is Nothing OrElse astrIndice Is Nothing OrElse astrFiltro Is Nothing Then
            Dim lstrArgumentoNull As String = String.Empty
            Select Case True
                Case adstDataSet Is Nothing
                    lstrArgumentoNull = "adstDataSet"
                Case astrSqlSelect Is Nothing
                    lstrArgumentoNull = "astrSqlSelect"
                Case astrNombreTabla Is Nothing
                    lstrArgumentoNull = "astrNombreTabla"
                Case astrCamposTabla Is Nothing
                    lstrArgumentoNull = "astrCamposTabla"
                Case astrIndice Is Nothing
                    lstrArgumentoNull = "astrCamposTabla"
                Case astrFiltro Is Nothing
                    lstrArgumentoNull = "astrFiltro"
            End Select
            Throw New ArgumentNullException(lstrArgumentoNull)
        End If
        Dim lblnNoHayError = False
        Dim ldapAdaptador As Object
        Dim lstrSql = "SELECT " & astrCamposTabla & " FROM " & astrNombreTabla
        If Not String.IsNullOrEmpty(astrFiltro) Then
            lstrSql &= " WHERE " & astrFiltro
        End If
        If Not String.IsNullOrEmpty(astrIndice) Then
            lstrSql &= " ORDER BY " & astrIndice
        End If
        Dim ldtbDataTable As DataTable
        Dim lsclColumnasIndice() As DataColumn
        Try
            SControleProcesoObj(True)
            SAbraConexionBd()
            adstDataSet.Tables.Add("DataTable1")
            ldapAdaptador = FdapAdaptador(GenuProveedorBD, astrSqlSelect, CnnConexionBd_App,
                        EnuTipoSentencia.enuSelect)
            ldapAdaptador.Fill(adstDataSet, "DataTable1")
            ldapAdaptador = FdapAdaptador(GenuProveedorBD, lstrSql, CnnConexionBd_App,
                    EnuTipoSentencia.enuSelect)
            ldapAdaptador.Fill(adstDataSet, astrNombreTabla)
            ldtbDataTable = adstDataSet.Tables(astrNombreTabla)
            ReDim lsclColumnasIndice(0)
            lsclColumnasIndice(0) = ldtbDataTable.Columns(astrIndice)
            ldtbDataTable.PrimaryKey = lsclColumnasIndice
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                SControleProcesoObj(False)
            Else
                SAborteTransaccion()
                SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    ''' <summary>
    ''' Puebla el DataSet "adstDataSet" con dos tablas: La primera con la tabla resultante de ejecutar la
    ''' expresion sql "astrSqlSelect" y la segunda con la tabla cuyo nombre es "astrNombreTabla" y que contiene los campos
    ''' expresados en el parametro "astrCamposTabla" ordenados según la expresion contenida en el parametro "astrIndice".
    ''' </summary>
    ''' <param name="adstDataSet">DataSet a ser poblado</param>
    ''' <param name="astrSqlSelect1">Expresion SELECT a partir de la cual se genera la primera tabla</param>
    ''' <param name="astrNombreTabla1">Nombre de la tabla generada por astrSqlSelect1.</param>
    ''' <param name="astrSqlSelect2">Expresion SELECT a partir de la cual se genera la segunda Tabla.</param>
    ''' <param name="astrNombreTabla2">Nombre de la tabla generada por astrSqlSelect2.</param>
    ''' <remarks></remarks>
    Friend Sub SdsDataSet(adstDataSet As DataSet, astrSqlSelect1 As String,
            astrNombreTabla1 As String, astrSqlSelect2 As String,
            astrNombreTabla2 As String)
        If adstDataSet Is Nothing OrElse String.IsNullOrEmpty(astrSqlSelect1) OrElse
                String.IsNullOrEmpty(astrNombreTabla1) OrElse
                String.IsNullOrEmpty(astrSqlSelect2) OrElse
                String.IsNullOrEmpty(astrNombreTabla2) Then
            Dim lstrArgumentoNull As String = String.Empty
            Select Case True
                Case adstDataSet Is Nothing
                    lstrArgumentoNull = "adstDataSet"
                Case String.IsNullOrEmpty(astrSqlSelect1)
                    lstrArgumentoNull = "astrSqlSelect1"
                Case String.IsNullOrEmpty(astrNombreTabla1)
                    lstrArgumentoNull = "astrNombreTabla1"
                Case String.IsNullOrEmpty(astrSqlSelect2)
                    lstrArgumentoNull = "astrSqlSelect2"
                Case String.IsNullOrEmpty(astrNombreTabla2)
                    lstrArgumentoNull = "astrNombreTabla2"
            End Select
            Throw New ArgumentNullException(lstrArgumentoNull)
        End If
        Dim lblnNoHayError = False
        Dim ldapAdaptador As Object
        Try
            SControleProcesoObj(True)
            SAbraConexionBd()
            adstDataSet.Tables.Add(astrNombreTabla1)
            ldapAdaptador = FdapAdaptador(GenuProveedorBD, astrSqlSelect1, CnnConexionBd_App,
                        EnuTipoSentencia.enuSelect)
            ldapAdaptador.Fill(adstDataSet, astrNombreTabla1)
            adstDataSet.Tables.Add(astrNombreTabla2)
            ldapAdaptador = FdapAdaptador(GenuProveedorBD, astrSqlSelect2, CnnConexionBd_App,
                    EnuTipoSentencia.enuSelect)
            ldapAdaptador.Fill(adstDataSet, astrNombreTabla2)
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                SControleProcesoObj(False)
            Else
                SAborteTransaccion()
                SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    ''' <summary>
    ''' Puebla el DataSet "adsDataSet" con las tablas resultantes de ejecutar las expresiones SELECT contenidas
    ''' en la coleccion acolExpresionesSql. Los nombres de las tablas estan contenidas en la coleccion 
    ''' acolNombresTablas.
    ''' </summary>
    ''' <param name="adsDataSet">DataSet a ser poblado</param>
    ''' <param name="acolExpresionesSql">Coleccion que contiene las expresiones SELECT para cada tabla que va
    ''' a contener el dataset</param>
    ''' <param name="acolNombresTablas">Coleccion que contiene los Nombres de las tablas que contendra 
    ''' el DataSet.</param>
    ''' <remarks></remarks>
    Friend Sub SdsDataSet(adsDataSet As DataSet, acolExpresionesSql As Collection,
                           acolNombresTablas As Collection)
        Dim lblnNoHayError = False
        Dim ldapAdaptador As Object
        Try
            SControleProcesoObj(True)
            If acolExpresionesSql.Count <> acolNombresTablas.Count Then
                Throw New ErrorInesperadoPanDatException("Colecciones con longitudes diferentes")
            End If
            SAbraConexionBd()
            For i = 1 To acolExpresionesSql.Count
                adsDataSet.Tables.Add(acolNombresTablas(i))
                ldapAdaptador = FdapAdaptador(GenuProveedorBD, acolExpresionesSql(i),
                        CnnConexionBd_App, EnuTipoSentencia.enuSelect)
                ldapAdaptador.Fill(adsDataSet, acolNombresTablas(i))
            Next
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                SControleProcesoObj(False)
            Else
                SAborteTransaccion()
                SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    ''' <summary>
    ''' Puebla el DatSet "adstDataSet" con el contenido de una tabla perteneciente a
    ''' una base de datos Access.
    ''' </summary>
    ''' <param name="adstDataSet">DataSet a ser poblado.</param>
    ''' <param name="astrArchivoOrigen">Trayectoria completa del archivo que contiene la base de datos
    ''' Access</param>
    ''' <param name="astrContrasena">Contraseña de la base de datos Access</param>
    ''' <param name="astrNombreTabla">Nombre de la tabla que poblará el DataSet o una instrucción SELECT</param>
    ''' <remarks></remarks>
    Friend Sub SdstDataSetAccess(adstDataSet As DataSet, astrArchivoOrigen As String,
            astrContrasena As String, astrNombreTabla As String)
        Dim lblnNoHayError = False
        Dim lblnExcel As Boolean = astrArchivoOrigen.EndsWith(".xls")
        Dim lblnExcel10 As Boolean = astrArchivoOrigen.EndsWith(".xlsx")
        Dim lstrTrayectoria As String = FstrObtengaUnidad(astrArchivoOrigen) &
                FstrObtengaTrayectoria(astrArchivoOrigen)
        Dim lstrNombreArchivo As String = FstrObtengaNombre(astrArchivoOrigen)
        Dim lstrSql As String
        If astrNombreTabla.ToUpper.StartsWith("SELECT") Then
            lstrSql = astrNombreTabla
        Else
            If Not astrNombreTabla.EndsWith("$") Then
                astrNombreTabla += "$"
            End If
            lstrSql = "SELECT * FROM [" & astrNombreTabla & "]"
        End If
        Dim lcnnConexion As OleDb.OleDbConnection = Nothing
        Try
            If lblnExcel10 Then
                lcnnConexion = MobjCnnDat.FcnnConexion(EnuProveedorBD.enuExcel10,
                        EnuTipoAutenticacion.None, "", lstrTrayectoria, lstrNombreArchivo, "",
                        astrContrasena, "", False, False)
            Else
                lcnnConexion = MobjCnnDat.FcnnConexion(EnuProveedorBD.enuOleDb,
                        EnuTipoAutenticacion.None, "", lstrTrayectoria, lstrNombreArchivo, "",
                        astrContrasena, "", lblnExcel, False)
            End If
            If lcnnConexion.State = ConnectionState.Closed Then
                lcnnConexion.Open()
            End If
            adstDataSet.EnforceConstraints = True
            adstDataSet.Tables.Add(astrNombreTabla)
            Using ldapAdaptador = FdapAdaptador(EnuProveedorBD.enuOleDb, lstrSql, lcnnConexion,
                    EnuTipoSentencia.enuSelect)
                ldapAdaptador.Fill(adstDataSet, astrNombreTabla)
            End Using
            lblnNoHayError = True
        Catch ex As OleDbException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As ParametrosConexionBDPanException
            Throw
        Catch ex As ProveedorBdPanException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                If lcnnConexion.State = ConnectionState.Open Then
                    lcnnConexion.Close()
                End If
            End If
        End Try
    End Sub
#End Region

#Region "Operaciones con Registros a traves de la conexion general que siempre esta disponible"
    Friend Function SInserteRegistro(astrNombreTabla As String,
                acolNombreCampos As Collection, acolDatos As Collection) As Integer
        Dim lentRegistrosAfectados As Integer = 0
        Dim lstrSql As String
        Dim lcmdComando = FcmdNewComando()
        Dim lblnNoHayError = False
        Try
            SControleProcesoObj(True)
            lstrSql = FstrConstruyaExpSqlInsertar(astrNombreTabla, acolNombreCampos)
            SAbraConexionBd()
            If Not IsNothing(MtraTransaccion) Then
                lcmdComando.Transaction = MtraTransaccion
            End If
            lcmdComando.Connection = CnnConexionBd_App
            lcmdComando.CommandText = lstrSql
            lcmdComando.CommandType = CommandType.Text
            For i As Byte = 1 To acolNombreCampos.Count
                lcmdComando.Parameters.AddWithValue("@" & acolNombreCampos(i),
                        acolDatos(acolNombreCampos(i)))
            Next
            lentRegistrosAfectados = lcmdComando.ExecuteNonQuery()
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As SqlClient.SqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OleDb.OleDbException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OracleClient.OracleException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lcmdComando) Then
                lcmdComando.Dispose()
            End If
            If lblnNoHayError Then
                SControleProcesoObj(False)
            Else
                SAborteTransaccion()
                SControleProcesoObj(False, True)
            End If
        End Try
        Return lentRegistrosAfectados
    End Function
    Friend Function SActualiceRegistro(astrNombreTabla As String, acolNombresCamposCambio As Collection,
            acolDatosNuevos As Collection) As Integer
        Dim lentRegistrosAfectados As Integer
        If acolNombresCamposCambio.Count <> acolDatosNuevos.Count Then
            Throw New ArgumentoInvalidoPanException("Alguna de las colecciones difiere en longitud!")
        End If
        lentRegistrosAfectados = SActualiceReg(MtraTransaccion, astrNombreTabla, acolNombresCamposCambio,
                    acolDatosNuevos)
        Return lentRegistrosAfectados
    End Function
    Friend Function SActualiceRegistro(astrNombreTabla As String, acolNombresCamposCambio As Collection,
            acolDatosNuevos As Collection, acolNombresCamposRef As Collection,
            acolDatosRef As Collection) As Integer
        Dim lentRegistrosAfectados As Integer
        If acolNombresCamposCambio.Count <> acolDatosNuevos.Count OrElse acolNombresCamposRef.Count <>
                    acolDatosRef.Count Then
            Throw New ArgumentoInvalidoPanException("Alguna de las colecciones difiere en longitud!")
        End If
        lentRegistrosAfectados = SActualiceReg(MtraTransaccion, astrNombreTabla,
                    acolNombresCamposCambio, acolDatosNuevos, acolNombresCamposRef,
                    acolDatosRef)
        Return lentRegistrosAfectados
    End Function
    Friend Function SActualiceRegistro(astrNombreTabla As String, acolNombresCamposCambio As Collection,
            acolDatosNuevos As Collection, acolNombresCamposRef As Collection,
            acolDatosRef As Collection, acolComparadores As Collection) As Integer
        Dim lentRegistrosAfectados As Integer
        If String.IsNullOrEmpty(astrNombreTabla) OrElse acolNombresCamposCambio Is Nothing OrElse
                acolNombresCamposCambio.Count = 0 OrElse acolDatosNuevos Is Nothing OrElse
                acolNombresCamposRef Is Nothing OrElse acolNombresCamposRef.Count = 0 OrElse
                acolDatosRef Is Nothing OrElse acolComparadores Is Nothing Then
            Dim lstrArgumentoNull As String = String.Empty
            Select Case True
                Case String.IsNullOrEmpty(astrNombreTabla)
                    lstrArgumentoNull = "astrNombreTabla"
                Case acolNombresCamposCambio Is Nothing OrElse acolNombresCamposCambio.Count = 0
                    lstrArgumentoNull = "acolNombresCamposCambio"
                Case acolDatosNuevos Is Nothing
                    lstrArgumentoNull = "acolDatosNuevos"
                Case acolNombresCamposRef Is Nothing OrElse acolNombresCamposRef.Count = 0
                    lstrArgumentoNull = "acolNombresCamposRef"
                Case acolDatosRef Is Nothing
                    lstrArgumentoNull = "acolDatosRef"
                Case acolComparadores Is Nothing
                    lstrArgumentoNull = "acolComparadores"
            End Select
            Throw New ArgumentNullException(lstrArgumentoNull)
        End If
        If acolNombresCamposCambio.Count <> acolDatosNuevos.Count OrElse
                acolNombresCamposRef.Count <> acolDatosRef.Count OrElse
                acolNombresCamposRef.Count <> acolComparadores.Count Then
            Throw New ArgumentoInvalidoPanException("Alguna de las colecciones difiere en longitud!")
        End If
        lentRegistrosAfectados = SActualiceReg(MtraTransaccion, astrNombreTabla,
                    acolNombresCamposCambio, acolDatosNuevos, acolNombresCamposRef,
                    acolDatosRef, acolComparadores)
        Return lentRegistrosAfectados
    End Function
    Private Function SActualiceReg(atraTrans As Object, astrNombreTabla As String,
                acolNombresCamposCambio As Collection, acolDatosNuevos As Collection) As Integer
        Dim lentRegistrosAfectados As Integer = 0
        Dim lstrSql As String
        Dim lcmdComando As Object = FcmdNewComando()
        Dim lblnNoHayError = False
        Try
            SControleProcesoObj(True)
            lstrSql = FstrConstruyaExpSqlActualizar(astrNombreTabla, acolNombresCamposCambio)
            SAbraConexionBd()
            If Not IsNothing(atraTrans) Then
                lcmdComando.Transaction = atraTrans
            End If
            lcmdComando.Connection = CnnConexionBd_App
            lcmdComando.CommandText = lstrSql
            lcmdComando.CommandType = CommandType.Text
            For i As Byte = 1 To acolNombresCamposCambio.Count
                lcmdComando.Parameters.AddWithValue("@" & acolNombresCamposCambio(i),
                        acolDatosNuevos(acolNombresCamposCambio(i)))
            Next
            lentRegistrosAfectados = lcmdComando.ExecuteNonQuery()
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As SqlClient.SqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OleDb.OleDbException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OracleClient.OracleException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As ProveedorBdPanException
            Throw
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lcmdComando) Then
                lcmdComando.Dispose()
            End If
            If lblnNoHayError Then
                SControleProcesoObj(False)
            Else
                SAborteTransaccion()
                SControleProcesoObj(False, True)
            End If
        End Try
        Return lentRegistrosAfectados
    End Function
    Private Function SActualiceReg(atraTrans As Object, astrNombreTabla As String,
                acolNombresCamposCambio As Collection, acolDatosNuevos As Collection,
                acolNombresCamposRef As Collection, acolDatosRef As Collection) As Integer
        Dim lentRegistrosAfectados = 0
        Dim lstrSql As String
        Dim lcmdComando As Object = FcmdNewComando()
        Dim lblnNoHayError = False
        Try
            SControleProcesoObj(True)
            lstrSql = FstrConstruyaExpSqlActualizar(astrNombreTabla, acolNombresCamposCambio,
                        acolNombresCamposRef)
            SAbraConexionBd()
            If Not IsNothing(atraTrans) Then
                lcmdComando.Transaction = atraTrans
            End If
            lcmdComando.Connection = CnnConexionBd_App
            lcmdComando.CommandText = lstrSql
            lcmdComando.CommandType = CommandType.Text
            For i As Byte = 1 To acolNombresCamposCambio.Count
                lcmdComando.Parameters.AddWithValue("@" & acolNombresCamposCambio(i),
                        acolDatosNuevos(acolNombresCamposCambio(i)))
            Next
            For i = 1 To acolNombresCamposRef.Count
                lcmdComando.Parameters.AddWithValue("@Ref" & acolNombresCamposRef(i),
                        acolDatosRef(i))
            Next
            lentRegistrosAfectados = lcmdComando.ExecuteNonQuery()
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As SqlClient.SqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OleDb.OleDbException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OracleClient.OracleException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As ProveedorBdPanException
            Throw
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lcmdComando) Then
                lcmdComando.Dispose()
            End If
            If lblnNoHayError Then
                SControleProcesoObj(False)
            Else
                SAborteTransaccion()
                SControleProcesoObj(False, True)
            End If
        End Try
        Return lentRegistrosAfectados
    End Function
    Private Function SActualiceReg(atraTrans As Object, astrNombreTabla As String,
            acolNombresCamposCambio As Collection, acolDatosNuevos As Collection,
            acolNombresCamposRef As Collection, acolDatosRef As Collection,
            acolComparadores As Collection) As Integer
        Dim lentRegistrosAfectados As Integer = 0
        Dim lstrSql As String
        Dim lcmdComando As Object = FcmdNewComando()
        Dim lblnNoHayError = False
        Try
            SControleProcesoObj(True)
            lstrSql = FstrConstruyaExpSqlActualizar(astrNombreTabla, acolNombresCamposCambio,
                        acolNombresCamposRef, acolComparadores)
            SAbraConexionBd()
            If Not IsNothing(atraTrans) Then
                lcmdComando.Transaction = atraTrans
            End If
            lcmdComando.Connection = CnnConexionBd_App
            lcmdComando.CommandText = lstrSql
            lcmdComando.CommandType = CommandType.Text
            For i As Byte = 1 To acolNombresCamposCambio.Count
                lcmdComando.Parameters.AddWithValue("@" & acolNombresCamposCambio(i),
                        acolDatosNuevos(acolNombresCamposCambio(i)))
            Next
            For i = 1 To acolNombresCamposRef.Count
                lcmdComando.Parameters.AddWithValue("@Ref" & acolNombresCamposRef(i),
                        acolDatosRef(acolNombresCamposRef(i)))
            Next
            lentRegistrosAfectados = lcmdComando.ExecuteNonQuery()
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As SqlClient.SqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OleDb.OleDbException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OracleClient.OracleException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As ProveedorBdPanException
            Throw
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lcmdComando) Then
                lcmdComando.Dispose()
            End If
            If lblnNoHayError Then
                SControleProcesoObj(False)
            Else
                SAborteTransaccion()
                SControleProcesoObj(False, True)
            End If
        End Try
        Return lentRegistrosAfectados
    End Function
    ''' <summary>
    ''' Elimina los registros de la tabla "astrNombreTabla" cuyos campos contenidos en "acolNombresCamposRef"
    ''' tengan los valores contenidos en "acolDatosRef"
    ''' </summary>
    ''' <param name="astrNombreTabla">Nombre de la tabla de la cual se eliminaran los registros</param>
    ''' <param name="acolNombresCamposRef">Colección que contiene los nombres de los campos que referencian
    ''' registros a ser eliminado.</param>
    ''' <param name="acolDatosRef">Colección que contiene los datos que deben tener los registros en las
    ''' columnas de referencia para ser borrados.</param>
    ''' <returns>Cantidad de registros borrados</returns>
    ''' <remarks>La items de la coleccion "acolNombresCamposRef" y los items de la coleccion "acolDatosRef"
    ''' de deben corrersponder uno a uno. </remarks>
    Friend Function SElimineRegistro(astrNombreTabla As String, acolNombresCamposRef As Collection,
            acolDatosRef As Collection) As Integer
        Dim lblnNoHayError = False
        Dim lentRegistrosAfectados As Integer = 0
        Dim lstrSql As String
        Dim lcmdComando As Object = FcmdNewComando()
        Try
            SControleProcesoObj(True)
            lstrSql = FstrConstruyaExpSqlEliminar(astrNombreTabla,
                        acolNombresCamposRef)
            SAbraConexionBd()
            lcmdComando.Connection = CnnConexionBd_App
            lcmdComando.CommandText = lstrSql
            lcmdComando.CommandType = CommandType.Text
            For i = 1 To acolNombresCamposRef.Count
                lcmdComando.Parameters.AddWithValue("@Ref" & acolNombresCamposRef(i),
                        acolDatosRef(i))
            Next
            lentRegistrosAfectados = lcmdComando.ExecuteNonQuery()
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As SqlClient.SqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OleDb.OleDbException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OracleClient.OracleException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As ProveedorBdPanException
            Throw
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lcmdComando) Then
                lcmdComando.Dispose()
            End If
            If lblnNoHayError Then
                SControleProcesoObj(False)
            Else
                SAborteTransaccion()
                SControleProcesoObj(False, True)
            End If
        End Try
        Return lentRegistrosAfectados
    End Function
#End Region

#Region "Operciones con registros a traves de conexion especificada en parametro"
    Friend Function SInserteRegistro(acnnConexion As Object, aenuProveedorBD As EnuProveedorBD,
                astrNombreTabla As String, acolNombreCampos As Collection,
                acolDatos As Collection) As Integer
        If acolNombreCampos.Count <> acolDatos.Count Then
            Throw New ArgumentoInvalidoPanException("Alguna de las colecciones difiere en longitud!")
        End If
        Dim lentRegistrosAfectados As Integer = 0
        Dim lcmdComando As Object = FcmdNewComando(aenuProveedorBD)
        Dim lblnNoHayError = False
        Try
            Dim lstrSql = FstrConstruyaExpSqlInsertar(astrNombreTabla, acolNombreCampos)
            If acnnConexion.State = ConnectionState.Closed Then
                acnnConexion.Open()
            End If
            If Not IsNothing(MtraTransaccion) AndAlso MtraTransaccion.GetType.Name() <>
                   "MySqlTransaction" Then
                lcmdComando.Transaction = MtraTransaccion
            End If
            lcmdComando.Connection = acnnConexion
            lcmdComando.CommandText = lstrSql
            lcmdComando.CommandType = CommandType.Text
            For i As Byte = 1 To acolNombreCampos.Count
                lcmdComando.Parameters.AddWithValue("@" & acolNombreCampos(i),
                        acolDatos(i))
            Next
            lentRegistrosAfectados = lcmdComando.ExecuteNonQuery()
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As SqlClient.SqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OleDb.OleDbException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OracleClient.OracleException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lcmdComando) Then
                lcmdComando.Dispose()
            End If
            If Not lblnNoHayError Then
                SAborteTransaccion()
                SControleProcesoObj(False, True)
            End If
        End Try
        Return lentRegistrosAfectados
    End Function
    Friend Function SLimpieTabla(acnnConexion As Object, aenuProveedorBD As EnuProveedorBD,
            astrNombreTabla As String) As Integer
        If acnnConexion Is Nothing OrElse aenuProveedorBD = EnuProveedorBD.None OrElse
                String.IsNullOrEmpty(astrNombreTabla) Then
            Dim lstrArgumentoNull As String = String.Empty
            Select Case True
                Case acnnConexion Is Nothing
                    lstrArgumentoNull = "acnnConexion"
                Case aenuProveedorBD = EnuProveedorBD.None
                    lstrArgumentoNull = "aenuProveedorBD"
                Case String.IsNullOrEmpty(astrNombreTabla)
                    lstrArgumentoNull = "astrNombreTabla"
            End Select
            Throw New ArgumentNullException(lstrArgumentoNull)
        End If
        Dim lblnNoHayError = False
        Dim lentRegistrosAfectados As Integer = 0
        Dim lstrSql As String = "DELETE FROM " & astrNombreTabla
        Dim lcmdComando As Object = FcmdNewComando(aenuProveedorBD)
        Try
            If acnnConexion.State = ConnectionState.Closed Then
                acnnConexion.Open()
            End If
            lcmdComando.Connection = acnnConexion
            lcmdComando.CommandText = lstrSql
            lcmdComando.CommandType = CommandType.Text
            lentRegistrosAfectados = lcmdComando.ExecuteNonQuery()
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As SqlClient.SqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OleDb.OleDbException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OracleClient.OracleException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As ProveedorBdPanException
            Throw
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lcmdComando) Then
                lcmdComando.Dispose()
            End If
            If Not lblnNoHayError Then
                SAborteTransaccion()
                SControleProcesoObj(False, True)
            End If
        End Try
        Return lentRegistrosAfectados
    End Function
#End Region

#Region "Operaciones con BD"
#Region "Manejo Conexiones"
    Friend Sub SAbraConexionBd()
        If CnnConexionBd_App.State = ConnectionState.Closed Then
            Select Case GenuProveedorBD
                Case EnuProveedorBD.enuMySql
                    Try
                        CnnConexionBd_App.Open()
                    Catch ex As MySqlException
                        Throw New ConexionBdPanException(ex.Message)
                    Catch ex As Exception
                        Throw
                    End Try
                Case EnuProveedorBD.enuOracle
                Case EnuProveedorBD.enuSQLServer
            End Select
        End If
    End Sub
    Private Sub SCierreConexionBd()
        If Not IsNothing(CnnConexionBd_App) Then
            If CnnConexionBd_App.State = ConnectionState.Open Then
                CnnConexionBd_App.Close()
            End If
        End If
    End Sub
    ' Despertador de la conexión
    Friend Sub SDespierteCnn()
        Try
            Dim lblnNoHayError = False
            Do While True
                If CnnConexionBd_App.State = ConnectionState.Closed Then
                    Dim lcnnCnnControl As MySqlConnection = MobjCnnDat.FcnnNewMySqlCon
                    Dim lentRegistrosAfectados As Integer = 0
                    Dim lcmdComando As Object = FcmdNewComando()
                    Dim lstrTabla = "PanCarpetas"
                    Dim lstrCampSel As String() = {"*"}
                    Dim lstrExpSql = FstrConstruyaExpSqlSelect(lstrTabla, lstrCampSel,
                            {{"", ""}}, String.Empty, {})
                    Try
                        lcnnCnnControl.Open()
                        lcmdComando.Connection = lcnnCnnControl
                        lcmdComando.CommandText = lstrExpSql
                        lcmdComando.CommandType = CommandType.Text
                        lentRegistrosAfectados = lcmdComando.ExecuteNonQuery()
                        SEspere()
                        lcnnCnnControl.Close()
                        lblnNoHayError = True
                    Catch ex As MySqlException
                        Throw New ProveedorBdPanException(ex.Message)
                    Catch ex As SqlClient.SqlException
                        Throw New ProveedorBdPanException(ex.Message)
                    Catch ex As OleDb.OleDbException
                        Throw New ProveedorBdPanException(ex.Message)
                    Catch ex As OracleClient.OracleException
                        Throw New ProveedorBdPanException(ex.Message)
                    Catch ex As PanDatException
                        Throw
                    Catch ex As Exception
                        Throw
                    Finally
                        If Not IsNothing(lcmdComando) Then
                            lcmdComando.Dispose()
                        End If
                    End Try
                End If
            Loop
        Catch ex As ConexionBdPanException
            Throw
        End Try
    End Sub
    Private Sub SEspere()
        Dim ldtmInicio As DateTime
        Dim ldtmFin = Now.AddSeconds(1)
        Do While ldtmInicio <= ldtmFin
            Exit Sub
            ldtmInicio = Now
        Loop
    End Sub

#End Region
    ''' <summary>
    ''' Devuelve una DataTable a partir del archivo de origen que puede ser un archivo de Access o uno de Excel
    ''' </summary>
    ''' <param name="astrArchivoOrigen">Nombre completo del archivo de origen.</param>
    ''' <param name="astrContrasena">La contraseña que tiene el archivo.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function FdtbTablasAccess(astrArchivoOrigen As String,
            astrContrasena As String) As DataTable
        If String.IsNullOrEmpty(astrArchivoOrigen) OrElse astrContrasena Is Nothing Then
            Dim lstrNombreArgumentoNulo As String
            If String.IsNullOrEmpty(astrArchivoOrigen) Then
                lstrNombreArgumentoNulo = "astrArchivoOrigen"
            Else
                lstrNombreArgumentoNulo = "astrContrasena"
            End If
            Throw New ArgumentNullException(lstrNombreArgumentoNulo)
        End If
        Dim ldtbTablas As DataTable
        Dim lblnEsOleDbExcel As Boolean = (astrArchivoOrigen.EndsWith(".xls"))
        Dim lblnEsExcel10 As Boolean = (astrArchivoOrigen.EndsWith(".xlsx"))
        Dim lstrTrayectoria As String = FstrObtengaUnidad(astrArchivoOrigen) &
                FstrObtengaTrayectoria(astrArchivoOrigen)
        Dim lcnnConexion As OleDb.OleDbConnection
        Try
            If lblnEsExcel10 Then
                lcnnConexion = MobjCnnDat.FcnnConexion(EnuProveedorBD.enuExcel10,
                        EnuTipoAutenticacion.None, "", lstrTrayectoria,
                        FstrObtengaNombre(astrArchivoOrigen), "", astrContrasena, "",
                        lblnEsOleDbExcel, False)
            Else
                lcnnConexion = MobjCnnDat.FcnnConexion(EnuProveedorBD.enuOleDb,
                        EnuTipoAutenticacion.None, "", lstrTrayectoria,
                        FstrObtengaNombre(astrArchivoOrigen), "", astrContrasena, "",
                        lblnEsOleDbExcel, False)
            End If
            lcnnConexion.Open()
            ldtbTablas = lcnnConexion.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
                    New Object() {Nothing, Nothing, Nothing, "Table"})
            lcnnConexion.Close()
        Catch ex As ParametrosConexionBDPanException
            Throw
        Catch ex As ProveedorBdPanException
            Throw
        Catch ex As Exception
            Throw
        End Try
        Return ldtbTablas
    End Function
    ''' <summary>
    ''' Ejecuta la sentencia SQL espresada en el argumento astrSql y devuelve la cantidad de registros afectados
    ''' </summary>
    ''' <param name="astrSql">Cadena de caracteres que expresa la sentencia a ejecutar.</param>
    ''' <remarks></remarks>
    Friend Function SEjecuteSentenciaSql(astrSql As String) As Integer
        If String.IsNullOrEmpty(astrSql) Then
            Throw New ArgumentNullException(NameOf(astrSql))
        End If
        Dim lblnNoHayError = False
        Dim lentRegistrosAfectados As Integer = 0
        Dim lcmdComando As Object = FcmdNewComando()
        SControleProcesoObj(True)
        Try
            SAbraConexionBd()
            lcmdComando.Connection = CnnConexionBd_App
            lcmdComando.CommandText = astrSql
            lcmdComando.CommandType = CommandType.Text
            lentRegistrosAfectados = lcmdComando.ExecuteNonQuery()
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As SqlClient.SqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OleDb.OleDbException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As OracleClient.OracleException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As PanDatException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lcmdComando) Then
                lcmdComando.Dispose()
            End If
            If lblnNoHayError Then
                SControleProcesoObj(False)
            Else
                SAborteTransaccion()
                SControleProcesoObj(False, True)
            End If
        End Try
        Return lentRegistrosAfectados
    End Function
    Friend Function FobjMaxValorCampo(astrNombreTabla As String, astrNombreCampo As String,
                astrFiltro As String) As Object
        Dim lblnNoHayError = False
        Dim lobjResultado As Object = Nothing
        Dim lcmdComando As Object = FcmdNewComando()
        Try
            SControleProcesoObj(True)
            Dim lstrSql As String = FstrConstruyaExpSqlSelectMax(astrNombreTabla,
                astrNombreCampo, astrFiltro)
            SAbraConexionBd()
            lcmdComando.Connection = CnnConexionBd_App
            lcmdComando.CommandText = lstrSql
            lcmdComando.CommandType = CommandType.Text
            lobjResultado = lcmdComando.ExecuteScalar()
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New SintaxisInvalidaPanException(ex.Message)
        Catch ex As SqlClient.SqlException
            Throw New SintaxisInvalidaPanException(ex.Message)
        Catch ex As OleDb.OleDbException
            Throw New SintaxisInvalidaPanException(ex.Message)
        Catch ex As OracleClient.OracleException
            Throw New SintaxisInvalidaPanException(ex.Message)
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lcmdComando) Then
                lcmdComando.Dispose()
            End If
            If lblnNoHayError Then
                SControleProcesoObj(False)
            Else
                SAborteTransaccion()
                SControleProcesoObj(False, True)
            End If
        End Try
        Return lobjResultado
    End Function
    Friend Function FshrCantidadRegistros(astrNombreTabla As String, astrNombreCampo As String,
                astrCondicion As String)
        If String.IsNullOrEmpty(astrNombreTabla) OrElse String.IsNullOrEmpty(astrNombreCampo) Then
            Dim lstrNombreArgumentoNulo As String = String.Empty
            Select Case True
                Case String.IsNullOrEmpty(astrNombreTabla)
                    lstrNombreArgumentoNulo = "astrNombreTabla"
                Case String.IsNullOrEmpty(astrNombreCampo)
                    lstrNombreArgumentoNulo = "astrNombreCampo"
            End Select
            Throw New ArgumentNullException(lstrNombreArgumentoNulo)
        End If
        Dim lblnNoHayError = False
        Dim lshrResultado As Short = 0
        Dim lcmdComando As Object = FcmdNewComando()
        Try
            SControleProcesoObj(True)
            Dim lstrSql As String = FstrConstruyaExpSqlCount(astrNombreTabla, astrNombreCampo, astrCondicion)
            SAbraConexionBd()
            lcmdComando.Connection = CnnConexionBd_App
            lcmdComando.CommandText = lstrSql
            lcmdComando.CommandType = CommandType.Text
            lshrResultado = CType(lcmdComando.ExecuteScalar(), Short)
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New SintaxisInvalidaPanException(ex.Message)
        Catch ex As SqlClient.SqlException
            Throw New SintaxisInvalidaPanException(ex.Message)
        Catch ex As OleDb.OleDbException
            Throw New SintaxisInvalidaPanException(ex.Message)
        Catch ex As OracleClient.OracleException
            Throw New SintaxisInvalidaPanException(ex.Message)
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lcmdComando) Then
                lcmdComando.Dispose()
            End If
            If lblnNoHayError Then
                SControleProcesoObj(False)
            Else
                SAborteTransaccion()
                SControleProcesoObj(False, True)
            End If
        End Try
        Return lshrResultado
    End Function
    ''' <summary>
    ''' Devuelve la Versión de la base de datos de la aplicación.
    ''' </summary>
    ''' <returns>Entero que identifica la Versión de la base de datos de la aplicación</returns>
    ''' <remarks></remarks>
    Friend Function FentVersionBD(ashrIdAplicacion As Short) As Integer
        Dim lentVersionBD As Integer, lstrNomTabla As String
        Dim lcmdComando As Object = FcmdNewComando()
        Dim lblnNoHayError = False
        Dim lstrSql As String
        If GshrIdAplicacion = 999 Then
            lstrNomTabla = "TraVersiones"
        Else
            lstrNomTabla = "PanVersiones"
        End If
        lstrSql = "SELECT Version FROM " & lstrNomTabla & " WHERE IdAplicacion = " &
                ashrIdAplicacion
        Try
            SAbraConexionBd()
            lcmdComando.Connection = CnnConexionBd_App
            lcmdComando.CommandText = lstrSql
            lcmdComando.CommandType = CommandType.Text
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New SintaxisInvalidaPanException(ex.Message)
        Catch ex As SqlClient.SqlException
            Throw New SintaxisInvalidaPanException(ex.Message)
        Catch ex As OleDb.OleDbException
            Throw New SintaxisInvalidaPanException(ex.Message)
        Catch ex As OracleClient.OracleException
            Throw New SintaxisInvalidaPanException(ex.Message)
        Catch ex As PanDatException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lcmdComando) Then
                lcmdComando.Dispose()
            End If
            If lblnNoHayError Then
                lentVersionBD = lcmdComando.ExecuteScalar()
            Else
                SControleProcesoObj(False, True)
                lentVersionBD = 0
            End If
        End Try
        If ashrIdAplicacion = GshrIdAplicacion Then
            MentVersionBD = lentVersionBD
        End If
        Return lentVersionBD
    End Function
    ''' <summary>
    ''' Devuelve un array de string donde cada elemento esta compuesto por el nombre de la tabla y el nombre 
    ''' de la columna separados por una coma para todas aquellas tablas que contienen una columna cuyo nombre
    ''' es igual al pasado en el argumento "astmrNombreColumna"
    ''' </summary>
    ''' <param name="astrNombreColumna">Nombre de la columna que sera buscado en todas las tablas 
    ''' de la base de datos</param>
    ''' <param name="ablnComo">Indica si el nombre de la columna debe coincidir exactamente o contiene el
    ''' string pasado en el parametro "astrNombreColumna"</param>
    ''' <param name="ablnPan">Indica si se buscará en las tablas de Panorama o en las tablas de la Aplicación.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function FstrNombreTablasContienenColumna(astrNombreColumna As String,
            ablnComo As Boolean, ablnPan As Boolean) As String()
        If astrNombreColumna Is Nothing Then
            Throw New ArgumentNullException(NameOf(astrNombreColumna))
        End If
        Dim lobjColumnaTbl As ClsColumna = Nothing
        Dim lblnIncluir As Boolean = False
        Dim lstrNomTbls As String() = Nothing
        Dim j As Short = -1
        Dim lobjEstructuraBD As ClsBaseDatos = Nothing
        If ablnPan Then
            lobjEstructuraBD = FobjEstructuraBD(EnuListaAplicaciones.EnuAdministrador)
        Else
            lobjEstructuraBD = FobjEstructuraBD(EnuListaAplicaciones.EnuOrionCop)
        End If
        If Not IsNothing(lobjEstructuraBD) Then
            With lobjEstructuraBD
                For Each lobjTabla As ClsTabla In .ColTablas
                    For i As UShort = 1 To lobjTabla.ColColumnas.Count
                        lobjColumnaTbl = lobjTabla.ColColumnas(i)
                        If ablnComo Then
                            lblnIncluir = (lobjColumnaTbl.StrNombre.ToUpper.Contains(astrNombreColumna.ToUpper))
                        Else
                            lblnIncluir = (lobjColumnaTbl.StrNombre.ToUpper = astrNombreColumna.ToUpper)
                        End If
                        If lblnIncluir Then
                            j += 1
                            ReDim Preserve lstrNomTbls(j)
                            lstrNomTbls(j) = lobjTabla.StrNombre & "," & lobjColumnaTbl.StrNombre
                            Exit For
                        End If
                    Next
                Next
            End With
        End If
        Return lstrNomTbls
    End Function
    ''' <summary>
    ''' Devuelve un array de arrays donde los elementos de cada array secundario esta compuesto por el nombre de 
    ''' la tabla y el nombre de la columna separados por una coma, siempre y cuando la tabla  
    ''' contenga todas las columnas pasadas en el argumento "astrNombresColumnas"
    ''' </summary>
    ''' <param name="astrNombresColumnas">Nombres de las columnas que seran buscados en todas las tablas 
    ''' de la base de datos</param>
    ''' <param name="ablnComo">Indica si el nombre de la columna debe coincidir exactamente o contiene el
    ''' string pasado en el elemento correrspondiente del argumentos "astrNombresColumnas"</param>
    ''' <param name="ablnPan">Indica si se buscará en las tablas de Panorama o en las tablas de la Aplicación.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function FstrNombreTablasContienenColumnas(astrNombresColumnas As String(),
            ablnComo As Boolean, ablnPan As Boolean) As String()()
        If astrNombresColumnas Is Nothing Then
            Throw New ArgumentNullException(NameOf(astrNombresColumnas))
        End If
        Dim lstrColumnasTabla As String() = Nothing
        Dim lstrTblsCols As String()() = Nothing
        Dim j As Short = -1
        Dim lobjEstructuraBD As ClsBaseDatos = Nothing
        If ablnPan Then
            lobjEstructuraBD = FobjEstructuraBD(EnuListaAplicaciones.EnuAdministrador)
        Else
            lobjEstructuraBD = FobjEstructuraBD(EnuListaAplicaciones.EnuOrionCop)
        End If
        If Not IsNothing(lobjEstructuraBD) Then
            With lobjEstructuraBD
                For Each lobjTabla As ClsTabla In .ColTablas
                    lstrColumnasTabla = FstrColumnasEnTabla(lobjTabla, astrNombresColumnas, ablnComo)
                    If Not IsNothing(lstrColumnasTabla) Then
                        j += 1
                        ReDim Preserve lstrTblsCols(j)
                        lstrTblsCols(j) = lstrColumnasTabla
                    End If
                Next
            End With
        End If
        Return lstrTblsCols
    End Function
    ''' <summary>
    ''' Devuelve un array de string donde cada elemento esta compuesto por el nombre de la tabla y el nombre 
    ''' de la columna separados por una coma, siempre y cuando la tabla pasada en el argumento 'aobjTabla' 
    ''' contenga todas las columnas pasadas en el argumento "astrNombresColumnas", de lo contrario devuelve
    ''' 'Nothing'
    ''' </summary>
    ''' <param name="aobjTabla">Tabla en la cual se buscaran los nombres de las columnas</param>
    ''' <param name="astrNombresColumnas">Array que contiene los nombres de las columnas a buscar.</param>
    ''' <param name="ablnComo">Indica si el nombre de la columna debe coincidir exactamente o contiene el
    ''' string pasado en el elemento correrspondiente del argumento "astrNombresColumnas"</param>
    ''' <returns>Un array o Nothing</returns>
    ''' <remarks></remarks>
    Private Shared Function FstrColumnasEnTabla(aobjTabla As ClsTabla, astrNombresColumnas As String(),
            ablnComo As Boolean) As String()
        Dim lstrNomCols As String() = Nothing
        Dim lobjColumnaTbl As ClsColumna
        Dim lblnIncluir As Boolean
        Dim i As Integer, j = -1
        For Each lstrNomColBuscada As String In astrNombresColumnas
            For i = 1 To aobjTabla.ColColumnas.Count
                lobjColumnaTbl = aobjTabla.ColColumnas(i)
                If ablnComo Then
                    lblnIncluir = (lobjColumnaTbl.StrNombre.ToUpper.Contains(lstrNomColBuscada.ToUpper))
                Else
                    lblnIncluir = (lobjColumnaTbl.StrNombre.ToUpper = lstrNomColBuscada.ToUpper)
                End If
                If lblnIncluir Then
                    j += 1
                    ReDim Preserve lstrNomCols(j)
                    lstrNomCols(j) = aobjTabla.StrNombre & "," & lobjColumnaTbl.StrNombre
                    Exit For
                End If
            Next
        Next
        If (j + 1) = astrNombresColumnas.Length Then
            Return lstrNomCols
        Else
            Return Nothing
        End If
    End Function
    '''<summary>
    ''' Devuelve un array con los nombres de las tablas que contienen al menos un campo cuyo nombre
    ''' sea iagual o empiece por el nombre pasado en "astrNombreColumna".
    ''' </summary>
    ''' <returns>Array con los nombres de las tablas</returns>
    ''' <remarks>Se busca en la base de datos correspondiente a la conexión actual (CnnConexionBd_App).</remarks>
    Friend Function FstrNombreTablasContienenColumna(astrNombreColumnaComo As String) As String()
        If String.IsNullOrEmpty(astrNombreColumnaComo) Then
            Throw New ArgumentNullException(NameOf(astrNombreColumnaComo))
        End If
        Dim lstrNombreTablas() As String = Nothing
        Dim i As Short = -1, lblnNoHayError = False
        Select Case GenuProveedorBD
            Case EnuProveedorBD.enuMySql
                Try
                    SControleProcesoObj(True)
                    SAbraConexionBd()
                    Dim lstrNomBD = StrNombreBD
                    Dim ldtbCampos As DataTable = CnnConexionBd_App.GetSchema("Columns")
                    Dim lstrFiltro As String = ("TABLE_CATALOG = 'def' AND TABLE_SCHEMA ='" &
                            lstrNomBD & "' AND " & "COLUMN_NAME LIKE '" & astrNombreColumnaComo & "'")
                    Dim ldrwDefColumnas As DataRow() = ldtbCampos.Select(lstrFiltro)
                    For Each ldrwDefCol As DataRow In ldrwDefColumnas
                        If IsNothing(lstrNombreTablas) Then
                            i += 1
                            ReDim lstrNombreTablas(i)
                            lstrNombreTablas(i) = ldrwDefCol("TABLE_NAME")
                        Else
                            If Not lstrNombreTablas.Contains(ldrwDefCol("TABLE_NAME")) Then
                                i += 1
                                ReDim Preserve lstrNombreTablas(i)
                                lstrNombreTablas(i) = ldrwDefCol("TABLE_NAME")
                            End If
                        End If
                    Next
                    lblnNoHayError = True
                Catch ex As MySqlException
                    Throw New SintaxisInvalidaPanException(ex.Message)
                Catch ex As PanDatException
                    Throw
                Catch ex As Exception
                    Throw
                Finally
                    If lblnNoHayError Then
                        SControleProcesoObj(False)
                    Else
                        SAborteTransaccion()
                        SControleProcesoObj(False, True)
                    End If
                End Try
        End Select
        Return lstrNombreTablas
    End Function
    ''' <summary>
    ''' Devuelve un array con los nombres de las tablas que contienen al menos un campo cuyo nombre
    ''' sea iagual o empiece por el nombre pasado en "astrNombreColumna".
    ''' </summary>
    ''' <param name="acnnConexion"></param>
    ''' Conexión a la base de datos que se quiere examinar
    ''' <param name="astrNombreColumnaComo">String contenido en el nombre de la columna a seleccionar</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Shared Function FstrNombreTablasContienenColumna(acnnConexion As Object,
            astrNombreColumnaComo As String) As String()
        If String.IsNullOrEmpty(astrNombreColumnaComo) OrElse acnnConexion Is Nothing Then
            Dim lstrArgumentoNulo As String = String.Empty
            Select Case True
                Case String.IsNullOrEmpty(astrNombreColumnaComo)
                    lstrArgumentoNulo = "astrNombreColumnaComo"
                Case acnnConexion Is Nothing
                    lstrArgumentoNulo = "acnnConexion"
            End Select
            Throw New ArgumentNullException(lstrArgumentoNulo)
        End If
        Dim lstrNombreTablas() As String = Nothing
        Dim i As Short = -1
        Select Case GenuProveedorBD
            Case EnuProveedorBD.enuMySql
                Try
                    If acnnConexion.State = ConnectionState.Closed Then
                        acnnConexion.Open()
                    End If
                    Dim ldtbCampos As DataTable = acnnConexion.GetSchema("Columns")
                    Dim lstrNomBD = StrNombreBD
                    acnnConexion.Close()
                    Dim lstrFiltro As String = ("TABLE_CATALOG = 'def' AND TABLE_SCHEMA ='" &
                            lstrNomBD & "' AND " & "COLUMN_NAME LIKE '" & astrNombreColumnaComo & "'")
                    Dim ldrwDefColumnas As DataRow() = ldtbCampos.Select(lstrFiltro)
                    For Each ldrwDefCol As DataRow In ldrwDefColumnas
                        If IsNothing(lstrNombreTablas) Then
                            i += 1
                            ReDim lstrNombreTablas(i)
                            lstrNombreTablas(i) = ldrwDefCol("TABLE_NAME")
                        Else
                            If Not lstrNombreTablas.Contains(ldrwDefCol("TABLE_NAME")) Then
                                i += 1
                                ReDim Preserve lstrNombreTablas(i)
                                lstrNombreTablas(i) = ldrwDefCol("TABLE_NAME")
                            End If
                        End If
                    Next
                Catch ex As MySqlException
                    Throw New ProveedorBdPanException(ex.Message)
                Catch ex As Exception
                    Throw
                End Try
        End Select
        Return lstrNombreTablas
    End Function
    ''' <summary>
    ''' Devuelve un array con los nombres de las columnas de la tabla "astrNombreTabla" que empiezan 
    ''' o contienen el string pasado en el argumento "astrNombreColumnaComo"
    ''' </summary>
    ''' <param name="astrNombreTabla">Nombre de la tabala que se va a examinar</param>
    ''' <param name="astrNombreColumnaComo">String que debe contener el nombre de la columna para formar parte del array
    ''' a ser devuelto.</param>
    ''' <returns>Array con los nombres de los campos</returns>
    ''' <remarks>La busqueda se hace en la base de datos corrrespondiente a la conexión actual:(CnnConexionBd_App)</remarks>
    Friend Function FstrNombresColumnasEnTabla(astrNombreTabla As String,
            astrNombreColumnaComo As String) As String()
        If String.IsNullOrEmpty(astrNombreTabla) OrElse String.IsNullOrEmpty(astrNombreColumnaComo) Then
            Dim lstrArgumentoNulo As String = String.Empty
            Select Case True
                Case String.IsNullOrEmpty(astrNombreTabla)
                    lstrArgumentoNulo = "astrNombreTabla"
                Case String.IsNullOrEmpty(astrNombreColumnaComo)
                    lstrArgumentoNulo = "astrNombreColumnaComo"
            End Select
            Throw New ArgumentNullException(lstrArgumentoNulo)
        End If
        Dim lstrNombreCampos() As String = Nothing
        Dim i As Short = -1, lblnNoHayError = False
        Select Case GenuProveedorBD
            Case EnuProveedorBD.enuMySql
                Try
                    SControleProcesoObj(True)
                    SAbraConexionBd()
                    Dim lstrNomBD = StrNombreBD
                    Dim ldtbCampos As DataTable = CnnConexionBd_App.GetSchema("Columns")
                    Dim lstrFiltro As String = ("TABLE_CATALOG = 'def' AND TABLE_SCHEMA ='" &
                            lstrNomBD & "' AND " & "TABLE_NAME = '" & astrNombreTabla &
                            "' AND " & "COLUMN_NAME LIKE '" & astrNombreColumnaComo & "'")
                    Dim ldrwDefColumnas As DataRow() = ldtbCampos.Select(lstrFiltro)
                    For Each ldrwDefCol In ldrwDefColumnas
                        If IsNothing(lstrNombreCampos) Then
                            i += 1
                            ReDim lstrNombreCampos(i)
                            lstrNombreCampos(i) = ldrwDefCol("COLUMN_NAME")
                        Else
                            If Not lstrNombreCampos.Contains(ldrwDefCol("TABLE_NAME")) Then
                                i += 1
                                ReDim Preserve lstrNombreCampos(i)
                                lstrNombreCampos(i) = ldrwDefCol("COLUMN_NAME")
                            End If
                        End If
                    Next
                    lblnNoHayError = True
                Catch ex As MySqlException
                    Throw New ProveedorBdPanException(ex.Message)
                Catch ex As PanDatException
                    Throw
                Catch ex As Exception
                    Throw
                Finally
                    If lblnNoHayError Then
                        SControleProcesoObj(False)
                    Else
                        SControleProcesoObj(False, True)
                    End If
                End Try
        End Select
        Return lstrNombreCampos
    End Function
    ''' <summary>
    ''' Devuelve un array con los nombres de las columnas de la tabla "astrNombreTabla" que empiezan 
    ''' o contienen el string pasado en el argumento "astrNombreColumnaComo"
    ''' </summary>
    ''' <param name="acnnConexion">Conexion a la base de datos donde se buscaran las tablas.</param>
    ''' <param name="astrNombreTabla">Nombre de la tabala que se va a examinar</param>
    ''' <param name="astrNombreColumnaComo">String que debe contener el nombre de la columna para formar parte del array
    ''' a ser devuelto.</param>
    ''' <returns>Array con los nombres de los campos</returns>
    ''' <remarks></remarks>
    Friend Function FstrNombresColumnasEnTabla(acnnConexion As Object, astrNombreTabla As String,
            astrNombreColumnaComo As String) As String()
        If String.IsNullOrEmpty(astrNombreTabla) OrElse String.IsNullOrEmpty(astrNombreColumnaComo) OrElse
                acnnConexion Is Nothing Then
            Dim lstrArgumentoNulo As String = String.Empty
            Select Case True
                Case String.IsNullOrEmpty(astrNombreTabla)
                    lstrArgumentoNulo = "astrNombreTabla"
                Case String.IsNullOrEmpty(astrNombreColumnaComo)
                    lstrArgumentoNulo = "astrNombreColumnaComo"
                Case acnnConexion Is Nothing
                    lstrArgumentoNulo = "acnnConexion"
            End Select
            Throw New ArgumentNullException(lstrArgumentoNulo)
        End If
        Dim lstrNombreCampos() As String = Nothing
        Dim i As Short = -1
        Select Case GenuProveedorBD
            Case EnuProveedorBD.enuMySql
                Try
                    If acnnConexion.State = ConnectionState.Closed Then
                        acnnConexion.Open()
                    End If
                    Dim lstrNomBD = StrNombreBD
                    Dim ldtbCampos As DataTable = CnnConexionBd_App.GetSchema("Columns")
                    acnnConexion.Close()
                    Dim lstrFiltro As String = ("TABLE_CATALOG = 'def' AND TABLE_SCHEMA ='" &
                            lstrNomBD & "' AND " & "TABLE_NAME = '" & astrNombreTabla &
                            "' AND " & "COLUMN_NAME LIKE '" & astrNombreColumnaComo & "'")
                    Dim ldrwDefColumnas As DataRow() = ldtbCampos.Select(lstrFiltro)
                    For Each ldrwDefCol In ldrwDefColumnas
                        If IsNothing(lstrNombreCampos) Then
                            i += 1
                            ReDim lstrNombreCampos(i)
                            lstrNombreCampos(i) = ldrwDefCol("COLUMN_NAME")
                        Else
                            If Not lstrNombreCampos.Contains(ldrwDefCol("TABLE_NAME")) Then
                                i += 1
                                ReDim Preserve lstrNombreCampos(i)
                                lstrNombreCampos(i) = ldrwDefCol("COLUMN_NAME")
                            End If
                        End If
                    Next
                Catch ex As MySqlException
                    Throw New ProveedorBdPanException(ex.Message)
                Catch ex As Exception
                    Throw
                End Try
        End Select
        Return lstrNombreCampos
    End Function
#End Region

#Region "Manejo DataAdapters"
    Private Shared Function FdapAdaptador(aenuProveedorBD As EnuProveedorBD,
            astrSql As String, acnnConexion As Object,
            aenuTipoSentencia As EnuTipoSentencia) As Object
        If aenuProveedorBD = EnuProveedorBD.None OrElse String.IsNullOrEmpty(astrSql) OrElse
                acnnConexion Is Nothing Then
            Dim lstrNombreArgumentoNulo As String = String.Empty
            Select Case True
                Case aenuProveedorBD = EnuProveedorBD.None
                    lstrNombreArgumentoNulo = "aenuProveedorBD"
                Case String.IsNullOrEmpty(astrSql)
                    lstrNombreArgumentoNulo = "astrSql"
                Case acnnConexion Is Nothing
                    lstrNombreArgumentoNulo = "acnnConexion"
            End Select
            Throw New ArgumentNullException(lstrNombreArgumentoNulo)
        End If
        Dim ldapDataAdaptador As Object = Nothing
        Select Case aenuProveedorBD
            Case EnuProveedorBD.enuMySql
                ldapDataAdaptador = FdapAdaptadorMySql(astrSql, acnnConexion, aenuTipoSentencia)
            Case EnuProveedorBD.enuOracle
            Case EnuProveedorBD.enuSQLServer
            Case EnuProveedorBD.enuOleDb
                ldapDataAdaptador = FdapAdaptadorOleDb(astrSql, acnnConexion, aenuTipoSentencia)
        End Select
        Return ldapDataAdaptador
        ldapDataAdaptador.Dispose()
    End Function
    Private Shared Function FdapAdaptadorMySql(astrSql As String, acnnConexion As Object,
            aenuTipoSentencia As EnuTipoSentencia) As MySqlDataAdapter
        Dim ldapDataAdaptador As MySqlDataAdapter = Nothing
        Dim lblnNoHayError = False
        Try
            Dim lcnnConexion As MySqlConnection = acnnConexion
            ldapDataAdaptador = New MySqlDataAdapter
            Select Case aenuTipoSentencia
                Case EnuTipoSentencia.enuSelect
                    ldapDataAdaptador.SelectCommand = New MySqlCommand(astrSql, lcnnConexion)
                Case EnuTipoSentencia.enuUpdate
                    ldapDataAdaptador.UpdateCommand = New MySqlCommand(astrSql, lcnnConexion)
                Case EnuTipoSentencia.enuInsert
                    ldapDataAdaptador.InsertCommand = New MySqlCommand(astrSql, lcnnConexion)
                Case EnuTipoSentencia.enuDelete
                    ldapDataAdaptador.DeleteCommand = New MySqlCommand(astrSql, lcnnConexion)
            End Select
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New PanDatException(ex.Message)
        Catch ex As Exception
            Throw
        Finally
            If Not lblnNoHayError Then
                ldapDataAdaptador.Dispose()
            End If
        End Try
        Return ldapDataAdaptador
    End Function
    Private Shared Function FdapAdaptadorOleDb(astrSql As String, acnnConexion As Object,
            aenuTipoSentencia As EnuTipoSentencia) As OleDbDataAdapter
        Dim ldapDataAdaptador As OleDbDataAdapter = Nothing
        Dim lblnNoHayError = False
        Try
            Dim lcnnConexion As OleDb.OleDbConnection = acnnConexion
            ldapDataAdaptador = New OleDb.OleDbDataAdapter
            Select Case aenuTipoSentencia
                Case EnuTipoSentencia.enuSelect
                    ldapDataAdaptador.SelectCommand = New OleDb.OleDbCommand(astrSql, lcnnConexion)
                Case EnuTipoSentencia.enuUpdate
                    ldapDataAdaptador.UpdateCommand = New OleDb.OleDbCommand(astrSql, lcnnConexion)
                Case EnuTipoSentencia.enuInsert
                    ldapDataAdaptador.InsertCommand = New OleDb.OleDbCommand(astrSql, lcnnConexion)
                Case EnuTipoSentencia.enuDelete
                    ldapDataAdaptador.DeleteCommand = New OleDb.OleDbCommand(astrSql, lcnnConexion)
            End Select
            lblnNoHayError = True
        Catch ex As OleDbException
            Throw New PanDatException(ex.Message)
        Catch ex As Exception
            Throw
        Finally
            If Not lblnNoHayError Then
                ldapDataAdaptador.Dispose()
            End If
        End Try
        Return ldapDataAdaptador
    End Function
#End Region

#Region "Manejo de transacciones y operaciones de control"
    Private Shared Sub SVerifiqueCarpetas()
        If Not My.Computer.FileSystem.DirectoryExists(GstrTrayDat) Then
            My.Computer.FileSystem.CreateDirectory(GstrTrayDat)
        End If
        If Not My.Computer.FileSystem.DirectoryExists(GstrTrayDatPrg) Then
            My.Computer.FileSystem.CreateDirectory(GstrTrayDatPrg)
        End If
        Dim lstrCarpCopiasSeg = GstrTrayDat & "CopiasSeguridad"
        If Not My.Computer.FileSystem.DirectoryExists(lstrCarpCopiasSeg) Then
            My.Computer.FileSystem.CreateDirectory(lstrCarpCopiasSeg)
        End If
    End Sub
    ''' <summary>
    ''' Inicia una transacción (BeginTransaction) sobre la conexion actual(mcnnConexion)
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub SInicialiceTransaccion()
        Dim lblnNoHayError = False
        Try
            SAbraConexionBd()
            If MshrIdTransaccion = -1 Then
                MtraTransaccion = CnnConexionBd_App.BeginTransaction()
            End If
            MshrIdTransaccion += 1
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New ProveedorBdPanException(ex.ToString)
        Catch ex As SqlClient.SqlException
            Throw New ProveedorBdPanException(ex.ToString)
        Catch ex As OleDbException
            Throw New ProveedorBdPanException(ex.ToString)
        Catch ex As OracleClient.OracleException
            Throw New ProveedorBdPanException(ex.ToString)
        Catch ex As PanDatException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not lblnNoHayError Then
                SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Friend Sub SConfirmeTransaccion()
        If MshrIdTransaccion >= 0 Then
            Dim lblnNoHayError = False
            Try
                If MshrIdTransaccion = 0 Then
                    MtraTransaccion.Commit()
                    MtraTransaccion = Nothing
                End If
                MshrIdTransaccion -= 1
                lblnNoHayError = True
            Catch ex As MySqlException
                Throw New ProveedorBdPanException(ex.ToString)
            Catch ex As SqlClient.SqlException
                Throw New ProveedorBdPanException(ex.ToString)
            Catch ex As OleDbException
                Throw New ProveedorBdPanException(ex.ToString)
            Catch ex As OracleClient.OracleException
                Throw New ProveedorBdPanException(ex.ToString)
            Catch ex As Exception
                Throw
            Finally
                If Not lblnNoHayError Then
                    SAborteTransaccion()
                    SControleProcesoObj(False, True)
                End If
            End Try
        End If
    End Sub
    Friend Sub SAborteTransaccion()
        If MshrIdTransaccion >= 0 Then
            Dim lblnNoHayError = False
            Try
                If CnnConexionBd_App.State = ConnectionState.Open Then
                    If Not IsNothing(MtraTransaccion) Then
                        MtraTransaccion.Rollback()
                    End If
                End If
                lblnNoHayError = True
            Catch ex As MySqlException
                Throw New ProveedorBdPanException(ex.ToString)
            Catch ex As SqlClient.SqlException
                Throw New ProveedorBdPanException(ex.ToString)
            Catch ex As OleDbException
                Throw New ProveedorBdPanException(ex.ToString)
            Catch ex As OracleClient.OracleException
                Throw New ProveedorBdPanException(ex.ToString)
            Catch ex As Exception
                Throw
            Finally
                MtraTransaccion = Nothing
                MshrIdTransaccion = -1
                If Not lblnNoHayError Then
                    SControleProcesoObj(False, True)
                End If
            End Try
        End If
    End Sub
    Public Sub SControleProcesoObj(ablnProcesando As Boolean)
        If ablnProcesando Then
            MshrProcesos += 1
        Else
            If MshrProcesos > 0 Then
                MshrProcesos -= 1
            End If
            If MshrProcesos = 0 Then
                SCierreConexionBd()
            End If
        End If
    End Sub
    Public Sub SControleProcesoObj(ablnProcesando As Boolean, ablnForzarCierre As Boolean)
        If Not ablnProcesando AndAlso ablnForzarCierre Then
            MshrProcesos = 0
            SCierreConexionBd()
        Else
            If ablnProcesando Then
                MshrProcesos += 1
            Else
                MshrProcesos -= 1
            End If
            If MshrProcesos = 0 Then
                SCierreConexionBd()
            End If
        End If
    End Sub
    Public Function FentProceso() As Integer
        Return MshrProcesos
    End Function
#End Region

#Region "Copia de Seguridad"
    Friend Sub SRestaureBkPan(astrArchivoCopia As String)
        Using lcnnNewConeccion As MySqlConnection = MobjCnnDat.FcnnNewMySqlCon
            lcnnNewConeccion.Open()
            Dim lcmdComando = New MySql.Data.MySqlClient.MySqlCommand With {
                .Connection = lcnnNewConeccion
            }
            MmbBackup = New MySqlBackup
            MmbBackup.ImportInfo.IntervalForProgressReport = 100
            MmbBackup.Command = lcmdComando
            MmbBackup.ImportFromFile(astrArchivoCopia)
            MmbBackup.Dispose()
        End Using
    End Sub
    Private Sub EvnBackup_ImporChange(sender As Object, e As ImportProgressArgs) _
            Handles MmbBackup.ImportProgressChanged
        Dim llngCantBytes As Long = CLng(e.TotalBytes)
        Dim llngCantProc As Long = CLng(e.CurrentBytes)
        If llngCantBytes > 0 Then
            EntProcientoAvance = Int((llngCantProc / llngCantBytes) * 100)
        Else
            EntProcientoAvance = 0
        End If
    End Sub
    Friend Sub SGenereBkPan(astrArchivoCopia As String)
        Dim lstrRutaArchivoSQL As String = astrArchivoCopia
        Dim lstrArchivoCopiaZip = astrArchivoCopia.Substring(0, astrArchivoCopia.Length - 3) & "zip"
        Using lcnnNewConeccion As MySqlConnection = MobjCnnDat.FcnnNewMySqlCon
            lcnnNewConeccion.Open()
            Dim lcmdComando = New MySql.Data.MySqlClient.MySqlCommand With {
                .Connection = lcnnNewConeccion
            }
            MmbBackup = New MySqlBackup
            MmbBackup.ExportInfo.IntervalForProgressReport = 100
            MmbBackup.ExportInfo.MaxSqlLength = 250000
            MmbBackup.Command = lcmdComando
            MmbBackup.ExportToFile(astrArchivoCopia)
            MmbBackup.Dispose()
            If My.Computer.FileSystem.FileExists(lstrArchivoCopiaZip) Then
                My.Computer.FileSystem.DeleteFile(lstrArchivoCopiaZip)
            End If
            Using lzaArchivoZip As ZipArchive = ZipFile.Open(lstrArchivoCopiaZip,
                    ZipArchiveMode.Create)
                lzaArchivoZip.CreateEntryFromFile(lstrRutaArchivoSQL,
                        Path.GetFileName(lstrRutaArchivoSQL), CompressionLevel.Optimal)
            End Using
            My.Computer.FileSystem.DeleteFile(lstrRutaArchivoSQL)
        End Using
    End Sub
    Private Sub EvnBackup_ExporChange(sender As Object, e As ExportProgressArgs) _
            Handles MmbBackup.ExportProgressChanged
        Dim llngCantRows As Long = CLng(e.TotalRowsInAllTables)
        Dim llngCantProc As Long = CLng(e.CurrentRowIndexInAllTables)
        EntProcientoAvance = Int((llngCantProc / llngCantRows) * 100)
    End Sub
    Friend Sub SCanceleAccionCopia()
        If Not IsNothing(MmbBackup) Then
            MmbBackup.StopAllProcess()
            MmbBackup.Command.Cancel()
            MmbBackup.Dispose()
        End If
    End Sub
    Friend Property EntProcientoAvance As Integer = 0
#End Region
#End Region

#Region "Funciones Varias"
    Private Shared Function FcmdNewComando(aenuProveedorBD As EnuProveedorBD) As Object
        Dim lcmdComando As Object = Nothing
        Select Case aenuProveedorBD
            Case EnuProveedorBD.enuSQLServer
                lcmdComando = New SqlClient.SqlCommand
            Case EnuProveedorBD.enuOracle
                lcmdComando = New OleDb.OleDbCommand
            Case EnuProveedorBD.enuOleDb
                lcmdComando = New OleDb.OleDbCommand
            Case EnuProveedorBD.enuMySql
                lcmdComando = New MySql.Data.MySqlClient.MySqlCommand
            Case EnuProveedorBD.enuOracle
                'lcmdComando = New System.Data.OracleClient.OracleCommand
        End Select
        Return lcmdComando
    End Function
    Private Shared Function FcmdNewComando() As Object
        Dim lcmdComando As Object = Nothing
        Select Case GenuProveedorBD
            Case EnuProveedorBD.enuSQLServer
                lcmdComando = New SqlClient.SqlCommand
            Case EnuProveedorBD.enuOleDb
                lcmdComando = New OleDb.OleDbCommand
            Case EnuProveedorBD.enuMySql
                lcmdComando = New MySql.Data.MySqlClient.MySqlCommand
            Case EnuProveedorBD.enuOracle
                'lcmdComando = New System.Data.OracleClient.OracleCommand
        End Select
        Return lcmdComando
    End Function
    ''' <summary>
    ''' Normaliza el nombre del campo
    ''' </summary>
    ''' <remarks>Excepciones:SintaxisInvalidaPanException </remarks>
    ''' 
    Friend Shared Sub SNormaliceNombreCampo(ByRef astrNombreCampo As String, aenuTipoTabla As EnuTipoTabla)
        Dim lstrNombreCampo As String
        Dim lstrPrefijo As String = String.Empty
        Dim lstrSufijo As String = String.Empty
        Dim lblnCampoSinCalificar As Boolean = False
        Dim lbytPosParentesis As Byte
        astrNombreCampo = astrNombreCampo.Trim()
        If astrNombreCampo.ToUpper.Trim.StartsWith("SUM(") OrElse
                astrNombreCampo.ToUpper.StartsWith("MIN(") OrElse
                astrNombreCampo.ToUpper.StartsWith("MAX(") OrElse
                astrNombreCampo.ToUpper.StartsWith("COUNT(") Then
            lstrNombreCampo = FstrNombreCampo(astrNombreCampo, lstrPrefijo, lstrSufijo)
        ElseIf astrNombreCampo.ToUpper.StartsWith("FPERIODO(") Then
            lbytPosParentesis = astrNombreCampo.IndexOf("(")
            If astrNombreCampo.Contains("=") Then
                lstrSufijo = astrNombreCampo.Substring(astrNombreCampo.IndexOf("=") + 1)
                lstrNombreCampo = astrNombreCampo.Substring(lbytPosParentesis + 1, astrNombreCampo.IndexOf("=") -
                        (lbytPosParentesis + 2))
                lstrNombreCampo = FstrNombreCampoPerido(lstrNombreCampo)
            Else
                Throw New SintaxisInvalidaPanException("El argumento '" & astrNombreCampo & "' no esta bien escrito!")
            End If
        ElseIf astrNombreCampo.ToUpper.StartsWith("IIF(") Then
            Select Case GenuProveedorBD
                Case EnuProveedorBD.enuMySql
                    lstrNombreCampo = astrNombreCampo.Replace("IIF", "If")
                Case Else
                    Throw New PanDatException("Proveedor de Datos no esperado!")
            End Select
        ElseIf astrNombreCampo.ToUpper.StartsWith("CONCAT(") Then
            lstrNombreCampo = astrNombreCampo
            lblnCampoSinCalificar = True
        ElseIf astrNombreCampo.ToUpper.StartsWith("DISTINCT ") Then
            If Not astrNombreCampo.Contains("'") Then
                Dim lbytPosEspacio = astrNombreCampo.Trim.IndexOf(" ")
                lstrPrefijo = astrNombreCampo.Substring(0, lbytPosEspacio + 1)
                lstrNombreCampo = astrNombreCampo.Substring(lbytPosEspacio + 1, astrNombreCampo.Length -
                        (lbytPosEspacio + 1))
            Else
                lstrNombreCampo = astrNombreCampo
            End If
        ElseIf astrNombreCampo.Contains("=") Then
            lstrNombreCampo = astrNombreCampo.Substring(0, astrNombreCampo.IndexOf("="))
            lstrSufijo = astrNombreCampo.Substring(astrNombreCampo.IndexOf("=") + 1)
            lblnCampoSinCalificar = True
        ElseIf astrNombreCampo.Contains("+") OrElse astrNombreCampo.Contains("-") Then
            lblnCampoSinCalificar = True
            lstrNombreCampo = astrNombreCampo
        Else
            lstrNombreCampo = astrNombreCampo
        End If
        lstrNombreCampo = FstrNombreCampo(aenuTipoTabla, lstrPrefijo, lstrNombreCampo,
                lblnCampoSinCalificar)
        If Not String.IsNullOrEmpty(lstrSufijo) Then
            lstrNombreCampo = lstrNombreCampo & " AS " & lstrSufijo
        End If
        astrNombreCampo = lstrNombreCampo
    End Sub
    Private Shared Function FstrNombreCampo(astrNombreCampo As String,
            ByRef astrPrefijo As String, ByRef astrSufijo As String) As String
        Dim lstrNomCam = String.Empty
        Dim lbytPosParentesis As Byte = astrNombreCampo.IndexOf("(")
        astrPrefijo = astrNombreCampo.Substring(0, lbytPosParentesis + 1)
        If astrNombreCampo.Contains("=") Then
            lstrNomCam = astrNombreCampo.Substring(lbytPosParentesis + 1,
                    astrNombreCampo.IndexOf("=") - (lbytPosParentesis + 2))
            astrSufijo = astrNombreCampo.Substring(astrNombreCampo.IndexOf("=") + 1)
        Else
            If astrNombreCampo.ToUpper.Contains(" AS ") Then
                lstrNomCam = astrNombreCampo.Substring(lbytPosParentesis + 1,
                        astrNombreCampo.Length - (lbytPosParentesis + 1))
                astrSufijo = String.Empty
            Else
                If lstrNomCam.Contains("-") OrElse lstrNomCam.Contains("+") OrElse
                        lstrNomCam.Contains("*") OrElse lstrNomCam.Contains("/") Then
                    astrSufijo = " AS Total "
                Else
                    lstrNomCam = astrNombreCampo.Substring(lbytPosParentesis + 1,
                            astrNombreCampo.Length - (lbytPosParentesis + 2))
                    astrSufijo = lstrNomCam
                End If
            End If
        End If
        Return lstrNomCam
    End Function
    Private Shared Function FstrNombreCampo(aenuTipoTabla As EnuTipoTabla,
            astrPrefijo As String, astrNombreCam As String,
            ablnCamposSinCali As Boolean) As String
        Dim lstrNomCam = astrNombreCam
        Select Case aenuTipoTabla
            Case EnuTipoTabla.enuUnica
                If Not String.IsNullOrEmpty(astrPrefijo) Then
                    lstrNomCam = astrPrefijo & astrNombreCam
                    If Not lstrNomCam.Contains("AS") AndAlso lstrNomCam.Contains("(") Then
                        lstrNomCam &= ")"
                    End If
                End If
            Case EnuTipoTabla.enuPrimaria
                If Not String.IsNullOrEmpty(astrPrefijo) Then
                    lstrNomCam = astrPrefijo & "P." & astrNombreCam
                    If astrPrefijo <> "DISTINCT " Then
                        lstrNomCam &= ")"
                    End If
                Else
                    If Not ablnCamposSinCali Then
                        If Not astrNombreCam.Contains("'") Then
                            lstrNomCam = "P." & astrNombreCam
                        End If
                    End If
                End If
            Case EnuTipoTabla.enuSecundaria
                If Not String.IsNullOrEmpty(astrPrefijo) Then
                    If Not astrNombreCam.Contains("DISTINCT") Then
                        lstrNomCam = astrPrefijo & "S." & astrNombreCam & ")"
                    End If
                Else
                    If Not ablnCamposSinCali Then
                        lstrNomCam = "S." & astrNombreCam
                    End If
                End If
        End Select
        Return lstrNomCam
    End Function
    ''' <summary>
    ''' Descompone el nombre del archivo pasado en el argumento "astrNombreArchivo" en sus diferentes partes 
    ''' devolviendolos en un array donde el primer elemento es la unidad, sigue la trayectoria, sigue al Nombre
    ''' del archivo y termina con la extension.
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Function FstrComponentesRutaArchivo(astrRutaArchvo As String) As String()
        Dim i As Integer
        Dim lstrComRutArch() As String
        Dim lstrUnidad As String
        Dim lstrNombre As String = String.Empty
        Dim lstrTray As String
        Dim lstrExten As String
        Try
            'Determina la astrUnidad
            If astrRutaArchvo.StartsWith("\\") Then
                If astrRutaArchvo.IndexOf("\", 2) > 0 Then
                    lstrUnidad = astrRutaArchvo.Substring(0, astrRutaArchvo.IndexOf("\", 2))
                Else
                    lstrUnidad = astrRutaArchvo
                End If
            ElseIf astrRutaArchvo.IndexOf(":") = 1 Then
                lstrUnidad = astrRutaArchvo.Substring(0, 2)
            Else
                lstrUnidad = String.Empty
            End If
            'Determina el astrRutaArchvo
            If astrRutaArchvo.IndexOf("\", lstrUnidad.Length) > 0 Then
                For i = astrRutaArchvo.Length To 1 Step -1
                    If astrRutaArchvo.Substring(i - 1, 1) = "\" Then
                        lstrNombre = astrRutaArchvo.Substring(i)
                        Exit For
                    End If
                Next
            Else
                lstrNombre = astrRutaArchvo
            End If
            'Determina la astrExtension
            If lstrNombre.IndexOf(".") > 0 Then
                lstrExten = lstrNombre.Substring(lstrNombre.IndexOf(".") + 1)
            Else
                lstrExten = String.Empty
            End If
            'Determina astrTrayectoria
            lstrTray = astrRutaArchvo.Substring(lstrUnidad.Length, astrRutaArchvo.Length -
                    lstrUnidad.Length - lstrNombre.Length)
            lstrComRutArch = {lstrUnidad, lstrTray, lstrNombre, lstrExten}
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        End Try
        Return lstrComRutArch
    End Function
    ''' <summary>
    ''' Devuelve la que forma parte de la trayectoria de archivo pasada en el argumento "astrNombreArchivo"
    ''' </summary>
    ''' <param name="astrNombreArchivo">Nombre de archivo completo.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Shared Function FstrObtengaUnidad(astrNombreArchivo As String) As String
        'Retorna la lstrUnidad o Servidor a partir de un lstrNombre de Archivo
        Dim lstrRutaArch As String()
        lstrRutaArch = FstrComponentesRutaArchivo(astrNombreArchivo)
        Return lstrRutaArch(0)
    End Function
    Friend Shared Function FstrObtengaTrayectoria(astrNombreArchivo As String) As String
        'Retorna la lstrTrayectoria o Path a partir de un lstrNombre de Archivo
        Dim lstrRutaArch As String()
        lstrRutaArch = FstrComponentesRutaArchivo(astrNombreArchivo)
        Return lstrRutaArch(1)
    End Function
    Friend Shared Function FstrObtengaNombre(astrNombreArchivo As String) As String
        'Retorna exclusivamente el lstrNombre a partir de un lstrNombre de Archivo
        Dim lstrRutaArch As String()
        lstrRutaArch = FstrComponentesRutaArchivo(astrNombreArchivo)
        Return lstrRutaArch(2)
    End Function
    ''' <summary>
    ''' Devuelve una expresion para ser usada en una expresion Sql con el fin de obtener a partir de un campo DATE
    ''' una cadena que represente el periodo del año conformado por "aaaamm" donde "aaaa" es el año y "mm" el mes.
    ''' </summary>
    ''' <param name="astrNombreCampo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function FstrNombreCampoPerido(astrNombreCampo As String) As String
        Dim lstrNombreCampo As String = String.Empty
        Select Case GenuProveedorBD
            Case EnuProveedorBD.enuMySql
                lstrNombreCampo = "CONCAT(CAST(YEAR(" & astrNombreCampo & ") AS CHAR), " &
                        "RIGHT(CONCAT('00', CAST(MONTH(" & astrNombreCampo & ") as CHAR)),2))"
            Case Else
                '
        End Select
        Return lstrNombreCampo
    End Function
    Private Shared Function FblnHayIndice(astrCamposIndice(,) As String) As Boolean
        If Not IsNothing(astrCamposIndice) AndAlso astrCamposIndice.Length > 0 Then
            For i = 0 To astrCamposIndice.GetUpperBound(0)
                If Not String.IsNullOrEmpty(astrCamposIndice(i, 0)) Then
                    Return True
                End If
            Next
        End If
        Return False
    End Function
    Private Shared Function FblnHayCamposRef(acolNombresCamposRef As Collection) As Boolean
        If Not IsNothing(acolNombresCamposRef) Then
            For i As Byte = 1 To acolNombresCamposRef.Count
                If Not String.IsNullOrEmpty(acolNombresCamposRef(i).ToString) Then
                    Return True
                End If
            Next
        End If
        Return False
    End Function
    Friend Shared Function FsrStreamReader(astrArchivo As String) As StreamReader
        Dim lsrArchivo As StreamReader
        Try
            lsrArchivo = File.OpenText(astrArchivo)
        Catch ex As FileNotFoundException
            Throw New PanDatException(ex.Message)
        Catch ex As IOException
            Throw New PanDatException(ex.Message)
        Catch ex As Exception
            Throw
        End Try
        Return lsrArchivo
    End Function
    Public Shared Function FswStreamWriter(astrArchivo As String) As StreamWriter
        If astrArchivo Is Nothing Then
            Throw New ArgumentNullException(NameOf(astrArchivo))
        End If
        Dim lblnNoHayError = False
        Dim lswArchivo As StreamWriter = Nothing
        Try
            lswArchivo = File.CreateText(astrArchivo)
            lblnNoHayError = True
        Catch ex As ArgumentException
            Throw
        Catch ex As PathTooLongException
            Throw
        Catch ex As DirectoryNotFoundException
            Throw
        Catch ex As NotSupportedException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not lblnNoHayError Then
                lswArchivo.Dispose()
            End If
        End Try
        Return lswArchivo
        lswArchivo.Dispose()
    End Function
#Region "Funciones que devuelven expresiones sql"
    Private Shared Function FstrConstruyaExpSqlInsertar(astrNombreTabla As String,
                acolNombreCampos As Collection) As String
        Dim lstbExpresion As New StringBuilder
        Dim lstrSql As String
        Try
            lstbExpresion.EnsureCapacity(500)
            Select Case GenuProveedorBD
                Case EnuProveedorBD.enuMySql
                    lstbExpresion.Append("INSERT INTO ")
                    lstbExpresion.Append(astrNombreTabla)
                    lstbExpresion.Append(" (")
                    For Each lobjObjeto In acolNombreCampos
                        lstbExpresion.Append(lobjObjeto.ToString)
                        lstbExpresion.Append(", ")
                    Next
                    lstrSql = lstbExpresion.ToString
                    lstrSql = lstrSql.Remove(lstrSql.Length - 2, 2)
                    lstbExpresion.Clear()
                    lstbExpresion.Append(lstrSql)
                    lstbExpresion.Append(") VALUES (")
                    For Each lobjObjeto In acolNombreCampos
                        lstbExpresion.Append("@")
                        lstbExpresion.Append(lobjObjeto.ToString)
                        lstbExpresion.Append(", ")
                    Next
                    lstrSql = lstbExpresion.ToString
                    lstrSql = lstrSql.Remove(lstrSql.Length - 2, 2) & ")"
                Case Else
                    lstrSql = String.Empty
            End Select
            Return lstrSql
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
    Friend Shared Function FstrConstruyaExpSqlActualizar(astrNombreTabla As String,
            acolNombresCamposCambio As Collection) As String
        Dim lstrNombreCampoCambio As String
        Dim lstbExpresion As New StringBuilder
        Dim lstrSql As String
        lstbExpresion.EnsureCapacity(500)
        Dim i As Integer
        Try
            Select Case GenuProveedorBD
                Case EnuProveedorBD.enuMySql
                    lstbExpresion.Append("UPDATE " & astrNombreTabla & " SET ")
                    For i = 1 To acolNombresCamposCambio.Count
                        lstrNombreCampoCambio = acolNombresCamposCambio(i)
                        lstbExpresion.Append(lstrNombreCampoCambio)
                        lstbExpresion.Append(" = ")
                        lstbExpresion.Append("@")
                        lstbExpresion.Append(lstrNombreCampoCambio)
                        lstbExpresion.Append(", ")
                    Next i
                    lstrSql = lstbExpresion.ToString
                    lstrSql = lstrSql.Substring(0, lstrSql.Length - 2)
                Case Else
                    lstrSql = String.Empty
            End Select
            Return lstrSql
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
    Friend Shared Function FstrConstruyaExpSqlActualizar(astrNombreTabla As String,
            acolNombresCamposCambio As Collection, acolNombresCamposRef As Collection) As String
        Dim lstrNombreCampoCambio As String
        Dim lstbExpresion As New StringBuilder
        Dim lblnHayWhere As Boolean = False
        Dim lstrSql As String
        lstbExpresion.EnsureCapacity(500)
        Dim i As Integer
        Try
            Select Case GenuProveedorBD
                Case EnuProveedorBD.enuMySql
                    lstbExpresion.Append("UPDATE " & astrNombreTabla & " SET ")
                    For i = 1 To acolNombresCamposCambio.Count
                        lstrNombreCampoCambio = acolNombresCamposCambio(i)
                        lstbExpresion.Append(lstrNombreCampoCambio)
                        lstbExpresion.Append(" = ")
                        lstbExpresion.Append("@")
                        lstbExpresion.Append(lstrNombreCampoCambio)
                        lstbExpresion.Append(", ")
                    Next i
                    lstrSql = lstbExpresion.ToString
                    lstrSql = lstrSql.Substring(0, lstrSql.Length - 2)
                    lstbExpresion.Clear()
                    lstbExpresion.Append(lstrSql)
                    If FblnHayCamposRef(acolNombresCamposRef) Then
                        lblnHayWhere = True
                        lstbExpresion.Append(" WHERE ")
                        For i = 1 To acolNombresCamposRef.Count
                            lstbExpresion.Append(acolNombresCamposRef(i))
                            lstbExpresion.Append(" = ")
                            lstbExpresion.Append("@Ref")
                            lstbExpresion.Append(acolNombresCamposRef(i))
                            lstbExpresion.Append(" AND ")
                        Next
                    End If
                    lstrSql = lstbExpresion.ToString
                    If lblnHayWhere Then
                        lstrSql = lstrSql.Substring(0, lstrSql.Length - 5)
                    End If
                Case Else
                    lstrSql = String.Empty
            End Select
            Return lstrSql
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch exep As Exception
            Throw
        End Try
    End Function
    Friend Shared Function FstrConstruyaExpSqlActualizar(astrNombreTabla As String,
            acolNombresCamposCambio As Collection, acolNombresCamposRef As Collection,
            acolComparadores As Collection) As String
        Dim lstrNombreCampoCambio As String
        Dim lstbExpresion As New StringBuilder
        Dim lblnHayWhere As Boolean = False
        Dim lstrSql As String
        lstbExpresion.EnsureCapacity(500)
        Dim i As Integer
        Try
            Select Case GenuProveedorBD
                Case EnuProveedorBD.enuMySql
                    lstbExpresion.Append("UPDATE " & astrNombreTabla & " SET ")
                    For i = 1 To acolNombresCamposCambio.Count
                        lstrNombreCampoCambio = acolNombresCamposCambio(i)
                        lstbExpresion.Append(lstrNombreCampoCambio)
                        lstbExpresion.Append(" = ")
                        lstbExpresion.Append("@")
                        lstbExpresion.Append(lstrNombreCampoCambio)
                        lstbExpresion.Append(", ")
                    Next i
                    lstrSql = lstbExpresion.ToString
                    lstrSql = lstrSql.Substring(0, lstrSql.Length - 2)
                    lstbExpresion.Clear()
                    lstbExpresion.Append(lstrSql)
                    If FblnHayCamposRef(acolNombresCamposRef) Then
                        lblnHayWhere = True
                        lstbExpresion.Append(" WHERE ")
                        For i = 1 To acolNombresCamposRef.Count
                            lstbExpresion.Append(acolNombresCamposRef(i))
                            If IsNothing(acolComparadores) Then
                                lstbExpresion.Append(" = ")
                            Else
                                lstbExpresion.Append(acolComparadores(i))
                            End If
                            lstbExpresion.Append("@Ref")
                            lstbExpresion.Append(acolNombresCamposRef(i))
                            lstbExpresion.Append(" AND ")
                        Next
                    End If
                    lstrSql = lstbExpresion.ToString
                    If lblnHayWhere Then
                        lstrSql = lstrSql.Substring(0, lstrSql.Length - 5)
                    End If
                Case Else
                    lstrSql = String.Empty
            End Select
            Return lstrSql
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
    Friend Shared Function FstrConstruyaExpSqlEliminar(astrNombreTabla As String,
            acolNombresCamposRef As Collection) As String
        Dim lstbExpresion As New StringBuilder
        Dim lstrSql As String
        lstbExpresion.EnsureCapacity(150)
        Try
            Select Case GenuProveedorBD
                Case EnuProveedorBD.enuMySql
                    lstbExpresion.Append("DELETE FROM ")
                    lstbExpresion.Append(astrNombreTabla)
                    lstbExpresion.Append(" WHERE ")
                    For i = 1 To acolNombresCamposRef.Count
                        lstbExpresion.Append(acolNombresCamposRef(i))
                        lstbExpresion.Append(" = ")
                        lstbExpresion.Append("@Ref")
                        lstbExpresion.Append(acolNombresCamposRef(i))
                        lstbExpresion.Append(" AND ")
                    Next
                    lstrSql = lstbExpresion.ToString
                    lstrSql = lstrSql.Substring(0, lstrSql.Length - 5)
                Case Else
                    lstrSql = String.Empty
            End Select
            Return lstrSql
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
    Friend Shared Function FstrConstruyaExpSqlEliminar(astrNombreTabla As String,
            astrFiltro As String) As String
        Dim lstbExpresion As New StringBuilder
        Dim lstrSql As String
        lstbExpresion.EnsureCapacity(150)
        Try
            Select Case GenuProveedorBD
                Case EnuProveedorBD.enuMySql
                    lstbExpresion.Append("DELETE FROM ")
                    lstbExpresion.Append(astrNombreTabla)
                    If Not String.IsNullOrEmpty(astrFiltro) Then
                        lstbExpresion.Append(" WHERE ")
                        lstbExpresion.Append(astrFiltro)
                    End If
                    lstrSql = lstbExpresion.ToString
                Case Else
                    lstrSql = String.Empty
            End Select
            Return lstrSql
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' Construye uns expresión SELECT
    ''' </summary>
    ''' <param name="astrNombreTabla">Nombre de la tabla de la cual se seleccionaran los registros.</param>
    ''' <param name="astrCampos">Array que contiene los nombres de los campos que conformaran la nueva tabla.</param>
    ''' <param name="astrIndice">Array que contiene los campos por los cuales se ordenaran los registros de la nueva tabla</param>
    ''' <param name="astrFiltro">Expresion que indica los registros que seran tenidos en cuenta.</param>
    ''' <param name="astrCamposGrupo">Array que contiene los nombres de los campos por los cuales se agruparan los
    ''' registros de la nueva tabla.</param>
    ''' <returns>Espresion SELECT</returns>
    ''' <remarks>Excepciones: ArgumentNullException, SintaxisInvalidaPanException, Exception</remarks>
    Friend Shared Function FstrConstruyaExpSqlSelect(astrNombreTabla As String, astrCampos() As String,
            astrIndice(,) As String, astrFiltro As String, astrCamposGrupo() As String) As String
        Dim lstbExpresion As New StringBuilder
        Dim lstrSql As String
        Dim i As Byte
        Dim lstrCampo As String = String.Empty
        lstbExpresion.EnsureCapacity(500)
        Try
            Select Case GenuProveedorBD
                Case EnuProveedorBD.enuMySql
                    lstbExpresion.Append("SELECT ")
                    For i = 0 To astrCampos.GetUpperBound(0)
                        lstrCampo = astrCampos(i)
                        SNormaliceNombreCampo(lstrCampo, EnuTipoTabla.enuUnica)
                        lstbExpresion.Append(lstrCampo)
                        lstbExpresion.Append(", ")
                    Next
                    lstrSql = lstbExpresion.ToString
                    lstrSql = lstrSql.Substring(0, lstrSql.Length - 2)
                    lstbExpresion.Clear()
                    lstbExpresion.Append(lstrSql)
                    lstbExpresion.Append(" FROM ")
                    lstbExpresion.Append(astrNombreTabla)
                    If Not String.IsNullOrEmpty(astrFiltro) Then
                        lstbExpresion.Append(" WHERE ")
                        lstbExpresion.Append(astrFiltro)
                        lstbExpresion.Append(" ")
                    End If
                    If astrCamposGrupo.Length > 0 AndAlso Not String.IsNullOrEmpty(astrCamposGrupo(0)) Then
                        lstbExpresion.Append(" GROUP BY ")
                        For i = 0 To astrCamposGrupo.GetUpperBound(0)
                            lstbExpresion.Append(astrCamposGrupo(i))
                            lstbExpresion.Append(", ")
                        Next
                        lstrSql = lstbExpresion.ToString
                        lstrSql = lstrSql.Substring(0, lstrSql.Length - 2)
                        lstbExpresion.Clear()
                        lstbExpresion.Append(lstrSql)
                    End If
                    If FblnHayIndice(astrIndice) Then
                        lstbExpresion.Append(" ORDER BY ")
                        For i = 0 To astrIndice.GetUpperBound(0)
                            lstbExpresion.Append(astrIndice(i, 0))
                            lstbExpresion.Append(" ")
                            If Not String.IsNullOrEmpty(astrIndice(i, 1)) Then
                                lstbExpresion.Append(astrIndice(i, 1))
                            Else
                                lstbExpresion.Append("ASC")
                            End If
                            lstbExpresion.Append(", ")
                        Next
                        lstrSql = lstbExpresion.ToString
                        lstrSql = lstrSql.Substring(0, lstrSql.Length - 2)
                    Else
                        lstrSql = lstbExpresion.ToString
                    End If
                Case Else
                    lstrSql = String.Empty
            End Select
            Return lstrSql
        Catch exep As SintaxisInvalidaPanException
            Throw
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch exep As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' Devuelve una expresion "SELECT" compuesta por campos de dos tablas relacionada mediane "INNER JOIN".
    ''' </summary>
    ''' <param name="astrNombreTablaPri">Nombre de la tabla primaria</param>
    ''' <param name="astrCamposTablaPri">Array que contiene los nombres de los campos de la tabla primaria que formaran
    ''' parte de la nueva tabla</param>
    ''' <param name="astrNombreTablaSec">Nombre de la tabla secundaria</param>
    ''' <param name="astrCamposTablaSec">Array que contiene los nombres de los campos de la tabla secundaria que formaran
    ''' parte de la nueva tabla</param>
    ''' <param name="astrCamposPriRel">Array que contiene los nombres de la tabla primaria que determinan la relación
    ''' con la tabla secundaria.</param>
    ''' <param name="astrCamposSecRel">Array que contiene los nombres de la tabla secundaria que determinan la relación
    ''' con la tabla primaria.</param>
    ''' <param name="astrCamposIndice">Array de dos dimensiones que establece el indice de la nueva tabla. 
    ''' Si este array se pasa vacio los registros no se indexan.</param>
    ''' <param name="astrFiltro">String que contiene una expresion para filtrar los registros que formaran la
    ''' nueva tabla. Para no filtrar el resultado este string debe ser vacio:("")</param>
    ''' <param name="astrCamposGrupo">Array que contiene los nombres de los campos que se utilizan para agrupar los
    ''' registros. Si este array se pasa vacio no se hara agrupamiento.</param>
    ''' <returns></returns>
    ''' <remarks>Excepciones:SintaxisInvalidaPanException, Exception</remarks>
    Friend Shared Function FstrConstruyaExpSqlSelect(astrNombreTablaPri As String, astrCamposTablaPri() As String,
            astrNombreTablaSec As String, astrCamposTablaSec() As String, astrCamposPriRel() As String,
            astrCamposSecRel() As String, astrCamposIndice(,) As String, astrFiltro As String,
            astrCamposGrupo() As String) As String
        Dim lstbExpresion As New StringBuilder
        Dim lstrCampo As String = String.Empty
        Dim lstrSql As String
        Dim i As Byte
        lstbExpresion.EnsureCapacity(500)
        Try
            Select Case GenuProveedorBD
                Case EnuProveedorBD.enuMySql
                    lstbExpresion.Append("SELECT ")
                    If astrCamposTablaPri.Length > 0 Then
                        For i = 0 To astrCamposTablaPri.GetUpperBound(0)
                            lstrCampo = astrCamposTablaPri(i)
                            SNormaliceNombreCampo(lstrCampo, EnuTipoTabla.enuPrimaria)
                            lstbExpresion.Append(lstrCampo)
                            lstbExpresion.Append(", ")
                        Next
                    End If
                    If astrCamposTablaSec.Length > 0 Then
                        For i = 0 To astrCamposTablaSec.GetUpperBound(0)
                            lstrCampo = astrCamposTablaSec(i)
                            SNormaliceNombreCampo(lstrCampo, EnuTipoTabla.enuSecundaria)
                            lstbExpresion.Append(lstrCampo)
                            lstbExpresion.Append(", ")
                        Next
                    End If
                    lstrSql = lstbExpresion.ToString
                    lstrSql = lstrSql.Substring(0, lstrSql.Length - 2)
                    lstbExpresion.Clear()
                    lstbExpresion.Append(lstrSql)
                    lstbExpresion.Append(" FROM ")
                    lstbExpresion.Append(astrNombreTablaPri)
                    lstbExpresion.Append(" AS P INNER JOIN ")
                    lstbExpresion.Append(astrNombreTablaSec)
                    lstbExpresion.Append(" AS S ON P.")
                    For i = 0 To astrCamposPriRel.GetUpperBound(0)
                        lstbExpresion.Append(astrCamposPriRel(i))
                        lstbExpresion.Append(" = S.")
                        lstbExpresion.Append(astrCamposSecRel(i))
                        If i < astrCamposPriRel.GetUpperBound(0) Then
                            lstbExpresion.Append(" AND P.")
                        End If
                    Next
                    If Not String.IsNullOrEmpty(astrFiltro) Then
                        lstbExpresion.Append(" WHERE ")
                        lstbExpresion.Append(astrFiltro)
                        lstbExpresion.Append(" ")
                    End If
                    If astrCamposGrupo.Length > 0 Then
                        lstbExpresion.Append(" GROUP BY ")
                        For i = 0 To astrCamposGrupo.GetUpperBound(0)
                            lstbExpresion.Append(astrCamposGrupo(i))
                            lstbExpresion.Append(", ")
                        Next
                        lstrSql = lstbExpresion.ToString
                        lstrSql = lstrSql.Substring(0, lstrSql.Length - 2)
                        lstbExpresion.Clear()
                        lstbExpresion.Append(lstrSql)
                    End If
                    If FblnHayIndice(astrCamposIndice) Then
                        lstbExpresion.Append(" ORDER BY ")
                        For i = 0 To astrCamposIndice.GetUpperBound(0)
                            lstbExpresion.Append(astrCamposIndice(i, 0))
                            lstbExpresion.Append(" ")
                            If Not String.IsNullOrEmpty(astrCamposIndice(i, 0)) Then
                                lstbExpresion.Append(astrCamposIndice(i, 1))
                            Else
                                lstbExpresion.Append("ASC")
                            End If
                            lstbExpresion.Append(", ")
                        Next
                        lstrSql = lstbExpresion.ToString
                        lstrSql = lstrSql.Substring(0, lstrSql.Length - 2)
                    Else
                        lstrSql = lstbExpresion.ToString
                    End If
                Case Else
                    lstrSql = String.Empty
            End Select
            lstrSql = lstrSql.Replace("P.''", "''")
            lstrSql = lstrSql.Replace("P.0", "0")
            If lstrSql.Contains("S.DATE_FORMAT") Then
                lstrSql = lstrSql.Replace("S.DATE_FORMAT", "DATE_FORMAT")
            End If
            If lstrSql.Contains("P.DATE_FORMAT") Then
                lstrSql = lstrSql.Replace("P.DATE_FORMAT", "DATE_FORMAT")
            End If
            Return lstrSql
        Catch exep As SintaxisInvalidaPanException
            Throw
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch exep As Exception
            Throw
        End Try
    End Function
    Friend Shared Function FstrConstruyaExpSqlCount(astrNombreTabla As String, astrNombreCampo As String,
                astrCondicion As String) As String
        Dim lstrSql As String
        Dim lstbExpresion As New StringBuilder
        lstbExpresion.EnsureCapacity(150)
        Try
            Select Case GenuProveedorBD
                Case EnuProveedorBD.enuMySql
                    lstbExpresion.Append("SELECT COUNT(")
                    lstbExpresion.Append(astrNombreCampo)
                    lstbExpresion.Append(") AS Cant FROM ")
                    lstbExpresion.Append(astrNombreTabla)
                    If Not String.IsNullOrEmpty(astrCondicion) Then
                        lstbExpresion.Append(" WHERE ")
                        lstbExpresion.Append(astrCondicion)
                    End If
                    lstrSql = lstbExpresion.ToString
                Case Else
                    lstrSql = String.Empty
            End Select
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch exep As Exception
            Throw
        End Try
        Return lstrSql
    End Function
    Friend Shared Function FstrConstruyaExpSqlSelectMax(astrNombreTabla As String, astrNombreCampo As String,
                Optional astrCondicion As String = "") As String
        Dim lstbExpresion As New StringBuilder
        Dim lstrSql As String
        lstbExpresion.EnsureCapacity(150)
        Try
            Select Case GenuProveedorBD
                Case EnuProveedorBD.enuMySql
                    lstbExpresion.Append("SELECT MAX(")
                    lstbExpresion.Append(astrNombreCampo)
                    lstbExpresion.Append(") FROM ")
                    lstbExpresion.Append(astrNombreTabla)
                    If Not String.IsNullOrEmpty(astrCondicion) Then
                        lstbExpresion.Append(" WHERE ")
                        lstbExpresion.Append(astrCondicion)
                    End If
                    lstrSql = lstbExpresion.ToString
                Case Else
                    lstrSql = String.Empty
            End Select
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch exep As Exception
            Throw
        End Try
        Return lstrSql
    End Function
#End Region
#Region "Funciones de fecha"
    ''' <summary>
    ''' Devuelve la fecha normalizada  expresada en un string según el proveedor de base de datos actual
    ''' </summary>
    ''' <param name="astrFecha">String que representa la fecha que se quiere normalizar</param>
    ''' <remarks></remarks>
    Friend Shared Function FstrFechaNormalizada(astrFecha As String) As String
        Dim lstrFechaNormailizada As String = String.Empty
        If IsDate(astrFecha) Then
            Select Case GenuProveedorBD
                Case EnuProveedorBD.enuMySql
                    lstrFechaNormailizada = FstrFechaNormalizadaMySql(astrFecha)
                Case EnuProveedorBD.enuOracle
                Case EnuProveedorBD.enuSQLServer
                Case Else
                    '
            End Select
        Else
            Throw New PanDatException("Texto no corresponde a una fecha!")
        End If
        Return lstrFechaNormailizada
    End Function
    ''' <summary>
    ''' Devuelve la fecha y la hora normalizada  expresada en un string según el proveedor de base de datos actual
    ''' </summary>
    ''' <param name="astrFecha">String que representa la fecha que se quiere normalizar</param>
    ''' <remarks></remarks>
    Friend Shared Function FstrFechaHoraNormalizada(astrFecha As String) As String
        Dim lstrFechaNormailizada As String = String.Empty
        If IsDate(astrFecha) Then
            Select Case GenuProveedorBD
                Case EnuProveedorBD.enuMySql
                    lstrFechaNormailizada = FstrFechaHoraNormalizadaMySql(astrFecha)
                Case EnuProveedorBD.enuOracle
                    '
                Case EnuProveedorBD.enuSQLServer
                    '
            End Select
        Else
            Throw New PanDatException("Texto no corresponde a una fecha!")
        End If
        Return lstrFechaNormailizada
    End Function
    Private Shared Function FstrFechaNormalizadaMySql(astrfecha As String) As String
        Dim ldtmFecha As Date = CType(astrfecha, Date)
        Dim lstrfecha As String = CStr(ldtmFecha.Year) & "-" & Format(ldtmFecha.Month, "0#") & "-" &
                Format(ldtmFecha.Day, "0#")
        Return lstrfecha
    End Function
    Private Shared Function FstrFechaHoraNormalizadaMySql(astrfecha As String) As String
        Dim lstrHora As String
        Dim lstrHoras As String
        Dim lbytHoras As Byte
        Dim lstrPartesFecha As String()
        Dim ldtmFecha As Date = CType(astrfecha, Date)
        Dim lstrfecha As String = CStr(ldtmFecha.Year) & "-" & CStr(ldtmFecha.Month) & "-" & CStr(ldtmFecha.Day)
        If astrfecha.Contains(" ") Then
            lstrPartesFecha = astrfecha.Split(" ")
            lstrHora = astrfecha.Split(" ")(1)
            If lstrPartesFecha.Length > 2 AndAlso astrfecha.Split(" ")(2).ToUpper.StartsWith("P") Then
                lstrHoras = lstrHora.Split(":")(0)
                If CByte(lstrHoras) <> 12 Then
                    lbytHoras = CByte(lstrHoras) + 12
                Else
                    lbytHoras = CByte(lstrHoras)
                End If
                If lbytHoras = 24 Then
                    lbytHoras = 0
                End If
                If CStr(lbytHoras).Length = 1 Then
                    lstrHora = "0" & CStr(lbytHoras) & lstrHora.Substring(2)
                Else
                    lstrHora = CStr(lbytHoras) & lstrHora.Substring(lstrHora.IndexOf(":"))
                End If
            Else
                If lstrHora.StartsWith("12") Then
                    lstrHora = "00" & lstrHora.Substring(2)
                End If
            End If
            lstrfecha = lstrfecha & " " & lstrHora
        End If
        Return lstrfecha
    End Function
    Friend Shared Function FstrFechaHoraFinDia(astrFecha As String) As String
        Dim lstrFechaFinDia As String = String.Empty
        If IsDate(astrFecha) Then
            Select Case GenuProveedorBD
                Case EnuProveedorBD.enuMySql
                    lstrFechaFinDia = FstrFechaHoraFinDiaMySql(astrFecha)
                Case EnuProveedorBD.enuOracle
                    '
                Case EnuProveedorBD.enuSQLServer
                    '
            End Select
        Else
            Throw New PanDatException("Texto no corresponde a una fecha!")
        End If
        Return lstrFechaFinDia
    End Function
    Friend Shared Function FstrFechaHoraInicioDia(astrFecha As String) As String
        Dim lstrFechaFinDia As String = String.Empty
        If IsDate(astrFecha) Then
            Select Case GenuProveedorBD
                Case EnuProveedorBD.enuMySql
                    lstrFechaFinDia = FstrFechaHoraInicioDiaMySql(astrFecha)
                Case EnuProveedorBD.enuOracle
                    '
                Case EnuProveedorBD.enuSQLServer
                    '
            End Select
        Else
            Throw New PanDatException("Texto no corresponde a una fecha!")
        End If
        Return lstrFechaFinDia
    End Function
    Private Shared Function FstrFechaHoraFinDiaMySql(astrFecha As String) As String
        If astrFecha IsNot Nothing Then
            Dim lushPosEspacio As UShort = astrFecha.IndexOf(" ")
            Dim lstrFechaFinDiaMySql As String = astrFecha.Substring(0, lushPosEspacio) & " 23:59:59"
            Return lstrFechaFinDiaMySql
        Else
            Throw New ArgumentNullException(NameOf(astrFecha))
        End If
    End Function
    Private Shared Function FstrFechaHoraInicioDiaMySql(astrFecha As String) As String
        If astrFecha IsNot Nothing Then
            Dim lushPosEspacio As UShort = astrFecha.IndexOf(" ")
            Dim lstrFechaFinDiaMySql As String = astrFecha.Substring(0, lushPosEspacio) & " 00:00:01"
            Return lstrFechaFinDiaMySql
        Else
            Throw New ArgumentNullException(NameOf(astrFecha))
        End If
    End Function
#End Region
#End Region

#Region "Dispose"
    Protected Overridable Sub Dispose(ablnDisposing As Boolean)
        If Not MblnDisposed Then
            If ablnDisposing Then
                CnnConexionBd_App.Dispose()
                CnnConexionBd_App = Nothing
            End If
            MblnDisposed = True
        End If
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class