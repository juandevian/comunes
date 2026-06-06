Imports System.IO
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports MySql.Data.MySqlClient
Public Class ClsCnnDat
#Region "Definiciones"
    ' Constantes
    Private Const MCOBJREGISTRO As Object = "A0b1f9*hjBó^23ö~"
    ' Variables
    Private MstrNombreServidor As String = String.Empty
    Private MstrPuerto As String = String.Empty
    Private MblnCnnExcel As Boolean = False
    Private MstrTrayectoriaBD As String = String.Empty
    Private MenuProveedorBD As enuProveedorBD = enuProveedorBD.None
    Private MenuTipoAuten As enuTipoAutenticacion = enuTipoAutenticacion.None
    Private MstrUsuario As String = String.Empty
    Private MstrNombreBD As String = String.Empty
    Private MstrContraseña As String = String.Empty
    Private MblnAdoDb As Boolean = False
#End Region
#Region "Constructores"
    Public Sub New()
        If Not IsNothing(gobjPanDat) Then
            If Not gobjPanDat.blnRegistrado Then
                Throw New ModuloNoRegistradoPanException
            End If
        End If
    End Sub
#End Region
#Region "Propiedades"
    Friend ReadOnly Property StrNombreServidor As String
        Get
            Return MstrNombreServidor
        End Get
    End Property
#End Region
#Region "Metodos y funciones"
    Friend Function FcnnMySqlCon_Au() As MySqlConnection
        Dim lstrCadenaConexion As String, lblnExisteAuriga As Boolean
        Dim lcnnConexionMySql As MySqlConnection
        Dim lstrConex = "Padipa1", lstrServer As String
#If DES = 0 Then
        lstrServer = "8.8.246.41"
#Else
        lstrServer = "localhost"
#End If
        lstrCadenaConexion = "server=" & lstrServer & ";uid=ftzgrngl_soporte" & ";pwd=" &
                lstrConex & ";PORT=" & MstrPuerto & ";"
        lcnnConexionMySql = FcnnConexionMySql(lstrCadenaConexion)
        Dim lstrFiltro(3) As String
        lstrFiltro(1) = "ftzgrngl_auriga"
        Dim ldtbDatabasesBD = lcnnConexionMySql.GetSchema("DataBases", lstrFiltro)
        If ldtbDatabasesBD.Rows.Count > 0 Then
            Dim lstrFiltro1 = "database_name = '" & lstrFiltro(1) & "'"
            Dim ldrwDataBases() As DataRow = ldtbDatabasesBD.Select(lstrFiltro1)
            lblnExisteAuriga = ldrwDataBases.Count > 0
        End If
        If lblnExisteAuriga Then
            lcnnConexionMySql.Close()
            lstrCadenaConexion &= "server=" & lstrServer & ";uid=ftzgrngl_soporte" & ";pwd=" &
                    lstrConex & "; database=ftzgrngl_auriga;"
            lstrCadenaConexion &= "PORT=" & MstrPuerto & ";"
            lcnnConexionMySql = FcnnConexionMySql(lstrCadenaConexion)
        End If
        If lblnExisteAuriga Then
            lcnnConexionMySql.Close()
            Return lcnnConexionMySql
        Else
            Return Nothing
        End If
    End Function
    ''' <summary>
    ''' Establece la conexion a la base de datos de Orión y Admin Orión
    ''' </summary>
    Friend Sub SEstablezcaConexionBd_App()
        SDetermineProveedorBD()
        Select Case GenuProveedorBD
            Case EnuProveedorBD.enuMySql
                SConecteMySql_App()
            Case EnuProveedorBD.enuOracle
            Case EnuProveedorBD.enuSQLServer
        End Select
    End Sub
    ''' <summary>
    ''' Segunda conexion a la BD de la aplicación para ser utilizada en segundo plano y para 
    ''' el manejo de las copias de seguridad
    ''' </summary>
    ''' <returns></returns>
    Friend Function FcnnNewMySqlCon() As MySqlConnection
        SDetermineProveedorBD()
        If GenuProveedorBD = EnuProveedorBD.enuMySql Then
            Return FcnnMySqlCon()
        Else
            Return Nothing
        End If
    End Function
    Private Sub SDetermineProveedorBD()
        Dim lstrLinea As String = String.Empty
        Dim lstrArchBD As String = GstrTrayAppDat & StrNombreBD
        Dim lstrArchivo As String = String.Empty
        If My.Computer.FileSystem.FileExists(lstrArchBD & ".udlx") Then
            lstrArchivo = lstrArchBD & ".udlx"
        ElseIf My.Computer.FileSystem.FileExists(lstrArchBD & ".udl") Then
            lstrArchivo = lstrArchBD & ".udl"
        ElseIf GshrIdAplicacion = 803 OrElse GshrIdAplicacion = 100 OrElse
                GshrIdAplicacion = 999 Then
            GenuProveedorBD = EnuProveedorBD.enuMySql
            MstrNombreServidor = "localhost"
            MstrUsuario = "root"
            MstrPuerto = "3306"
        Else
            Throw New PanDatException("No se encontro el archivo de direccionamiento al repositorio!")
        End If
        If Not String.IsNullOrEmpty(lstrArchivo) Then
            Dim lsrPanoramaUdl As StreamReader = Nothing
            Try
                lsrPanoramaUdl = ClsPanoramaDat.FsrStreamReader(lstrArchivo)
            Catch ex As FileNotFoundException
                Throw
            Catch ex As IOException
                Throw
            Catch ex As Exception
                Throw
            End Try
            If lstrArchivo.EndsWith("udlx") Then
                lstrLinea = lsrPanoramaUdl.ReadLine
            Else
                lstrLinea = lsrPanoramaUdl.ReadLine
                lstrLinea = lsrPanoramaUdl.ReadLine
                lstrLinea = lsrPanoramaUdl.ReadLine
            End If
            lsrPanoramaUdl.Close()
            SLeaParametrosCon(lstrLinea)
        End If
    End Sub
    Private Sub SLeaParametrosCon(astrLinea As String)
        Dim lstrLinea As String = astrLinea
        Dim lstrNombreBd = String.Empty
        Dim lstrArg() As String
        lstrArg = lstrLinea.Split(";")
        For i = 0 To lstrArg.GetUpperBound(0)
            Select Case lstrArg(i).Split("=")(0)
                Case "Provider"
                    Select Case lstrArg(i).Split("=")(1)
                        Case "SQLOLEDB.1"
                            GenuProveedorBD = EnuProveedorBD.enuOleDb
                        Case "OraOLEDB.Oracle.1"
                            GenuProveedorBD = EnuProveedorBD.enuOracle
                        Case "MySql"
                            GenuProveedorBD = EnuProveedorBD.enuMySql
                        Case "SqlServer"
                            GenuProveedorBD = EnuProveedorBD.enuSQLServer
                        Case "Oracle"
                            GenuProveedorBD = EnuProveedorBD.enuOracle
                    End Select
                Case "Data Source"
                    MstrNombreServidor = lstrArg(i).Split("=")(1)
                Case "Initial Catalog"
                    lstrNombreBd = lstrArg(i).Split("=")(1)
                Case "Integrated Security"
                    If GenuProveedorBD = EnuProveedorBD.enuSQLServer Then
                        If lstrArg(i).Split("=")(i) = "SSPI" Then
                            GenuTipoAutenticacion = EnuTipoAutenticacion.enuWindows
                        End If
                    End If
                Case "Usuario"
                    MstrUsuario = lstrArg(i).Split("=")(1)
                Case "Puerto"
                    MstrPuerto = lstrArg(i).Split("=")(1)
            End Select
        Next i
        If lstrNombreBd.ToUpper <> StrNombreBD.ToUpper Then
            Dim lstrMens = "El nombre de la Base de Datos en el Archivo de Direccinamiento" &
                    " o no existe o es incorrecto!"
            Throw New ArgumentoInvalidoPanException(lstrMens)
        End If
    End Sub
    Friend Function FcnnConexion(aenuProveedorBD As EnuProveedorBD,
            aenuTipoAuten As EnuTipoAutenticacion,
            astrNombreServidor As String, astrTrayectoriaBD As String,
            astrNombreBD As String, astrUsuario As String, astrContrasena As String,
            astrPuerto As String, ablnExcel As Boolean, ablnAdodb As Boolean) As Object
        Dim lstrMens As String = String.Empty
        Dim lcnnConexion As Object
        MblnAdoDb = ablnAdodb
        MenuProveedorBD = aenuProveedorBD
        MenuTipoAuten = aenuTipoAuten
        If aenuProveedorBD = GenuProveedorBD Then
            MstrNombreServidor = astrNombreServidor
        End If
        MstrTrayectoriaBD = astrTrayectoriaBD
        MstrNombreBD = astrNombreBD
        MstrUsuario = astrUsuario
        MstrContraseña = astrContrasena
        MstrPuerto = astrPuerto
        MblnCnnExcel = ablnExcel
        If MenuProveedorBD = EnuProveedorBD.enuSQLServer Then
            If MenuTipoAuten = EnuTipoAutenticacion.None Then
                lstrMens &= "Falta el tipo de autenticación"
            End If
        End If
        If String.IsNullOrEmpty(lstrMens) Then
            If String.IsNullOrEmpty(MstrNombreServidor) Then
                If MenuProveedorBD <> EnuProveedorBD.enuOleDb AndAlso MenuProveedorBD <>
                        EnuProveedorBD.enuExcel10 Then
                    lstrMens &= "Falta el nombre del servidor"
                End If
            End If
            If String.IsNullOrEmpty(MstrUsuario) Then
                If MenuProveedorBD <> EnuProveedorBD.enuOleDb AndAlso MenuProveedorBD <>
                        EnuProveedorBD.enuExcel10 Then
                    If MenuTipoAuten = EnuTipoAutenticacion.enuBaseDatos Then
                        If String.IsNullOrEmpty(lstrMens) Then
                            lstrMens &= "Falta el dato correspondiente al Usuario"
                        Else
                            lstrMens &= ", el dato correspondiente al Usuario"
                        End If
                    End If
                End If
            End If
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            lstrMens &= "."
            Throw New ParametrosConexionBDPanException(lstrMens)
        End If
        lcnnConexion = FcnnConexion()
        Return lcnnConexion
    End Function
    Private Function FcnnConexion() As Object
        Dim lcnnConexion As Object = Nothing
        Select Case MenuProveedorBD
            Case EnuProveedorBD.enuSQLServer
                If MenuTipoAuten = EnuTipoAutenticacion.enuWindows Then
                    lcnnConexion = FcnnConexionSqlWin()
                ElseIf MenuTipoAuten = EnuTipoAutenticacion.enuBaseDatos Then
                    lcnnConexion = FcnnConexionSqlUsu()
                End If
            Case EnuProveedorBD.enuOracle
                lcnnConexion = FcnnConexionOdbc()
            Case EnuProveedorBD.enuMySql
                lcnnConexion = FcnnConexionMySql()
            Case EnuProveedorBD.enuOleDb
                lcnnConexion = FcnnConexionOleDb()
            Case EnuProveedorBD.enuExcel10
                lcnnConexion = FcnnConexionExcel10()
        End Select
        Return lcnnConexion
    End Function
    Private Function FcnnConexionSqlWin() As SqlConnection
        Dim lcnnConexionSql = New SqlConnection
        Dim lstrCadenaConexion As String = String.Empty
        lstrCadenaConexion &= "Data Source=" & MstrNombreServidor & ";"
        lstrCadenaConexion &= "Initial Catalog=" & StrNombreBD & ";"
        lstrCadenaConexion &= "Trusted_Connection=yes;"
        Try
            lcnnConexionSql.ConnectionString = lstrCadenaConexion
            lcnnConexionSql.Open()
        Catch ex As SqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As Exception
            Throw
        Finally
            lcnnConexionSql.Close()
        End Try
        Return lcnnConexionSql
    End Function
    Private Function FcnnConexionSqlUsu() As SqlConnection
        Dim lcnnConexion As New SqlConnection
        Dim lstrCadenaConexion As String
        lstrCadenaConexion = "Data Source=" & MstrNombreServidor & ";"
        lstrCadenaConexion &= "Initial Catalog=" & StrNombreBD & ";"
        lstrCadenaConexion &= "Password=" & """" & MstrContraseña & """" & ";"
        lstrCadenaConexion &= "User ID=" & MstrUsuario & ";"
        Try
            lcnnConexion.ConnectionString = lstrCadenaConexion
            lcnnConexion.Open()
        Catch ex As SqlException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As Exception
            Throw
        Finally
            lcnnConexion.Close()
        End Try
        Return lcnnConexion
    End Function
    Private Function FcnnConexionExcel10() As OleDbConnection 'Microsoft Excel 2010
        Dim lcnnConexion = New OleDbConnection
        If Not MstrTrayectoriaBD.EndsWith("\") Then
            MstrTrayectoriaBD &= ("\")
        End If
        Dim lstrBaseDatos As String = MstrTrayectoriaBD & MstrNombreBD
        Dim lstrCadenaConexion As String = "Provider=Microsoft.ACE.OLEDB.12.0;"
        lstrCadenaConexion &= "Data Source=" & lstrBaseDatos & ";"
        lstrCadenaConexion &= "Extended Properties = ""Excel 12.0 Xml; HDR=YES"""
        If Not String.IsNullOrEmpty(MstrContraseña) Then
            lstrCadenaConexion &= "Persist Security Info=False;"
            lstrCadenaConexion &= "Microsoft.ACE.OLEDB:Database Password=" & MstrContraseña
        End If
        lcnnConexion.ConnectionString = lstrCadenaConexion
        Try
            lcnnConexion.Open()
        Catch ex As OleDbException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As Exception
            Throw
        Finally
            lcnnConexion.Close()
        End Try
        Return lcnnConexion
    End Function
    'Microsoft.Jet.OLEDB.4.0 Access o Excel
    Private Function FcnnConexionOleDb() As Object
        Dim lstrCadenaConexion As String = String.Empty, lblnNoHayError = False
        If Not MstrTrayectoriaBD.EndsWith("\") Then
            MstrTrayectoriaBD &= ("\")
        End If
        Dim lstrBaseDatos As String = MstrTrayectoriaBD & MstrNombreBD
        Dim lcnnConexion As Object = Nothing
        If MblnAdoDb Then
            lcnnConexion = New ADODB.Connection
        Else
            lcnnConexion = New OleDbConnection
        End If
#If BIT64 = 1 Then
        lstrCadenaConexion = "Provider=Microsoft.ACE.OLEDB.12.0;"
        lstrCadenaConexion &= "Data Source=" & lstrBaseDatos & ";"
        If MblnCnnExcel Then
            lstrCadenaConexion &= "Extended Properties = ""Excel 14.0; HDR=YES"""
        End If
#Else
        lstrCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0;"
        lstrCadenaConexion &= "Data Source=" & lstrBaseDatos & ";"
        If mblnCnnExcel Then
            lstrCadenaConexion &= "Extended Properties = ""Excel 8.0; HDR=YES"""
        End If
#End If
        If Not String.IsNullOrEmpty(MstrContraseña) Then
            lstrCadenaConexion &= "Persist Security Info=False;"
            lstrCadenaConexion &= "Jet OLEDB:Database Password=" & MstrContraseña
        End If
        Try
            lcnnConexion.ConnectionString = lstrCadenaConexion
            lcnnConexion.Open()
            lblnNoHayError = True
        Catch ex As OleDbException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                lcnnConexion.Close()
            End If
        End Try
        Return lcnnConexion
    End Function
    Private Shared Function FcnnConexionOdbc() As OdbcConnection
        Dim lstrCadenaConexion As String = String.Empty
        Dim lcnnConexion = New OdbcConnection
        'Aurelio: Falta construir la cade de conexion
        Try
            lcnnConexion.ConnectionString = lstrCadenaConexion

        Catch ex As OdbcException
            Throw New ProveedorBdPanException(ex.Message)
        Catch ex As Exception
            Throw
        Finally
            lcnnConexion.Close()
        End Try
        Return lcnnConexion
    End Function
    Private Function FcnnConexionMySql() As MySqlConnection
        Dim lstrCadenaConexion As String = "server=" & MstrNombreServidor.Trim & ";uid=" &
                MstrUsuario.Trim & ";pwd=" & MstrContraseña & ";database=" & MstrNombreBD.Trim & ";"
        If Not String.IsNullOrEmpty(MstrPuerto) Then
            lstrCadenaConexion &= "PORT=" & MstrPuerto & ";"
        End If
        Dim lcnnConexionMySql = FcnnConexionMySql(lstrCadenaConexion)
        lcnnConexionMySql.Close()
        Return lcnnConexionMySql
    End Function
    Private Function FcnnMySqlCon() As MySqlConnection
        Dim lstrCadenaConexion As String
        Dim lblnExisteBdMySql As Boolean = False
        Dim lcnnConexionMySql As MySqlConnection
        Dim lstrConex = "padipa"
        lstrCadenaConexion = "server=" & MstrNombreServidor.Trim & ";uid=" & MstrUsuario.Trim & ";pwd=" &
                lstrConex & ";"
        If Not String.IsNullOrEmpty(MstrPuerto) Then
            lstrCadenaConexion &= "PORT=" & MstrPuerto & ";"
        End If
        lcnnConexionMySql = FcnnConexionMySql(lstrCadenaConexion)
        Dim lstrFiltro(3) As String
        lstrFiltro(1) = StrNombreBD.ToLower
        Dim ldtbDatabasesBD = lcnnConexionMySql.GetSchema("DataBases", lstrFiltro)
        If ldtbDatabasesBD.Rows.Count > 0 Then
            Dim lstrFiltro1 = "database_name = '" & lstrFiltro(1) & "'"
            Dim ldrwDataBases() As DataRow = ldtbDatabasesBD.Select(lstrFiltro1)
            lblnExisteBdMySql = (ldrwDataBases.Count > 0)
        End If
        If lblnExisteBdMySql Then
            lcnnConexionMySql.Close()
            lstrCadenaConexion = "server=" & MstrNombreServidor.Trim & ";uid=" & MstrUsuario.Trim &
                    ";pwd=" & lstrConex & ";database=" & StrNombreBD.Trim & ";"
            If Not String.IsNullOrEmpty(MstrPuerto) Then
                lstrCadenaConexion &= "PORT=" & MstrPuerto & ";"
            End If
            lcnnConexionMySql = FcnnConexionMySql(lstrCadenaConexion)
        End If
        lcnnConexionMySql.Close()
        If lblnExisteBdMySql Then
            Return lcnnConexionMySql
        Else
            Return Nothing
        End If
    End Function
    Private Sub SConecteMySql_App()
        Dim lstrCadenaConexion As String
        Dim lblnExisteBdMySql As Boolean = False
        Dim lcnnConexionMySql As MySqlConnection
        Dim lstrConex = "padipa"
        lstrCadenaConexion = "server=" & MstrNombreServidor.Trim & ";uid=" & MstrUsuario.Trim & ";pwd=" &
                lstrConex & ";"
        If Not String.IsNullOrEmpty(MstrPuerto) Then
            lstrCadenaConexion &= "PORT=" & MstrPuerto & ";"
        End If
        lcnnConexionMySql = FcnnConexionMySql(lstrCadenaConexion)
        Dim lstrFiltro(3) As String
        lstrFiltro(1) = StrNombreBD.ToLower
        Dim ldtbDatabasesBD = lcnnConexionMySql.GetSchema("DataBases", lstrFiltro)
        If ldtbDatabasesBD.Rows.Count > 0 Then
            Dim lstrFiltro1 = "database_name = '" & lstrFiltro(1) & "'"
            Dim ldrwDataBases() As DataRow = ldtbDatabasesBD.Select(lstrFiltro1)
            lblnExisteBdMySql = (ldrwDataBases.Count > 0)
        End If
        If lblnExisteBdMySql Then
            lcnnConexionMySql.Close()
            lstrCadenaConexion = "server=" & MstrNombreServidor.Trim & ";uid=" & MstrUsuario.Trim &
                    ";pwd=" & lstrConex & ";database=" & StrNombreBD.Trim & ";"
            If Not String.IsNullOrEmpty(MstrPuerto) Then
                lstrCadenaConexion &= "PORT=" & MstrPuerto & ";"
            End If
            lcnnConexionMySql = FcnnConexionMySql(lstrCadenaConexion)
        End If
        lcnnConexionMySql.Close()
        GobjPanDat.BlnExisteBdPanorama = lblnExisteBdMySql
        GobjPanDat.CnnConexionBd_App = lcnnConexionMySql
    End Sub
    Private Shared Function FcnnConexionMySql(astrCadenaConexion As String) As MySqlConnection
        Dim lcnnConexionMySql As New MySqlConnection
        Dim lstrCadenaConexion As String = astrCadenaConexion
        Dim lblnNoHayError = False
        Try
            lcnnConexionMySql.ConnectionString = lstrCadenaConexion
            lcnnConexionMySql.Open()
            lblnNoHayError = True
        Catch ex As MySqlException
            Throw New ConexionBdPanException(ex.Message)
        Catch ex As Exception
            Throw
        Finally
            If Not lblnNoHayError Then
                lcnnConexionMySql.Dispose()
            End If
        End Try
        Return lcnnConexionMySql
    End Function
#End Region
End Class
