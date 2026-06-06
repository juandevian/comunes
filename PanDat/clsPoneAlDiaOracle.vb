Imports System.Data.OracleClient
Imports TyS.Panorama.GesDat.clsEstructuraBD
Friend Class clsPoneAlDiaOracle
    Inherits clsCBPoneAlDiaBD
    ' Variables
    Private mobjPadre As clsActualizaBD = Nothing

    Public Sub New(ByVal aobjPadre As clsActualizaBD,ByVal astrArchivoXml As String)
        If gobjPanDat.blnRegistrado Then
            mobjPadre = aobjPadre
            hobjEstructuraBD = New clsEstructuraBD
            sActualiceBaseDatos(astrArchivoXml)
        Else
            Throw New System.InvalidOperationException("El módulo no esta registrado")
        End If
    End Sub

    '    Friend Sub sejemplo()
    'Dim lstrCS As String
    'Dim lstrsql As String = "CREATE TABLE Item (Numero number, Nombre varchar2(20))"
    'lstrsql = "DROP TABLE Corte"
    'lstrCS = "Data Source=(DESCRIPTION=" _
    '   & "(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=OFICINA)(PORT=1521)))" _
    '   & "(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=ORCL)));" _
    '   & "User Id=Aurelio;Password=auvivi*6259"
    'lstrCS = "Data Source=(DESCRIPTION=" _
    '   & "(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=OFICINA)(PORT=1521))));" _
    '   & "User Id=Aurelio;Password=auvivi*6259"
    'genuProveedorBD = enuPBDProveedorBDDef.enuPBDOracle
    'Dim lcnn As New OracleConnection()
    'lcnn.ConnectionString = lstrCS
    'lcnn.Open()
    'sEjecuteSentenciaSql(lstrsql)
    'lcnn.Close()
    'lcnn.Open()
    'lcnn.Dispose()

    'lstrCS = "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=HR;Password=auvivi*6259;Data Source=Oficina"
    'lstrCS = "Provider=MSDAORA.1;User ID=Aurelio;Password=auvivi*6259;Data Source=Oficina;Persist Security Info=False"
    'lcnnAdodb.ActiveConnection.Execute(lstrsql, , CommandTypeEnum.adCmdText)
    'End Sub
#Region "Procedimientos sobreescritos"
    Protected Overrides Sub sGenereEstructuraActualBD()
        Stop
    End Sub

    Protected Overrides Sub sCreeBaseDatos()
        Stop
    End Sub

    Protected Overrides Sub sActualiceBD()

    End Sub

    Protected Overrides Sub sVerifiqueCollationTabla(ByVal astrCoolationNameOriginalBD As String)

    End Sub

    Protected Overrides Sub sVerifiqueCollationColumna(ByVal aobjColumnaXml As clsEstructuraBD.clsColumna,
            ByRef ablnCambiarColumna As Boolean)
        '
    End Sub

    Protected Overrides Sub sCreeTabla()
        '
    End Sub

    Protected Overrides Sub sRenombreTabla(ByVal astrNombreTablaOri As String,
            ByVal astrNombreTablaDes As String)
        '
    End Sub

    Protected Overrides Sub sVinculeTablas(ByVal astrNombreTabla As String,
        ByVal astrBaseDatosExterna As String)

    End Sub

    Protected Overrides Sub sVersioneBD(ByVal ablnActualiceTamano As Boolean)
        Stop
    End Sub

    Protected Overrides Sub sCompareIndices(ByVal aobjIndiceXml As clsEstructuraBD.clsIndice,
        ByRef ablnDifieren As Boolean)
        '
    End Sub

    Protected Overrides Sub sElimineIndice(ByVal astrNombreTabla As String, ByVal astrNombreIndice As String)
        Stop
    End Sub

    Protected Overrides Sub sCreeIndice(ByVal aobjIndiceXml As clsIndice)
        Stop
    End Sub
#End Region
End Class