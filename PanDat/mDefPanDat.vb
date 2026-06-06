Module mDefPanDat
#Region "Definiciones globales de proyecto"
#Region "Constantes"
    Friend Const GCSTRFMTIDTERCERO As String = "###,###,###,##0"
    Friend Const GCDBLTERCERONULO As Double = 999999999999
    ' Fechas
    Friend Const GCDTMFECHANULA As Date = #1/1/1900#
    Friend Const GCDTMTIMENULA As DateTime = #1/1/1900#
    Friend Const GCDTMFECHAMAXI As Date = #12/31/3000#
    Friend Const GCSTRFMTFECHASIMPLE = "dd/MM/yyyy"
    ' Periodo nulo
    Friend Const GCSTRPERIODONULO As String = "000000"
    ' Usuario universal
    Friend Const GCSTRUSUARIOU As String = "OPT"
#End Region
#Region "Variables"
    Friend GenuIdAplicacion As EnuListaAplicaciones = EnuListaAplicaciones.None
    Public GenuTipoInstanciamiento As EnuTipoInstanciamiento = EnuTipoInstanciamiento.enuNormal
    Friend WithEvents GobjPanDat As ClsPanoramaDat = Nothing
    Friend GentVerBDEnProg As Integer = 0
    Friend GshrIdAplicacion As Short = 0
    Friend GenuProveedorBD As EnuProveedorBD = EnuProveedorBD.enuMySql
    Friend GenuTipoAutenticacion As EnuTipoAutenticacion = EnuTipoAutenticacion.enuBaseDatos
    Friend GstrTrayDatos As String = String.Empty
    Friend GstrTrayReportes As String = String.Empty
    Friend GstrTrayEmails As String = String.Empty
    Friend GstrTrayEFac As String = String.Empty
    Friend GstrTrayInterfContable As String = String.Empty
    Friend GstrTrayFacturasPdf As String = String.Empty
    Friend GstrTrayRecibosCajaPdf As String = String.Empty
    Private ReadOnly MstrUniIns As String =
            My.Computer.FileSystem.SpecialDirectories.ProgramFiles.Substring(0, 3)
    Friend GstrTraySegApp As String = MstrUniIns & "Users\Default\AppData\Local\Sumitpo\"
    Friend GstrTrayAppDat As String = MstrUniIns & "ProgramData\OPTIMUSOFT\"
    Friend GstrTrayDat As String = MstrUniIns & "Panorama.net\Dat\"
    Friend GstrTrayDatPrg As String = MstrUniIns & "Panorama.Net\DatPrg\"
#End Region
#End Region
End Module