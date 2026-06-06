Imports Microsoft.Win32
Friend Module MDefPubPanL
#Region "Definiciones globales de solución"
#Region "Variables globales de Proyecto"
    Friend GblnImportando As Boolean = False
    Friend GblnActualizandoApp As Boolean = False
    Friend GblnPosteando As Boolean = False
    Friend GblnEnviandoEmail As Boolean = False
    Friend GblnNotiSonoras As Boolean = False
    Friend GblnOK As Boolean = True
    Friend GstrVerAntApp As String = String.Empty
    '
    Friend GshrIdCarpeta As Short = 0
    Friend GshrIdCentroUtil As Short = 0
    Friend GstrIdUsuario As String = String.Empty
    Friend GstrOrigenActual As String = String.Empty
    Friend GenuTipoNomTer As EnuTipoNomTercero = EnuTipoNomTercero.None
    Public Property GenuTamanoIcono As EnuTamanoIconos = EnuTamanoIconos.EnuMediano
    Public Property GblnMostrandoTitulos As Boolean = False
    '
    Friend GobjAdministrador As ClsAdministrador = Nothing
    Friend GobjPanorama As ClsPanorama = Nothing
    Friend StrCampoCarpeta As String = ClsIdCarpetaShr.SstrNombreCampoBd
    Friend StrCampoCentroUtil As String = ClsIdCentroUtilShr.SstrNombreCampoBd
#End Region
#Region "Constantes globales"
    ' Constante Golbales Panorama
    Friend Const GCBYTMINDPTO As Byte = 1 'Departamento
    Friend Const GCBYTMAXDPTO As Byte = 99 'Departamento
    Friend Const GCSHRMINCIUD As Short = 1 'Ciudad
    Friend Const GCSHRMAXCIUD As Short = 999 'Ciudad
    Friend Const GCDBLMINTERC As Double = 1 'Tercero
    Friend Const GCDBLMAXTERC As Double = Double.MaxValue 'Tercero
    Friend Const GCSTRADMIN As String = "Admin"
    ' Constantes del ensamblado
    Friend Const GCOBJREGISTRO As Object = "A0b1f9*hjBó^23ö~"
#End Region
#Region "Enumeradores"
    Public Enum EnuTamanoIconos As Integer
        EnuPequeño = 0
        EnuMediano
        EnuGrande
    End Enum
#End Region
#Region "Métodos"
    Friend Function FblnEstaAdminInstalado() As Boolean
        Dim lblnEstaIns As Boolean
        Dim lobjLine = Nothing
        Dim lstrSoftwareInstallPath = String.Empty
        Dim lstrRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        Using baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,
                Microsoft.Win32.RegistryView.Registry64)
            Using key = baseKey.OpenSubKey(lstrRegistryKey)
                For Each subkey_name As String In key.GetSubKeyNames()
                    Using subKey = key.OpenSubKey(subkey_name)
                        lobjLine = subKey.GetValue("DisplayName")
                        If Not IsNothing(lobjLine) Then
                            If lobjLine.ToString().ToUpper().Contains("ADMINIU") Then
                                lstrSoftwareInstallPath = subKey.GetValue("InstallLocation").ToString()
                                Exit For
                            End If
                        End If
                    End Using
                Next
            End Using
        End Using
        lblnEstaIns = Not String.IsNullOrEmpty(lstrSoftwareInstallPath)
        Return lblnEstaIns
    End Function
    Friend Function FstrNombreDoc(aenuIdClase As EnuIdClasesPanDef)
        Dim lstrNomDoc = String.Empty
        Select Case aenuIdClase
            Case EnuIdClasesPanDef.enuFactura
                lstrNomDoc = "La Factura "
            Case EnuIdClasesPanDef.enuNotaCr
                lstrNomDoc = "La Nota Crédito "
            Case EnuIdClasesPanDef.enuNotaDevAnt
                lstrNomDoc = "La Nota Reintegro Anticipo "
            Case EnuIdClasesPanDef.enuNotaReversaCr
                lstrNomDoc = "La Nota Reversión Crédito "
            Case EnuIdClasesPanDef.enuReciboCaja
                lstrNomDoc = "El Recibo de Caja "
            Case EnuIdClasesPanDef.enuEstadoCuenta
                lstrNomDoc = "La Cuenta de Cobro"
        End Select
        Return lstrNomDoc
    End Function
    Friend Function FstrNombreTercero(astrParteNom As String)
        Dim lstrNomTer As String = String.Empty
        If Not String.IsNullOrEmpty(astrParteNom) Then
            Dim i = 0
            Dim lstrPartesNom As String() = astrParteNom.Trim.Split(" ")
            Dim lenuTipoNom As EnuTipoNomTercero = GenuTipoNomTer
            For Each lstrParte As String In lstrPartesNom
                i += 1
                lstrParte = lstrParte.Trim
                Select Case lenuTipoNom
                    Case EnuTipoNomTercero.EnuMaySostenida
                        lstrParte = lstrParte.ToUpper
                    Case EnuTipoNomTercero.EnuMayInicial
                        lstrParte = FstrCambieInicialAMAyu(lstrParte)
                    Case Else
                        lstrParte = lstrParte
                End Select
                lstrNomTer &= lstrParte
                If Not i = lstrPartesNom.Length Then
                    lstrNomTer &= " "
                End If
            Next
        End If
        Return lstrNomTer
    End Function
    Private Function FstrCambieInicialAMAyu(astrCadena As String)
        Dim lstrPrimCadena = astrCadena.Substring(0, 1).ToUpper
        Dim lstrCadena = lstrPrimCadena & astrCadena.Substring(1).ToLower
        Return lstrCadena
    End Function
    Friend Sub SLeaArchivoAdminIni(ablnAdmin As Boolean)
        Dim lsrArchivoIni As StreamReader
        Dim lstrLinea As String
        Dim lstrArg() As String
        Dim lstrArchINI As String = GstrTrayDatPrg & "Panorama.ini"
        lsrArchivoIni = ClsPanorama.FsrStreamReader(lstrArchINI)
        If IsNothing(lsrArchivoIni) Then Exit Sub
        lstrLinea = lsrArchivoIni.ReadLine
        Do While Not IsNothing(lstrLinea)
            lstrArg = lstrLinea.Split("=")
            Select Case lstrArg(0)
                Case "TamanoIcono"
                    Select Case lstrArg(1)
                        Case "L"
                            GenuTamanoIcono = EnuTamanoIconos.EnuGrande
                        Case "M"
                            GenuTamanoIcono = EnuTamanoIconos.EnuMediano
                        Case "S"
                            GenuTamanoIcono = EnuTamanoIconos.EnuPequeño
                    End Select
                Case "TitulosMenu"
                    If lstrArg(1) = "S" Then
                        GblnMostrandoTitulos = True
                    End If
                Case "Usuario"
                    If ablnAdmin Then
                        GstrIdUsuario = lstrArg(1)
                    End If
                Case "TipoNomTer"
                    Select Case lstrArg(1)
                        Case "1"
                            GenuTipoNomTer = EnuTipoNomTercero.EnuMaySostenida
                        Case "2"
                            GenuTipoNomTer = EnuTipoNomTercero.EnuMayInicial
                        Case "3"
                            GenuTipoNomTer = EnuTipoNomTercero.EnuComoIngreso
                    End Select
            End Select
            lstrLinea = lsrArchivoIni.ReadLine
        Loop
        lsrArchivoIni.Close()
    End Sub
    ''' <summary>
    ''' Indica si una tabla de origen de datos para importar servicios corresponde a un
    ''' servicio de consumo
    ''' </summary>
    ''' <param name="adtbOrigen"></param>
    ''' <returns></returns>
    Friend Function FblnEsConsumo(adtbOrigen As DataTable) As Boolean
        Const lstrLectAct = "lectura_actual"
        Const lstrLectAnt = "lectura_anterior"
        Const lstrVlrUni = "valor_unitario"
        Dim lblnEsConsumo = adtbOrigen.Columns.Contains(lstrLectAct) AndAlso
                adtbOrigen.Columns.Contains(lstrLectAnt) AndAlso
                adtbOrigen.Columns.Contains("valor_unitario")
        Dim ldecLectAct As Decimal, ldecLectAnt As Decimal, ldecValorUni As Decimal
        If lblnEsConsumo Then
            For Each ldrwOrigen As DataRow In adtbOrigen.Rows
                ldecLectAct = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrLectAct),
                        EnuTipoValor.EnuDecimal)
                ldecLectAnt = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrLectAnt),
                        EnuTipoValor.EnuDecimal)
                ldecValorUni = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrVlrUni),
                        EnuTipoValor.EnuDecimal)
                lblnEsConsumo = ldecLectAct > 0 AndAlso ldecLectAnt > 0 AndAlso ldecValorUni > 0
                If lblnEsConsumo Then Exit For
            Next
        End If
        Return lblnEsConsumo
    End Function
    Friend Function FstrColumnaOrigen(astrColumnasRelacionadas As String(),
            astrCampoDestino As String) As String
        Dim lstrCampOrig = String.Empty
        For Each lstrCampRela As String In astrColumnasRelacionadas
            If lstrCampRela.Split("=")(0).ToLower = astrCampoDestino.ToLower Then
                lstrCampOrig = lstrCampRela.Split("=")(1)
                Exit For
            End If
        Next
        Return lstrCampOrig
    End Function
    Friend Function FstrColumnaDestino(astrColumnasRelacionadas As String(),
            astrCampoOrigen As String) As String
        Dim lstrCampDes = String.Empty
        For Each lstrCampRela As String In astrColumnasRelacionadas
            If lstrCampRela.Split("=")(1).ToLower = astrCampoOrigen.ToLower Then
                lstrCampDes = lstrCampRela.Split("=")(0)
                Exit For
            End If
        Next
        Return lstrCampDes
    End Function
#End Region
#End Region
End Module