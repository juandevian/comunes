Friend Class ClsImportar
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = ""
    ' Variables
    Public Event EvnImportadoObj As EventHandler(Of ClsPanEventArgs)
    Public Event EvnFinImportar As EventHandler(Of ClsPanEventArgs)
    Private ReadOnly MobjArgumentoEventoPan As ClsPanEventArgs = Nothing
    '
    Private MdtbTablaDatosOrigen As DataTable = Nothing
    Private MstrColumnasOrigen As String() = {}
    Private MstrColumnasRequeridas As String() = Nothing
    Private MstrColumnasRelacionadas As String() = {}
    Private MstrUltimaImportacion As String = String.Empty
    Private MblnHayReporte As Boolean = False
    Private MshrNumeroregistro As Short = 0
    Private MshrExportados As Short = 0
    Private MshrNoExportados As Short = 0
    Private MstrTiempoImpor As String = String.Empty
    Private MblnEstaVacioOrigenDatos As Boolean = True
    Private MblnHayDatosImportacion As Boolean = False
    Private ReadOnly MobjObjetoDestino As ClsCBObjetoPan = Nothing
    Private MblnDatosOrigenOk As Boolean = False
#End Region

#Region "Constructores"
    Public Sub New(aobjObjetoDestino As Object)
        HobjPadre = Nothing
        HblnEsAnulable = False
        EnuPermisosObj = EnuPermisosDef.EnuTodos
        HenuTipoObjeto = EnuModoInstanciaObjDef.EnuUnico
        MobjObjetoDestino = aobjObjetoDestino
        SCargueDatos()
    End Sub
#End Region

#Region "Propiedades"
#Region "Propiedades indentificadoras"
    Protected Overrides ReadOnly Property HstrNombreTabla As String
        Get
            Return MCSTRNOMBRETABLA
        End Get
    End Property
    Protected Friend Overrides ReadOnly Property HenuIdClase As EnuIdClasesPanDef
        Get
            Return EnuIdClasesPanDef.EnuImportar
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Importar"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjArgumentoEventoPan As New ClsPanEventArgs
    Friend ReadOnly Property ObjArchivoOrigenStr As New ClsArchivoOrigenStr(Me)
    Friend ReadOnly Property ObjTablaOrigenStr As New ClsTablaOrigenStr(Me)
    Friend ReadOnly Property ObjExigeRequeridosBln As New ClsExigeRequeridosBln(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjArchivoOrigenStr)
                HcolPropiedades.Add(ObjTablaOrigenStr)
                HcolPropiedades.Add(ObjExigeRequeridosBln)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras Propiedades"
    Friend Function FstrColumnasObjetoDes() As String()
        Dim lstrColumnasDes As String() = {}
        Dim i As Short = -1, j As Short = -1
        MstrColumnasRequeridas = Nothing
        For Each lobjPropiedad As ClsCBPropiedad In MobjObjetoDestino.ColPropiedades
            If Not String.IsNullOrEmpty(lobjPropiedad.StrNombreCampoBD) Then
                i += 1
                ReDim Preserve lstrColumnasDes(i)
                lstrColumnasDes(i) = lobjPropiedad.StrNombreCampoBD
                If lobjPropiedad.BlnEsRequerido Then
                    j += 1
                    ReDim Preserve MstrColumnasRequeridas(j)
                    MstrColumnasRequeridas(j) = lobjPropiedad.StrNombreCampoBD
                End If
            End If
        Next
        Return lstrColumnasDes
    End Function
    Friend ReadOnly Property StrColumnasRequeridas As String()
        Get
            Return MstrColumnasRequeridas
        End Get
    End Property
    Friend ReadOnly Property ObjObjetoDestino As ClsCBObjetoPan
        Get
            Return MobjObjetoDestino
        End Get
    End Property
    Friend ReadOnly Property EntCantidadObjetos As Integer
        Get
            If Not IsNothing(DtbTablaDatosOrigen) Then
                Return DtbTablaDatosOrigen.Rows.Count
            Else
                Return 0
            End If
        End Get
    End Property
    Friend ReadOnly Property BlnHayDatosImportacion As Boolean
        Get
            Return MblnHayDatosImportacion
        End Get
    End Property
    Friend ReadOnly Property StrNombreArchivoResultados As String
        Get
            Dim lstrArchivo As String = GstrTrayDatPrg &
                    "Imp" & MobjObjetoDestino.StrNombreClase & ".txt"
            Return lstrArchivo
        End Get
    End Property
    Friend Property StrColumnasRelacionadas() As String()
        Get
            If MstrColumnasRelacionadas.Length = 0 OrElse MstrColumnasRelacionadas Is Nothing Then
                MstrColumnasRelacionadas = FstrColumnasRelacionadas()
            End If
            Return MstrColumnasRelacionadas
        End Get
        Set(value As String())
            MstrColumnasRelacionadas = value
        End Set
    End Property
    Friend Property DtbTablaDatosOrigen As DataTable
        Get
            If MdtbTablaDatosOrigen Is Nothing Then
                MdtbTablaDatosOrigen = FdtbTablaDatosOrigen()
            End If
            Return MdtbTablaDatosOrigen
        End Get
        Set(value As DataTable)
            MdtbTablaDatosOrigen = value
        End Set
    End Property
#End Region
#End Region

#Region "Procedimientos y funciones invalidantes"
    Protected Friend Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        SEscribaArchivoPli()
        EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando
    End Sub
    Protected Friend Overrides Sub SVacie()
        MyBase.SVacie()
        MdtbTablaDatosOrigen = Nothing
        MstrColumnasRequeridas = Nothing
        MstrColumnasRelacionadas = {}
    End Sub
#End Region

#Region "Procedimientos del objeto"
    Friend Overrides Function FblnEstaVacioOrigenDatos() As Boolean
        Return MblnEstaVacioOrigenDatos
    End Function

    Private Sub SCargueDatos()
        Dim lstrArchivoPli As String = GstrTrayDatPrg &
                    "Imp" & MobjObjetoDestino.StrNombreClase & ".pli"
        MstrUltimaImportacion = lstrArchivoPli
        If My.Computer.FileSystem.FileExists(lstrArchivoPli) Then
            SCargueDatosArchivo()
            MblnHayDatosImportacion = True
        Else
            SVaciePropiedades()
        End If
    End Sub

    Private Sub SEscribaArchivoPli()
        Dim lswArchivoPli As StreamWriter
        Dim lstrLinea As String
        Dim lstrArchivoPli As String
        lstrArchivoPli = GstrTrayDatPrg & "Imp" & MobjObjetoDestino.StrNombreClase & ".pli"
        lswArchivoPli = ClsPanorama.FswStreamWriter(lstrArchivoPli)
        lstrLinea = "[Encabezado]"
        lswArchivoPli.WriteLine(lstrLinea)
        lstrLinea = "ArchivoOrigen=" & ObjArchivoOrigenStr.ObjValorPro
        lswArchivoPli.WriteLine(lstrLinea)
        lstrLinea = "TablaOrigen=" & ObjTablaOrigenStr.ObjValorPro
        lswArchivoPli.WriteLine(lstrLinea)
        lstrLinea = "ExigeRequeridos=" & ObjExigeRequeridosBln.ObjValorPro.ToString
        lswArchivoPli.WriteLine(lstrLinea)
        lstrLinea = "[Relacionados]"
        lswArchivoPli.WriteLine(lstrLinea)
        For Each lstrRelacion As String In StrColumnasRelacionadas()
            lswArchivoPli.WriteLine(lstrRelacion)
        Next
        lswArchivoPli.Close()
    End Sub

    Private Sub SCargueDatosArchivo()
        Dim lsrArchivoPli As StreamReader
        Dim lstrLinea As String
        Dim lstrPropiedades() As String
        Dim lblnArchBienConformado As Boolean = True
        lsrArchivoPli = ClsPanorama.FsrStreamReader(MstrUltimaImportacion)
        lstrLinea = lsrArchivoPli.ReadLine
        If lstrLinea <> "[Encabezado]" Then
            lblnArchBienConformado = False
        End If
        If lblnArchBienConformado Then
            lstrLinea = lsrArchivoPli.ReadLine
            Do While Not lstrLinea.StartsWith("[")
                lstrPropiedades = lstrLinea.Split("=")
                Select Case lstrPropiedades(0)
                    Case "ArchivoOrigen"
                        ObjArchivoOrigenStr.BlnLeyendoOrigen = True
                        ObjArchivoOrigenStr.ObjValorPro = lstrPropiedades(1)
                    Case "TablaOrigen"
                        ObjTablaOrigenStr.BlnLeyendoOrigen = True
                        ObjTablaOrigenStr.ObjValorPro = lstrPropiedades(1)
                    Case "ExigeRequeridos"
                        ObjExigeRequeridosBln.BlnLeyendoOrigen = True
                        ObjExigeRequeridosBln.ObjValorPro = (lstrPropiedades(1).ToUpper = "TRUE")
                    Case Else
                        lblnArchBienConformado = False
                        Exit Do
                End Select
                lstrLinea = lsrArchivoPli.ReadLine
            Loop
            If lblnArchBienConformado Then
                If lstrLinea <> "[Relacionados]" Then
                    lblnArchBienConformado = False
                End If
            End If
        End If
        MstrColumnasRelacionadas = {}
        If lblnArchBienConformado AndAlso ObjArchivoOrigenStr.BlnEsValido AndAlso
                ObjTablaOrigenStr.BlnEsValido Then
            lstrLinea = lsrArchivoPli.ReadLine
            Dim i As Short = -1
            Do While Not IsNothing(lstrLinea)
                i += 1
                ReDim Preserve MstrColumnasRelacionadas(i)
                MstrColumnasRelacionadas(i) = lstrLinea
                lstrLinea = lsrArchivoPli.ReadLine
            Loop
            MblnEstaVacioOrigenDatos = False
        Else
            SVaciePropiedades()
        End If
        lsrArchivoPli.Close()
        If Not lblnArchBienConformado Then
            Dim lstrMens = "La plantilla de importación no esta bien conformada!"
            SLevanteEventoNot(lstrMens, "", 0, EnuSeveridadNot.EnuError)
            SVaciePropiedades()
        End If
    End Sub

    Private Sub SVaciePropiedades()
        ObjArchivoOrigenStr.BlnLeyendoOrigen = True
        ObjArchivoOrigenStr.ObjValorPro = Nothing
        ObjTablaOrigenStr.BlnLeyendoOrigen = True
        ObjTablaOrigenStr.ObjValorPro = Nothing
        ObjExigeRequeridosBln.BlnLeyendoOrigen = True
        ObjExigeRequeridosBln.ObjValorPro = True
    End Sub

    Friend Function FblnImportoDatos() As Boolean
        Dim ltspDuracion As TimeSpan
        Dim dtmFin As Date, lblnEsElUltimo As Boolean
        Dim dtmInicio As Date = Now, lblnImporto = False
        Try
            If Not IsNothing(DtbTablaDatosOrigen) Then
                Dim lstrMens = String.Empty
                If MblnDatosOrigenOk Then
                    GobjPanDat.SControleProcesoObj(True)
                    MshrExportados = 0
                    MshrNoExportados = 0
                    MshrNumeroregistro = 0
                    Try
                        For Each ldrwRegistroOrigen As DataRow In DtbTablaDatosOrigen.Rows
                            MshrNumeroregistro += 1
                            MobjObjetoDestino.SVacie()
                            lblnEsElUltimo = MshrNumeroregistro = DtbTablaDatosOrigen.Rows.Count
                            If FblnContieneDatos(ldrwRegistroOrigen) Then
                                SCreeObjeto(ldrwRegistroOrigen, lblnEsElUltimo)
                                ObjArgumentoEventoPan.BlnCancele = False
                                RaiseEvent EvnImportadoObj(Me, ObjArgumentoEventoPan)
                            End If
                        Next
                    Catch ex As PanDatException
                        Throw
                    Catch ex As PanLException
                        Throw
                    Catch ex As Exception
                        Throw
                    Finally
                        lblnImporto = True
                        dtmFin = Date.Now
                        ltspDuracion = dtmFin - dtmInicio
                        MstrTiempoImpor = ltspDuracion.ToString
                        SEscribaInformeImp(False, False, True)
                        ObjArgumentoEventoPan.BlnCancele = False
                        RaiseEvent EvnFinImportar(Me, ObjArgumentoEventoPan)
                        MobjObjetoDestino.SNormaliceEstado(True)
                        GobjPanorama.SRegistreAccionLogApp(HstrNombreClase, "Importación " &
                            ObjArchivoOrigenStr.ToString)
                        GobjPanDat.SControleProcesoObj(False)
                    End Try
                Else
                    lblnImporto = False
                    SLevanteEventoNot(lstrMens, "", 0, EnuSeveridadNot.EnuDatoInvalido)
                End If
            End If
        Catch ex As ProveedorBdPanException
            Throw
        Catch ex As ParametrosConexionBDPanException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            GblnImportando = False
        End Try
        Return lblnImporto
    End Function

    Private Shared Function FblnContieneDatos(adrwOrigen As DataRow) As Boolean
        Dim lblnContDat = False
        For Each ldclColuma As DataColumn In adrwOrigen.Table.Columns
            If Not IsNothing(ClsPanorama.FobjValorCampo(adrwOrigen(ldclColuma),
                    EnuTipoValor.EnuObjeto)) Then
                lblnContDat = True
                Exit For
            End If
        Next
        Return lblnContDat
    End Function

    Private Sub SCreeObjeto(adrwRegistroOrigen As DataRow, ablnElUltimo As Boolean)
        Dim lstrColumnaOri As String
        Dim lobjVlrCampOri As Object
        Dim lblnEsAdicionable = False
        If Not IsNothing(StrColumnasRelacionadas) Then
            If MobjObjetoDestino.EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuConsultando Then
                MobjObjetoDestino.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando
            End If
            Dim lblnEsConsumo = FblnEsConsumo(DtbTablaDatosOrigen)
            MobjObjetoDestino.SCreeObj(Nothing)
            MobjObjetoDestino.SPrepareParaImportacion()
            MobjObjetoDestino.BlnImportoUltimo = ablnElUltimo
            If lblnEsConsumo Then
                SRegistreValoresPropsConsumo(adrwRegistroOrigen)
            Else
                For Each lobjProp As ClsCBPropiedad In MobjObjetoDestino.ColPropiedades
                    lstrColumnaOri = FstrColumnaOrigen(StrColumnasRelacionadas,
                            lobjProp.StrNombreCampoBD)
                    If Not String.IsNullOrEmpty(lstrColumnaOri) Then
                        lobjVlrCampOri = ClsPanorama.FobjValorCampo(adrwRegistroOrigen(
                                lstrColumnaOri), lobjProp.EnuTipoValor)
                        lobjProp.ObjValorPro = lobjVlrCampOri
                    End If
                Next
            End If
            If Not lblnEsAdicionable Then
                lblnEsAdicionable = MobjObjetoDestino.FblnEsCreable
            End If
            If lblnEsAdicionable Then
                If ObjExigeRequeridosBln.ObjValorPro Then
                    If Not MobjObjetoDestino.FblnEstanTodosOk Then
                        SEscribaInformeImp(False, False, False)
                        MobjObjetoDestino.SVacie()
                        MobjObjetoDestino.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando
                        MshrNoExportados += 1
                    Else
                        MobjObjetoDestino.SActualice(True)
                        MshrExportados += 1
                    End If
                Else
                    MobjObjetoDestino.SActualice(False)
                    MshrExportados += 1
                End If
            Else
                SEscribaInformeImp(False, True, False)
                MobjObjetoDestino.SVacie()
                MobjObjetoDestino.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando
                MshrNoExportados += 1
            End If
        End If
    End Sub

    ''' <summary>
    ''' registra en el objeto los valores de las propiedades cuando la importación son
    ''' servicios con consumo
    ''' </summary>
    ''' <param name="adrwRegistroOrigen"></param>
    Private Sub SRegistreValoresPropsConsumo(adrwRegistroOrigen As DataRow)
        Dim lblnRegistrado As Boolean, lstrColumnaOri As String, lobjVlrColumnaOri As Object
        For Each lobjProp As ClsCBPropiedad In MobjObjetoDestino.ColPropiedades
            lblnRegistrado = False
            If lobjProp.StrNombreCampoBD = ClsIdCentroUtilShr.SstrNombreCampoBd OrElse
                            lobjProp.StrNombreCampoBD = "CantidadPeriodos" Then
                lobjProp.ObjValorPro = 1
                lblnRegistrado = True
            ElseIf lobjProp.StrNombreCampoBD = "IdAno" Then
                lobjProp.ObjValorPro = 0
                lblnRegistrado = True
            End If
            If Not lblnRegistrado Then
                lstrColumnaOri = FstrColumnaOrigen(StrColumnasRelacionadas,
                        lobjProp.StrNombreCampoBD)
                If Not String.IsNullOrEmpty(lstrColumnaOri) Then
                    lobjVlrColumnaOri = ClsPanorama.FobjValorCampo(adrwRegistroOrigen(
                        lstrColumnaOri), lobjProp.EnuTipoValor)
                    lobjProp.ObjValorPro = lobjVlrColumnaOri
                End If
            End If
        Next
    End Sub

    Friend Sub SEscribaInformeImp(ablnImportado As Boolean, ablnYaExiste As Boolean, ablnFin As Boolean)
        Dim lswNoImportados As StreamWriter
        Dim lstrLinea As String
        Dim lstrArchivo As String = StrNombreArchivoResultados
        If Not MblnHayReporte Then
            lswNoImportados = ClsPanorama.FswStreamWriter(lstrArchivo)
            MblnHayReporte = True
        Else
            lswNoImportados = File.AppendText(lstrArchivo)
        End If
        If Not ablnFin Then
            lstrLinea = "Registro número " & MshrNumeroregistro.ToString
            For Each lobjProp As ClsCBPropiedad In MobjObjetoDestino.ColPropiedades
                If lobjProp.BlnEsLlave Then
                    If IsNothing(lobjProp.ObjValorPro) Then
                        lstrLinea &= ", " & lobjProp.StrNombre & " = " & "Sin valor"
                    Else
                        lstrLinea &= ", " & lobjProp.StrNombre & " = " & lobjProp.ObjValorPro.ToString
                    End If
                End If
            Next
            If ablnImportado Then
                lstrLinea &= " Objeto Importado"
            Else
                lstrLinea &= " Objeto No Importado"
            End If
            If ablnYaExiste Then
                lstrLinea &= " Este objeto ya existe o no existe el servicio a importar!" & vbCrLf
            Else
                lstrLinea &= vbCrLf & vbTab & "Campos invalidos:" & vbCrLf
                For Each lobjProp As ClsCBPropiedad In MobjObjetoDestino.ColPropiedades
                    If Not lobjProp.BlnEsValido Then
                        If IsNothing(lobjProp.ObjValorPro) Then
                            lstrLinea &= vbTab & vbTab & "" & lobjProp.StrNombre & " = " & " " & vbCrLf
                        Else
                            lstrLinea &= vbTab & vbTab & "" & lobjProp.StrNombre & " = " &
                                    lobjProp.ObjValorPro.ToString & vbCrLf
                        End If
                    End If
                Next
            End If
        Else
            lstrLinea = vbCrLf & vbCrLf & "Objetos Importados: " & CType(MshrExportados, String) & vbCrLf
            lstrLinea &= "Objetos no Importados: " & CType(MshrNoExportados, String) & vbCrLf
            lstrLinea &= "Tiempo de importación: " & MstrTiempoImpor
        End If
        lswNoImportados.WriteLine(lstrLinea)
        lswNoImportados.Close()
    End Sub

    Friend Function FblnSonValidosDatosTabla(ByRef astrMens As String) As Boolean
        MblnDatosOrigenOk = ObjObjetoDestino.FblnSonValidosDatosOrigen(DtbTablaDatosOrigen,
                StrColumnasRelacionadas, True, astrMens)
        Return MblnDatosOrigenOk
    End Function

#Region "Manejo datos de origen"
    Friend Function FstrTablasOrigen()
        Dim ldtbTablasArchivoOrigen As DataTable
        Dim lstrTablasArchOrigen As String() = {}
        If ObjArchivoOrigenStr.BlnEsValido Then
            Dim lstrArchOrig As String = ObjArchivoOrigenStr.ObjValorPro
            ldtbTablasArchivoOrigen = GobjPanDat.FdtbTablasAccess(lstrArchOrig, String.Empty)
            If ldtbTablasArchivoOrigen IsNot Nothing Then
                ReDim lstrTablasArchOrigen(0 To ldtbTablasArchivoOrigen.Rows.Count - 1)
                Dim i = 0
                For Each ldrwTabla As DataRow In ldtbTablasArchivoOrigen.Rows
                    lstrTablasArchOrigen(i) = ldrwTabla("TABLE_NAME")
                    i += 1
                Next
            Else
                lstrTablasArchOrigen = {}
            End If
        End If
        Return lstrTablasArchOrigen
    End Function
    Private Function FstrColumnasRelacionadas() As String()
        Dim i = 0, lstrColumnasRelacionadas As String() = {}
        Dim lstrColumnasObjeto As String() = FstrColumnasObjetoDes()
        If MstrColumnasOrigen IsNot Nothing AndAlso FstrColumnasOrigen.Count > 0 Then
            For Each lstrCampoObjeto As String In lstrColumnasObjeto
                If MstrColumnasOrigen.Contains(lstrCampoObjeto.ToLower) Then
                    ReDim Preserve lstrColumnasRelacionadas(i)
                    lstrColumnasRelacionadas(i) = lstrCampoObjeto & "=" & lstrCampoObjeto
                    i += 1
                End If
            Next
        End If
        Return lstrColumnasRelacionadas
    End Function
    Private Sub SCargueDatosOrigen()
        Try
            If ObjArchivoOrigenStr.BlnEsValido Then
                Dim lstrArchOrig As String = ObjArchivoOrigenStr.ObjValorPro
                Dim ldtbDatosOrigen = GobjPanDat.FdtbTablasAccess(lstrArchOrig, String.Empty)
                DtbTablaDatosOrigen = ldtbDatosOrigen
            Else
                DtbTablaDatosOrigen = Nothing
            End If
        Catch ex As ParametrosConexionBDPanException
            Throw
        Catch ex As ProveedorBdPanException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Sub
    Friend Function FstrColumnasOrigen() As String()
        Dim ldtbTablaOrigen = DtbTablaDatosOrigen()
        If ldtbTablaOrigen IsNot Nothing Then
            Dim i = 0
            ReDim MstrColumnasOrigen(0 To ldtbTablaOrigen.Columns.Count - 1)
            For Each lclmCampoTabla As DataColumn In ldtbTablaOrigen.Columns
                MstrColumnasOrigen(i) = lclmCampoTabla.ColumnName.ToLower
                i += 1
            Next
        Else
            MstrColumnasOrigen = {}
            MstrColumnasRelacionadas = {}
        End If
        Return MstrColumnasOrigen
    End Function
    Private Function FdtbTablaDatosOrigen() As DataTable
        Dim ldtbTablaOrigen As DataTable = Nothing
        Try
            If ObjArchivoOrigenStr.BlnEsValido AndAlso ObjTablaOrigenStr.BlnEsValido Then
                ldtbTablaOrigen = ClsPanorama.FdtbTablaAccess(
                        ObjArchivoOrigenStr.ObjValorPro, String.Empty,
                        ObjTablaOrigenStr.ObjValorPro)
                SlimpieTblorigen(ldtbTablaOrigen)
            End If
        Catch ex As ParametrosConexionBDPanException
            Throw
        Catch ex As ProveedorBdPanException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        End Try
        Return ldtbTablaOrigen
    End Function
    Private Sub SlimpieTblorigen(adtbOrigen As DataTable)
        Dim ldrwVacias As DataRow() = {}, i = 0
        For Each ldrwOrigen As DataRow In adtbOrigen.Rows
            If IsDBNull(ldrwOrigen(0)) AndAlso IsDBNull(ldrwOrigen(1)) AndAlso
                    IsDBNull(ldrwOrigen(3)) AndAlso IsDBNull(ldrwOrigen(4)) Then
                ReDim Preserve ldrwVacias(i)
                ldrwVacias(i) = ldrwOrigen
                i += 1
            End If
        Next
        If ldrwVacias.Length > 0 Then
            For Each ldrwVacia As DataRow In ldrwVacias
                adtbOrigen.Rows.Remove(ldrwVacia)
            Next
        End If
    End Sub
    Private Function FblnRegistroVacio(adrwTblori As DataRow)
        Dim lblnRegVacio = True
        For Each ldclColumna As DataColumn In adrwTblori.Table.Columns
            If Not IsDBNull(adrwTblori(ldclColumna.ColumnName)) Then
                If Not String.IsNullOrEmpty(adrwTblori(ldclColumna.ColumnName)) Then
                    lblnRegVacio = False
                    Exit For
                End If
            End If
        Next
        Return lblnRegVacio
    End Function
#End Region
#End Region

#Region "Notificaciones"
    Friend Overrides Function FblnNotificaOk(aenuIdMensNot As EnuIdMens) As Boolean
        Dim lstrMens = String.Empty
        If aenuIdMensNot = EnuIdMens.EnuDatosTablaOri Then
            Dim ldtbTablaorigen = DtbTablaDatosOrigen()
            If ldtbTablaorigen Is Nothing Then
                MblnDatosOrigenOk = False
            Else
                If Not MblnDatosOrigenOk Then
                    MblnDatosOrigenOk = ObjObjetoDestino.FblnSonValidosDatosOrigen(
                            ldtbTablaorigen, StrColumnasRelacionadas, False, lstrMens)
                End If
            End If
        End If
        Return MblnDatosOrigenOk
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsArchivoOrigenStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "ArchivoOrigen"
        HshrLongitud = 250
        HenuTipoValor = EnuTipoValor.EnuString
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            HblnEsValido = My.Computer.FileSystem.FileExists(HobjValorNew)
        End If
        If HblnEsValido Then
            If Not (HobjValorNew.ToString.EndsWith(".xls") OrElse HobjValorNew.ToString.EndsWith(".xlsx") OrElse
                    HobjValorNew.ToString.EndsWith(".mdb")) Then
                HblnEsValido = False
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class

Friend Class ClsTablaOrigenStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TablaOrigen"
        HshrLongitud = 250
        HenuTipoValor = EnuTipoValor.EnuString
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HstrMens = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        HobjValorNew = Trim(HobjValorNew)
        HblnEsValido = Not HobjValorNew.ToString.Contains(" "c)
        If Not HblnEsValido Then
            HstrMens = "El nombre de la tabla no puede contener espacios!"
        End If
        If HblnEsValido Then
            HblnEsValido = HobjValorNew.ToString().Length <= 31
            If Not HblnEsValido Then
                HstrMens = "La longitud máxima del nombre de una hoja de Excel es de 31 caracteres!"
            End If
        End If
        If Not String.IsNullOrEmpty(HstrMens) Then
            SNotifiqueDatInv()
        End If
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class

Friend Class ClsExigeRequeridosBln
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "ExigeRequeridos"
        HenuTipoValor = EnuTipoValor.EnuBoolean
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = (ClsPanorama.FblnEsValidoBuleano(HobjValorNew))
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
#End Region