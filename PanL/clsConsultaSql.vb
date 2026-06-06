Friend Class ClsConsultaSql
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    '
    Private ReadOnly MstrNombreTabla As String = String.Empty
    Private McolNombresConsultas As Collection = Nothing
    Private McolNombresConsultasDeTipo As Collection = Nothing
    Private MdtbResultado As DataTable = Nothing
    Private MdtbNombresConsultas As DataTable = Nothing
    Private MintCantidadRegistros As Integer = 0
    Private MstrTiempoEjecucion As String = String.Empty
    Private ReadOnly MblnPermisoAlto As Boolean = False
#End Region
#Region "Constructores"
    Public Sub New(ablnPermisoAlto As Boolean)
        HobjPadre = Nothing
        hblnEsAnulable = False
        '
        Dim lstrCamposSelect() As String = {clsIdConsultaShr.sstrNombreCampoBd}
        MstrNombreTabla = "PanConsultasSql"
        HcolTablas.Add(MstrNombreTabla)
        HcolCamposSelect.Add(lstrCamposSelect)

        MblnPermisoAlto = ablnPermisoAlto
    End Sub
#End Region
#Region "Propiedades"
#Region "Propiedades indentificadoras"
    Protected Overrides ReadOnly Property HstrNombreTabla As String
        Get
            Return mstrNombreTabla
        End Get
    End Property
    Friend ReadOnly Property SstrNombreTabla As String
        Get
            Return MstrNombreTabla
        End Get
    End Property
    Protected Friend Overrides ReadOnly Property HenuIdClase As EnuIdClasesPanDef
        Get
            Return EnuIdClasesPanDef.enuConsultaSql
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Consultas SQL"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjCamposActualizarStr As New ClsCamposActualizarStr(Me)
    Friend ReadOnly Property ObjCamposAgrupamientoStr As New ClsCamposAgrupamientoStr(Me)
    Friend ReadOnly Property ObjCamposInsertarStr As New ClsCamposInsertarStr(Me)
    Friend ReadOnly Property ObjCamposIndiceStr As New ClsCamposIndiceStr(Me)
    Friend ReadOnly Property ObjCamposPrimariosStr As New ClsCamposPrimariosStr(Me)
    Friend ReadOnly Property ObjCamposRefActualizacionStr As New ClsCamposRefActualizacionStr(Me)
    Friend ReadOnly Property ObjCamposRefEliminacionStr As New ClsCamposRefEliminacionStr(Me)
    Friend ReadOnly Property ObjCamposRelPrimariosStr As New ClsCamposRelPrimariosStr(Me)
    Friend ReadOnly Property ObjCamposRelSecundariosStr As New ClsCamposRelSecundariosStr(Me)
    Friend ReadOnly Property ObjCamposSecundariosStr As New ClsCamposSecundariosStr(Me)
    Friend ReadOnly Property ObjExpresionSqlStr As New ClsExpresionSqlStr(Me)
    Friend ReadOnly Property ObjEsIndiceUnicoBln As New ClsEsIndiceUnicoBln(Me)
    Friend ReadOnly Property ObjFiltroStr As New ClsFiltroStr(Me)
    Friend ReadOnly Property ObjIdAppConsultaShr As New ClsIdAppConsultaShr(Me)
    Friend ReadOnly Property ObjIdConsultaShr As New ClsIdConsultaShr(Me)
    Friend ReadOnly Property ObjTipoConsultaByt As New ClsTipoConsultaByt(Me)
    Friend ReadOnly Property ObjNombreConsultaStr As New ClsNombreConsultaStr(Me)
    Friend ReadOnly Property ObjTablaActualizarStr As New ClsTablaActualizarStr(Me)
    Friend ReadOnly Property ObjTablaEliminarStr As New ClsTablaEliminarStr(Me)
    Friend ReadOnly Property ObjTablaInsertarStr As New ClsTablaInsertarStr(Me)
    Friend ReadOnly Property ObjTablaPrimariaStr As New ClsTablaPrimariaStr(Me)
    Friend ReadOnly Property ObjTablaSecundariaStr As New ClsTablaSecundariaStr(Me)
    Friend ReadOnly Property ObjValoresActualizacionStr As New ClsValoresActualizacionStr(Me)
    Friend ReadOnly Property ObjValoresInsertarStr As New ClsValoresInsertarStr(Me)
    Friend ReadOnly Property ObjValoresRefActualizacionStr As New ClsValoresRefActualizacionStr(Me)
    Friend ReadOnly Property ObjValoresRefEliminacionStr As New ClsValoresRefEliminacionStr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjCamposActualizarStr)
                HcolPropiedades.Add(ObjCamposAgrupamientoStr)
                HcolPropiedades.Add(ObjCamposIndiceStr)
                HcolPropiedades.Add(ObjCamposInsertarStr)
                HcolPropiedades.Add(ObjCamposPrimariosStr)
                HcolPropiedades.Add(ObjCamposRefActualizacionStr)
                HcolPropiedades.Add(ObjCamposRefEliminacionStr)
                HcolPropiedades.Add(ObjCamposRelPrimariosStr)
                HcolPropiedades.Add(ObjCamposRelSecundariosStr)
                HcolPropiedades.Add(ObjCamposSecundariosStr)
                HcolPropiedades.Add(ObjEsIndiceUnicoBln)
                HcolPropiedades.Add(ObjExpresionSqlStr)
                HcolPropiedades.Add(ObjFiltroStr)
                HcolPropiedades.Add(ObjIdAppConsultaShr)
                HcolPropiedades.Add(ObjIdConsultaShr)
                HcolPropiedades.Add(ObjNombreConsultaStr)
                HcolPropiedades.Add(ObjTablaActualizarStr)
                HcolPropiedades.Add(ObjTablaEliminarStr)
                HcolPropiedades.Add(ObjTablaInsertarStr)
                HcolPropiedades.Add(ObjTablaPrimariaStr)
                HcolPropiedades.Add(ObjTablaSecundariaStr)
                HcolPropiedades.Add(ObjTipoConsultaByt)
                HcolPropiedades.Add(ObjValoresActualizacionStr)
                HcolPropiedades.Add(ObjValoresInsertarStr)
                HcolPropiedades.Add(ObjValoresRefActualizacionStr)
                HcolPropiedades.Add(ObjValoresRefEliminacionStr)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras Propiedades"
    Friend ReadOnly Property DtbTablaResultado As DataTable
        Get
            Return mdtbResultado
        End Get
    End Property
    Friend ReadOnly Property StrResultado As String
        Get
            Dim lstrResultado As String = String.Empty
            If MintCantidadRegistros > 0 Then
                lstrResultado = Format(MintCantidadRegistros, "###,##0") &
                        " Registros en " & MstrTiempoEjecucion & " Horas."
            End If
            Return lstrResultado
        End Get
    End Property
    Friend ReadOnly Property StrTiempoEjecucion As String
        Get
            Return MstrTiempoEjecucion
        End Get
    End Property
    Friend ReadOnly Property EntCantidadRegistros As Integer
        Get
            Return MintCantidadRegistros
        End Get
    End Property
    Friend ReadOnly Property ColNombresConsultasDeTipo As Collection
        Get
            SPuebleColNombresConsultasDeTipo()
            Return McolNombresConsultasDeTipo
        End Get
    End Property
    Friend ReadOnly Property DtbNombresConsultas As DataTable
        Get
            SCargueDtbNombresConsultas()
            Return MdtbNombresConsultas
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Friend Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        MyBase.sActualice(ablnExigeRequeridos)
        mdtbNombresConsultas = Nothing
        sPuebleColNombresConsultas()
    End Sub
    Protected Friend Overrides Sub SVacie()
        MyBase.SVacie()
        McolNombresConsultasDeTipo = Nothing
    End Sub
#End Region
#Region "Procedimientos del objeto"
    Friend Function FalsNombresConsultas() As ArrayList
        Dim lalsNombresCols As New ArrayList
        Dim lenuTipoConsulta As enuTipoConsultaDef = enuTipoConsultaDef.enuExpresionSql
        If Not IsNothing(ObjTipoConsultaByt.ObjValorPro) Then
            lenuTipoConsulta = ObjTipoConsultaByt.ObjValorPro
        End If
        SCargueDtbNombresConsultas()
        Dim lstrFiltro As String = clsTipoConsultaByt.sstrNombreCampoBd & " = " & lenuTipoConsulta
        Dim ldrwConsultas() As DataRow = mdtbNombresConsultas.Select(lstrFiltro)
        Dim lstrNomCon As String
        If ldrwConsultas.Length > 0 Then
            For Each ldrwConsulta As DataRow In ldrwConsultas
                lstrNomCon = ClsPanorama.FobjValorCampo(ldrwConsulta(
                            ObjNombreConsultaStr.StrNombreCampoBD), EnuTipoValor.enuString)
                lalsNombresCols.Add(lstrNomCon)
            Next
        End If
        lalsNombresCols.Sort()
        Return lalsNombresCols
    End Function
    Private Sub SPuebleColNombresConsultasDeTipo()
        SCargueDtbNombresConsultas()
        If IsNothing(McolNombresConsultasDeTipo) Then
            McolNombresConsultasDeTipo = New Collection
        Else
            McolNombresConsultasDeTipo.Clear()
        End If
        Dim lstrFiltro As String = ObjTipoConsultaByt.StrNombreCampoBD & " = " & ObjTipoConsultaByt.ObjValorPro
        Dim ldrwConsultas As DataRow() = Nothing
        If Not IsNothing(MdtbNombresConsultas) Then
            ldrwConsultas = MdtbNombresConsultas.Select(lstrFiltro)
        End If
        If Not IsNothing(ldrwConsultas) AndAlso ldrwConsultas.Length > 0 Then
            For Each ldrwConsulta As DataRow In ldrwConsultas
                Dim lstrNomCon As String = ClsPanorama.FobjValorCampo(ldrwConsulta(
                        ObjNombreConsultaStr.StrNombreCampoBD), EnuTipoValor.enuString)
                McolNombresConsultasDeTipo.Add(lstrNomCon)
            Next
        End If
    End Sub
    Friend Sub SPuebleColNombresConsultas()
        SCargueDtbNombresConsultas()
        If IsNothing(McolNombresConsultas) Then
            McolNombresConsultas = New Collection
        Else
            McolNombresConsultas.Clear()
        End If
        If Not IsNothing(MdtbNombresConsultas) Then
            Dim ldrwConsultas As DataRow() = MdtbNombresConsultas.Select()
            If ldrwConsultas.Count > 0 Then
                For Each ldrwConsulta As DataRow In ldrwConsultas
                    Dim lstrNomCon As String = ClsPanorama.FobjValorCampo(ldrwConsulta(
                                ObjNombreConsultaStr.StrNombreCampoBD), EnuTipoValor.enuString)
                    McolNombresConsultas.Add(lstrNomCon)
                Next
            End If
        End If
    End Sub
    Private Sub SCargueDtbNombresConsultas()
        MdtbNombresConsultas = Nothing
        Dim lstrFiltro = ObjIdAppConsultaShr.StrNombreCampoBD & " = " & GenuIdAplicacion
        Using ldstNombresConsultas As New DataSet
            GobjPanDat.SdsDataSet(ldstNombresConsultas, SstrNombreTabla, {ObjIdConsultaShr.StrNombreCampoBD,
                    ObjNombreConsultaStr.StrNombreCampoBD, ObjTipoConsultaByt.StrNombreCampoBD},
                    {{ObjIdAppConsultaShr.StrNombreCampoBD, "ASC"},
                     {ObjIdConsultaShr.StrNombreCampoBD, "ASC"}}, lstrFiltro, False, Array.Empty(Of String))
            If ldstNombresConsultas.Tables.Count > 0 Then
                MdtbNombresConsultas = ldstNombresConsultas.Tables(0)
            Else
                MdtbNombresConsultas = Nothing
            End If
        End Using
    End Sub
    Public Sub SEjecute()
        MdtbResultado = Nothing
        MstrTiempoEjecucion = String.Empty
        MintCantidadRegistros = 0
        If FblnEstanTodosOk() Then
            Select Case ObjTipoConsultaByt.ObjValorPro
                Case EnuTipoConsultaDef.enuExpresionSql
                    SEjecuteExpresionSql()
                Case EnuTipoConsultaDef.enuSelect
                    SEjecuteSelect()
                Case EnuTipoConsultaDef.enuUpdate
                    SEjecuteUpdate()
                Case EnuTipoConsultaDef.enuInsert
                    SEjecuteInsert()
                Case EnuTipoConsultaDef.enuDelete
                    SEjecuteDelete()
            End Select
        End If
    End Sub
    Private Sub SEjecuteExpresionSql()
        Dim ltspDuracion As TimeSpan
        Dim dtmFin As Date
        Dim dtmInicio As Date = Now
        MstrTiempoEjecucion = String.Empty
        MintCantidadRegistros = 0
        MdtbResultado = Nothing
        Dim lstrExpresionSql As String = ObjExpresionSqlStr.ObjValorPro.ToString.Replace(Chr(34), "'")
        If lstrExpresionSql.ToUpper.StartsWith("SELECT") Then
            Using ldsExpresionSql As New DataSet
                GobjPanDat.SdsDataSet(ldsExpresionSql, ObjExpresionSqlStr.ObjValorPro)
                MdtbResultado = ldsExpresionSql.Tables(0)
                MintCantidadRegistros = MdtbResultado.Rows.Count
            End Using
        Else
            If MblnPermisoAlto Then
                MintCantidadRegistros = GobjPanDat.SEjecuteSentenciaSql(lstrExpresionSql)
            End If
        End If
        dtmFin = Date.Now
        ltspDuracion = dtmFin - dtmInicio
        MstrTiempoEjecucion = ltspDuracion.ToString
    End Sub
    Private Sub SEjecuteSelect()
        Dim ltspDuracion As TimeSpan
        Dim dtmFin As Date
        Dim dtmInicio As Date = Now
        Dim lblnSelectCompuesto As Boolean = (ObjTablaSecundariaStr.ObjValorPro <> "")
        Dim lstrNombreTablaPri As String = ObjTablaPrimariaStr.ObjValorPro
        Dim lstrCamposPri As String() = ObjCamposPrimariosStr.ObjValorPro.ToString.Split(",")
        Dim lstrIndice(,) As String = FstrIndice()
        Dim lstrFiltro As String = ObjFiltroStr.ObjValorPro
        Dim lstrCamposGrupo As String()
        Dim lblnIndiceUnico As Boolean = ObjEsIndiceUnicoBln.ObjValorPro
        MdtbResultado = Nothing
        MstrTiempoEjecucion = String.Empty
        MintCantidadRegistros = 0
        If Not IsNothing(lstrFiltro) Then
            lstrFiltro = lstrFiltro.Replace(Chr(34), "'")
        Else
            lstrFiltro = String.Empty
        End If
        If Not String.IsNullOrEmpty(ObjCamposAgrupamientoStr.ToString) Then
            lstrCamposGrupo = ObjCamposAgrupamientoStr.ObjValorPro.ToString.Split(",")
        Else
            lstrCamposGrupo = {""}
        End If
        If lblnSelectCompuesto Then
            Dim lstrNombreTablaSec As String = ObjTablaSecundariaStr.ObjValorPro
            Dim lstrCamposRelPri() As String = ObjCamposRelPrimariosStr.ObjValorPro.ToString.Split(",")
            Dim lstrCamposRelSec() As String = ObjCamposRelSecundariosStr.ObjValorPro.ToString.Split(",")
            Dim lstrCamposSec As String() = ObjCamposSecundariosStr.ObjValorPro.ToString.Split(",")
            Using ldsSelect As New DataSet
                GobjPanDat.SdsDataSet(ldsSelect, lstrNombreTablaPri, lstrCamposPri, lstrNombreTablaSec,
                                lstrCamposSec, lstrCamposRelPri, lstrCamposRelSec, lstrIndice, lstrFiltro,
                                lblnIndiceUnico, lstrCamposGrupo)
                MdtbResultado = ldsSelect.Tables(0)
            End Using
        Else
            Using ldsSelect As New DataSet
                GobjPanDat.SdsDataSet(ldsSelect, lstrNombreTablaPri, lstrCamposPri,
                            lstrIndice, lstrFiltro, lblnIndiceUnico, lstrCamposGrupo)
                MdtbResultado = ldsSelect.Tables(0)
            End Using
        End If
        dtmFin = Date.Now
        ltspDuracion = dtmFin - dtmInicio
        MstrTiempoEjecucion = ltspDuracion.ToString
        If Not IsNothing(MdtbResultado) Then
            MintCantidadRegistros = MdtbResultado.Rows.Count
        End If
    End Sub
    Private Sub SEjecuteUpdate()
        Dim ltspDuracion As TimeSpan
        Dim dtmFin As Date
        Dim dtmInicio As Date = Now
        Dim lcolCamposAct As New Collection
        Dim lcolDatosAct As New Collection
        Dim lcolCamposRef As New Collection
        Dim lcolDatosRef As New Collection
        Dim lstrNombreTablaAct As String = ObjTablaActualizarStr.ObjValorPro
        Dim lstrCamposAct As String() = ObjCamposActualizarStr.ObjValorPro.ToString.Split(",")
        Dim lstrValoresAct As String()
        Dim lstrCamposRef As String() = ObjCamposRefActualizacionStr.ObjValorPro.ToString.Split(",")
        Dim lstrValoresRef As String()
        Dim lstrValAct As String = ObjValoresActualizacionStr.ObjValorPro.ToString.Replace(Chr(34), "'")
        Dim lstrValRef As String = ObjValoresRefActualizacionStr.ObjValorPro.ToString.Replace(Chr(34), "'")
        MstrTiempoEjecucion = String.Empty
        MintCantidadRegistros = 0
        lstrValoresAct = lstrValAct.Split(",")
        lstrValoresRef = lstrValRef.Split(",")
        For Each lstrCamAct As String In lstrCamposAct
            lcolCamposAct.Add(lstrCamAct)
        Next
        For Each lstrDatAct As String In lstrValoresAct
            lcolDatosAct.Add(lstrDatAct)
        Next
        For Each lstrCamRef As String In lstrCamposRef
            lcolCamposRef.Add(lstrCamRef)
        Next
        For Each lstrDatRef As String In lstrValoresRef
            lcolDatosRef.Add(lstrDatRef)
        Next
        Using ldsSelect As New DataSet
            MintCantidadRegistros = GobjPanDat.SActualiceRegistro(lstrNombreTablaAct, lcolCamposAct, lcolDatosAct,
                        lcolCamposRef, lcolDatosRef, Nothing)
        End Using
        dtmFin = Date.Now
        ltspDuracion = dtmFin - dtmInicio
        MstrTiempoEjecucion = ltspDuracion.ToString
    End Sub
    Private Sub SEjecuteInsert()
        Dim ltspDuracion As TimeSpan
        Dim dtmFin As Date
        Dim dtmInicio As Date = Now
        Dim lstrNombreTabla As String = ObjTablaInsertarStr.ObjValorPro
        Dim lcolCampos As New Collection
        Dim lcolDatos As New Collection
        MintCantidadRegistros = 0
        MstrTiempoEjecucion = String.Empty
        For Each lstrCampo In ObjCamposInsertarStr.ObjValorPro.ToString.Split(",")
            lcolCampos.Add(lstrCampo.Trim)
        Next
        For Each lobjValor As Object In ObjValoresInsertarStr.ObjValorPro.ToString.Split(",")
            If IsNumeric(lobjValor) Then
                lobjValor = Val(lobjValor)
            ElseIf lobjValor.ToString.ToUpper = "TRUE" Then
                lobjValor = True
            ElseIf lobjValor.ToString.ToUpper = "FALSE" Then
                lobjValor = False
            ElseIf IsDate(lobjValor) Then
                lobjValor = CType(lobjValor, Date)
            Else
                lobjValor = lobjValor.ToString.Trim
            End If
            lcolDatos.Add(lobjValor)
        Next
        MintCantidadRegistros = GobjPanDat.SInserteRegistro(lstrNombreTabla, lcolCampos, lcolDatos)
        dtmFin = Date.Now
        ltspDuracion = dtmFin - dtmInicio
        MstrTiempoEjecucion = ltspDuracion.ToString
    End Sub
    Private Sub SEjecuteDelete()
        Dim ltspDuracion As TimeSpan
        Dim dtmFin As Date
        Dim dtmInicio As Date = Now
        Dim lstrNombreTabla As String = ObjTablaEliminarStr.ObjValorPro
        Dim lcolCamposRef As New Collection
        Dim lcolDatosRef As New Collection
        MintCantidadRegistros = 0
        MstrTiempoEjecucion = String.Empty
        For Each lstrCampo In ObjCamposRefEliminacionStr.ObjValorPro.ToString.Split(",")
            lcolCamposRef.Add(lstrCampo.Trim)
        Next
        For Each lobjValor As Object In ObjValoresRefEliminacionStr.ObjValorPro.ToString.Split(",")
            If IsNumeric(lobjValor) Then
                lobjValor = Val(lobjValor)
            ElseIf lobjValor.ToString.ToUpper = "TRUE" Then
                lobjValor = True
            ElseIf lobjValor.ToString.ToUpper = "FALSE" Then
                lobjValor = False
            ElseIf IsDate(lobjValor) Then
                lobjValor = CType(lobjValor, Date)
            Else
                lobjValor = lobjValor.ToString.Trim
            End If
            lcolDatosRef.Add(lobjValor)
        Next
        MintCantidadRegistros = GobjPanDat.SElimineRegistro(lstrNombreTabla, lcolCamposRef, lcolDatosRef)
        dtmFin = Date.Now
        ltspDuracion = dtmFin - dtmInicio
        MstrTiempoEjecucion = ltspDuracion.ToString
    End Sub
    Private Function FstrIndice() As String(,)
        Dim lstrCamposIndice As String = ObjCamposIndiceStr.ObjValorPro
        Dim lstrIndices As String()
        Dim lstrIndice(,) As String = Nothing
        Dim i As Byte = 0
        If Not String.IsNullOrEmpty(lstrCamposIndice) Then
            If lstrCamposIndice.Contains(",") Then
                lstrIndices = lstrCamposIndice.Split(",")
                ReDim lstrIndice(lstrIndices.GetUpperBound(0), 1)
                For Each lstrCampoIndice As String In lstrIndices
                    lstrIndice(i, 0) = lstrCampoIndice.Split(";")(0)
                    If lstrCampoIndice.Contains(";") Then
                        lstrIndice(i, 1) = lstrCampoIndice.Split(";")(1)
                    Else
                        lstrIndice(i, 1) = "ASC"
                    End If
                    i += 1
                Next
            Else
                ReDim lstrIndice(0, 1)
                If lstrCamposIndice.Contains(";") Then
                    lstrIndice(0, 0) = lstrCamposIndice.Split(";")(0)
                    lstrIndice(0, 1) = lstrCamposIndice.Split(";")(1)
                Else
                    lstrIndice(0, 0) = lstrCamposIndice
                    lstrIndice(0, 1) = "ASC"
                End If
            End If
        End If
        Return lstrIndice
    End Function
    Shared Function FstrTiposConsulta() As String()
        Dim ldrwTiposConsulta As DataRow()
        Dim lstrTiposConsulta As String()
        Dim i As Byte = 0
        ldrwTiposConsulta = ClsAdministrador.FdrwConstantesPan(
                EnuGrupoConstantesPanDef.enuTipoConsulta)
        ReDim lstrTiposConsulta(ldrwTiposConsulta.Count - 1)
        For Each ldrwTipoCon As DataRow In ldrwTiposConsulta
            lstrTiposConsulta(i) = ldrwTipoCon("Dato")
            i += 1
        Next
        Return lstrTiposConsulta
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsCamposActualizarStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CamposActualizar"
        HshrLongitud = 300
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "CamposActualizar"
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsConsultaSql = ObjPadre
        HblnEsRequerido = (lobjPadre.ObjTipoConsultaByt.ObjValorPro = EnuTipoConsultaDef.enuUpdate)
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsCamposAgrupamientoStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CamposAgrupamiento"
        HshrLongitud = 100
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "CamposAgrupamiento"
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsCamposInsertarStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CamposInsertar"
        HshrLongitud = 300
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "CamposInsertar"
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsConsultaSql = ObjPadre
        HblnEsRequerido = (lobjPadre.ObjTipoConsultaByt.ObjValorPro = EnuTipoConsultaDef.enuInsert)
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsCamposPrimariosStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CamposPrimarios"
        HshrLongitud = 300
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "CamposPrimarios"
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsConsultaSql = ObjPadre
        HblnEsRequerido = (lobjPadre.ObjTipoConsultaByt.ObjValorPro = EnuTipoConsultaDef.enuSelect)
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsCamposRefActualizacionStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CamposRefActualizacion"
        HshrLongitud = 100
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "CamposRefActualizacion"
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsConsultaSql = ObjPadre
        HblnEsRequerido = (lobjPadre.ObjTipoConsultaByt.ObjValorPro = EnuTipoConsultaDef.enuUpdate)
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return CByte(ObjValorPro).ToString
    End Function
End Class
Friend Class ClsCamposRefEliminacionStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CamposRefEliminacion"
        HshrLongitud = 100
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "CamposRefEliminacion"
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsConsultaSql = ObjPadre
        HblnEsRequerido = (lobjPadre.ObjTipoConsultaByt.ObjValorPro = EnuTipoConsultaDef.enuDelete)
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return CByte(ObjValorPro).ToString
    End Function
End Class
Friend Class ClsCamposRelPrimariosStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CamposRelPrimarios"
        HshrLongitud = 100
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "CamposRelPrimarios"
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsConsultaSql = ObjPadre
        HblnEsRequerido = (Not IsNothing(lobjPadre.ObjTablaSecundariaStr.ObjValorPro) AndAlso
                lobjPadre.ObjTablaSecundariaStr.ObjValorPro <> "")
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return CByte(ObjValorPro).ToString
    End Function
End Class
Friend Class ClsCamposRelSecundariosStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CamposRelSecundarios"
        HshrLongitud = 100
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "CamposRelSecundarios"
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsConsultaSql = ObjPadre
        HblnEsRequerido = (Not IsNothing(lobjPadre.ObjTablaSecundariaStr.ObjValorPro) AndAlso
                lobjPadre.ObjTablaSecundariaStr.ObjValorPro <> "")
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return CByte(ObjValorPro).ToString
    End Function
End Class
Friend Class ClsCamposSecundariosStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CamposSecundarios"
        HshrLongitud = 300
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "CamposSecundarios"
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsConsultaSql = ObjPadre
        HblnEsRequerido = (Not IsNothing(lobjPadre.ObjTablaSecundariaStr.ObjValorPro) AndAlso
                lobjPadre.ObjTablaSecundariaStr.ObjValorPro <> "")
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return CByte(ObjValorPro).ToString
    End Function
End Class
Friend Class ClsExpresionSqlStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "ExpresionSql"
        HshrLongitud = 3000
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "ExpresionSql"
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsConsultaSql = ObjPadre
        HblnEsRequerido = (lobjPadre.ObjTipoConsultaByt.ObjValorPro = EnuTipoConsultaDef.enuExpresionSql)
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 10, ShrLongitud, BlnEsRequerido)
        If HblnEsValido AndAlso BlnEsRequerido Then
            If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                HstrMens = String.Empty
                If HobjValorNew.ToString.Trim.ToUpper.StartsWith("UPDATE ") OrElse
                        HobjValorNew.ToString.Trim.ToUpper.StartsWith("DELETE ") OrElse
                        HobjValorNew.ToString.Trim.ToUpper.StartsWith("INSERT ") Then
                    HblnEsValido = (GstrIdUsuario = GCSTRUSUARIOU)
                    If Not HblnEsValido Then
                        HstrMens = "Esta tipo de Consulta solo es posible para el Usuario OPT!"
                    End If
                ElseIf Not HobjValorNew.ToString.Trim.ToUpper.StartsWith("SELECT ") Then
                    HblnEsValido = False
                    If Not HblnEsValido Then
                        HstrMens = "La Expresión no es una expresion Sql válida!"
                    End If
                End If
                If Not String.IsNullOrEmpty(HstrMens) Then
                    SNotifiqueDatInv()
                End If
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsEsIndiceUnicoBln
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "EsIndiceUnico"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = "EsIndiceUnico"
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return CByte(ObjValorPro).ToString
    End Function
End Class
Friend Class ClsFiltroStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Filtro"
        HshrLongitud = 150
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "Filtro"
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 5, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return CByte(ObjValorPro).ToString
    End Function

End Class
Friend Class ClsIdAppConsultaShr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdApp"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = "IdAplicacion"
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = (HobjValorNew = GenuIdAplicacion)
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdConsultaShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdConsultaSql"
    Private ReadOnly MobjPadre As ClsConsultaSql = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdConsulta"
        HshrLongitud = 0
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 0
        HblnEsAutonumerico = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        HblnEsValido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = True
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando AndAlso
                Not BlnLeyendoOrigen Then
            HstrMens = String.Empty
            lblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Short.MaxValue, BlnEsRequerido,
                                EnuTipoValor.enuShort)
            If lblnEsValido Then
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                    If MobjPadre.ObjIdConsultaShr.ObjValorPro <> HobjValorNew OrElse (Not MobjPadre.BlnExiste) Then
                        Dim lobjLlavePrincipal As Object() = {HobjValorNew}
                        MobjPadre.SAbra(lobjLlavePrincipal)
                    End If
                    If Not MobjPadre.BlnExiste Then
                        HstrMens = "La Id. de la Consulta ingresada no existe!"
                        lblnEsValido = False
                    End If
                End If
            Else
                HstrMens = "La Id. de la Consulta ingresada no es válida!"
            End If
            If Not String.IsNullOrEmpty(HstrMens) Then
                SNotifiqueDatInv()
            End If
        End If
        HblnEsValido = lblnEsValido
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return CByte(ObjValorPro).ToString
    End Function
End Class
Friend Class ClsCamposIndiceStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Indice"
        HshrLongitud = 100
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "Indice"
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 5, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return CByte(ObjValorPro).ToString
    End Function
End Class
Friend Class ClsNombreConsultaStr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsConsultaSql = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Nombre"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "NombreConsulta"
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        With MobjPadre
            If HblnEsValido Then
                If .EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                    If HobjValorNew <> HobjValorPro Then
                        Dim ldtbNombreConsultas As DataTable = .DtbNombresConsultas
                        If Not IsNothing(ldtbNombreConsultas) AndAlso ldtbNombreConsultas.Rows.Count > 0 Then
                            Dim lshrIdConsulta As Short = CType(ClsPanorama.FobjValorCampoDataRow(
                                    ldtbNombreConsultas.Select, StrNombreCampoBD,
                                    .ObjIdConsultaShr.StrNombreCampoBD, HobjValorNew), Short)
                            If lshrIdConsulta >= 0 AndAlso Not BlnLeyendoOrigen Then
                                .ObjIdConsultaShr.ObjValorPro = lshrIdConsulta
                            End If
                        End If
                    End If
                ElseIf .EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                    Dim lalsNombreConsultas As ArrayList = MobjPadre.FalsNombresConsultas
                    If lalsNombreConsultas.Count > 0 Then
                        HblnEsValido = Not lalsNombreConsultas.Contains(HobjValorNew)
                    End If
                    If Not HblnEsValido Then
                        HstrMens = "El nombre de la Consulta ya existe!"
                        SNotifiqueDatInv()
                    End If
                End If
            End If
        End With
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsTablaActualizarStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TablaActualizar"
        HshrLongitud = 25
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "TablaActualizar"
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsConsultaSql = ObjPadre
        HblnEsRequerido = (lobjPadre.ObjTipoConsultaByt.ObjValorPro = EnuTipoConsultaDef.enuUpdate)
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsTablaEliminarStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TablaEliminar"
        HshrLongitud = 25
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "TablaEliminar"
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsConsultaSql = ObjPadre
        HblnEsRequerido = (lobjPadre.ObjTipoConsultaByt.ObjValorPro = EnuTipoConsultaDef.enuDelete)
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsTablaInsertarStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TablaInsertar"
        HshrLongitud = 25
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "TablaInsertar"
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsConsultaSql = ObjPadre
        HblnEsRequerido = (lobjPadre.ObjTipoConsultaByt.ObjValorPro = EnuTipoConsultaDef.enuInsert)
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsTablaPrimariaStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TablaPrimaria"
        HshrLongitud = 25
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "TablaPrimaria"
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsConsultaSql = ObjPadre
        HblnEsRequerido = (lobjPadre.ObjTipoConsultaByt.ObjValorPro = EnuTipoConsultaDef.enuSelect)
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsTablaSecundariaStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TablaSecundaria"
        HshrLongitud = 25
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "TablaSecundaria"
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Private Sub EPosSetValor() Handles Me.EvnPosSetValor
        Dim lobjPadre As ClsConsultaSql = ObjPadre
        With lobjPadre
            If BlnEsValido Then
                .ObjCamposRelPrimariosStr.SValide()
                .ObjCamposRelSecundariosStr.SValide()
                .ObjCamposSecundariosStr.SValide()
            End If
        End With
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsTipoConsultaByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTblTipoConsulta"
    Private MshrIdConsultaAnt As Short = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TipoConsulta"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuTipoConsultaDef.enuSelect,
                EnuTipoConsultaDef.enuExpresionSql, HblnEsRequerido)
    End Sub
    Private Sub EPosCambio() Handles Me.EvnPosCambio
        Dim lobjPadre As ClsConsultaSql = ObjPadre
        Dim lshrIdConsulta As Short = Nothing
        With lobjPadre
            If BlnEsValido AndAlso .EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                If lobjPadre.ColNombresConsultasDeTipo.Count > 0 Then
                    Dim lstrNombreConsulta As String = lobjPadre.ColNombresConsultasDeTipo(1)
                    Dim ldtbNombresConsultas As DataTable = lobjPadre.DtbNombresConsultas
                    lshrIdConsulta = CType(ClsPanorama.FobjValorCampoDataRow(ldtbNombresConsultas.Select,
                            lobjPadre.ObjNombreConsultaStr.StrNombreCampoBD, lobjPadre.ObjIdConsultaShr.StrNombreCampoBD,
                            lstrNombreConsulta), Short)
                    If lshrIdConsulta <> 0 AndAlso MshrIdConsultaAnt <> lshrIdConsulta Then
                        MshrIdConsultaAnt = lshrIdConsulta
                        .ObjIdConsultaShr.ObjValorPro = lshrIdConsulta
                    End If
                End If
            End If
        End With
    End Sub
    Private Sub EPosSetValor() Handles Me.EvnPosSetValor
        Dim lobjPadre As ClsConsultaSql = ObjPadre
        lobjPadre.ObjCamposActualizarStr.SValide()
        lobjPadre.ObjCamposInsertarStr.SValide()
        lobjPadre.ObjCamposPrimariosStr.SValide()
        lobjPadre.ObjCamposRefActualizacionStr.SValide()
        lobjPadre.ObjCamposRefEliminacionStr.SValide()
        lobjPadre.ObjExpresionSqlStr.SValide()
        lobjPadre.ObjTablaActualizarStr.SValide()
        lobjPadre.ObjTablaEliminarStr.SValide()
        lobjPadre.ObjTablaInsertarStr.SValide()
        lobjPadre.ObjTablaPrimariaStr.SValide()
        lobjPadre.ObjValoresActualizacionStr.SValide()
        lobjPadre.ObjValoresInsertarStr.SValide()
        lobjPadre.ObjValoresRefActualizacionStr.SValide()
        lobjPadre.ObjValoresRefEliminacionStr.SValide()
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return CByte(ObjValorPro).ToString
    End Function
End Class
Friend Class ClsValoresActualizacionStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "ValoresActualizacion"
        HshrLongitud = 300
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "ValoresActualizacion"
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsConsultaSql = ObjPadre
        HblnEsRequerido = (lobjPadre.ObjTipoConsultaByt.ObjValorPro = EnuTipoConsultaDef.enuUpdate)
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsValoresRefActualizacionStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "ValoresRefActualizacion"
        HshrLongitud = 100
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "ValoresRefActualizacion"
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsConsultaSql = ObjPadre
        HblnEsRequerido = (lobjPadre.ObjTipoConsultaByt.ObjValorPro = EnuTipoConsultaDef.enuUpdate)
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsValoresRefEliminacionStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "ValoresRefEliminacion"
        HshrLongitud = 100
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "ValoresRefEliminacion"
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsConsultaSql = ObjPadre
        HblnEsRequerido = (lobjPadre.ObjTipoConsultaByt.ObjValorPro = EnuTipoConsultaDef.enuDelete)
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsValoresInsertarStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "ValoresInsertar"
        HshrLongitud = 300
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "ValoresInsertar"
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsConsultaSql = ObjPadre
        HblnEsRequerido = (lobjPadre.ObjTipoConsultaByt.ObjValorPro = EnuTipoConsultaDef.enuInsert)
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
#End Region