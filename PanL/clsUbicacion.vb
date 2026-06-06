Namespace Ubicacion
    Friend Class ClsUbicacion
#Region "Definiciones"
        ' Herencia
        Inherits ClsCBObjetoPan
        ' Constantes
        Private Const MCSTRNOMBRETABLA As String = ""
        ' Colecciones y objetos
        Private McolPaises As Collection = Nothing
        Private MobjPais As ClsPais = Nothing
        Private MobjNuevoPais As ClsPais = Nothing
        'Arrays de nombres
        Private MstrPaises As String() = Array.Empty(Of String)
        Private MstrDptos As String() = Array.Empty(Of String)
        Private MstrCiudades As String() = Array.Empty(Of String)
        '
        Private MstrIdPais As String = String.Empty
        Private MstrNomPais As String = String.Empty
        Private MbytIdDpto As Byte = 0
        Private MstrNomDpto As String = String.Empty
        Private MshrIdCiud As Short = 0S
        Private MstrNomCiu As String = String.Empty
        '
        Private MstrIdPaisDpt As String = String.Empty
        Private MstrIdPaisCiu As String = String.Empty

        Private MbytIdDptoCiu As Byte = 0
        '
        Private MenuTipoAccion As EnuTipoAccionDef = EnuTipoAccionDef.None
#End Region
#Region "Constructores"
        Friend Sub New()
            HobjPadre = Nothing
            HblnEsCreable = False
            HblnEsSuprimible = False
            HblnEsAnulable = False
            HenuTipoObjeto = EnuModoInstanciaObjDef.None
            EnuPermisosObj = EnuPermisosDef.enuTodos
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
                Return EnuIdClasesPanDef.enuUbicacion
            End Get
        End Property

        Protected Overrides ReadOnly Property HstrNombreClase As String
            Get
                Return "Ubicacion"
            End Get
        End Property
#End Region
#Region "Clases de Propiedad"
        Friend Overrides ReadOnly Property ColPropiedades As Collection
            Get
                Return HcolPropiedades
            End Get
        End Property
#End Region
#Region "Otras Propiedades"
        Friend ReadOnly Property ColPaises() As Collection
            Get
                If McolPaises Is Nothing Then
                    Dim lobjPais As ClsPais
                    Dim ldtbPaises = FdtbLocalizacion(EnuTipoLocalizacionDef.enuPais)
                    McolPaises = New Collection
                    For Each ldrwPais As DataRow In ldtbPaises.Rows
                        lobjPais = New ClsPais(Me, ldrwPais)
                        lobjPais.SLeaValores(True)
                        McolPaises.Add(lobjPais, lobjPais.ObjIdPaisStr.ToString)
                    Next
                End If
                Return McolPaises
            End Get
        End Property
        Friend ReadOnly Property ObjPais(astrIdPais As String) As ClsPais
            Get
                McolPaises = ColPaises()
                If McolPaises.Contains(astrIdPais) Then
                    MobjPais = McolPaises(astrIdPais)
                Else
                    MobjPais.SVacie()
                End If
                Return MobjPais
            End Get
        End Property
        ''' <summary>
        ''' Devuekve el objeto Pais actual
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend ReadOnly Property ObjPais As ClsPais
            Get
                Return MobjPais
            End Get
        End Property
        ''' <summary>
        ''' Devuelve una matriz que contiene todos los nombres de los paises ordenados en forma ascendente.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Function FstrPaises() As String()
            SPuebleStrPaises()
            Return MstrPaises
        End Function
        Friend ReadOnly Property StrNombrePais(astrIdPais As String) As String
            Get
                Dim lstrNombre = String.Empty
                If Not String.IsNullOrEmpty(astrIdPais) Then
                    If ColPaises.Contains(astrIdPais) Then
                        Dim lobjPais As ClsPais = ColPaises(astrIdPais)
                        lstrNombre = lobjPais.ObjNombrePaisStr.ToString
                    End If
                End If
                Return lstrNombre
            End Get
        End Property
        Friend ReadOnly Property StrIdPais(astrNombrePais As String) As String
            Get
                If MstrNomPais <> astrNombrePais Then
                    MstrNomPais = astrNombrePais
                    Dim ldtbPaises = FdtbLocalizacion(EnuTipoLocalizacionDef.EnuPais)
                    Dim ldrwPaises() As DataRow = ldtbPaises.Select("Nombre = '" & astrNombrePais & "'")
                    If ldrwPaises.Count > 0 Then
                        MstrIdPais = ClsPanorama.FobjValorCampo(ldrwPaises(0)("IdPais"),
                                EnuTipoValor.EnuString)
                    Else
                        MstrIdPais = String.Empty
                    End If
                End If
                Return MstrIdPais
            End Get
        End Property
        Friend ReadOnly Property DtbPaises() As DataTable
            Get
                Dim ldtbPaises = FdtbLocalizacion(EnuTipoLocalizacionDef.enuPais)
                Return ldtbPaises
            End Get
        End Property
        Friend Function FstrDptos(astrIdPais As String) As String()
            If MstrIdPaisDpt <> astrIdPais Then
                MstrIdPaisDpt = astrIdPais
            End If
            SPuebleStrDptos()
            Return MstrDptos
        End Function
        Friend ReadOnly Property StrNombreDpto(astrIdPais As String,
                    abytIdDpto As Byte) As String
            Get
                If astrIdPais <> MstrIdPaisDpt Then
                    MstrIdPaisDpt = astrIdPais
                End If
                Dim ldtbDptos = FdtbLocalizacion(EnuTipoLocalizacionDef.enuDpto)
                Dim lstrCondicion As String = "IdDepartamento = " & abytIdDpto
                Dim lstrNombrDpto As String = String.Empty
                Dim ldrwDptos As DataRow() = ldtbDptos.Select(lstrCondicion)
                If ldrwDptos.Count > 0 Then
                    lstrNombrDpto = ClsPanorama.FobjValorCampo(ldrwDptos(0)("Nombre"), EnuTipoValor.enuString)
                End If
                Return lstrNombrDpto
            End Get
        End Property
        Friend ReadOnly Property BytIdDpto(astrIdPais As String, astrNombreDpto As String) As Byte
            Get
                If astrIdPais <> MstrIdPaisDpt OrElse astrNombreDpto <> MstrNomDpto Then
                    If astrIdPais <> MstrIdPaisDpt Then
                        MstrIdPaisDpt = astrIdPais
                    End If
                    MstrNomDpto = astrNombreDpto
                    Dim ldtbDpto = FdtbLocalizacion(EnuTipoLocalizacionDef.enuDpto)
                    Dim lstrCondicion As String = "Nombre = '" & astrNombreDpto & "'"
                    Dim ldrwDptos As DataRow() = ldtbDpto.Select(lstrCondicion)
                    If ldrwDptos.Count > 0 Then
                        MbytIdDpto = ClsPanorama.FobjValorCampo(ldrwDptos(0)("IdDepartamento"),
                                EnuTipoValor.EnuByte)
                    Else
                        MbytIdDpto = 0
                    End If
                End If
                Return MbytIdDpto
            End Get
        End Property
        Friend Function FstrCiudades(astrIdPais As String, abytIdDpto As Byte) As String()
            If astrIdPais <> MstrIdPaisCiu OrElse abytIdDpto <> MbytIdDptoCiu Then
                MstrIdPaisCiu = astrIdPais
                MbytIdDptoCiu = abytIdDpto
            End If
            SPuebleStrCiudades()
            Return MstrCiudades
        End Function
        Friend ReadOnly Property StrNombreCiudad(astrIdPais As String, abytIdDpto As Byte,
                ashrIdCiudad As Short) As String
            Get
                If MstrIdPaisCiu <> astrIdPais OrElse MbytIdDptoCiu <> abytIdDpto Then
                    MstrIdPaisCiu = astrIdPais
                    MbytIdDptoCiu = abytIdDpto
                End If
                Dim ldtbCiudades = FdtbLocalizacion(EnuTipoLocalizacionDef.enuCiudad)
                Dim lstrCondicion As String = "IdPais = '" & astrIdPais & "' AND IdDepartamento = " & abytIdDpto &
                        " AND IdCiudad = " & ashrIdCiudad
                Dim lstrNombrCiudad As String = String.Empty
                Dim ldrwCiudades As DataRow() = ldtbCiudades.Select(lstrCondicion)
                If ldrwCiudades.Count > 0 Then
                    lstrNombrCiudad = ClsPanorama.FobjValorCampo(ldrwCiudades(0)("Nombre"),
                            EnuTipoValor.enuString)
                End If
                Return lstrNombrCiudad
            End Get
        End Property
        Friend ReadOnly Property ShrIdCiudad(astrIdPais As String, abytIdDpto As Byte,
                astrNombreCiudad As String) As Short
            Get
                If MstrIdPaisCiu <> astrIdPais OrElse MbytIdDptoCiu <> abytIdDpto OrElse
                        MstrNomCiu <> astrNombreCiudad Then
                    If MstrIdPaisCiu <> astrIdPais OrElse MbytIdDptoCiu <> abytIdDpto Then
                        MstrIdPaisCiu = astrIdPais
                        MbytIdDptoCiu = abytIdDpto
                    End If
                    MstrNomCiu = astrNombreCiudad
                    Dim ldtbCiudades = FdtbLocalizacion(EnuTipoLocalizacionDef.enuCiudad)
                    Dim lstrCondicion As String = "IdPais = '" & astrIdPais & "' AND IdDepartamento = " & abytIdDpto &
                            " AND Nombre = '" & astrNombreCiudad & "'"
                    Dim ldrwCiudades As DataRow() = ldtbCiudades.Select(lstrCondicion)
                    If ldrwCiudades.Count > 0 Then
                        MshrIdCiud = ClsPanorama.FobjValorCampo(ldrwCiudades(0)("IdCiudad"),
                                EnuTipoValor.EnuShort)
                    Else
                        MshrIdCiud = 0
                    End If
                End If
                Return MshrIdCiud
            End Get
        End Property
        Friend ReadOnly Property DtbDptos(astrIdPais As String) As DataTable
            Get
                If MstrIdPaisDpt <> astrIdPais Then
                    MstrIdPaisDpt = astrIdPais
                End If
                Dim ldtbDptos = FdtbLocalizacion(EnuTipoLocalizacionDef.enuDpto)
                Return ldtbDptos
            End Get
        End Property
        Friend ReadOnly Property DtbCiudades(astrIdPais As String, abytIdDpto As Byte) As DataTable
            Get
                If MstrIdPaisCiu <> astrIdPais OrElse MbytIdDptoCiu <> abytIdDpto Then
                    MstrIdPaisCiu = astrIdPais
                    MbytIdDptoCiu = abytIdDpto
                End If
                Dim ldtbCiudades = FdtbLocalizacion(EnuTipoLocalizacionDef.enuCiudad)
                Return ldtbCiudades
            End Get
        End Property
        Friend Property EnuTipoAccion As EnuTipoAccionDef
            Get
                Return MenuTipoAccion
            End Get
            Set(value As EnuTipoAccionDef)
                MenuTipoAccion = value
                If MenuTipoAccion <> EnuTipoAccionDef.None Then
                    Select Case MenuTipoAccion
                        Case EnuTipoAccionDef.enuNuevoPais
                            '
                        Case EnuTipoAccionDef.enuModifPais
                            ObjPais.EnuTipoAccion = EnuTipoAccionDef.enuModifPais
                        Case Else
                            ObjPais.EnuTipoAccion = MenuTipoAccion
                    End Select
                End If
            End Set
        End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
        Protected Friend Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                Select Case MenuTipoAccion
                    Case EnuTipoAccionDef.enuNuevoPais
                        SAdicionePais()
                    Case EnuTipoAccionDef.enuModifPais
                        MobjPais.SActualice(ablnExigeRequeridos)
                    Case Else
                        MobjPais.SActualice(ablnExigeRequeridos)
                End Select
            End If
        End Sub

        Friend Overrides Function FblnEstaVacioOrigenDatos() As Boolean
            Return (DtbPaises.Rows.Count = 0)
        End Function

        Friend Overrides Sub SRefresqueObj()
            McolPaises = ColPaises()
            Select Case MenuTipoAccion
                Case EnuTipoAccionDef.enuNuevoPais
                    If Not IsNothing(MobjNuevoPais) Then
                        MobjNuevoPais = Nothing
                    End If
                Case EnuTipoAccionDef.None
                    '
                Case Else
                    MobjPais.SNormaliceEstado(True)
                    STengoCambios(False, False, False)
            End Select
            MenuTipoAccion = EnuTipoAccionDef.None
        End Sub

        Private Sub SRefresqueObjeto()
            Dim lstrIdPaisActual As String
            Dim lbytIdDptoActual As Byte = 0
            Dim lshrIdCiudadActual As Short = 0
            Dim lshrIdBarrioActual As Short = 0
            Dim lobjDptoActual As ClsDepartamento
            Dim lobjCiudadActual As ClsCiudad
            If Not IsNothing(MobjPais) Then
                lstrIdPaisActual = MobjPais.ObjIdPaisStr.ObjValorPro
                lobjDptoActual = ObjPais.ObjDpto
                If Not IsNothing(lobjDptoActual) Then
                    lbytIdDptoActual = lobjDptoActual.ObjIdDptoByt.ObjValorPro
                End If
                If lbytIdDptoActual > 0 Then
                    lobjCiudadActual = lobjDptoActual.ObjCiudad
                    If Not IsNothing(lobjCiudadActual) Then
                        lshrIdCiudadActual = lobjCiudadActual.ObjIdCiudadShr.ObjValorPro
                    End If
                End If
                McolPaises = ColPaises()
                If Not String.IsNullOrEmpty(lstrIdPaisActual) Then
                    MobjPais = McolPaises(lstrIdPaisActual)
                    If lbytIdDptoActual > 0 Then
                        MobjPais.SRefresqueObj()
                        lobjDptoActual = MobjPais.ObjDpto(lbytIdDptoActual)
                        If lshrIdCiudadActual > 0 Then
                            lobjDptoActual.SRefresqueObj()
                            lobjCiudadActual = lobjDptoActual.ObjCiudad(lshrIdCiudadActual)
                            If lshrIdBarrioActual > 0 Then
                                lobjCiudadActual.SRefresqueObj()
                            End If
                        End If
                    End If
                End If
            End If
        End Sub
#End Region
#Region "Procedimientos y funciones del objeto"
        Private Sub SPuebleStrPaises()
            Dim i As Short = -1
            If Not IsNothing(MstrPaises) Then
                Array.Clear(MstrPaises, 0, MstrPaises.Length)
            End If
            Dim ldtbPaises = FdtbLocalizacion(EnuTipoLocalizacionDef.enuPais)
            If ldtbPaises.Rows.Count > 0 Then
                ReDim MstrPaises(0 To ldtbPaises.Rows.Count - 1)
                For Each ldrwPais As DataRow In ldtbPaises.Rows
                    i += 1
                    MstrPaises(i) = ldrwPais("Nombre")
                Next
            End If
        End Sub
        Private Sub SPuebleStrDptos()
            Dim i As Short = -1
            If Not IsNothing(MstrDptos) Then
                MstrDptos = {}
            End If
            Dim ldtbDptos = FdtbLocalizacion(EnuTipoLocalizacionDef.enuDpto)
            If ldtbDptos.Rows.Count > 0 Then
                ReDim MstrDptos(0 To ldtbDptos.Rows.Count - 1)
                For Each ldrwDpto As DataRow In ldtbDptos.Rows
                    i += 1
                    MstrDptos(i) = ldrwDpto("Nombre")
                Next
            End If
        End Sub
        Private Sub SPuebleStrCiudades()
            Dim i As Short = -1
            If Not IsNothing(MstrCiudades) Then
                Array.Clear(MstrCiudades, 0, MstrCiudades.Length)
            End If
            Dim ldtbCiud = FdtbLocalizacion(EnuTipoLocalizacionDef.enuCiudad)
            If ldtbCiud.Rows.Count > 0 Then
                ReDim MstrCiudades(0 To ldtbCiud.Rows.Count - 1)
                For Each ldrwCiudad As DataRow In ldtbCiud.Rows
                    i += 1
                    MstrCiudades(i) = ldrwCiudad("Nombre")
                Next
            End If
        End Sub
        Friend Function FdtbLocalizacion(aenuTipoLocalizacion As EnuTipoLocalizacionDef) _
                As DataTable
            Dim lstrPref = String.Empty, lstrFiltro As String
            Dim ldtbLocaliza As DataTable = Nothing
            If GshrIdAplicacion = 803 OrElse GshrIdAplicacion = 100 Then
                lstrPref = ""
            ElseIf GshrIdAplicacion = 999 Then
                lstrPref = "Tra"
            End If
            Dim lstrTablaPais = lstrPref & "PanTblPaises"
            Dim lstrTablaDpto = lstrPref & "PanTblDepartamentos"
            Dim lstrTablaCiud = lstrPref & "PanTblCiudades"
            Select Case aenuTipoLocalizacion
                Case EnuTipoLocalizacionDef.enuPais
                    ldtbLocaliza = ClsPanorama.FdtbDataTable(lstrTablaPais, {"*"}, {{"Nombre", "ASC"}}, "")
                Case EnuTipoLocalizacionDef.enuDpto
                    lstrFiltro = "IDPais = '" & MstrIdPaisDpt & "'"
                    ldtbLocaliza = ClsPanorama.FdtbDataTable(lstrTablaDpto, {"*"},
                                {{"Nombre", "ASC"}}, lstrFiltro)
                Case EnuTipoLocalizacionDef.enuCiudad
                    lstrFiltro = "IDPais = '" & MstrIdPaisCiu & "' AND IDDepartamento = " & MbytIdDptoCiu
                    ldtbLocaliza = ClsPanorama.FdtbDataTable(lstrTablaCiud, {"*"},
                                {{"Nombre", "ASC"}}, lstrFiltro)
            End Select
            Return ldtbLocaliza
        End Function
        Friend Function FblnEsAdicionablePais(aobjLlave() As Object) As Boolean
            Dim lblnEsAdicionable As Boolean = True
            If Not IsNothing(aobjLlave) Then
                Dim lstrIdPais As String = aobjLlave(0)
                Dim ldtbPaises = FdtbLocalizacion(EnuTipoLocalizacionDef.enuPais)
                lblnEsAdicionable = (ldtbPaises.Select("IdPais = '" & lstrIdPais & "'").Count = 0)
            End If
            Return lblnEsAdicionable
        End Function
        Private Sub SAdicionePais()
            MobjNuevoPais.SActualice(True)
            If IsNothing(McolPaises) Then
                McolPaises = ColPaises()
            End If
            MobjPais = MobjNuevoPais
            MobjPais.SLeaValores(True)
            McolPaises.Add(MobjPais, MobjPais.ObjIdPaisStr.ToString)
            MobjNuevoPais = Nothing
        End Sub
        Friend Function FobjNuevoPais() As ClsPais
            Dim ldtbPaises = FdtbLocalizacion(EnuTipoLocalizacionDef.enuPais)
            Dim ldrwPais As DataRow = ldtbPaises.NewRow
            MobjNuevoPais = New ClsPais(Me, ldrwPais)
            MobjNuevoPais.SCreeObj(Nothing)
            Return MobjNuevoPais
        End Function
        Friend Function FblnEliminoPais()
            Dim lblnSuprimio = False
            If MobjPais.FblnEsSuprimible() Then
                McolPaises.Remove(MobjPais.ObjIdPaisStr.ToString)
                lblnSuprimio = MobjPais.FblnSuprimio()
                If lblnSuprimio Then
                    If ColPaises.Count > 0 Then
                        MobjPais = McolPaises(1)
                    Else
                        MobjPais = Nothing
                    End If
                End If
                SRefresqueObjeto()
            End If
            Return lblnSuprimio
        End Function
#End Region
#Region "Validacion"
        Friend Sub SGenereEventoNot(astrMens As String, aenuOrigen As EnuIdMens,
                aenuSeveridadNot As EnuSeveridadNot)
            SLevanteEventoNot(astrMens, "", aenuOrigen, aenuSeveridadNot)
        End Sub
        Friend Overrides Function FblnNotificaOk(aenuIdOrigenNot As EnuIdMens) As Boolean
            Dim lblnOk As Boolean
            Select Case aenuIdOrigenNot
                Case EnuIdMens.EnuIdPais
                    If MenuTipoAccion = EnuTipoAccionDef.enuNuevoPais Then
                        lblnOk = MobjNuevoPais.ObjIdPaisStr.BlnEsValido
                    Else
                        lblnOk = MobjPais.ObjIdPaisStr.BlnEsValido
                    End If
                Case EnuIdMens.EnuNomPais
                    If MenuTipoAccion = EnuTipoAccionDef.enuNuevoPais Then
                        lblnOk = MobjNuevoPais.ObjNombrePaisStr.BlnEsValido
                    Else
                        lblnOk = MobjPais.ObjNombrePaisStr.BlnEsValido
                    End If
                Case EnuIdMens.EnuIdDpto
                    If MenuTipoAccion = EnuTipoAccionDef.enuNuevoDpto Then
                        lblnOk = MobjPais.ObjNuevoDpto.ObjIdDptoByt.BlnEsValido
                    Else
                        lblnOk = MobjPais.ObjDpto.ObjIdDptoByt.BlnEsValido
                    End If
                Case EnuIdMens.EnuNomDpto
                    If MenuTipoAccion = EnuTipoAccionDef.enuNuevoDpto Then
                        lblnOk = MobjPais.ObjNuevoDpto.ObjNombreDptoStr.BlnEsValido
                    Else
                        lblnOk = MobjPais.ObjDpto.ObjNombreDptoStr.BlnEsValido
                    End If
                Case EnuIdMens.EnuIdCiudad
                    If MenuTipoAccion = EnuTipoAccionDef.enuNuevoCiud Then
                        lblnOk = MobjPais.ObjDpto.ObjNuevoCiudad.ObjIdCiudadShr.BlnEsValido
                    Else
                        lblnOk = MobjPais.ObjDpto.ObjCiudad.ObjIdCiudadShr.BlnEsValido
                    End If
                Case EnuIdMens.EnuNomCiudad
                    If MenuTipoAccion = EnuTipoAccionDef.enuNuevoCiud Then
                        lblnOk = MobjPais.ObjDpto.ObjNuevoCiudad.ObjNombreCiudadStr.BlnEsValido
                    Else
                        lblnOk = MobjPais.ObjDpto.ObjCiudad.ObjNombreCiudadStr.BlnEsValido
                    End If
            End Select
            Return lblnOk
        End Function
#End Region
    End Class
End Namespace