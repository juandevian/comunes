Namespace Ubicacion
    Friend Class ClsPais
#Region "Definiciones"
        Inherits ClsCBObjetoPan
        ' Constantes
        Private Const MCSTRNOMBRETABLA As String = "PanTblPaises"
        ' Variables
        Private ReadOnly MobjPadre As ClsUbicacion = Nothing
        Private McolDptos As Collection = Nothing
        Private MobjDpto As ClsDepartamento = Nothing
        Private MobjNuevoDpto As ClsDepartamento = Nothing
        '
        Private MenuTipoAccion As EnuTipoAccionDef = EnuTipoAccionDef.None
#End Region
#Region "Constructores"
        Friend Sub New(aobjPadre As ClsUbicacion, adrwPais As DataRow)
            HobjPadre = aobjPadre
            MobjPadre = aobjPadre
            HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
            HblnEsAnulable = False
            '
            DrwRegistroActual = adrwPais
            DtbTablaColeccion = DrwRegistroActual.Table
        End Sub
#End Region
#Region "Propiedades"
#Region "Propiedades indentificadoras"
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
                Return EnuIdClasesPanDef.enuPais
            End Get
        End Property
        Protected Overrides ReadOnly Property HstrNombreClase As String
            Get
                Return "País"
            End Get
        End Property
#End Region
#Region "Propiedades Prop"
        Friend ReadOnly Property ObjIdPaisStr As New ClsIdPaisStr(Me)
        Friend ReadOnly Property ObjNombrePaisStr As New ClsNombrePaisStr(Me)
        Friend Overrides ReadOnly Property ColPropiedades As Collection
            Get
                If HcolPropiedades.Count = 0 Then
                    HcolPropiedades.Add(ObjIdPaisStr)
                    HcolPropiedades.Add(ObjNombrePaisStr)
                End If
                Return HcolPropiedades
            End Get
        End Property
#End Region
#Region "Otras Propiedades"
        Friend ReadOnly Property ColDepartamentos() As Collection
            Get
                If McolDptos Is Nothing Then
                    Dim ldtbDptos = MobjPadre.DtbDptos(ObjIdPaisStr.ObjValorPro)
                    Dim lobjDpto As ClsDepartamento
                    McolDptos = New Collection
                    For Each ldrwDpto As DataRow In ldtbDptos.Rows
                        lobjDpto = New ClsDepartamento(Me, ldrwDpto)
                        lobjDpto.SLeaValores(True)
                        McolDptos.Add(lobjDpto, lobjDpto.ObjIdDptoByt.ToString)
                    Next
                End If
                Return McolDptos
            End Get
        End Property
        Friend ReadOnly Property ObjDpto(abytIdDpto As Byte) As ClsDepartamento
            Get
                McolDptos = ColDepartamentos()
                If McolDptos.Contains(abytIdDpto.ToString) Then
                    MobjDpto = McolDptos(abytIdDpto.ToString)
                Else
                    MobjDpto.SVacie()
                End If
                Return MobjDpto
            End Get
        End Property
        Friend ReadOnly Property ObjDpto As ClsDepartamento
            Get
                Return MobjDpto
            End Get
        End Property
        Friend ReadOnly Property ObjNuevoDpto As ClsDepartamento
            Get
                Return MobjNuevoDpto
            End Get
        End Property
        Public Property EnuTipoAccion As EnuTipoAccionDef
            Get
                Return MenuTipoAccion
            End Get
            Set(value As EnuTipoAccionDef)
                MenuTipoAccion = value
                If MenuTipoAccion <> EnuTipoAccionDef.None Then
                    Select Case MenuTipoAccion
                        Case EnuTipoAccionDef.enuModifPais
                            EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                        Case EnuTipoAccionDef.enuNuevoDpto
                            '
                        Case EnuTipoAccionDef.enuModifDpto
                            ObjDpto.EnuTipoAccion = EnuTipoAccionDef.enuModifDpto
                        Case Else
                            ObjDpto.EnuTipoAccion = MenuTipoAccion
                    End Select
                End If
            End Set
        End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
        Protected Friend Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
            Select Case menuTipoAccion
                Case enuTipoAccionDef.None, enuTipoAccionDef.enuModifPais
                    MyBase.sActualice(ablnExigeRequeridos)
                Case enuTipoAccionDef.enuNuevoDpto
                    sAdicioneDpto()
                Case enuTipoAccionDef.enuModifDpto
                    MobjDpto.sActualice(True)
                    MyBase.sActualice(ablnExigeRequeridos)
                Case Else
                    MobjDpto.sActualice(True)
                    MyBase.sActualice(ablnExigeRequeridos)
            End Select
        End Sub
        Friend Overrides Function FblnEsCreable(aobjLlave() As Object) As Boolean
            Return MobjPadre.FblnEsAdicionablePais(aobjLlave)
        End Function
        Friend Overrides Function FblnEsSuprimible() As Boolean
            Dim lblnEsSuprimible = FblnPermitidoSuprimir()
            If lblnEsSuprimible Then
                Dim lstrTablasExcluir As String() = {SstrNombreTabla}
                ' La condicion no debe incluir el Nombre de la columna
                Dim lstrCondicion As String = " = " & ObjIdPaisStr.ObjValorPro
                lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg(lstrTablasExcluir,
                        ObjIdPaisStr.StrNombreCampoBD, lstrCondicion, True, False)
                If lblnEsSuprimible Then
                    lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg(lstrTablasExcluir,
                            ObjIdPaisStr.StrNombreCampoBD, lstrCondicion, True, True)
                End If
            End If
            Return lblnEsSuprimible
        End Function
        Friend Overrides Sub SRefresqueObj()
            McolDptos = ColDepartamentos()
            Select Case MenuTipoAccion
                Case EnuTipoAccionDef.enuNuevoPais, EnuTipoAccionDef.None
                    '
                Case EnuTipoAccionDef.enuModifPais
                    SNormaliceEstado(False)
                    MyBase.SRefresqueObj()
                Case EnuTipoAccionDef.enuNuevoDpto
                    If Not IsNothing(MobjNuevoDpto) Then
                        MobjNuevoDpto = Nothing
                    End If
                Case Else
                    If Not IsNothing(MobjDpto) Then
                        MobjDpto.SNormaliceEstado(True)
                    End If
                    STengoCambios(False, False, True)
            End Select
            MenuTipoAccion = EnuTipoAccionDef.None
        End Sub
#End Region
#Region "Procedimientos y funciones del objeto"
        Friend Function FblnEsValidoNombrePais(astrNombrePais As String) As Boolean
            Dim lblnEsValido As Boolean = DtbTablaColeccion.Rows.Count = 0
            If Not lblnEsValido Then
                Dim ldrwPaises As DataRow()
                ldrwPaises = DtbTablaColeccion.Select("Nombre = '" & astrNombrePais & "'")
                lblnEsValido = (ldrwPaises.Count = 0)
            End If
            Return lblnEsValido
        End Function
        Friend Function FblnEsAdicionableDpto(aobjLlave() As Object) As Boolean
            Dim lblnEsAdicionable As Boolean = True
            If Not IsNothing(aobjLlave) Then
                Dim lstrIdPais As String = CType(aobjLlave(0), String)
                If Not String.IsNullOrEmpty(lstrIdPais) Then
                    If lstrIdPais = ObjIdPaisStr.ObjValorPro Then
                        Dim lbytIdDpto As Byte = CType(aobjLlave(1), Byte)
                        Dim ldtbDeptos As DataTable = MobjPadre.DtbDptos(ObjIdPaisStr.ObjValorPro)
                        If ldtbDeptos.Rows.Count > 0 Then
                            lblnEsAdicionable = (ldtbDeptos.Select("IdDepartamento = " & lbytIdDpto).Count = 0)
                        End If
                    Else
                        lblnEsAdicionable = False
                    End If
                Else
                    lblnEsAdicionable = False
                End If
            End If
            Return lblnEsAdicionable
        End Function
        Private Sub SAdicioneDpto()
            MobjNuevoDpto.SActualice(True)
            If IsNothing(McolDptos) Then
                McolDptos = ColDepartamentos()
            End If
            MobjDpto = MobjNuevoDpto
            If Not McolDptos.Contains(MobjDpto.ObjIdDptoByt.ToString()) Then
                McolDptos.Add(MobjDpto, MobjDpto.ObjIdDptoByt.ToString())
            End If
            MobjNuevoDpto = Nothing
        End Sub
        Friend Function FobjNuevoDpto() As ClsDepartamento
            Dim ldtbDptos As DataTable = MobjPadre.DtbDptos(ObjIdPaisStr.ObjValorPro)
            Dim ldrwDpto As DataRow = ldtbDptos.NewRow
            MobjNuevoDpto = New ClsDepartamento(Me, ldrwDpto)
            MobjNuevoDpto.SCreeObj(Nothing)
            MobjNuevoDpto.ObjIdPaisDptoStr.ObjValorPro = ObjIdPaisStr.ObjValorPro
            Return MobjNuevoDpto
        End Function
        Friend Function FblnSuprimioDpto()
            Dim lblnSuprimio As Boolean
            lblnSuprimio = MobjDpto.FblnEsSuprimible()
            If lblnSuprimio Then
                McolDptos.Remove(MobjDpto.ObjIdDptoByt.ToString)
                lblnSuprimio = MobjDpto.FblnSuprimio()
                If lblnSuprimio Then
                    If McolDptos.Count > 0 Then
                        MobjDpto = McolDptos(1)
                    Else
                        MobjDpto = Nothing
                    End If
                End If
            End If
            Return lblnSuprimio
        End Function
#End Region
    End Class
#Region "Clases de propiedad"
    Friend Class ClsIdPaisStr
        Inherits ClsCBPropiedad
        Private ReadOnly MobjPadre As ClsPais = Nothing
        Private Const MCSTRNOMBRECAMPOBD As String = "IdPais"
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            MobjPadre = aobjPadre
            HstrNombre = "IdPais"
            HenuTipoValor = EnuTipoValor.enuString
            HshrLongitud = 2
            HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
            HblnEsRequerido = True
            HblnEsLlave = True
            HbytPosicionLlave = 0
        End Sub
        Public Overrides Sub SValide()
            Dim lstrMens = String.Empty
            HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, HshrLongitud,
                    HshrLongitud, BlnEsRequerido)
            If HblnEsValido Then
                If Not BlnLeyendoOrigen Then
                    HobjValorNew = HobjValorNew.ToString().ToUpper()
                    If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                        Dim lobjLlavePrincipal As Object() = {HobjValorNew}
                        If Not MobjPadre.FblnEsCreable(lobjLlavePrincipal) Then
                            lstrMens = "El Código del País ingresado, ya existe!"
                            HblnEsValido = False
                        End If
                    End If
                End If
            Else
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso
                            Not String.IsNullOrEmpty(HobjValorNew) Then
                    If HobjValorNew.ToString().Length <> HshrLongitud Then
                        lstrMens = "El Código del País debe tener una longitud de dos letras!"
                    End If
                End If
            End If
            If Not String.IsNullOrEmpty(lstrMens) Then
                Dim lobjAbuelo As ClsUbicacion = MobjPadre.ObjPadre
                lobjAbuelo.SGenereEventoNot(lstrMens, EnuIdMens.EnuIdPais,
                            EnuSeveridadNot.EnuDatoInvalido)
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
    Friend Class ClsNombrePaisStr
        Inherits ClsCBPropiedad
        Private ReadOnly MobjPadre As ClsPais = Nothing
        Private Const MCSTRNOMBRECAMPOBD As String = "Nombre"
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            MobjPadre = aobjPadre
            HstrNombre = "NombrePais"
            HshrLongitud = 60
            HenuTipoValor = EnuTipoValor.enuString
            HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
            HblnEsRequerido = True
            HblnRegistrarLogCambio = True
        End Sub
        Public Overrides Sub SValide()
            Dim lstrMens = String.Empty
            HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 2,
                    ShrLongitud, BlnEsRequerido)
            If HblnEsValido Then
                If Not BlnLeyendoOrigen Then
                    HobjValorNew = HobjValorNew.ToString().ToUpper()
                    If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                        If HobjValorNew <> HobjValorOriginal Then
                            HblnEsValido = MobjPadre.FblnEsValidoNombrePais(HobjValorNew)
                        End If
                        If Not HblnEsValido Then
                            lstrMens = "El Nombre del País ingresado, ya existe!"
                        End If
                    End If
                End If
            Else
                Dim lentLargo = HobjValorNew().ToString().Length()
                If lentLargo < 2 OrElse lentLargo > HshrLongitud Then
                    lstrMens = "El Nombre del País, debe tener una longitud entre 2 " &
                                "y " & ShrLongitud.ToString() & " letras!"
                End If
            End If
            If Not String.IsNullOrEmpty(lstrMens) Then
                Dim lobjAbuelo As ClsUbicacion = MobjPadre.ObjPadre
                lobjAbuelo.SGenereEventoNot(lstrMens, EnuIdMens.EnuNomPais,
                                EnuSeveridadNot.EnuDatoInvalido)
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