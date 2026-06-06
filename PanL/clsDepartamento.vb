Namespace Ubicacion
    Friend Class ClsDepartamento
#Region "Definiciones"
        Inherits ClsCBObjetoPan
        ' Constantes
        Private Const MCSTRNOMBRETABLA As String = "PanTblDepartamentos"
        ' Variables
        Private ReadOnly MobjPadre As ClsPais = Nothing
        Private McolCiudades As Collection = Nothing
        Private MobjCiudad As ClsCiudad = Nothing
        Private MobjNuevoCiudad As ClsCiudad = Nothing
        '
        Private MenuTipoAccion As EnuTipoAccionDef = EnuTipoAccionDef.None
#End Region
#Region "Constructores"
        Friend Sub New(aobjPadre As ClsPais, adrwDpto As DataRow)
            HobjPadre = aobjPadre
            MobjPadre = aobjPadre
            HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
            HblnEsAnulable = False
            '
            DrwRegistroActual = adrwDpto
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
                Return EnuIdClasesPanDef.enuDepartamento
            End Get
        End Property
        Protected Overrides ReadOnly Property HstrNombreClase As String
            Get
                Return "Departamento"
            End Get
        End Property
#End Region
#Region "Propiedades Prop"
        Friend ReadOnly Property ObjIdPaisDptoStr As New ClsIdPaisDptoStr(Me)
        Friend ReadOnly Property ObjIdDptoByt As New ClsIdDptoByt(Me)
        Friend ReadOnly Property ObjNombreDptoStr As New ClsNombreDptoStr(Me)
        Friend Overrides ReadOnly Property ColPropiedades As Collection
            Get
                If HcolPropiedades.Count = 0 Then
                    HcolPropiedades.Add(ObjIdDptoByt)
                    HcolPropiedades.Add(ObjIdPaisDptoStr)
                    HcolPropiedades.Add(ObjNombreDptoStr)
                End If
                Return HcolPropiedades
            End Get
        End Property
#End Region
#Region "Otras Propiedades"
        Friend ReadOnly Property ObjAbuelo As ClsUbicacion
            Get
                Return MobjPadre.ObjPadre
            End Get
        End Property
        Friend ReadOnly Property ColCiudades() As Collection
            Get
                If McolCiudades Is Nothing Then
                    Dim ldtbCiudades = ObjAbuelo.DtbCiudades(ObjIdPaisDptoStr.ObjValorPro,
                                ObjIdDptoByt.ObjValorPro)
                    McolCiudades = New Collection
                    Dim lobjCiudad As ClsCiudad
                    For Each ldrwCiudad As DataRow In ldtbCiudades.Rows
                        lobjCiudad = New ClsCiudad(Me, ldrwCiudad)
                        lobjCiudad.SLeaValores(True)
                        McolCiudades.Add(lobjCiudad, lobjCiudad.ObjIdCiudadShr.ToString)
                    Next
                End If
                Return McolCiudades
            End Get
        End Property
        Friend ReadOnly Property ObjCiudad(aShrIdCiudad As Short) As ClsCiudad
            Get
                If IsNothing(aShrIdCiudad) Then
                    Throw New ArgumentNullException(NameOf(aShrIdCiudad))
                End If
                McolCiudades = ColCiudades()
                If McolCiudades.Contains(aShrIdCiudad.ToString) Then
                    MobjCiudad = McolCiudades(aShrIdCiudad.ToString)
                Else
                    MobjCiudad = Nothing
                End If
                Return MobjCiudad
            End Get
        End Property
        Friend ReadOnly Property ObjCiudad() As ClsCiudad
            Get
                Return MobjCiudad
            End Get
        End Property
        Friend ReadOnly Property ObjNuevoCiudad As ClsCiudad
            Get
                Return MobjNuevoCiudad
            End Get
        End Property
        Public Property EnuTipoAccion() As EnuTipoAccionDef
            Get
                Return MenuTipoAccion
            End Get
            Set(value As EnuTipoAccionDef)
                MenuTipoAccion = value
                If MenuTipoAccion <> EnuTipoAccionDef.None Then
                    Select Case MenuTipoAccion
                        Case EnuTipoAccionDef.enuModifDpto
                            EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                        Case EnuTipoAccionDef.enuNuevoCiud
                            '
                        Case EnuTipoAccionDef.enuModifCiud
                            ObjCiudad.EnuTipoAccion = EnuTipoAccionDef.enuModifCiud
                        Case Else
                            ObjCiudad.EnuTipoAccion = MenuTipoAccion
                    End Select
                End If
            End Set
        End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
        Protected Friend Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
            Select Case MenuTipoAccion
                Case EnuTipoAccionDef.None, EnuTipoAccionDef.enuModifDpto
                    MyBase.SActualice(ablnExigeRequeridos)
                Case EnuTipoAccionDef.enuNuevoCiud
                    SAdicioneCiudad()
                Case EnuTipoAccionDef.enuModifCiud
                    MobjCiudad.SActualice(True)
                    MyBase.SActualice(ablnExigeRequeridos)
                Case Else
                    MobjCiudad.SActualice(True)
                    MyBase.SActualice(ablnExigeRequeridos)
            End Select
        End Sub
        Friend Overrides Function FblnEsCreable(aobjLlave() As Object) As Boolean
            Return MobjPadre.FblnEsAdicionableDpto(aobjLlave)
        End Function
        Friend Overrides Function FblnEsSuprimible() As Boolean
            Dim lblnEsSuprimible = FblnPermitidoSuprimir()
            If lblnEsSuprimible Then
                Dim lstrNombreColumnas As String() = {ObjIdPaisDptoStr.StrNombreCampoBD,
                        ObjIdDptoByt.StrNombreCampoBD}
                Dim lcolValoresRef As New Collection From {
                    {ObjIdPaisDptoStr.ToString, ObjIdPaisDptoStr.StrNombreCampoBD},
                    {ObjIdDptoByt.ToString, ObjIdDptoByt.StrNombreCampoBD}
                }
                Dim lstrTablasExcluir As String() = {SstrNombreTabla}
                lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg(lstrTablasExcluir,
                        lstrNombreColumnas, lcolValoresRef, True, True)
                If lblnEsSuprimible Then
                    lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg(lstrTablasExcluir,
                            lstrNombreColumnas, lcolValoresRef, True, False)
                End If
            End If
            Return lblnEsSuprimible
        End Function
        Friend Overrides Sub SRefresqueObj()
            McolCiudades = ColCiudades()
            Select Case MenuTipoAccion
                Case EnuTipoAccionDef.enuModifDpto
                    SNormaliceEstado(False)
                    MyBase.SRefresqueObj()
                Case EnuTipoAccionDef.enuNuevoCiud
                    If Not IsNothing(MobjNuevoCiudad) Then
                        MobjNuevoCiudad = Nothing
                    End If
                Case Else
                    If Not IsNothing(MobjCiudad) Then
                        MobjCiudad.SNormaliceEstado(True)
                    End If
                    STengoCambios(False, False, True)
            End Select
            MenuTipoAccion = EnuTipoAccionDef.None
        End Sub
#End Region
#Region "Procedimientos y funciones del objeto"
        Friend Function FblnEsValidoNombreDpto(astrNombreDpto As String) As Boolean
            Dim lblnEsValido As Boolean = DtbTablaColeccion.Rows.Count = 0
            If Not lblnEsValido Then
                Dim ldrwDptos As DataRow()
                ldrwDptos = DtbTablaColeccion.Select("Nombre = '" & astrNombreDpto & "'")
                lblnEsValido = (ldrwDptos.Count = 0)
            End If
            Return lblnEsValido
        End Function
        Friend Function FblnEsAdicionableCiudad(aobjLlave() As Object) As Boolean
            Dim lblnEsAdicionable As Boolean = True
            If Not IsNothing(aobjLlave) Then
                Dim lstrIdPais As String = CType(aobjLlave(0), String)
                Dim lbytIdDpto As Byte = CType(aobjLlave(1), Byte)
                If Not String.IsNullOrEmpty(lstrIdPais) AndAlso lbytIdDpto > 0 Then
                    If lstrIdPais = ObjIdPaisDptoStr.ObjValorPro AndAlso lbytIdDpto = ObjIdDptoByt.ObjValorPro Then
                        Dim lshrIdCiudad As Short = CType(aobjLlave(2), Short)
                        Dim ldtbCiudades As DataTable = ObjAbuelo.DtbCiudades(ObjIdPaisDptoStr.ObjValorPro,
                                ObjIdDptoByt.ObjValorPro)
                        If ldtbCiudades.Rows.Count > 0 Then
                            lblnEsAdicionable = (ldtbCiudades.Select("IdCiudad = " & lshrIdCiudad).Count = 0)
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
        Friend Sub SAdicioneCiudad()
            MobjNuevoCiudad.SActualice(True)
            If IsNothing(McolCiudades) Then
                McolCiudades = ColCiudades()
            End If
            If Not McolCiudades.Contains(MobjNuevoCiudad.ObjIdCiudadShr.ToString()) Then
                MobjCiudad = MobjNuevoCiudad
                McolCiudades.Add(MobjCiudad, MobjCiudad.ObjIdCiudadShr.ToString())
            End If
            MobjNuevoCiudad = Nothing
        End Sub
        Public Function FobjNuevoCiudad() As ClsCiudad
            Dim ldtbCiudades As DataTable = ObjAbuelo.DtbCiudades(ObjIdPaisDptoStr.ObjValorPro,
                ObjIdDptoByt.ObjValorPro)
            Dim ldrwCiudad As DataRow = ldtbCiudades.NewRow
            MobjNuevoCiudad = New ClsCiudad(Me, ldrwCiudad)
            MobjNuevoCiudad.SCreeObj(Nothing)
            MobjNuevoCiudad.ObjIdPaisCiudadStr.ObjValorPro = ObjIdPaisDptoStr.ObjValorPro
            MobjNuevoCiudad.ObjIdDptoCiudadByt.ObjValorPro = ObjIdDptoByt.ObjValorPro
            Return MobjNuevoCiudad
        End Function
#End Region
    End Class
#Region "Clases de propiedad"
    Friend Class ClsIdPaisDptoStr
        Inherits ClsCBPropiedad
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            HstrNombre = "IdPais"
            HenuTipoValor = EnuTipoValor.enuString
            HstrNombreCampoBd = "IdPais"
            HshrLongitud = 2
            HblnEsRequerido = True
            HblnEsLlave = True
            HbytPosicionLlave = 0
        End Sub
        Public Overrides Sub SValide()
            HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, HshrLongitud,
                        HshrLongitud, BlnEsRequerido)
        End Sub
        Public Overrides Function ToString() As String
            If Not IsNothing(ObjValorPro) Then
                Return HobjValorPro.ToString
            Else
                Return ""
            End If
        End Function
    End Class
    Friend Class ClsIdDptoByt
        Inherits ClsCBPropiedad
        Friend MobjPadre As ClsDepartamento = Nothing
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            MobjPadre = ObjPadre
            HstrNombre = "IdDpto"
            HenuTipoValor = EnuTipoValor.enuByte
            HstrNombreCampoBd = "IdDepartamento"
            HblnEsRequerido = True
            HblnEsLlave = True
            HbytPosicionLlave = 1
        End Sub
        Public Overrides Sub SValide()
            Dim lstrMens = String.Empty
            Dim lstrEntrada = HobjValorNew.ToString()
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCBYTMINDPTO,
                    GCBYTMAXDPTO, BlnEsRequerido, EnuTipoValor)
            If HblnEsValido Then
                If Not BlnLeyendoOrigen Then
                    If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                        Dim lobjVlrLlave As Object() = {MobjPadre.ObjIdPaisDptoStr.ToString(),
                                HobjValorNew}
                        If Not MobjPadre.FblnEsCreable(lobjVlrLlave) Then
                            lstrMens = "El Código del Departamento ingresado, ya existe!"
                            HblnEsValido = False
                        End If
                    End If
                End If
            Else
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso
                            Not String.IsNullOrEmpty(lstrEntrada) Then
                    lstrMens = "El Código del Departamento ingresado, no es valido!"
                End If
            End If
            If Not String.IsNullOrEmpty(lstrMens) Then
                Dim lobjAbuelo As ClsPais = MobjPadre.ObjPadre
                Dim lobjBisAbuelo As ClsUbicacion = lobjAbuelo.ObjPadre
                lobjBisAbuelo.SGenereEventoNot(lstrMens, EnuIdMens.EnuIdDpto,
                        EnuSeveridadNot.EnuDatoInvalido)
            End If
        End Sub
        Public Overrides Function ToString() As String
            If Not IsNothing(ObjValorPro) Then
                Return HobjValorPro.ToString
            Else
                Return ""
            End If
        End Function
    End Class
    Friend Class ClsNombreDptoStr
        Inherits ClsCBPropiedad
        Friend MobjPadre As ClsDepartamento = Nothing
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            MobjPadre = ObjPadre
            HstrNombre = "NombreDpto"
            HshrLongitud = 50
            HenuTipoValor = EnuTipoValor.enuString
            HstrNombreCampoBd = "Nombre"
            HblnEsRequerido = True
            HblnRegistrarLogCambio = True
        End Sub
        Public Overrides Sub SValide()
            Dim lstrMens = String.Empty
            HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 2, ShrLongitud, BlnEsRequerido)
            If HblnEsValido Then
                HobjValorNew = HobjValorNew.ToUpper()
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                    HblnEsValido = MobjPadre.FblnEsValidoNombreDpto(HobjValorNew)
                    If Not HblnEsValido Then
                        lstrMens = "El Nombre del Departamento ingresado, ya existe!"
                    End If
                End If
            Else
                Dim lentLargo = HobjValorNew().ToString().Length()
                If lentLargo < 2 OrElse lentLargo > HshrLongitud Then
                    lstrMens = "El Nombre del Departamento, debe tener una longitud " &
                                "entre 2 y " & ShrLongitud.ToString() & " letras!"
                End If
            End If
            If Not String.IsNullOrEmpty(lstrMens) Then
                Dim lobjAbuelo As ClsPais = MobjPadre.ObjPadre
                Dim lobjBisAbuelo As ClsUbicacion = lobjAbuelo.ObjPadre
                lobjBisAbuelo.SGenereEventoNot(lstrMens, EnuIdMens.EnuNomDpto,
                            EnuSeveridadNot.EnuDatoInvalido)
            End If
        End Sub
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