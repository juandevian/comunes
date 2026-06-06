Namespace Ubicacion
    Friend Class ClsCiudad
#Region "Definiciones"
        Inherits ClsCBObjetoPan
        ' Constantes
        Private Const MCSTRNOMBRETABLA As String = "PanTblCiudades"
        ' Variables
        Private ReadOnly MobjPadre As ClsDepartamento = Nothing
        '
        Private MenuTipoAccion As EnuTipoAccionDef = EnuTipoAccionDef.None
#End Region
#Region "Constructores"
        Friend Sub New(aobjPadre As ClsDepartamento, adrwCiudad As DataRow)
            HobjPadre = aobjPadre
            MobjPadre = aobjPadre
            HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
            HblnEsAnulable = False
            '
            DrwRegistroActual = adrwCiudad
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
                Return EnuIdClasesPanDef.enuCiudad
            End Get
        End Property
        Protected Overrides ReadOnly Property HstrNombreClase As String
            Get
                Return "Ciudad"
            End Get
        End Property
#End Region
#Region "Propiedades Prop"
        Friend ReadOnly Property ObjIdPaisCiudadStr As New ClsIdPaisCiudadStr(Me)
        Friend ReadOnly Property ObjIdDptoCiudadByt As New ClsIdDptoCiudadByt(Me)
        Friend ReadOnly Property ObjIdCiudadShr As New ClsIdCiudadShr(Me)
        Friend ReadOnly Property ObjNombreCiudadStr As New ClsNombreCiudadStr(Me)
        Friend Overrides ReadOnly Property ColPropiedades As Collection
            Get
                If HcolPropiedades.Count = 0 Then
                    HcolPropiedades.Add(ObjIdCiudadShr)
                    HcolPropiedades.Add(ObjIdDptoCiudadByt)
                    HcolPropiedades.Add(ObjIdPaisCiudadStr)
                    HcolPropiedades.Add(ObjNombreCiudadStr)
                End If
                Return HcolPropiedades
            End Get
        End Property
#End Region
#Region "Otras Propiedades"
        Public Property EnuTipoAccion As EnuTipoAccionDef
            Get
                Return MenuTipoAccion
            End Get
            Set(value As EnuTipoAccionDef)
                MenuTipoAccion = value
                If MenuTipoAccion = EnuTipoAccionDef.enuModifCiud Then
                    EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                End If
            End Set
        End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
        Friend Overrides Function FblnEsCreable(aobjLlave() As Object) As Boolean
            Return MobjPadre.FblnEsAdicionableCiudad(aobjLlave)
        End Function
        Friend Overrides Function FblnEsSuprimible() As Boolean
            Dim lblnEsSuprimible = FblnPermitidoSuprimir()
            If lblnEsSuprimible Then
                Dim lstrNombreColumnas As String() = {ObjIdPaisCiudadStr.StrNombreCampoBD,
                        ObjIdDptoCiudadByt.StrNombreCampoBD, ObjIdCiudadShr.StrNombreCampoBD}
                Dim lcolValoresRef As New Collection From {
                    {ObjIdPaisCiudadStr.ToString, ObjIdPaisCiudadStr.StrNombreCampoBD},
                    {ObjIdDptoCiudadByt.ToString, ObjIdDptoCiudadByt.StrNombreCampoBD},
                    {ObjIdCiudadShr.ToString, ObjIdCiudadShr.StrNombreCampoBD}
                }
                Dim lstrTablasExcluir As String() = {SstrNombreTabla}
                lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg(lstrTablasExcluir,
                        lstrNombreColumnas, lcolValoresRef, True, False)
                If lblnEsSuprimible Then
                    lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg(lstrTablasExcluir,
                            lstrNombreColumnas, lcolValoresRef, True, True)
                End If
            End If
            Return lblnEsSuprimible
        End Function
        Friend Overrides Sub SRefresqueObj()
            If MenuTipoAccion = EnuTipoAccionDef.enuModifCiud Then
                SNormaliceEstado(False)
                MyBase.SRefresqueObj()
            End If
            MenuTipoAccion = EnuTipoAccionDef.None
        End Sub
#End Region
#Region "Procedimientos y funciones del objeto"
        Friend Function FblnEsValidoNombreCiudad(astrNombreCiudad As String) As Boolean
            Dim lblnEsValido = (DtbTablaColeccion.Rows.Count = 0)
            If Not lblnEsValido Then
                Dim ldrwCiudades As DataRow()
                ldrwCiudades = DtbTablaColeccion.Select("Nombre = '" & astrNombreCiudad & "'")
                lblnEsValido = (ldrwCiudades.Count = 0)
            End If
            Return lblnEsValido
        End Function
#End Region
    End Class
#Region "Clases de propiedad"
    Friend Class ClsIdPaisCiudadStr
        Inherits ClsCBPropiedad
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            HstrNombre = "IdPais"
            HenuTipoValor = EnuTipoValor.enuString
            HshrLongitud = 2
            HstrNombreCampoBd = "IdPais"
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
    Friend Class ClsIdDptoCiudadByt
        Inherits ClsCBPropiedad
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            HstrNombre = "IdDpto"
            HenuTipoValor = EnuTipoValor.enuByte
            HstrNombreCampoBd = "IdDepartamento"
            HblnEsRequerido = True
            HblnEsLlave = True
            HbytPosicionLlave = 1
        End Sub
        Public Overrides Sub SValide()
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCBYTMINDPTO, GCBYTMAXDPTO, BlnEsRequerido,
                    EnuTipoValor)
        End Sub
        Public Overrides Function ToString() As String
            If Not IsNothing(ObjValorPro) Then
                Return HobjValorPro.ToString
            Else
                Return ""
            End If
        End Function
    End Class
    Friend Class ClsIdCiudadShr
        Inherits ClsCBPropiedad
        Friend MobjPadre As ClsCiudad = Nothing
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            MobjPadre = aobjPadre
            HstrNombre = "IdCiudad"
            HenuTipoValor = EnuTipoValor.enuUShort
            HstrNombreCampoBd = "IdCiudad"
            HblnEsRequerido = True
            HblnEsLlave = True
            HbytPosicionLlave = 2
        End Sub
        Public Overrides Sub SValide()
            Dim lstrMens = String.Empty
            Dim lstrEntrada = HobjValorNew.ToString()
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCSHRMINCIUD,
                    GCSHRMAXCIUD, BlnEsRequerido, EnuTipoValor)
            If HblnEsValido Then
                If Not BlnLeyendoOrigen Then
                    If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                        Dim lobjVlrLlave As Object() =
                                {MobjPadre.ObjIdPaisCiudadStr.ToString(),
                                MobjPadre.ObjIdDptoCiudadByt.ObjValorPro, HobjValorNew}
                        If Not MobjPadre.FblnEsCreable(lobjVlrLlave) Then
                            lstrMens = "El Código de la Ciudad ingresado, ya existe!"
                            HblnEsValido = False
                        End If
                    End If
                End If
            Else
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso
                            Not String.IsNullOrEmpty(lstrEntrada) Then
                    lstrMens = "El Código de la Ciudad ingresado, no es válido!"
                End If
            End If
            If Not String.IsNullOrEmpty(lstrMens) Then
                Dim lobjAbuelo As ClsDepartamento = MobjPadre.ObjPadre
                Dim lobjBisAbuelo As ClsPais = lobjAbuelo.ObjPadre
                Dim lobjTataAbuelo As ClsUbicacion = lobjBisAbuelo.ObjPadre
                lobjTataAbuelo.SGenereEventoNot(lstrMens, EnuIdMens.EnuIdCiudad,
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
    Friend Class ClsNombreCiudadStr
        Inherits ClsCBPropiedad
        Friend MobjPadre As ClsCiudad = Nothing
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            MobjPadre = aobjPadre
            HstrNombre = "NombreCiudad"
            HshrLongitud = 50
            HenuTipoValor = EnuTipoValor.enuString
            HstrNombreCampoBd = "Nombre"
            HblnEsRequerido = True
            HblnRegistrarLogCambio = True
        End Sub
        Public Overrides Sub SValide()
            Dim lstrMens = String.Empty
            HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 2, ShrLongitud,
                    BlnEsRequerido)
            If HblnEsValido Then
                HobjValorNew = HobjValorNew.ToUpper()
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                    HblnEsValido = MobjPadre.FblnEsValidoNombreCiudad(HobjValorNew)
                    If Not HblnEsValido Then
                        lstrMens = "El Nombre de la Ciudad ingresado, ya existe!"
                    End If
                End If
            Else
                Dim lentLargo = HobjValorNew().ToString().Length()
                If lentLargo < 2 OrElse lentLargo > HshrLongitud Then
                    lstrMens = "El Nombre de la Ciudad, debe tener una longitud " &
                                    "entre 2 y " & ShrLongitud.ToString() & " letras!"
                End If
            End If
            If Not String.IsNullOrEmpty(lstrMens) Then
                Dim lobjAbuelo As ClsDepartamento = MobjPadre.ObjPadre
                Dim lobjBisAbuelo As ClsPais = lobjAbuelo.ObjPadre
                Dim lobjTataAbuelo As ClsUbicacion = lobjBisAbuelo.ObjPadre
                lobjTataAbuelo.SGenereEventoNot(lstrMens, EnuIdMens.EnuNomCiudad,
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