Namespace Ubicacion
    Friend Class ClsBarrio
#Region "Definiciones"
        Inherits clsCBObjetoPan
        ' Constantes
        Private Const MCSTRNOMBRETABLA As String = "PanBarrios"
        '
        Private ReadOnly MobjPadre As ClsCiudad = Nothing
        Private MenuTipoAccion As EnuTipoAccionDef = EnuTipoAccionDef.None
#End Region
#Region "Constructores"
        Friend Sub New(aobjPadre As ClsCiudad, adrwBarrio As DataRow)
            HobjPadre = aobjPadre
            MobjPadre = aobjPadre
            HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
            HblnEsAnulable = False
            '
            DrwRegistroActual = adrwBarrio
            DtbTablaColeccion = DrwRegistroActual.Table
        End Sub
#End Region
#Region "Propiedades"
#Region "Propiedades identificadoras"
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
                Return EnuIdClasesPanDef.enuBarrio
            End Get
        End Property
        Protected Overrides ReadOnly Property HstrNombreClase As String
            Get
                Return "Barrio"
            End Get
        End Property
#End Region
#Region "Propiedades Prop"
        Friend ReadOnly Property ObjIdPaisBarrioStr As New ClsIdPaisBarrioStr(Me)
        Friend ReadOnly Property ObjIdDptoBarrioByt As New ClsIdDptoBarrioByt(Me)
        Friend ReadOnly Property ObjIdCiudadBarrioShr As New ClsIdCiudadBarrioShr(Me)
        Friend ReadOnly Property ObjIdBarrioShr As New ClsIdBarrioShr(Me)
        Friend ReadOnly Property ObjNombreBarrioStr As New ClsNombreBarrioStr(Me)
        Friend Overrides ReadOnly Property ColPropiedades As Collection
            Get
                If HcolPropiedades.Count = 0 Then
                    HcolPropiedades.Add(ObjIdPaisBarrioStr)
                    HcolPropiedades.Add(ObjIdDptoBarrioByt)
                    HcolPropiedades.Add(ObjIdCiudadBarrioShr)
                    HcolPropiedades.Add(ObjIdBarrioShr)
                    HcolPropiedades.Add(ObjNombreBarrioStr)
                End If
                Return HcolPropiedades
            End Get
        End Property
#End Region
#Region "Otras Propiedades"
        Public Property EnuTipoAccion() As EnuTipoAccionDef
            Get
                Return MenuTipoAccion
            End Get
            Set(value As EnuTipoAccionDef)
                MenuTipoAccion = value
                If MenuTipoAccion <> EnuTipoAccionDef.None Then
                    EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                End If
            End Set
        End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
        Protected Friend Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                GobjPanDat.SControleProcesoObj(True)
                Try
                    SNumereObj()
                    MyBase.SActualice(ablnExigeRequeridos)
                Catch ex As PanLException
                    Throw
                Catch ex As PanDatException
                    Throw
                Catch ex As ArgumentNullException
                    Throw
                Catch ex As Exception
                    Throw
                Finally
                    GobjPanDat.SControleProcesoObj(False)
                End Try
            Else
                MyBase.SActualice(ablnExigeRequeridos)
            End If
        End Sub
        Friend Overrides Function FblnEsCreable(aobjLlave() As Object) As Boolean
            Return MobjPadre.FblnEsAdicionableBarrio(aobjLlave)
        End Function
        Friend Overrides Function FblnEsSuprimible() As Boolean
            Dim lblnEsSuprimible = FblnPermitidoSuprimir()
            If lblnEsSuprimible Then
                Dim lstrNombreColumnas As String() = {ObjIdPaisBarrioStr.StrNombreCampoBD,
                        ObjIdDptoBarrioByt.StrNombreCampoBD, ObjIdCiudadBarrioShr.StrNombreCampoBD,
                        ObjIdBarrioShr.StrNombreCampoBD}
                Dim lcolValoresRef As New Collection From {
                    {ObjIdPaisBarrioStr.ToString, ObjIdPaisBarrioStr.StrNombreCampoBD},
                    {ObjIdDptoBarrioByt.ToString, ObjIdDptoBarrioByt.StrNombreCampoBD},
                    {ObjIdCiudadBarrioShr.ToString, ObjIdCiudadBarrioShr.StrNombreCampoBD},
                    {ObjIdBarrioShr.ToString, ObjIdBarrioShr.StrNombreCampoBD}
                }
                Dim lstrTablasExcluir As String() = {SstrNombreTabla}
                lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg(lstrTablasExcluir,
                        lstrNombreColumnas, lcolValoresRef, True, True)
                If GenuIdAplicacion <> EnuListaAplicaciones.EnuAdministrador AndAlso lblnEsSuprimible Then
                    lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg(lstrTablasExcluir,
                            lstrNombreColumnas, lcolValoresRef, True, False)
                End If
            End If
            Return lblnEsSuprimible
        End Function
        Friend Overrides Sub SRefresqueObj()
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                SNormaliceEstado(False)
                MyBase.SRefresqueObj()
                MenuTipoAccion = EnuTipoAccionDef.None
            End If
        End Sub
#End Region
#Region "Procedimientos del objeto"
        Friend Sub SNumereObj()
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                Dim lstrFiltro As String = ObjIdPaisBarrioStr.StrNombreCampoBD & " = '" &
                        ObjIdPaisBarrioStr.ObjValorPro & "' AND " &
                        ObjIdDptoBarrioByt.StrNombreCampoBD & " = " & ObjIdDptoBarrioByt.ObjValorPro &
                        " AND " & ObjIdCiudadBarrioShr.StrNombreCampoBD & " = " & ObjIdCiudadBarrioShr.ObjValorPro
                Dim lobjUltimoValor As Object = ClsPanorama.FobjUltimaIdNumericaObjeto(SstrNombreTabla,
                        ObjIdBarrioShr.StrNombreCampoBD, ObjIdBarrioShr.EnuTipoValor, lstrFiltro)
                ObjIdBarrioShr.ObjValorPro = lobjUltimoValor + 1
            End If
        End Sub
        Friend Function FblnEsValidoNombreBarrio(astrNombreBarrio As String) As Boolean
            Dim lblnEsValido = (DtbTablaColeccion.Rows.Count = 0)
            If Not lblnEsValido Then
                Dim ldrwBarrios As DataRow()
                ldrwBarrios = DtbTablaColeccion.Select("Nombre = '" & astrNombreBarrio & "'")
                lblnEsValido = (ldrwBarrios.Count = 0)
            End If
            Return lblnEsValido
        End Function
#End Region
    End Class
#Region "Clases de Propiedad"
    Friend Class ClsIdBarrioShr
        Inherits ClsCBPropiedad
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            HstrNombre = "IdBarrio"
            HenuTipoValor = EnuTipoValorDef.enuShort
            HstrNombreCampoBd = "IdBarrio"
            HblnEsRequerido = True
            HblnEsLlave = True
            HbytPosicionLlave = 3
        End Sub

        Public Overrides Sub SValide()
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                        BlnEsRequerido, EnuTipoValor)
        End Sub
        Public Overrides Function ToString() As String
            If IsNothing(ObjValorPro) Then
                Return ""
            Else
                Return HobjValorPro.ToString
            End If
        End Function
    End Class
    Friend Class ClsIdCiudadBarrioShr
        Inherits ClsCBPropiedad
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            HstrNombre = "IdCiudad"
            HenuTipoValor = EnuTipoValorDef.enuShort
            HstrNombreCampoBd = "IdCiudad"
            HblnEsRequerido = True
            HblnEsLlave = True
            HbytPosicionLlave = 2
        End Sub
        Public Overrides Sub SValide()
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCSHRMINCIUD, GCSHRMAXCIUD, BlnEsRequerido,
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
    Friend Class ClsIdDptoBarrioByt
        Inherits ClsCBPropiedad
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            HstrNombre = "IdDpto"
            HenuTipoValor = EnuTipoValorDef.enuByte
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
    Friend Class ClsIdPaisBarrioStr
        Inherits ClsCBPropiedad
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            HstrNombre = "IdPais"
            HenuTipoValor = EnuTipoValorDef.enuString
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
    Friend Class ClsNombreBarrioStr
        Inherits ClsCBPropiedad
        Friend MobjPadre As ClsBarrio = Nothing
        Public Sub New(aobjPadre As ClsCBObjetoPan)
            MyBase.New(aobjPadre)
            MobjPadre = aobjPadre
            HstrNombre = "NombreBarrio"
            HshrLongitud = 50
            HenuTipoValor = EnuTipoValorDef.enuString
            HstrNombreCampoBd = "Nombre"
            HblnEsRequerido = True
            HblnRegistrarLogCambio = True
        End Sub
        Private ReadOnly Property ObjTataAbuelo As ClsUbicacion
            Get
                Dim lobAbuelo As ClsCiudad = MobjPadre.ObjPadre
                Dim lobjBisAbuelo As ClsDepartamento = lobAbuelo.ObjPadre
                Dim lobjTaAbuelo As ClsPais = lobjBisAbuelo.ObjPadre
                Return lobjTaAbuelo.ObjPadre
            End Get
        End Property
        Public Overrides Sub SValide()
            HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 2,
                    ShrLongitud, BlnEsRequerido)
            Dim lstrMens = String.Empty
            If HblnEsValido Then
                HobjValorNew = HobjValorNew.ToUpper
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                    HblnEsValido = MobjPadre.FblnEsValidoNombreBarrio(HobjValorNew)
                    If Not HblnEsValido Then
                        lstrMens = "El Nombre del Barrio ingresado, ya existe!"
                    End If
                End If
            Else
                If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando AndAlso
                            Not String.IsNullOrEmpty(HobjValorNew) Then
                    Dim lentLargo = HobjValorNew().ToString().Length()
                    If lentLargo < 2 OrElse lentLargo > HshrLongitud Then
                        lstrMens = "El Nombre del Barrio, debe tener una longitud " &
                                    "entre 2 y " & ShrLongitud.ToString() & " letras!"
                    End If
                End If
            End If
            If Not String.IsNullOrEmpty(lstrMens) Then
                ObjTataAbuelo.SGenereEventoNot(lstrMens, EnuIdMens.EnuNomBarrio,
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