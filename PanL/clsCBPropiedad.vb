Public MustInherit Class ClsCBPropiedad
#Region "Definiciones"
#Region "Variables"
    Protected HstrMens As String = String.Empty
    Private MblnCambio As Boolean = False
    ' Variables de Propiedad
#Region "Propiedades autoimplementadas"
    Protected Property HstrNombre As String = String.Empty
    Protected Property HshrLongitud As Short = 0
    Protected Property HenuTipoValor As EnuTipoValor = EnuTipoValor.EnuNothingNull
    Protected Property HStrNombreCampoBd As String = String.Empty
    Protected Property HblnEsRequerido As Boolean = False
    Protected Property HblnEsLlave As Boolean = False
    Protected Property HstrOrdenIndice As String = "ASC"
    Protected Property HbytPosicionLlave As Byte = 0
    Protected Property HblnRegistrarLogCambio As Boolean = False
    Protected Property HblnEsValido As Boolean = False
    ' Indica que el campo asociado a la propiedad es autonumérico por BD 
    Protected Property HblnEsAutonumerico As Boolean = False
    Protected Property HobjValorOriginal As Object = Nothing
    Protected Property HobjValorPro As Object = Nothing
    Protected Property HobjValorNew As Object = Nothing
    Protected Friend Property BlnLeyendoOrigen As Boolean = False
    Public Property ObjPadre As ClsCBObjetoPan = Nothing
#End Region
#End Region
#Region "Eventos"
    ' Eventos
    Public MustOverride Sub SValide()
    Protected Event EvnPreSetValor As EventHandler(Of ClsPanEventArgs)
    Protected Event EvnPosCambio As EventHandler(Of ClsPanEventArgs)
    Protected Event EvnPosSetValor As EventHandler(Of ClsPanEventArgs)
    Private ReadOnly objArgEventoPan As New ClsPanEventArgs
    Public Event EvnNotifica As EventHandler(Of ClsNotiEventArgs)
#End Region
#End Region
#Region "Constructores"
    Protected Sub New(aobjPadre As ClsCBObjetoPan)
        ObjPadre = aobjPadre
    End Sub
#End Region
#Region "Propiedades"
    Public ReadOnly Property StrNombre() As String
        Get
            Return HstrNombre
        End Get
    End Property
    Protected Friend ReadOnly Property ShrLongitud() As Short
        Get
            Return HshrLongitud
        End Get
    End Property
    Friend ReadOnly Property EnuTipoValor() As EnuTipoValor
        Get
            Return HenuTipoValor
        End Get
    End Property
    Friend ReadOnly Property StrNombreCampoBD() As String
        Get
            Return HstrNombreCampoBd
        End Get
    End Property
    Protected Friend ReadOnly Property BlnEsRequerido As Boolean
        Get
            Return HblnEsRequerido
        End Get
    End Property
    Protected Friend ReadOnly Property BlnEsLlave As Boolean
        Get
            Return HblnEsLlave
        End Get
    End Property
    Protected Friend ReadOnly Property StrOrdenIdice As String
        Get
            Return HstrOrdenIndice
        End Get
    End Property
    Protected Friend ReadOnly Property BytPosLlave As Byte
        Get
            Return HbytPosicionLlave
        End Get
    End Property
    Friend ReadOnly Property BlnRegistrarLogCambio As Boolean
        Get
            Return HblnRegistrarLogCambio
        End Get
    End Property
    Friend ReadOnly Property BlnCambio() As Boolean
        Get
            Return MblnCambio
        End Get
    End Property
    '''<summary>
    ''' Obtiene o establece el valor de la propiedad.
    ''' </summary>
    ''' <returns>Devuelve un objeto con el valor de la propiedad</returns>
    ''' <remarks></remarks>
    ''' 
    Public Property ObjValorPro() As Object
        Get
            Return HobjValorPro
        End Get
        Set(value As Object)
            HobjValorNew = Nothing
            If IsNothing(value) OrElse IsDBNull(value) Then
                HobjValorNew = ClsPanorama.FobjValorNuloPropiedad(Me)
            Else
                HobjValorNew = value
                objArgEventoPan.BlnCancele = False
                RaiseEvent EvnPreSetValor(Me, objArgEventoPan)
                If objArgEventoPan.BlnCancele Then
                    Exit Property
                End If
            End If
            If BlnLeyendoOrigen Then
                HobjValorOriginal = HobjValorNew
            End If
            SValide()
            SAsigneValor()
            If Not ObjPadre.BlnVaciandoObjeto Then
                objArgEventoPan.BlnCancele = False
                RaiseEvent EvnPosSetValor(Me, objArgEventoPan)
            End If
            BlnLeyendoOrigen = False
        End Set
    End Property
    Private Sub SAsigneValor()
        If TypeOf HobjValorPro Is String Then
            If String.IsNullOrEmpty(HobjValorPro) Then HobjValorPro = Nothing
        End If
        If BlnLeyendoOrigen Then
            HobjValorPro = HobjValorNew
            If HblnEsValido Then
                objArgEventoPan.BlnCancele = False
                RaiseEvent EvnPosCambio(Me, objArgEventoPan)
            End If
            MblnCambio = False
        Else
            If IsNothing(HobjValorOriginal) OrElse ClsPanorama.FblnEsFechaNula(HobjValorOriginal) Then
                If HobjValorPro <> HobjValorNew Then
                    HobjValorPro = HobjValorNew
                    SAsigneCambio(True, True)
                    If HblnEsValido Then
                        objArgEventoPan.BlnCancele = False
                        RaiseEvent EvnPosCambio(Me, objArgEventoPan)
                    End If
                Else
                    If (HobjValorPro Is Nothing AndAlso IsNumeric(HobjValorNew) AndAlso
                            HobjValorNew = 0) OrElse (HobjValorPro Is Nothing AndAlso
                            TypeOf HobjValorNew Is String AndAlso String.IsNullOrEmpty(HobjValorNew)) Then
                        HobjValorPro = HobjValorNew
                        SAsigneCambio(True, True)
                        If HblnEsValido Then
                            objArgEventoPan.BlnCancele = False
                            RaiseEvent EvnPosCambio(Me, objArgEventoPan)
                        End If
                    End If
                End If
            Else
                If HobjValorNew <> HobjValorOriginal Then
                    SAsigneCambio(True, True)
                Else
                    If BlnCambio AndAlso HobjValorNew <> HobjValorPro Then
                        SAsigneCambio(True, True)
                    End If
                End If
                If HobjValorPro Is Nothing OrElse HobjValorPro <> HobjValorNew Then
                    HobjValorPro = HobjValorNew
                    objArgEventoPan.BlnCancele = False
                    RaiseEvent EvnPosCambio(Me, objArgEventoPan)
                End If
            End If
        End If
    End Sub
    Friend Overridable ReadOnly Property ObjValorOriginal As Object
        Get
            If IsNothing(HobjValorOriginal) Then
                If FblnEsPropiedadNumerica() Then
                    HobjValorOriginal = 0
                ElseIf TypeOf HobjValorOriginal Is Date Then
                    HobjValorOriginal = GCDTMFECHANULA
                ElseIf TypeOf HobjValorOriginal Is Boolean Then
                    HobjValorOriginal = False
                Else
                    HobjValorOriginal = String.Empty
                End If
            End If
            Return HobjValorOriginal
        End Get
    End Property
    Protected Friend Overridable ReadOnly Property ObjValorNuevo As Object
        Get
            If IsNothing(HobjValorPro) Then
                Return ""
            Else
                Return HobjValorPro
            End If
        End Get
    End Property
    Friend ReadOnly Property BlnEsValido() As Boolean
        Get
            Return HblnEsValido
        End Get
    End Property
    Friend ReadOnly Property BlnEsAutonumerico As Boolean
        Get
            Return HblnEsAutonumerico
        End Get
    End Property
#End Region
#Region "Procedimientos"
    ''' <summary>
    ''' Asigna un valor buleano que indica si la propiedad cambio su valor
    ''' </summary>
    ''' <param name="ablnCambio">Establece si la propiedad cambio o no.</param>
    ''' <param name="ablnPropagueAPadre">Indica si el cambio debe propagarse a la propiedad blnTengoCambio del
    ''' objeto Padre de esta propiedad</param>
    ''' <remarks></remarks>
    Friend Sub SAsigneCambio(ablnCambio As Boolean, ablnPropagueAPadre As Boolean)
        Dim lobjPadre As ClsCBObjetoPan = ObjPadre
        If ablnCambio <> MblnCambio Then
            MblnCambio = ablnCambio
            If ablnPropagueAPadre Then
                lobjPadre.STengoCambios(MblnCambio, False, False)
            End If
        End If
    End Sub
    Protected Friend Sub SAsigneValorOriginal(aobjvalorOriginal As Object)
        HobjValorOriginal = aobjvalorOriginal
    End Sub
    Friend Sub SVacie(ablnVaciandoObjeto As Boolean)
        If ablnVaciandoObjeto Then
            SAsigneCambio(False, False)
            HobjValorOriginal = ClsPanorama.FobjValorNuloPropiedad(Me)
        Else
            SAsigneCambio(False, True)
        End If
        HobjValorNew = ClsPanorama.FobjValorNuloPropiedad(Me)
        HobjValorPro = HobjValorNew
        HblnEsValido = False
        HstrMens = String.Empty
        SVaciePropiedad()
        objArgEventoPan.BlnCancele = False
        objArgEventoPan.BlnVaciandoObjeto = True
        RaiseEvent EvnPosSetValor(Me, objArgEventoPan)
        objArgEventoPan.BlnVaciandoObjeto = False
    End Sub
    Protected Overridable Sub SVaciePropiedad()
        '
    End Sub
    Protected Sub SReinicieEstado()
        MblnCambio = False
    End Sub
    Protected Friend Function FblnEsPropiedadNumerica() As Boolean
        Select Case HenuTipoValor
            Case EnuTipoValor.enuByte, EnuTipoValor.enuDecimal, EnuTipoValor.enuDouble,
                    EnuTipoValor.enuInteger, EnuTipoValor.enuLong, EnuTipoValor.enuSByte,
                    EnuTipoValor.enuShort, EnuTipoValor.enuSingle, EnuTipoValor.enuUInteger,
                    EnuTipoValor.enuULong, EnuTipoValor.enuUShort
                Return True
            Case Else
                Return False
        End Select
    End Function
    ' Manejo notificaciones
    Protected Sub SLevanteEveNot(astrMensEx As String, aenuIdMensNot As EnuIdMens,
            aenuSeveridadNot As EnuSeveridadNot)
        ObjPadre.SLevanteEventoNot(Me, HstrMens, astrMensEx, aenuIdMensNot, aenuSeveridadNot)
    End Sub
    Protected Sub SNotifiqueDatInv()
        ObjPadre.SLevanteEventoNot(Me, HstrMens, "", 0, EnuSeveridadNot.EnuDatoInvalido)
    End Sub
    Friend Overridable Function FblnNotiInfoOk(aenuIdMensNot As EnuIdMens) As Boolean
        Return True
    End Function
#End Region
End Class