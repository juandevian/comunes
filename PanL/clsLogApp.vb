Friend Class ClsLogApp
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "PanLogApp"
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto padre al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsPanorama, adrwObjeto As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsModificable = False
        HblnEsSuprimible = False
        HblnEsAnulable = False
        '
        DrwRegistroActual = adrwObjeto
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
            Return EnuIdClasesPanDef.enuLogAplicacion
        End Get
    End Property

    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Log Aplicación"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjIdAppLogShr As New ClsIdAppLogShr(Me)
    Friend ReadOnly Property ObjIdCarpetaLogShr As New ClsIdCarpetaLogShr(Me)
    Friend ReadOnly Property ObjIdCentroUtilLogShr As New ClsIdCentroUtilLogShr(Me)
    Friend ReadOnly Property ObjIdLogAppInt As New ClsIdLogAppInt(Me)
    Friend ReadOnly Property ObjIdObjetoStr As New ClsIdObjetoStr(Me)
    Friend ReadOnly Property ObjIdUsuarioLogStr As New ClsIdUsuarioLogStr(Me)
    Friend ReadOnly Property ObjNombreClaseStr As New ClsNombreClaseStr(Me)
    Friend ReadOnly Property ObjNombreCampoLogStr As New ClsNombreCampoLogStr(Me)
    Friend ReadOnly Property ObjTipoLogByt As New ClsTipoLogByt(Me)
    Friend ReadOnly Property ObjValorAnteriorStr As New ClsValorAnteriorStr(Me)
    Friend ReadOnly Property ObjValorNuevoStr As New ClsValorNuevoStr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjFechaCreacionDtm)
                HcolPropiedades.Add(ObjIdAppLogShr)
                HcolPropiedades.Add(ObjIdCarpetaLogShr)
                HcolPropiedades.Add(ObjIdCentroUtilLogShr)
                HcolPropiedades.Add(ObjIdLogAppInt)
                HcolPropiedades.Add(ObjIdObjetoStr)
                HcolPropiedades.Add(ObjIdUsuarioLogStr)
                HcolPropiedades.Add(ObjNombreClaseStr)
                HcolPropiedades.Add(ObjNombreCampoLogStr)
                HcolPropiedades.Add(ObjTipoLogByt)
                HcolPropiedades.Add(ObjValorAnteriorStr)
                HcolPropiedades.Add(ObjValorNuevoStr)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    '
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"

#End Region
#Region "Procedimientos del objeto"
    '
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsIdAppLogShr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As clsCBObjetoPan)
        MyBase.New(aobjPadre)
        hstrNombre = "IdApp"
        henuTipoValor = EnuTipoValor.enuShort
        hstrNombreCampoBd = "IdAplicacion"
        hblnEsRequerido = True
    End Sub

    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoNumero(HobjValorNew,
                EnuListaAplicaciones.EnuAdministrador, EnuListaAplicaciones.EnuOrionCop, BlnEsRequerido,
                EnuTipoValor.enuShort)
        HblnEsValido = lblnEsValido
    End Sub

    Public Overrides Function ToString() As String
        If IsNothing(objValorPro) Then
            Return ""
        Else
            Return hobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdCarpetaLogShr
    Inherits clsCBPropiedad

    Public Sub New(aobjPadre As clsCBObjetoPan)
        MyBase.New(aobjPadre)
        hstrNombre = "IdCarpeta"
        henuTipoValor = EnuTipoValor.enuShort
        hstrNombreCampoBd = "IdCarpeta"
        hblnEsRequerido = True
    End Sub

    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = clsPanorama.fblnEsValidoNumero(hobjValorNew, 0, Short.MaxValue,
                    blnEsRequerido, enuTipoValor)
        hblnEsValido = lblnEsValido
    End Sub

    Public Overrides Function ToString() As String
        If IsNothing(objValorPro) Then
            Return ""
        Else
            Return hobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdCentroUtilLogShr
    Inherits ClsCBPropiedad

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdCentroUtil"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = "IdCentroUtil"
        HblnEsRequerido = True
    End Sub

    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Short.MaxValue,
                    BlnEsRequerido, EnuTipoValor)
        HblnEsValido = lblnEsValido
        HblnEsValido = lblnEsValido
    End Sub

    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdLogAppInt
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdLogApp"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = "IdLogApp"
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 0
        HblnEsAutonumerico = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsLogApp = ObjPadre
        Dim lblnEsValido As Boolean
        If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso HobjValorNew = 0 Then
            lblnEsValido = True
        Else
            lblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, Integer.MinValue,
                    Integer.MaxValue, BlnEsRequerido, EnuTipoValor)
        End If
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdObjetoStr
    Inherits ClsCBPropiedad

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdObjeto"
        HshrLongitud = 50
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "IdObjeto"
    End Sub

    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = True
        If Not String.IsNullOrEmpty(HobjValorNew) Then
            If IsNumeric(HobjValorNew) Then
                HobjValorNew = HobjValorNew.ToString
            End If
            lblnEsValido = HobjValorNew.GetType.Name = "String" OrElse IsNumeric(HobjValorNew)
            If lblnEsValido Then
                lblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
            End If
        End If
        HblnEsValido = lblnEsValido
    End Sub

    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdUsuarioLogStr
    Inherits ClsCBPropiedad

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdUsuario"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "IdUsuario"
        HblnEsRequerido = True
    End Sub

    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud,
                BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub

    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsNombreClaseStr
    Inherits ClsCBPropiedad

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "NombreClase"
        HshrLongitud = 50
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "NombreClase"
        HblnEsRequerido = True
    End Sub

    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud,
                BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub

    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsNombreCampoLogStr
    Inherits ClsCBPropiedad

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Nombre"
        HshrLongitud = 250
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "Nombre"
        HblnEsRequerido = True
    End Sub

    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud,
                BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub

    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function

End Class
Friend Class ClsTipoLogByt
    Inherits ClsCBPropiedad

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TipoLog"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = "IdTblTipoLog"
        HblnEsRequerido = True
    End Sub

    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuTipoLogDef.enuAccion,
                EnuTipoLogDef.enuCambio, HblnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub

    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function

End Class
Friend Class ClsValorAnteriorStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "ValorAnterior"
        HshrLongitud = 100
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "ValorAnterior"
    End Sub

    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = True
        If Not String.IsNullOrEmpty(HobjValorNew) Then
            If IsNumeric(HobjValorNew) Then
                HobjValorNew = HobjValorNew.ToString
            End If
            lblnEsValido = HobjValorNew.GetType.Name = "String"
            If lblnEsValido Then
                If HobjValorNew.ToString.Length > 100 Then
                    HobjValorNew = HobjValorNew.ToString.Substring(0, 99)
                End If
                lblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
            End If
            If lblnEsValido Then
                lblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
            End If
        End If
        HblnEsValido = lblnEsValido
    End Sub

    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsValorNuevoStr
    Inherits ClsCBPropiedad

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "ValorNuevo"
        HshrLongitud = 100
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "ValorNuevo"
    End Sub

    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = True
        If Not String.IsNullOrEmpty(HobjValorNew) Then
            If IsNumeric(HobjValorNew) Then
                HobjValorNew = HobjValorNew.ToString
            End If
            lblnEsValido = HobjValorNew.GetType.Name = "String"
            If lblnEsValido Then
                If HobjValorNew.ToString.Length > 100 Then
                    HobjValorNew = HobjValorNew.ToString.Substring(0, 99)
                End If
                lblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
            End If
        End If
        HblnEsValido = lblnEsValido
    End Sub

    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
#End Region