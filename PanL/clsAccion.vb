Friend Class ClsAccion ' Ok Constructores
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "PanTblAcciones"
    ' Variables de modulo
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto padre al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwAccion">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsAplicacion, adrwAccion As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsCreable = False
        HblnEsModificable = False
        HblnEsSuprimible = False
        HblnEsAnulable = False
        '
        DrwRegistroActual = adrwAccion
        DtbTablaColeccion = DrwRegistroActual.Table
        HenuTipoPermiso = EnuPermisosDef.enuHeredado
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
            Return EnuIdClasesPanDef.enuAccion
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Acción"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjIdAccionShr As New ClsIdAccionShr(Me)
    Friend ReadOnly Property ObjIdAppShr As New ClsIdAppShr(Me)
    Friend ReadOnly Property ObjIdObjetoShr As New ClsIdObjetoShr(Me)
    Friend ReadOnly Property ObjNivelByt As New ClsNivelByt(Me)
    Friend ReadOnly Property ObjNombreAccionStr As New ClsNombreAccionStr(Me)
    Friend ReadOnly Property ObjOrdinalShr As New ClsOrdinalShr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjIdAccionShr)
                HcolPropiedades.Add(ObjIdAppShr)
                HcolPropiedades.Add(ObjIdObjetoShr)
                HcolPropiedades.Add(ObjNivelByt)
                HcolPropiedades.Add(ObjNombreAccionStr)
                HcolPropiedades.Add(ObjOrdinalShr)
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
    '
#End Region
#Region "Procedimientos del objeto"
    '
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsIdAccionShr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdAccion"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = "IdAccion"
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                BlnEsRequerido, EnuTipoValor.enuShort)
        Dim lobjPadre As ClsAccion = ObjPadre
        If HblnEsValido Then
            HblnEsValido = (lobjPadre.ObjIdObjetoShr.BlnEsValido)
        End If
    End Sub
    Private Sub EPosSetValor() Handles Me.EvnPosSetValor
        Dim lobjPadre As ClsAccion = ObjPadre
        If BlnEsValido Then
            lobjPadre.ObjIdAccionShr.SValide()
        End If
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsIdAppShr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdApp"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = "IdAplicacion"
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 0
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsAccion = ObjPadre
        Dim lobjAbuelo As ClsAplicacion = lobjPadre.ObjPadre
        HblnEsValido = (HobjValorNew = lobjAbuelo.ObjIdAppShr.ObjValorPro)
    End Sub
    Private Sub EPosSetValor() Handles Me.EvnPosSetValor
        Dim lobjPadre As ClsAccion = ObjPadre
        If BlnEsValido Then
            lobjPadre.ObjIdAccionShr.SValide()
        End If
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdObjetoShr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdObjeto"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = "IdObjeto"
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoNumero(HobjValorNew, Short.MinValue,
                Short.MaxValue, BlnEsRequerido, EnuTipoValor)
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
Friend Class ClsNivelByt
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Nivel"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = "Nivel"
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoNumero(HobjValorNew, Byte.MinValue,
                Byte.MaxValue, BlnEsRequerido, EnuTipoValor.enuByte)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsNombreAccionStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "NombreAccion"
        HshrLongitud = 60
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "NombreAccion"
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
Friend Class ClsOrdinalShr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Ordinal"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = "Ordinal"
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 1
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoNumero(HobjValorNew, Short.MinValue,
                Short.MaxValue, BlnEsRequerido, EnuTipoValor)
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