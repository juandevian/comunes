Friend Class ClsOrigenInstancia
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "PanOrigenesInstancias"
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto padre al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwOrigenInstancia">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsAplicacion, adrwOrigenInstancia As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsSuprimible = False
        HblnEsAnulable = False
        '
        DrwRegistroActual = adrwOrigenInstancia
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
            Return EnuIdClasesPanDef.enuOrigenInstancia
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Origen Instancia"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjEstaOrigenActivoBln As New ClsEstaOrigenActivoBln(Me)
    Friend ReadOnly Property ObjFechaActivacionDtm As New ClsFechaActivacionDtm(Me)
    Friend ReadOnly Property ObjIdAppOrigenShr As New ClsIdAppOrigenShr(Me)
    Friend ReadOnly Property ObjIdUsuarioActivoStr As New ClsIdUsuarioActivoStr(Me)
    Friend ReadOnly Property ObjNombreOrigenStr As New ClsNombreOrigenStr(Me)
    Friend ReadOnly Property ObjOrigenInstanciaByt As New ClsOrigenInstanciaByt(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjEstaOrigenActivoBln)
                HcolPropiedades.Add(ObjFechaActivacionDtm)
                HcolPropiedades.Add(ObjFechaCreacionDtm)
                HcolPropiedades.Add(ObjIdAppOrigenShr)
                HcolPropiedades.Add(ObjIdUsuarioActivoStr)
                HcolPropiedades.Add(ObjNombreOrigenStr)
                HcolPropiedades.Add(ObjOrigenInstanciaByt)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    Friend ReadOnly Property StrNombreTipoOrigen As String
        Get
            Dim lstrNombreTipoOrigen As String = ClsAdministrador.FstrNombreDatoConstantePan(
                    EnuGrupoConstantesPanDef.enuOrigenInstancia,
                    ObjOrigenInstanciaByt.ObjValorPro)
            Return lstrNombreTipoOrigen
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return Me.ObjNombreOrigenStr.ToString
        End Get
    End Property
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsEstaOrigenActivoBln
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "EstaActiva"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = "EstaActiva"
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsFechaActivacionDtm
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "FechaActivacion"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = "FechaActivacion"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim ldtmFechaMin As Date = DateAdd(DateInterval.Minute, -1, Date.Now)
        Dim ldtmFechaMax As Date = DateAdd(DateInterval.Minute, 1, Date.Now)
        Dim lblnEsValido As Boolean = IsDate(HobjValorNew)
        If lblnEsValido Then
            If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                lblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                        BlnEsRequerido)
            End If
        End If
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdAppOrigenShr
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
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoNumero(HobjValorNew,
                EnuListaAplicaciones.EnuAdministrador, EnuListaAplicaciones.EnuOrionCop,
                BlnEsRequerido, EnuTipoValor.enuShort)
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
Friend Class ClsIdUsuarioActivoStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdUsuarioActivo"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "IdUsuarioActivo"
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud, BlnEsRequerido)
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsNombreOrigenStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Nombre"
        HshrLongitud = 200
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "Nombre"
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 1
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, HshrLongitud, HblnEsRequerido)
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
Friend Class ClsOrigenInstanciaByt
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TipoOrigenInstancia"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = "IdTblTipoOrigenInstancia"
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew,
                EnuOrigenInstanciamientoDef.enuEstacionTrabajo,
                EnuOrigenInstanciamientoDef.enuMovil, HblnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            HobjValorPro = CType(HobjValorPro, Byte)
            Return HobjValorPro.ToString
        End If
    End Function
End Class
#End Region