Friend Class ClsCarpetaUsuario
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "PanCarpetasUsuarios"
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto CarpetaUsuario.
    ''' </summary>
    ''' <param name="aenuModoInstanciaObj">Esta clase no sepuede instanciar como Navegable!</param>
    ''' <remarks>Si se instancia como un objeto navegable, se crea un datatable que contiene las columnas de
    ''' la llave con las llaves de todos los objetos y queda a la espera de que se indique que objeto abrir.
    ''' Si se instancia como un objeto único, queda a la espera de recibir el valor de los campos de la llave 
    ''' para abrir dicho objeto. </remarks>
    Public Sub New(aenuModoInstanciaObj As enuModoInstanciaObjDef)
        If aenuModoInstanciaObj = enuModoInstanciaObjDef.enuDeColeccion Then
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        ElseIf aenuModoInstanciaObj = enuModoInstanciaObjDef.enuNavegable Then
            Throw New ErrorInesperadoPanLException("Esta Clase no se puede instanciar como Navegable!")
        End If
        Dim lstrCamposSelect As String() = {"*"}
        HobjPadre = Nothing
        hblnEsCreable = False
        hblnEsModificable = False
        HblnEsSuprimible = False
        hblnEsAnulable = False
        henuTipoObjeto = enuModoInstanciaObjDef.enuUnico

        hcolTablas.Add(MCSTRNOMBRETABLA)
        hcolCamposSelect.Add(lstrCamposSelect)
        henuTipoPermiso = EnuPermisosDef.enuHeredado
    End Sub

    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto padre al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwCarpetaUsuario">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsUsuario, adrwCarpetaUsuario As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        '
        DrwRegistroActual = adrwCarpetaUsuario
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
            Return EnuIdClasesPanDef.enuCarpetaUsuario
        End Get
    End Property

    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Carpeta Usuario"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjIdCarpetaUsuarioShr As New ClsIdCarpetaUsuarioShr(Me)
    Friend ReadOnly Property ObjIdUsuarioCarpetaStr As New ClsIdUsuarioCarpetaStr(Me)
    Friend ReadOnly Property ObjNombreCarpetaStr As New ClsNombreCarpetaStr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjIdCarpetaUsuarioShr)
                HcolPropiedades.Add(ObjIdUsuarioCarpetaStr)
                HcolPropiedades.Add(ObjNombreCarpetaStr)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return objIdCarpetaUsuarioShr.ToString & "-" & objIdUsuarioCarpetaStr.ToString
        End Get
    End Property
#End Region
#Region "Procedimientos del objeto"
    '
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsIdCarpetaUsuarioShr
    Inherits clsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCarpeta"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdCarpeta"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 1
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean
        lblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                BlnEsRequerido, EnuTipoValor)
        HblnEsValido = lblnEsValido
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(objValorPro) Then
            Return ""
        Else
            Return hobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsIdUsuarioCarpetaStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdUsuario"
    Public Sub New(aobjPadre As clsCBObjetoPan)
        MyBase.New(aobjPadre)
        hstrNombre = "IdUsuario"
        hshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 0
    End Sub

    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean
        lblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class

Friend Class ClsNombreCarpetaStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "NombreCarpeta"

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "NombreCarpeta"
        HshrLongitud = 50
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        hblnEsRequerido = True
        hblnRegistrarLogCambio = True
    End Sub

    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = clsPanorama.fblnEsValidoString(hobjValorNew, 3,
                shrLongitud, blnEsRequerido)
        hblnEsValido = lblnEsValido
    End Sub

    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return hobjValorPro.ToString
    End Function
End Class
#End Region