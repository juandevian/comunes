Friend Class ClsPermisoAccion
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "PanPermisosAcciones"
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto padre al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwPermisoAccion">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsPerfil, adrwPermisoAccion As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        '
        DrwRegistroActual = adrwPermisoAccion
        DtbTablaColeccion = DrwRegistroActual.Table
        With DtbTablaColeccion
            Dim ldclIndice As DataColumn() = { .Columns(ObjIdAppPermisoAccionShr.StrNombreCampoBD),
                    .Columns(ObjIdPerfilPermisoAccionShr.StrNombreCampoBD),
                    .Columns(ObjOrdinalPermisoAccionShr.StrNombreCampoBD)}
            .PrimaryKey = ldclIndice
        End With
        HenuTipoPermiso = EnuPermisosDef.enuHeredado
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
            Return EnuIdClasesPanDef.enuPermisoAccion
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Permiso Accion"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjAccionPermitidaBln As New ClsAccionPermitidaBln(Me)
    Friend ReadOnly Property ObjIdAccionPermisoAccionShr As New ClsIdAccionPermisoAcionShr(Me)
    Friend ReadOnly Property ObjIdAppPermisoAccionShr As New ClsIdAppPermisoAccionShr(Me)
    Friend ReadOnly Property ObjIdObjetoPermisoAccionShr As New ClsIdObjetoPermisoAccionShr(Me)
    Friend ReadOnly Property ObjIdPerfilPermisoAccionShr As New ClsIdPerfilPermisoAccionShr(Me)
    Friend ReadOnly Property ObjNivelPermisoAccionByt As New ClsNivelPermisoAccionByt(Me)
    Friend ReadOnly Property ObjNombreAccionPermisoAccionStr As New ClsNombreAccionPermisoAccionStr(Me)
    Friend ReadOnly Property ObjOrdinalPermisoAccionShr As New ClsOrdinalPermisoAccionShr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjAccionPermitidaBln)
                HcolPropiedades.Add(ObjIdAccionPermisoAccionShr)
                HcolPropiedades.Add(ObjIdAppPermisoAccionShr)
                HcolPropiedades.Add(ObjIdObjetoPermisoAccionShr)
                HcolPropiedades.Add(ObjIdPerfilPermisoAccionShr)
                HcolPropiedades.Add(ObjNivelPermisoAccionByt)
                HcolPropiedades.Add(ObjNombreAccionPermisoAccionStr)
                HcolPropiedades.Add(ObjOrdinalPermisoAccionShr)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdPerfilPermisoAccionShr.ToString & "-" & ObjIdObjetoPermisoAccionShr.ToString & "-" &
                    ObjIdAccionPermisoAccionShr.ToString
        End Get
    End Property
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsAccionPermitidaBln
    Inherits ClsCBPropiedad

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Permitido"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = "Permitida"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsIdAccionPermisoAcionShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdAccion"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = MCSTRNOMBRECAMPOBD
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = "IdAccion"
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue, True,
                EnuTipoValor.enuShort)
        HblnEsValido = lblnEsValido
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdAppPermisoAccionShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdAplicacion"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdApp"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 0
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsPermisoAccion = ObjPadre
        Dim lblnEsValido As Boolean
        If Not IsNothing(lobjPadre.ObjPadre) Then
            Dim lobjAbuelo As ClsPerfil = lobjPadre.ObjPadre
            lblnEsValido = (HobjValorNew = lobjAbuelo.ObjIdAppPerfilShr.ObjValorPro)
        Else
            lblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, EnuListaAplicaciones.EnuAdministrador,
                    EnuListaAplicaciones.EnuOrionCop, BlnEsRequerido, EnuTipoValor.enuShort)
        End If
        HblnEsValido = lblnEsValido
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdObjetoPermisoAccionShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdObjeto"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdObjeto"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Short.MaxValue, True,
                EnuTipoValor.enuShort)
        HblnEsValido = lblnEsValido
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdPerfilPermisoAccionShr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdPerfil"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = "IdPerfil"
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 1
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsPermisoAccion = ObjPadre
        Dim lobjAbuelo As ClsPerfil = lobjPadre.ObjPadre
        Dim lblnEsValido As Boolean
        If Not IsNothing(lobjAbuelo) Then
            lblnEsValido = HobjValorNew = lobjAbuelo.ObjIdPerfilShr.ObjValorPro
        Else
            lblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue, BlnEsRequerido,
                    EnuTipoValor.enuShort)
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
Friend Class ClsNivelPermisoAccionByt
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Nivel"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = "Nivel"
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoNumero(HobjValorNew,
                Byte.MinValue, Byte.MaxValue, BlnEsRequerido, EnuTipoValor)
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
Friend Class ClsNombreAccionPermisoAccionStr
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
        Dim lblnEsvalido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 3,
                ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsvalido
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsOrdinalPermisoAccionShr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Ordinal"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = "Ordinal"
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 2
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