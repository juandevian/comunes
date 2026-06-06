Friend Class ClsPerfilUsuario
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "PanPerfilesUsuarios"
    '
    Private ReadOnly MobjPadre As ClsUsuario = Nothing
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto padre al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwPerfilUsuario">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsUsuario, adrwPerfilUsuario As DataRow)
        HobjPadre = aobjPadre
        MobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        '
        DrwRegistroActual = adrwPerfilUsuario
        DtbTablaColeccion = DrwRegistroActual.Table
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
            Return EnuIdClasesPanDef.enuPerfilUsuario
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Perfil Usuario"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjIdAppPerfilUsuarioShr As New ClsIdAppPerfilUsuarioShr(Me)
    Friend ReadOnly Property ObjIdPerfilUsuarioShr As New ClsIdPerfilUsuarioShr(Me)
    Friend ReadOnly Property ObjIdUsuarioPerfilStr As New ClsIdUsuarioPerfilStr(Me)
    Friend ReadOnly Property ObjNombreAppPerfilUsuarioStr As New ClsNombreAppPerfilUsuarioStr(Me)
    Friend ReadOnly Property ObjNombrePerfilUsuarioStr As New ClsNombrePerfilUsuarioStr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjIdAppPerfilUsuarioShr)
                HcolPropiedades.Add(ObjIdPerfilUsuarioShr)
                HcolPropiedades.Add(ObjIdUsuarioPerfilStr)
                HcolPropiedades.Add(ObjNombreAppPerfilUsuarioStr)
                HcolPropiedades.Add(ObjNombrePerfilUsuarioStr)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Friend Overrides Sub SCreeObj(aobjValorLlave() As Object)
        MyBase.SCreeObj(aobjValorLlave)
    End Sub
    Protected Friend Overrides Sub SLeaValores(ablnLeyendoOrigen As Boolean)
        MyBase.SLeaValores(ablnLeyendoOrigen)
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdUsuarioPerfilStr.ToString & "-" & ObjIdPerfilUsuarioShr.ToString
        End Get
    End Property
#End Region
#Region "Procedimientos del objeto"
    Friend Sub SDeterminePermisos()
        If MobjPadre.EnuPermisosObj > EnuPermisosDef.enuConsultar Then
            EnuPermisosObj = EnuPermisosDef.enuConCreModSup
        End If
    End Sub
#Region "Manejo PermisosAcciones"
    Friend Function FcolPermisosAcciones() As Collection
        Dim lcolPermisosAcciones As New Collection
        If ObjIdUsuarioPerfilStr.BlnEsValido AndAlso ObjIdAppPerfilUsuarioShr.BlnEsValido AndAlso
                    ObjIdPerfilUsuarioShr.BlnEsValido AndAlso EnuEstadoActualizacion <>
                    EnuEstadoObjetoDef.enuCreando Then
            Dim lstrKey As String
            Dim ldrwPermisosAcciones As DataRow() = FdrwPermisosAcciones()
            If Not IsNothing(ldrwPermisosAcciones) AndAlso ldrwPermisosAcciones.Count > 0 Then
                For Each ldrwPermisoAccion As DataRow In ldrwPermisosAcciones
                    Dim lobjPermisoAccion As New ClsPermisoAccion(Nothing, ldrwPermisoAccion)
                    lobjPermisoAccion.SLeaValores(True)
                    lstrKey = lobjPermisoAccion.ObjIdObjetoPermisoAccionShr.ToString &
                                lobjPermisoAccion.ObjIdAccionPermisoAccionShr.ToString
                    lcolPermisosAcciones.Add(lobjPermisoAccion, lstrKey)
                Next
            End If
        End If
        Return lcolPermisosAcciones
    End Function
    Private Function FdrwPermisosAcciones() As DataRow()
        Dim ldtbPermAcciones = FdtbPermisosAcciones()
        Dim ldrwPermiAssiones As DataRow() = Array.Empty(Of DataRow)
        If Not IsNothing(ldtbPermAcciones) Then
            ldrwPermiAssiones = ldtbPermAcciones.Select
        End If
        Return ldrwPermiAssiones
    End Function
    Friend Function FcolPermisosAccionesObj(ashrIdObjeto As Short) As Collection
        Dim lstrFiltro As String = "IdObjeto = " & ashrIdObjeto
        Dim lcolPermAccObj As New Collection
        Dim ldtbPermAcciones = FdtbPermisosAcciones()
        If Not IsNothing(ldtbPermAcciones) Then
            Dim ldrwPermAccObj = ldtbPermAcciones.Select(lstrFiltro)
            For Each ldrwPerAcc As DataRow In ldrwPermAccObj
                lcolPermAccObj.Add(ldrwPerAcc)
            Next
        End If
        Return lcolPermAccObj
    End Function
    Private Function FdtbPermisosAcciones() As DataTable
        Dim ldtbPermAccio As DataTable = Nothing
        If Not (IsNothing(ObjIdAppPerfilUsuarioShr.ObjValorPro) OrElse
                    IsNothing(ObjIdPerfilUsuarioShr.ObjValorPro)) Then
            Dim lstrFiltro As String = ObjIdAppPerfilUsuarioShr.StrNombreCampoBD & " = " &
                        ObjIdAppPerfilUsuarioShr.ObjValorPro & " AND " &
                        ObjIdPerfilUsuarioShr.StrNombreCampoBD & " = " &
                        ObjIdPerfilUsuarioShr.ObjValorPro
            ldtbPermAccio = ClsPanorama.FdtbDataTable("PanPermisosAcciones", {"*"},
                             {{"IdObjeto", "ASC"}, {"IdAccion", "ASC"}}, lstrFiltro)
        End If
        Return ldtbPermAccio
    End Function
#End Region
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsIdUsuarioPerfilStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdUsuario"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "IdUsuario"
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsPerfilUsuario = ObjPadre
        Dim lobjAbuelo As ClsUsuario = lobjPadre.ObjPadre
        Dim lblnEsValido As Boolean
        If Not IsNothing(lobjPadre.ObjPadre) Then
            If lobjAbuelo.ObjIdUsuarioStr.ObjValorPro = GCSTRUSUARIOU Then
                lblnEsValido = True
            Else
                lblnEsValido = (lobjAbuelo.ObjIdUsuarioStr.ObjValorPro = HobjValorNew)
            End If
        Else
            lblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud, BlnEsRequerido)
        End If
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsIdAppPerfilUsuarioShr
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
Friend Class ClsIdPerfilUsuarioShr
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
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
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
Friend Class ClsNombreAppPerfilUsuarioStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "NombreApp"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "NombreAplicacion"
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 4,
                ShrLongitud, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsNombrePerfilUsuarioStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "NombrePerfil"
        HshrLongitud = 20
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "NombrePerfil"
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 2, ShrLongitud,
                BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        Return ObjValorPro
    End Function
End Class
#End Region