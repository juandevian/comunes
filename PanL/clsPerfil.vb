Friend Class ClsPerfil
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "PanPerfiles"
    ' Objetos de propiedad
    Private ReadOnly MobjIdAppPerfilShr As New ClsIdAppPerfilShr(Me)
    Private ReadOnly MobjIdPerfilShr As New ClsIdPerfilShr(Me)
    Private ReadOnly MobjNombrePerfilStr As New ClsNombrePerfilStr(Me)
    Private ReadOnly MobjPadre As clsAplicacion = Nothing
    ' Variables de modulo
    Private MbytCantidadNivelesPermisos As Byte = 0
    Private MblnTodosPermitidos As Boolean = False
    Private MblnPermitirCambiosEnAdmin As Boolean = False
    ' Coleccion
    Private McolAcciones As Collection = Nothing
    Private McolPermisosAcciones As Collection = Nothing
    Private MblnColPermisosAccionesCargada As Boolean = False
    Private MdtbPermisosAcciones As DataTable = Nothing
    Private MblnDtblPermisosAccionesCargada As Boolean = False
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Perfil.
    ''' </summary>
    ''' <param name="aenuModoInstanciaObj">Esta clase no sepuede instanciar como Navegable!</param>
    ''' <remarks>Si se instancia como un objeto navegable, se crea un datatable que contiene las columnas de
    ''' la llave con las llaves de todos los objetos y queda a la espera de que se indique que objeto abrir.
    ''' Si se instancia como un objeto único, queda a la espera de recibir el valor de los campos de la llave 
    ''' para abrir dicho objeto. </remarks>
    Public Sub New(aobjApp As ClsAplicacion, aenuModoInstanciaObj As EnuModoInstanciaObjDef)
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuDeColeccion Then
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        End If
        Dim lstrCamposSelect As String()
        HobjPadre = aobjApp
        MobjPadre = HobjPadre
        HblnEsAnulable = False
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuNavegable Then
            lstrCamposSelect = {ObjIdAppPerfilShr.StrNombreCampoBD, ObjIdPerfilShr.StrNombreCampoBD}
            Dim lshrIdApp As Short = aobjApp.ObjIdAppShr.ObjValorPro
            Dim lstrFiltro = ClsIdAppPerfilShr.SstrNombreCampoBd & " = " & lshrIdApp
            HcolFiltros.Add(lstrFiltro)
        Else
            HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
            lstrCamposSelect = {"*"}
        End If
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
    End Sub
    ''' <summary>
    ''' Instancia un objeto Perfil de una aplicación con todos los permisos de las acciones en falso.
    ''' </summary>
    ''' <param name="aobjPadre">Objeto padre al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwPerfil">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <param name="acolAcciones">La colección de acciones correspondientes a la aplicación para la cual
    ''' se esá creando el perfil </param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsAplicacion, adrwPerfil As DataRow,
                acolAcciones As Collection)
        HobjPadre = aobjPadre
        MobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        '
        DrwRegistroActual = adrwPerfil
        DtbTablaColeccion = DrwRegistroActual.Table
        McolAcciones = acolAcciones
    End Sub
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto padre al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwPerfil">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsAplicacion, adrwPerfil As DataRow)
        HobjPadre = aobjPadre
        MobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsCreable = False
        HblnEsAnulable = False
        '
        DrwRegistroActual = adrwPerfil
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
            Return EnuIdClasesPanDef.EnuPerfil
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Perfil"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjIdAppPerfilShr As ClsIdAppPerfilShr
        Get
            Return MobjIdAppPerfilShr
        End Get
    End Property
    Friend ReadOnly Property ObjIdPerfilShr As ClsIdPerfilShr
        Get
            Return MobjIdPerfilShr
        End Get
    End Property
    Friend ReadOnly Property ObjNombrePerfilStr As ClsNombrePerfilStr
        Get
            Return MobjNombrePerfilStr
        End Get
    End Property
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(MobjIdAppPerfilShr, MobjIdAppPerfilShr.StrNombre)
                HcolPropiedades.Add(MobjIdPerfilShr, MobjIdPerfilShr.StrNombre)
                HcolPropiedades.Add(MobjNombrePerfilStr, MobjNombrePerfilStr.StrNombre)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    ''' <summary>
    ''' Devuelve la colección de PermisosAccion correspondiente al perfil 
    ''' </summary>
    ''' <param name="ablnKeyOrdinal">Determina si la llave asignada a cada elemento de la colección
    ''' esta conformada por el id del perfil y el ordinal, o por el id del objeto y el id de la acción</param>
    ''' <returns>Un acoleccion de objetos objPermisoAccion </returns>
    ''' <remarks></remarks>
    Friend ReadOnly Property ColPermisosAcciones(ablnKeyOrdinal As Boolean) As Collection
        Get
            Dim lbytKey As Byte = 1
            If Not ablnKeyOrdinal Then
                lbytKey = 2
            End If
            Static lbytKeyOrdinal As Byte
            If Not MblnColPermisosAccionesCargada OrElse lbytKey <> lbytKeyOrdinal Then
                If ObjIdAppPerfilShr.BlnEsValido Then
                    If IsNothing(McolPermisosAcciones) Then
                        McolPermisosAcciones = New Collection
                    Else
                        McolPermisosAcciones.Clear()
                    End If
                    Dim ldrwPermisosAcciones As DataRow() = FdrwPermisosAcciones()
                    If Not IsNothing(ldrwPermisosAcciones) AndAlso ldrwPermisosAcciones.Count > 0 Then
                        Dim lstrKey As String
                        For Each ldrwPermisoAccion As DataRow In ldrwPermisosAcciones
                            Dim lobjPermisoAccion As New ClsPermisoAccion(Me, ldrwPermisoAccion)
                            lobjPermisoAccion.SLeaValores(True)
                            If ablnKeyOrdinal Then
                                lstrKey = lobjPermisoAccion.ObjIdPerfilPermisoAccionShr.ToString &
                                        "," & lobjPermisoAccion.ObjOrdinalPermisoAccionShr.ToString
                            Else
                                lstrKey = lobjPermisoAccion.ObjIdObjetoPermisoAccionShr.ToString &
                                        "," & lobjPermisoAccion.ObjIdAccionPermisoAccionShr.ToString
                            End If
                            McolPermisosAcciones.Add(lobjPermisoAccion, lstrKey)
                        Next
                        MblnColPermisosAccionesCargada = True
                    End If
                    lbytKeyOrdinal = lbytKey
                Else
                    If Not IsNothing(McolPermisosAcciones) Then
                        McolPermisosAcciones.Clear()
                    End If
                End If
            End If
            Return McolPermisosAcciones
        End Get
    End Property
    ''' <summary>
    ''' Devuelve la cantidad de niveles en la colección "colPermisosAcciones" 
    ''' </summary>
    ''' <returns>Devuelve un byte</returns>
    ''' <remarks></remarks>
    Friend ReadOnly Property BytCantidadNivelesPermisos As Byte
        Get
            If Not String.IsNullOrEmpty(ObjIdAppPerfilShr.ToString) Then
                Dim lstrFiltro As String = ObjIdAppPerfilShr.StrNombreCampoBD & " = " & ObjIdAppPerfilShr.ToString
                Dim lbytMaxNivel As Byte = ClsPanorama.FobjValorCampo(GobjPanDat.FobjMaxValorCampo("PanPermisosAcciones",
                        "Nivel", lstrFiltro), EnuTipoValor.enuByte)
                MbytCantidadNivelesPermisos = lbytMaxNivel
            End If
            Return MbytCantidadNivelesPermisos
        End Get
    End Property
    Friend ReadOnly Property BlnTodosPermitidos As Boolean
        Get
            Dim lblnTodPer As Boolean = True
            If Not MblnColPermisosAccionesCargada Then
                McolPermisosAcciones = ColPermisosAcciones(True)
            End If
            If Not IsNothing(McolPermisosAcciones) AndAlso McolPermisosAcciones.Count > 0 Then
                For Each lobjPerAcc As ClsPermisoAccion In McolPermisosAcciones
                    If Not lobjPerAcc.ObjAccionPermitidaBln.ObjValorPro Then
                        lblnTodPer = False
                        Exit For
                    End If
                Next
            Else
                lblnTodPer = False
            End If
            MblnTodosPermitidos = lblnTodPer
            Return MblnTodosPermitidos
        End Get
    End Property
    Friend ReadOnly Property BlnTienePermitidos As Boolean
        Get
            Dim lblnTienePermitidos As Boolean = False
            If Not MblnColPermisosAccionesCargada Then
                McolPermisosAcciones = ColPermisosAcciones(True)
            End If
            If Not IsNothing(McolPermisosAcciones) AndAlso McolPermisosAcciones.Count > 0 Then
                For Each lobjPerAcc As ClsPermisoAccion In McolPermisosAcciones
                    If lobjPerAcc.ObjAccionPermitidaBln.ObjValorPro Then
                        lblnTienePermitidos = True
                        Exit For
                    End If
                Next
            Else
                lblnTienePermitidos = False
            End If
            Return lblnTienePermitidos
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Friend Overrides Sub SVacie()
        MyBase.SVacie()
        McolPermisosAcciones = Nothing
        MdtbPermisosAcciones = Nothing
        MblnColPermisosAccionesCargada = False
        MblnDtblPermisosAccionesCargada = False
    End Sub
    Protected Friend Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Dim lblnNoHayError = False
        GobjPanDat.SControleProcesoObj(True)
        Try
            GobjPanDat.SInicialiceTransaccion()
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                SNumereObj()
                MyBase.SActualice(ablnExigeRequeridos)
                SCreeColeccionPermisosAcciones()
            ElseIf EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                If ObjNombrePerfilStr.ToString.ToUpper = GCSTRADMIN.ToUpper Then
                    If Not MblnPermitirCambiosEnAdmin Then
                        Dim lstrMens = "El Perfil 'Admin' no puede ser modificado!"
                        SLevanteEventoNot(lstrMens, "", 0,
                                EnuSeveridadNot.EnuInformacion)
                        Exit Sub
                    Else
                        MyBase.SActualice(ablnExigeRequeridos)
                    End If
                End If
            End If
            ClsPanorama.SActualiceCol(McolPermisosAcciones)
            GobjPanDat.SConfirmeTransaccion()
            MblnDtblPermisosAccionesCargada = False
            MblnColPermisosAccionesCargada = False
            lblnNoHayError = True
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not lblnNoHayError Then
                GobjPanDat.SAborteTransaccion()
            End If
            GobjPanDat.SControleProcesoObj(False)
        End Try
    End Sub
    Protected Friend Overrides Sub SCreeObj(aobjValorLlave() As Object)
        MyBase.SCreeObj(aobjValorLlave)
        MblnColPermisosAccionesCargada = False
        MblnDtblPermisosAccionesCargada = False
    End Sub
    Protected Friend Overrides Function FblnSuprimio() As Boolean
        Dim lblnNoHayError = False
        Dim lblnSuprimio = FblnEsSuprimible()
        If lblnSuprimio Then
            Try
                GobjPanDat.SControleProcesoObj(True)
                GobjPanDat.SInicialiceTransaccion()
                lblnSuprimio = ClsPanorama.FblnSuprimioCol(ColPermisosAcciones(True))
                If lblnSuprimio Then
                    lblnSuprimio = MyBase.FblnSuprimio()
                End If
                lblnNoHayError = True
            Catch ex As PanLException
                Throw
            Catch ex As PanDatException
                Throw
            Catch ex As ArgumentNullException
                Throw
            Catch ex As Exception
                Throw
            Finally
                If lblnNoHayError Then
                    GobjPanDat.SConfirmeTransaccion()
                    GobjPanDat.SControleProcesoObj(False)
                Else
                    GobjPanDat.SAborteTransaccion()
                    GobjPanDat.SControleProcesoObj(False, True)
                End If
            End Try
        End If
        Return lblnSuprimio
    End Function
    Friend Overrides Function FblnEsSuprimible() As Boolean
        Dim lblnEsSuprimible As Boolean =
                Not (ObjNombrePerfilStr.ToString.ToUpper = GCSTRADMIN.ToUpper)
        lblnEsSuprimible = lblnEsSuprimible AndAlso FblnPermitidoSuprimir()
        If lblnEsSuprimible Then
            Dim lstrNombreColumnas As String() = {ObjIdAppPerfilShr.StrNombreCampoBD,
                    ObjIdPerfilShr.StrNombreCampoBD}
            Dim lstrCondiciones As String() = {" = " & ObjIdAppPerfilShr.ToString,
                    " = " & ObjIdPerfilShr.ToString}
            Dim lstrTablasExcluir As String() = {SstrNombreTabla,
                    ClsPermisoAccion.SstrNombreTabla}
            lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg(lstrTablasExcluir,
                    lstrNombreColumnas, lstrCondiciones, False, True)
            If GenuIdAplicacion <> EnuListaAplicaciones.EnuAdministrador And lblnEsSuprimible Then
                lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg(lstrTablasExcluir,
                    lstrNombreColumnas, lstrCondiciones, False, False)
            End If
        End If
        Return lblnEsSuprimible
    End Function
    Friend Overrides Sub SRefresqueObj()
        If Not IsNothing(McolPermisosAcciones) Then
            For Each lobjPermAcc As ClsPermisoAccion In McolPermisosAcciones
                If lobjPermAcc.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                    lobjPermAcc.SNormaliceEstado(True)
                End If
            Next
        End If
        MyBase.SRefresqueObj()
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdPerfilShr.ToString
        End Get
    End Property
#End Region
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Dim lstrFiltro As String = ObjIdAppPerfilShr.StrNombreCampoBD & " = " & ObjIdAppPerfilShr.ObjValorPro
            Dim lshrIdUltimoPerfil = ClsPanorama.FobjUltimaIdNumericaObjeto("PanPerfiles",
                    ObjIdPerfilShr.StrNombreCampoBD, ObjIdPerfilShr.EnuTipoValor, lstrFiltro)
            ObjIdPerfilShr.ObjValorPro = lshrIdUltimoPerfil + 1
        End If
    End Sub
    Public Sub SAsignePermiso(astrKeyPermisoAccion As String, ablnPermitido As Boolean)
        If McolPermisosAcciones.Contains(astrKeyPermisoAccion) Then
            Dim lobjPerAccion As ClsPermisoAccion = McolPermisosAcciones(astrKeyPermisoAccion)
            If lobjPerAccion.ObjAccionPermitidaBln.ObjValorPro <> ablnPermitido AndAlso
                    EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                If lobjPerAccion.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                    lobjPerAccion.SModifique()
                End If
                lobjPerAccion.ObjAccionPermitidaBln.ObjValorPro = ablnPermitido
            End If
        End If
    End Sub
#Region "Permisos Acciones"
    Friend Sub SCambiePermisosAdmin()
        If Not MblnColPermisosAccionesCargada Then
            McolPermisosAcciones = ColPermisosAcciones(True)
        End If
        MblnPermitirCambiosEnAdmin = True
        For Each lobjperAcc As ClsPermisoAccion In McolPermisosAcciones
            lobjperAcc.SModifique()
            lobjperAcc.ObjAccionPermitidaBln.ObjValorPro = True
        Next
        If McolPermisosAcciones.Count > 0 Then
            ClsPanorama.SActualiceCol(McolPermisosAcciones)
        End If
    End Sub
    Private Sub SCreeColeccionPermisosAcciones()
        SCargueDtbPermisosAcciones()
        If Not IsNothing(MdtbPermisosAcciones) Then
            Dim lstrMens = String.Empty
            If MdtbPermisosAcciones.Rows.Count = 0 Then
                McolAcciones = MobjPadre.ColAcciones
                If McolAcciones.Count > 0 Then
                    Dim lobjAccion As ClsAccion = Nothing
                    Dim ldrwPerAcc As DataRow = Nothing
                    Dim lstrKey As String = String.Empty
                    If IsNothing(McolPermisosAcciones) Then
                        McolPermisosAcciones = New Collection
                    Else
                        McolPermisosAcciones.Clear()
                    End If
                    For i = 1 To McolAcciones.Count
                        lobjAccion = McolAcciones(i)
                        ldrwPerAcc = MdtbPermisosAcciones.NewRow
                        Dim lobjPermisoAccion As New ClsPermisoAccion(Me, ldrwPerAcc)
                        lobjPermisoAccion.SCreeObj(Nothing)
                        With lobjPermisoAccion
                            .ObjIdAccionPermisoAccionShr.ObjValorPro = lobjAccion.ObjIdAccionShr.ObjValorPro
                            .ObjIdAppPermisoAccionShr.ObjValorPro = MobjPadre.ObjIdAppShr.ObjValorPro
                            .ObjIdObjetoPermisoAccionShr.ObjValorPro = lobjAccion.ObjIdObjetoShr.ObjValorPro
                            .ObjIdPerfilPermisoAccionShr.ObjValorPro = ObjIdPerfilShr.ObjValorPro
                            .ObjNivelPermisoAccionByt.ObjValorPro = lobjAccion.ObjNivelByt.ObjValorPro
                            .ObjNombreAccionPermisoAccionStr.ObjValorPro = lobjAccion.ObjNombreAccionStr.ObjValorPro
                            .ObjOrdinalPermisoAccionShr.ObjValorPro = lobjAccion.ObjOrdinalShr.ObjValorPro
                            .ObjAccionPermitidaBln.ObjValorPro = False
                            lstrKey = .ObjIdAppPermisoAccionShr.ToString &
                                    .ObjIdPerfilPermisoAccionShr.ToString &
                                    .ObjIdObjetoPermisoAccionShr.ToString &
                                    .ObjOrdinalPermisoAccionShr.ToString
                        End With
                        McolPermisosAcciones.Add(lobjPermisoAccion, lstrKey)
                    Next i
                    MblnColPermisosAccionesCargada = True
                Else
                    lstrMens = "La Coleccion de Acciones esta vacía!"
                End If
            Else
                lstrMens = "La Tabla PanPermisosAcciones ya tiene registros para la aplicación!"
            End If
            If Not String.IsNullOrEmpty(lstrMens) Then
                SLevanteEventoNot(lstrMens, "", 0, EnuSeveridadNot.EnuError)
            End If
        End If
    End Sub
    Private Sub SCargueDtbPermisosAcciones()
        If Not MblnDtblPermisosAccionesCargada Then
            If Not (IsNothing(ObjIdAppPerfilShr.ObjValorPro) OrElse IsNothing(ObjIdPerfilShr.ObjValorPro)) Then
                Dim lstrFiltro As String = ObjIdAppPerfilShr.StrNombreCampoBD & " = " & ObjIdAppPerfilShr.ObjValorPro &
                        " AND " & ObjIdPerfilShr.StrNombreCampoBD & " = " & ObjIdPerfilShr.ObjValorPro
                If Not IsNothing(MdtbPermisosAcciones) Then
                    MdtbPermisosAcciones.Clear()
                End If
                MdtbPermisosAcciones = ClsPanorama.FdtbDataTable("PanPermisosAcciones", {"*"},
                            {{"Ordinal", "ASC"}},
                            lstrFiltro)
                MblnDtblPermisosAccionesCargada = True
            End If
        End If
    End Sub
    Private Function FdrwPermisosAcciones() As DataRow()
        SCargueDtbPermisosAcciones()
        If Not IsNothing(MdtbPermisosAcciones) Then
            Return MdtbPermisosAcciones.Select
        Else
            Return Nothing
        End If
    End Function
#End Region
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsIdAppPerfilShr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsPerfil = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "IdAplicacion"
    Public Sub New(aobjPadre As ClsPerfil)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdApp"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 0
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean
        If Not IsNothing(MobjPadre.ObjPadre) Then
            Dim lobjAbuelo As ClsAplicacion = MobjPadre.ObjPadre
            lblnEsValido = (HobjValorNew = lobjAbuelo.ObjIdAppShr.ObjValorPro)
        Else
            lblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew,
                    EnuListaAplicaciones.EnuAdministrador, EnuListaAplicaciones.EnuOrionCop,
                    BlnEsRequerido, EnuTipoValor.enuShort)
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
Friend Class ClsIdPerfilShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdPerfil"
    Public Sub New(aobjPadre As ClsPerfil)
        MyBase.New(aobjPadre)
        HstrNombre = "IdPerfil"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 1
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue, BlnEsRequerido,
                EnuTipoValor)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
            If Not BlnLeyendoOrigen Then
                If Not HblnEsValido Then
                    If Not (ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso
                            (Not IsNumeric(HobjValorNew) OrElse IsNothing(HobjValorNew))) Then
                        HstrMens = "El código del Perfil ingresado no es valido!"
                        SNotifiqueDatInv()
                    End If
                End If
            End If
        Else
            HblnEsValido = True
        End If
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
Friend Class ClsNombrePerfilStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Nombre"
    Public Sub New(aobjPadre As ClsPerfil)
        MyBase.New(aobjPadre)
        HstrNombre = "NombrePerfil"
        HshrLongitud = 20
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 2, ShrLongitud,
                BlnEsRequerido)
        If lblnEsValido Then
            Dim lobjPadre As ClsPerfil = ObjPadre
            If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                Dim lobjAbuelo As ClsAplicacion = lobjPadre.ObjPadre
                lblnEsValido = Not lobjAbuelo.FblnExisteNombrePerfil(HobjValorNew)
                If Not lblnEsValido Then
                    HstrMens = "Ya existe un Perfil con este Nombre!"
                    SNotifiqueDatInv()
                End If
            End If
        End If
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        If ObjValorPro Is Nothing Then
            Return ""
        Else
            Return ObjValorPro
        End If
    End Function
End Class
#End Region
