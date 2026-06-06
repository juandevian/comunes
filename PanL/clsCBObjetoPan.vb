Public MustInherit Class ClsCBObjetoPan
#Region "Definiciones"
#Region "Variables"
    ' Colecciones utilizadas para construir instrucciónes SQL
    Private McolCamposIndicesTablas As Collection = Nothing
    Private McolNombresCamposRef As Collection = Nothing
    Private ReadOnly McolNombresCamposCambio As New Collection
    Private ReadOnly McolDatosNuevos As New Collection
    ' Coleccion de las colecciones que definen el indice de una instruccion SQL "ORDER BY"
    Private ReadOnly McolDatosRef As New Collection
    ' Objetos de datos
    Private MdtbTablaNavegacion As DataTable = Nothing
    Private MdtbTablaObjetoUnico As DataTable = Nothing
    Private MdtbTablaColeccion As DataTable = Nothing
    Private MdrwIdObjetoActual As DataRow = Nothing
    ' Variables de propiedades comunes
    Private MenuEstadoActualizacion As EnuEstadoObjetoDef = EnuEstadoObjetoDef.enuConsultando
    Protected Property HenuTipoPermiso As EnuPermisosDef = EnuPermisosDef.None
    Private MobjValorUltimaLlave() As Object = Nothing
    '
    Private MentCantCambios As Integer = 0
    Private MblnTengoCambios As Boolean = False
    Private MblnEsAutonumerico As Boolean = False
    '
    Private MblnExiste As Boolean = False
    Private MobjOrigenInstanciaStr As ClsOrigenInstanciaStr = Nothing
    Private MobjAnuladoBln As ClsAnuladoBln = Nothing
    Private MobjIdUsuarioAnuloStr As ClsIdUsuarioAnuloStr = Nothing
    Private MobjOrigenInstanciaAnuloStr As ClsOrigenInstanciaAnuloStr = Nothing
    Private MobjFechaCreacionDtm As ClsFechaCreacionDtm = Nothing
    ' 
    Private ReadOnly McolPropiedades As New Collection
    Private ReadOnly McolTablas As New Collection
    Private ReadOnly McolFiltros As New Collection
    Private ReadOnly McolCamposSelect As New Collection
#Region "Propiedades autoimplementadas"
    ' Propiedades heredables
    Friend MustOverride ReadOnly Property ColPropiedades As Collection
    Protected Friend MustOverride ReadOnly Property HenuIdClase As EnuIdClasesPanDef
    Protected MustOverride ReadOnly Property HstrNombreClase As String
    Protected MustOverride ReadOnly Property HstrNombreTabla As String
    Friend Overridable ReadOnly Property HstrNombreObj As String
    '
    Protected Property HobjPadre As Object = Nothing
    Protected Property HblnEsCreable As Boolean = True
    Protected Property HblnEsModificable As Boolean = True
    Protected Property HblnEsSuprimible As Boolean = True
    Protected Property HblnEsAnulable As Boolean = True
    Protected Property HenuTipoObjeto As EnuModoInstanciaObjDef =
            EnuModoInstanciaObjDef.EnuNavegable

    Protected Property DrwRegistroActual As DataRow = Nothing
    Protected Property ObjValorLlave As Object = Nothing
    ' Otras propiedades
    Friend Property HstrPropiedadNoValida As String = String.Empty
    Public Property BlnVaciandoObjeto As Boolean = False
    Friend Property BlnImportoUltimo As Boolean = False
#End Region
#End Region
#Region "Eventos"
    Friend Event EvnNotifica As EventHandler(Of ClsNotiEventArgs)
#End Region
#End Region

#Region "Constructores"
    Protected Sub New()
        If Not GobjPanorama.BlnRegistrado Then
            Throw New ModuloNoRegistradoPanException
        End If
    End Sub
#End Region

#Region "Propiedades comunes"
    Protected ReadOnly Property HcolPropiedades As Collection
        Get
            Return McolPropiedades
        End Get
    End Property
    Protected ReadOnly Property HcolTablas As Collection
        Get
            Return McolTablas
        End Get
    End Property
    Protected ReadOnly Property HcolFiltros As Collection
        Get
            Return McolFiltros
        End Get
    End Property
    Protected ReadOnly Property HcolCamposSelect As Collection
        Get
            Return McolCamposSelect
        End Get
    End Property
    ' Objetos de datos
    Protected ReadOnly Property DtbTablaNavegacion As DataTable
        Get
            If MdtbTablaNavegacion Is Nothing OrElse MdtbTablaNavegacion.Rows.Count = 0 Then
                Dim lblnNoHayError = False
                Using ldstDatSet As New DataSet
                    Try
                        GobjPanDat.SControleProcesoObj(True)
                        GobjPanDat.SdsDataSet(ldstDatSet, HcolTablas, HcolCamposSelect,
                                ColCamposIndexesTablas, HcolFiltros)
                        lblnNoHayError = True
                    Catch ex As PanDatException
                        Throw
                    Catch ex As PanLException
                        Throw
                    Catch ex As ArgumentNullException
                        Throw
                    Catch ex As Exception
                        Throw
                    Finally
                        If lblnNoHayError Then
                            If ldstDatSet.Tables.Count > 0 Then
                                MdtbTablaNavegacion = ldstDatSet.Tables(0)
                            End If
                        End If
                        GobjPanDat.SControleProcesoObj(False)
                    End Try
                End Using
            End If
            Return MdtbTablaNavegacion
        End Get
    End Property
    Private ReadOnly Property DtbObjetoUnico(aobjValorLlave() As Object) As DataTable
        Get
            If IsNothing(MdtbTablaObjetoUnico) OrElse Not MdtbTablaObjetoUnico.Rows.Contains(aobjValorLlave) Then
                Dim lcolFiltros As New Collection From {
                    FstrFiltro(aobjValorLlave)
                }
                Using ldstDatSet As New DataSet
                    Try
                        GobjPanDat.SControleProcesoObj(True) 'Ok
                        GobjPanDat.SdsDataSet(ldstDatSet, HcolTablas, HcolCamposSelect, ColCamposIndexesTablas,
                                    lcolFiltros)
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
                    If Not IsNothing(ldstDatSet) AndAlso ldstDatSet.Tables.Count > 0 Then
                        MdtbTablaObjetoUnico = ldstDatSet.Tables(0)
                        If MdtbTablaObjetoUnico.Rows.Count = 0 Then
                            MdtbTablaObjetoUnico = Nothing
                        End If
                    Else
                        MdtbTablaObjetoUnico = Nothing
                    End If
                End Using
            End If
            Return MdtbTablaObjetoUnico
        End Get
    End Property
    Protected Property DtbTablaColeccion() As DataTable
        Get
            Return MdtbTablaColeccion
        End Get
        Set(value As DataTable)
            MdtbTablaColeccion = value
            Dim lclmColsLlavePrimaria As DataColumn() = FclmPrimaryKey()
            If Not IsNothing(lclmColsLlavePrimaria) Then
                MdtbTablaColeccion.PrimaryKey = lclmColsLlavePrimaria
            End If
        End Set
    End Property
    ''' <summary>
    ''' Indica si un objeto que que se abrio con "sAbra(aobjIDLlave() As Object)" existe o no.
    ''' </summary>
    ''' <returns>Boleano</returns>
    ''' <remarks></remarks>
    Friend ReadOnly Property BlnExiste() As Boolean
        Get
            Return MblnExiste
        End Get
    End Property
    Friend ReadOnly Property ObjPadre As Object
        Get
            Return HobjPadre
        End Get
    End Property
    Public ReadOnly Property StrNombreClase As String
        Get
            Return HstrNombreClase
        End Get
    End Property
    Friend ReadOnly Property EnuTipoObjeto As EnuModoInstanciaObjDef
        Get
            Return HenuTipoObjeto
        End Get
    End Property
    Public Overridable ReadOnly Property BlnEsCreable As Boolean
        Get
            Return HblnEsCreable
        End Get
    End Property
    Public Overridable ReadOnly Property BlnEsModificable As Boolean
        Get
            Return HblnEsModificable
        End Get
    End Property
    Protected ReadOnly Property BlnEsSuprimible As Boolean
        Get
            Return HblnEsSuprimible
        End Get
    End Property
    Public Overridable ReadOnly Property BlnEsAnulable As Boolean
        Get
            Return HblnEsAnulable
        End Get
    End Property
    Public Overridable ReadOnly Property BlnEsNavegable As Boolean
        Get
            Return (HenuTipoObjeto = EnuModoInstanciaObjDef.enuNavegable)
        End Get
    End Property
    ' Colecciones utilizadas para construir instrucciónes SQL
    Protected ReadOnly Property ColCamposIndexesTablas As Collection
        Get
            If IsNothing(McolCamposIndicesTablas) Then
                Dim lstrCamposIndiceTabla As String(,) = FstrCamposIndiceTabla()
                If IsNothing(lstrCamposIndiceTabla) Then
                    ReDim lstrCamposIndiceTabla(0, 1)
                    lstrCamposIndiceTabla(0, 0) = String.Empty
                    lstrCamposIndiceTabla(0, 1) = String.Empty
                End If
                McolCamposIndicesTablas = New Collection From {
                    {lstrCamposIndiceTabla, HstrNombreTabla}
                }
            End If
            Return McolCamposIndicesTablas
        End Get
    End Property
    Protected ReadOnly Property ColNombresCamposRef As Collection
        Get
            If IsNothing(McolNombresCamposRef) Then
                McolNombresCamposRef = New Collection
                Dim lshrLimiteSupIndice As Short = FshrLimiteSupIndice()
                Dim lstrCamposIndiceTabla As String(,) = FstrCamposIndiceTabla()
                If lshrLimiteSupIndice > -1 Then
                    For i = 0 To lshrLimiteSupIndice
                        McolNombresCamposRef.Add(lstrCamposIndiceTabla(i, 0))
                    Next
                End If
            End If
            Return McolNombresCamposRef
        End Get
    End Property
    Friend ReadOnly Property ObjMatrizLlave() As Object
        Get
            Dim lshrLimiteMatriz As Short = FshrLimiteSupIndice()
            Dim lobjMatrizLlave(lshrLimiteMatriz) As Object
            For Each lobjPro As ClsCBPropiedad In ColPropiedades
                If lobjPro.BlnEsLlave Then
                    lobjMatrizLlave(lobjPro.BytPosLlave) = lobjPro.ObjValorPro
                End If
            Next
            Return lobjMatrizLlave
        End Get
    End Property
    Friend ReadOnly Property ObjValorUltimaLlave() As Object
        Get
            Return MobjValorUltimaLlave
        End Get
    End Property
    Friend Property ObjAnuladoBln As ClsAnuladoBln
        Get
            If IsNothing(MobjAnuladoBln) Then
                MobjAnuladoBln = New ClsAnuladoBln(Me)
            End If
            Return MobjAnuladoBln
        End Get
        Set(value As ClsAnuladoBln)
            If Not IsNothing(MobjAnuladoBln) Then
                MobjAnuladoBln = value
            End If
        End Set
    End Property
    Friend Property ObjOrigenInstanciaStr As ClsOrigenInstanciaStr
        Get
            If IsNothing(MobjOrigenInstanciaStr) Then
                MobjOrigenInstanciaStr = New ClsOrigenInstanciaStr(Me)
            End If
            Return MobjOrigenInstanciaStr
        End Get
        Set(value As ClsOrigenInstanciaStr)
            If Not IsNothing(MobjOrigenInstanciaStr) Then
                MobjOrigenInstanciaStr = value
            End If
        End Set
    End Property
    Friend Property ObjOrigenInstanciaAnuloStr As ClsOrigenInstanciaAnuloStr
        Get
            If IsNothing(MobjOrigenInstanciaAnuloStr) Then
                MobjOrigenInstanciaAnuloStr = New ClsOrigenInstanciaAnuloStr(Me)
            End If
            Return MobjOrigenInstanciaAnuloStr
        End Get
        Set(value As ClsOrigenInstanciaAnuloStr)
            If Not IsNothing(MobjOrigenInstanciaAnuloStr) Then
                MobjOrigenInstanciaAnuloStr = value
            End If
        End Set
    End Property
    Friend Property ObjIdUsuarioAnuloStr As ClsIdUsuarioAnuloStr
        Get
            If IsNothing(MobjIdUsuarioAnuloStr) Then
                MobjIdUsuarioAnuloStr = New ClsIdUsuarioAnuloStr(Me)
            End If
            Return MobjIdUsuarioAnuloStr
        End Get
        Set(value As ClsIdUsuarioAnuloStr)
            If Not IsNothing(MobjIdUsuarioAnuloStr) Then
                MobjIdUsuarioAnuloStr = value
            End If
        End Set
    End Property
    Friend Property ObjFechaCreacionDtm As ClsFechaCreacionDtm
        Get
            If IsNothing(MobjFechaCreacionDtm) Then
                MobjFechaCreacionDtm = New ClsFechaCreacionDtm(Me)
            End If
            Return MobjFechaCreacionDtm
        End Get
        Set(value As ClsFechaCreacionDtm)
            If Not IsNothing(MobjFechaCreacionDtm) Then
                MobjFechaCreacionDtm = value
            End If
        End Set
    End Property
    Friend Overridable ReadOnly Property StrIdObjeto As String
        Get
            Return Nothing
        End Get
    End Property
    Public Property EnuPermisosObj As EnuPermisosDef
        Get
            If GshrIdAplicacion = 803 OrElse GshrIdAplicacion = 100 Then
                If HenuTipoPermiso = EnuPermisosDef.None Then
                    HenuTipoPermiso = GobjPanorama.FenuTipoPermisos(HenuIdClase)
                ElseIf CType(HenuTipoPermiso And EnuPermisosDef.enuHeredado, Boolean) Then
                    If GstrIdUsuario = GCSTRUSUARIOU Then
                        HenuTipoPermiso = EnuPermisosDef.enuTodos
                    Else
                        HenuTipoPermiso = ObjPadre.EnuPermisosObj
                    End If
                End If
            Else
                HenuTipoPermiso = EnuPermisosDef.enuTodos
            End If
            Return HenuTipoPermiso
        End Get
        Set(value As EnuPermisosDef)
            HenuTipoPermiso = value
        End Set
    End Property
#End Region

#Region "Manejo estado objeto"
    Public Property EnuEstadoActualizacion() As EnuEstadoObjetoDef
        Get
            Return MenuEstadoActualizacion
        End Get
        Set(value As EnuEstadoObjetoDef)
            Try
                SActualiceEstado(value)
            Catch ex As ErrorInesperadoPanLException
                Throw
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property
    Private Sub SActualiceEstado(aenuEstadoNuevo As EnuEstadoObjetoDef)
        Select Case MenuEstadoActualizacion
            Case EnuEstadoObjetoDef.enuConsultando
                MenuEstadoActualizacion = aenuEstadoNuevo
            Case EnuEstadoObjetoDef.enuCreando
                Select Case aenuEstadoNuevo
                    Case EnuEstadoObjetoDef.enuConsultando
                        MenuEstadoActualizacion = aenuEstadoNuevo
                    Case Else
                        Throw New ErrorInesperadoPanLException("Estado inesperado del Objeto!")
                End Select
            Case EnuEstadoObjetoDef.enuModificando
                Select Case aenuEstadoNuevo
                    Case EnuEstadoObjetoDef.enuConsultando
                        MenuEstadoActualizacion = aenuEstadoNuevo
                    Case Else
                        Throw New ErrorInesperadoPanLException("Estado inesperado del Objeto!")
                End Select
            Case EnuEstadoObjetoDef.enuEliminando
                Throw New ErrorInesperadoPanLException("Estado inesperado del Objeto!")
        End Select
    End Sub
    ''' <summary>
    ''' Indica si el objeto ha tenido cambios en alguna de sus propiedades o en alguno de sus hijos cuando
    ''' es padre de objetos que conforman una colección
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property BlnTengoCambios As Boolean
        Get
            Return MblnTengoCambios
        End Get
    End Property
    ''' <summary>
    ''' Asigna un valor booleano indicando que una de las propiedades hija tuvo un cambio en su valor
    '''  o que alguno de los objetos hijos de alguna coleccion ha sufrido cambios. 
    ''' </summary>
    ''' <param name="ablnCambioEnObjetoHijo">Valor que indica que el cambio se efectuo en un objeto hijo
    ''' componente de una colección</param>
    ''' <param name="ablnReiniciar">Indica que se debe reiniciar el conteo de cambios en el hijo.</param>
    ''' <remarks></remarks>
    Public Sub STengoCambios(ablnTengoCambios As Boolean, ablnCambioEnObjetoHijo As Boolean,
            ablnReiniciar As Boolean)
        If ablnCambioEnObjetoHijo Then
            If ablnReiniciar Then
                MentCantCambios = 0
            Else
                If ablnTengoCambios Then
                    MentCantCambios += 1
                Else
                    MentCantCambios -= 1
                End If
            End If
            If MentCantCambios < 0 Then
                Throw New ErrorInesperadoPanLException("Cantidad de cambios inesperado")
            End If
            If MentCantCambios > 0 Then
                MblnTengoCambios = True
            Else
                MblnTengoCambios = FblnHayCambios()
            End If
        Else
            If MblnTengoCambios <> ablnTengoCambios Then
                If Not ablnTengoCambios Then
                    MblnTengoCambios = FblnHayCambios()
                Else
                    MblnTengoCambios = True
                End If
            End If
        End If
        If TypeOf HobjPadre Is ClsCBObjetoPan Then
            Dim lobjPadre As ClsCBObjetoPan = HobjPadre
            If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                lobjPadre.STengoCambios(ablnTengoCambios, True, ablnReiniciar)
            End If
        End If
    End Sub
    ''' <summary>
    ''' Establece el estado del objeto en "Consultando" lo que significa que los datos de la BD 
    ''' y los del objeto son iguales.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overridable Sub SNormaliceEstado(ablnRefresqueObjeto As Boolean)
        MenuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando
        If ablnRefresqueObjeto Then
            SRefresqueObj()
        End If
        If TypeOf HobjPadre Is ClsCBObjetoPan Then
            If HobjPadre.blnTengoCambios Then
                HobjPadre.sTengoCambios(False, True, True)
            End If
        End If
    End Sub
#End Region

#Region "Funciones"
    Friend Overridable Function FblnEsCreable() As Boolean
        Return FblnEsCreable(ObjMatrizLlave)
    End Function
    Friend Overridable Function FblnEsCreable(aobjValorLlave() As Object) As Boolean
        Dim lblnEsCreable As Boolean = False
        If BlnEsCreable Then
            If FblnValorLlaveEsNull(aobjValorLlave) Then
                lblnEsCreable = True
            Else
                lblnEsCreable = Not FblnExisteLlave(aobjValorLlave)
            End If
        End If
        Return lblnEsCreable
    End Function
    Friend Overridable Function FblnEsModificable() As Boolean
        Dim lblnEsModificable As Boolean = False
        If BlnEsModificable Then
            lblnEsModificable = Not (Not IsNothing(ObjAnuladoBln) AndAlso ObjAnuladoBln.ObjValorPro)
        End If
        Return lblnEsModificable
    End Function
    Friend Overridable Function FblnEsSuprimible() As Boolean
        Dim lblnEsSuprimible = FblnPermitidoSuprimir()
        Return lblnEsSuprimible
    End Function
    Protected Friend Function FblnPermitidoSuprimir() As Boolean
        Dim lblnPermitido = BlnEsSuprimible
        If Not lblnPermitido Then
            Return lblnPermitido
        End If
        lblnPermitido = MenuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando
        If lblnPermitido Then
            lblnPermitido = EnuPermisosObj And EnuPermisosDef.enuSuprimir
        End If
        Return lblnPermitido
    End Function
    Public Overridable Function FblnEsAnulable() As Boolean
        Return BlnEsAnulable
    End Function
    Friend Overridable Function FblnEstaVacioOrigenDatos() As Boolean
        Dim lblnEstaVacio As Boolean = True
        Select Case HenuTipoObjeto
            Case EnuModoInstanciaObjDef.enuNavegable
                If IsNothing(DtbTablaNavegacion) Then
                    Throw New ErrorInesperadoPanLException("'dtbTablaNavegacion' es Nothing en 'fblnEstaVacioOrigenDatos'")
                End If
                lblnEstaVacio = IsNothing(DtbTablaNavegacion) OrElse (DtbTablaNavegacion.Rows.Count = 0)
            Case EnuModoInstanciaObjDef.enuUnico
                lblnEstaVacio = True
            Case EnuModoInstanciaObjDef.enuDeColeccion
                If IsNothing(DtbTablaColeccion) Then
                    Throw New ErrorInesperadoPanLException("'DtbTablaColeccion' es Nothing en 'fblnEstaVacioOrigenDatos'")
                End If
                lblnEstaVacio = IsNothing(DtbTablaColeccion) OrElse (DtbTablaColeccion.Rows.Count = 0)
            Case EnuModoInstanciaObjDef.None
                Dim lstrMens = "Tipo de objeto no valido en 'fblnEstaVacioOrigenDatos'"
                SLevanteEventoNot(lstrMens, "", 0, EnuSeveridadNot.EnuError)
        End Select
        Return lblnEstaVacio
    End Function
    Public Function FblnEsNavegable() As Boolean
        If BlnEsNavegable AndAlso (Not FblnEstaVacioOrigenDatos()) Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function FblnEsElPrimerRegistro() As Boolean
        Return DtbTablaNavegacion.Rows.IndexOf(MdrwIdObjetoActual) = 0
    End Function
    Public Function FblnEsElUltimoRegistro() As Boolean
        Return DtbTablaNavegacion.Rows.IndexOf(MdrwIdObjetoActual) = DtbTablaNavegacion.Rows.Count - 1
    End Function
    Private Function FshrLimiteSupIndice() As Short
        Dim lshrLimSup As Short = -1
        For Each lobjProp As ClsCBPropiedad In ColPropiedades
            If lobjProp.BlnEsLlave Then
                If lobjProp.BytPosLlave > lshrLimSup Then
                    lshrLimSup = lobjProp.BytPosLlave
                End If
            End If
        Next
        Return lshrLimSup
    End Function
    ''' <summary>
    ''' Indica si los valores de todas las propiedades de la clase son validos 
    ''' </summary>
    ''' <returns>Boolean</returns>
    ''' <remarks></remarks>
    Public Function FblnEstanTodosOk() As Boolean
        Dim lblnOkTodos = True
        For Each lobjPropiedad As ClsCBPropiedad In ColPropiedades
            If Not lobjPropiedad.BlnEsValido Then
                Dim lstrObjeto = String.Empty
                For i = 0 To ObjMatrizLlave.length - 1
                    If Not IsNothing(ObjMatrizLlave(i)) Then
                        lstrObjeto += ObjMatrizLlave(i).ToString & ", "
                    End If
                Next
                HstrPropiedadNoValida = "Nombre Clase: " & HstrNombreClase & "; Id: " & lstrObjeto &
                        "; Propiedad: " & lobjPropiedad.StrNombre
#If DES = 1 Then
                Debug.Print(HstrPropiedadNoValida)
                Stop
#End If
                lblnOkTodos = False
                Exit For
            End If
        Next
        If Not lblnOkTodos Then
            SLevanteEventoNot("Aun hay propiedades por satisfacer!", "", 0,
                    EnuSeveridadNot.EnuCamInsatis)
        End If
        Return lblnOkTodos
    End Function
    Private Function FenuTipoDatoCampoBD(astrNombreCampo As String) As EnuTipoValor
        For Each lobjPropiedad As ClsCBPropiedad In ColPropiedades
            If lobjPropiedad.StrNombreCampoBD = astrNombreCampo Then
                Return lobjPropiedad.EnuTipoValor
            End If
        Next
        Return Nothing
    End Function
    ''' <summary>
    ''' Indica si el objeto identificado con la llave 'aobjValorLlave()' existe o no.
    ''' </summary>
    ''' <param name="aobjValorLlave">Matriz que contiene los valores que identifican de manera única un objeto.</param>
    ''' <returns>Buleano</returns>
    ''' <remarks></remarks>
    Protected Friend Function FblnExisteLlave(aobjValorLlave() As Object) As Boolean
        Dim lblnExiste = False
        Select Case HenuTipoObjeto
            Case EnuModoInstanciaObjDef.enuNavegable
                If Not IsNothing(DtbTablaNavegacion) Then
                    If Not IsNothing(aobjValorLlave) Then
                        lblnExiste = DtbTablaNavegacion.Rows.Contains(aobjValorLlave)
                    End If
                End If
            Case EnuModoInstanciaObjDef.enuUnico
                lblnExiste = (Not IsNothing(DtbObjetoUnico(aobjValorLlave)) AndAlso
                        DtbObjetoUnico(aobjValorLlave).Rows.Contains(aobjValorLlave))
            Case EnuModoInstanciaObjDef.enuDeColeccion
                If Not IsNothing(DtbTablaColeccion) Then
                    lblnExiste = DtbTablaColeccion.Rows.Contains(aobjValorLlave)
                End If
        End Select
        Return lblnExiste
    End Function
    Private Function FblnEsValidoValorLlave(aobjValorLlave As Object) As Boolean
        Dim lblnEsVal = False
        Dim lbytPos As Byte
        Dim lenuTipoValor As EnuTipoValor
        Dim lobjValor As Object = Nothing
        For Each lobjPropiedad As ClsCBPropiedad In ColPropiedades
            If lobjPropiedad.BlnEsLlave Then
                lenuTipoValor = lobjPropiedad.EnuTipoValor
                lbytPos = lobjPropiedad.BytPosLlave
                lobjValor = aobjValorLlave(lbytPos)
                Select Case lenuTipoValor
                    Case EnuTipoValor.enuBoolean
                        lblnEsVal = ClsPanorama.FblnEsValidoBuleano(lobjValor)
                    Case EnuTipoValor.enuByte
                        lblnEsVal = ClsPanorama.FblnEsValidoNumero(lobjValor, Byte.MinValue, Byte.MaxValue, True, EnuTipoValor.enuSByte)
                    Case EnuTipoValor.enuSByte
                        lblnEsVal = ClsPanorama.FblnEsValidoNumero(lobjValor, SByte.MinValue, SByte.MaxValue, True, EnuTipoValor.enuSByte)
                    Case EnuTipoValor.enuDate, EnuTipoValor.enuDateTime
                        lblnEsVal = ClsPanorama.FblnEsValidoFecha(lobjValor, DateSerial(2000, 1, 1), GCDTMFECHAMAXI, True)
                    Case EnuTipoValor.enuDecimal
                        lblnEsVal = ClsPanorama.FblnEsValidoNumero(lobjValor, Decimal.MinValue, Decimal.MaxValue, True, EnuTipoValor.enuDecimal)
                    Case EnuTipoValor.enuDouble
                        lblnEsVal = ClsPanorama.FblnEsValidoNumero(lobjValor, Double.MinValue, Double.MaxValue, True, EnuTipoValor.enuDouble)
                    Case EnuTipoValor.enuInteger
                        lblnEsVal = ClsPanorama.FblnEsValidoNumero(lobjValor, Integer.MinValue, Integer.MaxValue, True, EnuTipoValor.enuInteger)
                    Case EnuTipoValor.enuUInteger
                        lblnEsVal = ClsPanorama.FblnEsValidoNumero(lobjValor, UInteger.MinValue, UInteger.MaxValue, True, EnuTipoValor.enuUInteger)
                    Case EnuTipoValor.enuSingle
                        lblnEsVal = ClsPanorama.FblnEsValidoNumero(lobjValor, Single.MinValue, Single.MaxValue, True, EnuTipoValor.enuSingle)
                    Case EnuTipoValor.enuUShort
                        lblnEsVal = ClsPanorama.FblnEsValidoNumero(lobjValor, UShort.MinValue, UShort.MaxValue, True, EnuTipoValor.enuUShort)
                    Case EnuTipoValor.enuShort
                        lblnEsVal = ClsPanorama.FblnEsValidoNumero(lobjValor, Short.MinValue, Short.MaxValue, True, EnuTipoValor.enuUShort)
                    Case EnuTipoValor.enuLong
                        lblnEsVal = ClsPanorama.FblnEsValidoNumero(lobjValor, Long.MinValue, Long.MaxValue, True, EnuTipoValor.enuLong)
                    Case EnuTipoValor.enuULong
                        lblnEsVal = ClsPanorama.FblnEsValidoNumero(lobjValor, ULong.MinValue, ULong.MaxValue, True, EnuTipoValor.enuULong)
                    Case EnuTipoValor.enuString
                        lblnEsVal = ClsPanorama.FblnEsValidoString(lobjValor, 0, Short.MaxValue, True)
                End Select
                If Not lblnEsVal Then Exit For
            End If
        Next
        Return lblnEsVal
    End Function
    Friend Overridable Function FblnSonValidosDatosOrigen(adtbDatosOrigen As DataTable,
            astrColumnasRelacionadas As String(), ablnReinicie As Boolean,
            ByRef astrMens As String) As Boolean
        Return True
    End Function
#End Region

#Region "Procedimientos"
    Protected Friend Overridable Sub SActualice(ablnExigeRequeridos As Boolean)
        If BlnTengoCambios() Then
            Dim lblnNoHayError = False
            If ablnExigeRequeridos Then
                SAsigneVlaloresDefectoPro()
                If Not FblnEstanTodosOk() Then
#If DES = 0 Then
                    Throw New ErrorInesperadoPanLException("Propiedad no valida. " & HstrPropiedadNoValida & ".!!!")
#Else
                    Stop
                    Throw New ErrorInesperadoPanLException("Propiedad no valida. " & HstrPropiedadNoValida & ".!!!")
#End If
                End If
            End If
            Try
                GobjPanDat.SControleProcesoObj(True)
                If EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                    If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                        SDetermineReferencias()
                        SPuebleDataRowNuevo()
                        SInserteRegistro()
                    Else
                        GobjPanDat.SInicialiceTransaccion()
                        SActualiceDataRow()
                        SActualiceRegistro()
                        GobjPanDat.SConfirmeTransaccion()
                    End If
                End If
                STermineActualizacion()
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
                    If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                        GobjPanDat.SAborteTransaccion()
                    End If
                End If
                GobjPanDat.SControleProcesoObj(False)
            End Try
        ElseIf EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
            STermineActualizacion()
        End If
    End Sub
    ''' <summary>
    ''' Prepara la Clase para crear un nuevo objeto.
    ''' </summary>
    ''' <param name="aobjValorLlave">Array que identifica el nuevo objeto.</param>
    ''' <remarks>Solo es posible crear nuevos objetos desde un objeto navegable o para formar parte de
    ''' una colección.</remarks>
    Protected Friend Overridable Sub SCreeObj(aobjValorLlave() As Object)
        If (EnuPermisosObj And EnuPermisosDef.enuCrear) OrElse GblnImportando Then
            Try
                If FblnEsCreable(aobjValorLlave) Then
                    EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando
                    SVacie()
                    SInicialiceObj()
                    If HenuTipoObjeto = EnuModoInstanciaObjDef.enuNavegable OrElse
                            HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico Then
                        If Not IsNothing(DrwRegistroActual) Then
                            DrwRegistroActual = DrwRegistroActual.Table.NewRow
                        Else
                            If Not Me.HenuIdClase = EnuIdClasesPanDef.EnuImportar Then
                                Dim ldtbObj As DataTable =
                                        ClsPanorama.FdtbDataTable(HstrNombreTabla,
                                        {"*"}, {{"", ""}}, "")
                                DrwRegistroActual = ldtbObj.NewRow
                            End If
                        End If
                    End If
                    ObjValorLlave = aobjValorLlave
                End If
            Catch ex As PanDatException
                Throw
            Catch ex As PanLException
                Throw
            Catch ex As Exception
                Throw
            End Try
        End If
    End Sub
    Protected Friend Overridable Sub SInicialiceObj()
        ' Para asignar valores a las propiedades que tienen un valor predeterminado
        ' como la carpeta y la Copropiedad o aquellas otras propiedades las
        ' cuales requieren se inicializadas y validadas al momento de ser creadas
        For Each lobjPro As ClsCBPropiedad In ColPropiedades
            If Not lobjPro.BlnEsRequerido Then
                lobjPro.ObjValorPro = ClsPanorama.FobjValorNuloPropiedad(lobjPro)
            End If
        Next
    End Sub
    Protected Friend Overridable Sub SPrepareParaImportacion()
        GblnImportando = True
    End Sub
    Protected Friend Overridable Sub SModifique()
        If CType(EnuPermisosObj And EnuPermisosDef.enuModificar, Boolean) Then
            Try
                If FblnEsModificable() Then
                    EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                Else
                    Dim lstrMens = "El objeto no es modificable!"
                    SLevanteEventoNot(lstrMens, "", EnuIdMens.EnuNoModificable,
                            EnuSeveridadNot.EnuFalta)
                End If
            Catch ex As PanLException
                Throw
            Catch ex As PanDatException
                Throw
            Catch ex As Exception
                Throw
            End Try
        End If
    End Sub
    Public Sub SAnule()
        If MenuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            If CType(EnuPermisosObj And EnuPermisosDef.enuAnular, Boolean) Then
                EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                Dim lblnAnulo As Boolean
                Dim lblnNoHayError = False
                Try
                    GobjPanDat.SControleProcesoObj(True)
                    GobjPanDat.SInicialiceTransaccion()
                    lblnAnulo = SAnuleEnObj()
                    If lblnAnulo Then
                        SActualice(True)
                    Else
                        EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando
                    End If
                    SRefresqueObj()
                    lblnNoHayError = True
                Catch ex As PanLException
                    Throw
                Catch ex As PanDatException
                    Throw
                Catch ex As Exception
                    Throw
                Finally
                    If Not lblnNoHayError Then
                        SNormaliceEstado(True)
                        GobjPanDat.SAborteTransaccion()
                        GobjPanDat.SControleProcesoObj(False, True)
                    Else
                        GobjPanDat.SConfirmeTransaccion()
                        GobjPanDat.SControleProcesoObj(False)
                    End If
                End Try
            End If
        End If
    End Sub
    ''' <summary>
    ''' Devuelve True si se anulo el objeto
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Friend Overridable Function SAnuleEnObj() As Boolean
        Return False
        ' Puente para ejecutar el procedimiento sAnule en el objeto a ser anulado si es necesario
    End Function
    ''' <summary>
    ''' Suprime el presente objeto, de la base de datos.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Friend Overridable Function FblnSuprimio() As Boolean
        Dim lblnSuprimio = FblnPermitidoSuprimir()
        If lblnSuprimio Then
            Dim lblnNoHayError As Boolean
            GobjPanDat.SControleProcesoObj(True)
            Try
                If FblnEsSuprimible() Then
                    GobjPanDat.SElimineRegistro(HstrNombreTabla, ColNombresCamposRef, McolDatosRef)
                    If BlnEsNavegable Then
                        DtbTablaNavegacion.Rows.Remove(MdrwIdObjetoActual)
                    Else
                        DtbTablaColeccion.Rows.Remove(DrwRegistroActual)
                    End If
                    lblnSuprimio = True
                Else
                    lblnSuprimio = False
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
                    GobjPanDat.SControleProcesoObj(False)
                Else
                    GobjPanDat.SControleProcesoObj(False, True)
                End If
            End Try
        End If
        Return lblnSuprimio
    End Function
    Protected Friend Overridable Sub SLeaValores(ablnLeyendoOrigen As Boolean)
        If Not IsNothing(DrwRegistroActual) Then
            SVacie()
            For Each lobjProp As ClsCBPropiedad In ColPropiedades
                If Not String.IsNullOrEmpty(lobjProp.StrNombreCampoBD) Then
                    Dim lblnValorPropiedadInvalido = False
                    Dim lobjValorCampo As Object = Nothing
                    lobjProp.BlnLeyendoOrigen = ablnLeyendoOrigen
                    Try
                        lobjValorCampo = ClsPanorama.FobjValorCampo(DrwRegistroActual,
                                lobjProp)
                    Catch ex As ValorPropiedadInvalidoException
                        lblnValorPropiedadInvalido = True
                        Throw
                    Catch ex As Exception
                        Throw
                    Finally
                        If lblnValorPropiedadInvalido Then
                            lobjProp.ObjValorPro = ClsPanorama.FobjValorNuloPropiedad(lobjProp)
                        Else
                            lobjProp.ObjValorPro = lobjValorCampo
                        End If
                        If ablnLeyendoOrigen Then
                            lobjProp.BlnLeyendoOrigen = False
                        End If
                    End Try
                End If
            Next
            MblnExiste = FblnExisteObj()
            SDetermineReferencias()
        Else
            SVacie()
            MblnExiste = False
        End If
    End Sub
    Private Function FblnExisteObj() As Boolean
        Dim lblnExiste = True
        For Each lobjProp As ClsCBPropiedad In ColPropiedades
            If lobjProp.BlnEsLlave Then
                lblnExiste = lobjProp.BlnEsValido
                If Not lblnExiste Then Exit For
            End If
        Next
        Return lblnExiste
    End Function
    ''' <summary>
    ''' Lee de nuevo los valores del último objeto desde la base de datos.
    ''' </summary>
    ''' <remarks>El objeto debe estar en estado de Actualización</remarks>
    Friend Overridable Sub SRefresqueObj()
        If MenuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            Dim lblnNoHayError = False
            GobjPanDat.SControleProcesoObj(True)
            Try
                Select Case HenuTipoObjeto
                    Case EnuModoInstanciaObjDef.enuNavegable
                        MdtbTablaNavegacion = Nothing
                        If Not MblnEsAutonumerico Then
                            If FblnExisteLlave(MobjValorUltimaLlave) Then
                                SAbra(MobjValorUltimaLlave)
                            Else
                                SVayaAlPrimero()
                            End If
                        Else
                            SVayaAlPrimero()
                        End If
                    Case EnuModoInstanciaObjDef.enuUnico
                        MdtbTablaObjetoUnico = Nothing
                        SDetermineReferencias()
                        SAbra(MobjValorUltimaLlave)
                    Case EnuModoInstanciaObjDef.enuDeColeccion
                        ' En este caso solo refresca los valores de las propiedades a partir del datarow
                        ' sin leer la base de datos
                        SLeaValores(True)
                End Select
                lblnNoHayError = True
            Catch ex As PanLException
                Throw
            Catch ex As PanDatException
                Throw
            Catch ex As ArgumentNullException
                Throw
            Finally
                If lblnNoHayError Then
                    GobjPanDat.SControleProcesoObj(False)
                Else
                    GobjPanDat.SControleProcesoObj(False, True)
                End If
            End Try
        Else
            Throw New ErrorInesperadoPanLException("Estado inesperado del Objeto")
        End If
    End Sub
    Private Sub SResetPropiedades()
        For Each lobjProp As ClsCBPropiedad In ColPropiedades
            lobjProp.SAsigneCambio(False, False)
            lobjProp.SAsigneValorOriginal(lobjProp.ObjValorPro)
        Next
        MblnTengoCambios = False
    End Sub
    Protected Sub SAsigneVlaloresDefectoPro()
        For Each lobjProp As ClsCBPropiedad In ColPropiedades
            If Not lobjProp.BlnEsValido Then
                lobjProp.ObjValorPro = ClsPanorama.FobjValorNuloPropiedad(lobjProp)
            End If
        Next
    End Sub
    ''' <summary>
    ''' Deja todas las variables de los objetos de Propiedad en su valor Nulo y todas las demas variables
    ''' en su valor por defecto. Se debe utilizar para inicializar las variables de colecciones y de 
    ''' objetos de datos.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Friend Overridable Sub SVacie()
        BlnVaciandoObjeto = True
        MblnTengoCambios = False
        MblnExiste = False
        For Each lobjProp As ClsCBPropiedad In ColPropiedades
            lobjProp.SVacie(True)
        Next
        BlnVaciandoObjeto = False
    End Sub
    Protected Friend Overridable Sub SAbra(aobjIDLlave() As Object)
        GobjPanDat.SControleProcesoObj(True) ' Ok
        Dim lstrMens = String.Empty, lblnNoHayError = False
        Try
            Dim lblnIsNull = IsNothing(aobjIDLlave)
            If Not IsNothing(aobjIDLlave) Then
                For Each lobjObjeto As Object In aobjIDLlave
                    lblnIsNull = IsNothing(lobjObjeto)
                    If lblnIsNull Then Exit For
                Next
            End If
            If Not lblnIsNull Then
                Dim lstrPref = String.Empty, lentIdDcto = 0
                Dim lstrIdObjeto = String.Empty
                Dim lstrDatoIng = String.Empty
                If aobjIDLlave.Length = 4 Then
                    If Not String.IsNullOrEmpty(aobjIDLlave(2)) Then
                        lstrDatoIng = aobjIDLlave(2) & "-" & aobjIDLlave(3)
                    Else
                        lstrDatoIng = aobjIDLlave(3)
                    End If
                ElseIf aobjIDLlave.Length = 3 Then
                    lstrDatoIng = aobjIDLlave(2)
                Else
                    lstrDatoIng = aobjIDLlave(aobjIDLlave.Length - 1)
                End If
                If FblnEsValidoValorLlave(aobjIDLlave) Then
                    If aobjIDLlave.Length = 4 Then
                        lstrPref = aobjIDLlave(2)
                        lentIdDcto = aobjIDLlave(3)
                        lstrIdObjeto = ClsPanorama.FstrNumeroDcto(lstrPref, lentIdDcto)
                    ElseIf aobjIDLlave.Length = 3 Then
                        lstrIdObjeto = aobjIDLlave(2)
                    Else
                        lstrIdObjeto = aobjIDLlave(aobjIDLlave.Length - 1)
                    End If
                    If FblnExisteLlave(aobjIDLlave) Then
                        Select Case HenuTipoObjeto
                            Case EnuModoInstanciaObjDef.enuNavegable
                                Dim lstrFiltro As String = FstrFiltro(aobjIDLlave)
                                MdrwIdObjetoActual = DtbTablaNavegacion.Select(lstrFiltro)(0)
                                SObtegaDataRowObjetoNavegable(lstrFiltro)
                                SVacie()
                                If Not IsNothing(DrwRegistroActual) Then
                                    SLeaValores(True)
                                End If
                            Case EnuModoInstanciaObjDef.enuUnico
                                If Not IsNothing(DtbObjetoUnico(aobjIDLlave)) Then
                                    DrwRegistroActual = DtbObjetoUnico(aobjIDLlave).Rows(0)
                                    SLeaValores(True)
                                Else
                                    SVacie()
                                End If
                            Case EnuModoInstanciaObjDef.enuDeColeccion
                                'Un tipo de objeto de colección no se abre
                                SVacie()
                                Throw New ErrorInesperadoPanLException(
                                        "Tratando abrir un objeto de colección!")
                        End Select
                    Else
                        Dim lstrNomObj = FstrNombreDoc(HenuIdClase)
                        If String.IsNullOrEmpty(lstrNomObj) Then
                            lstrNomObj = HstrNombreClase
                        End If
                        SVacie()
                        lstrMens = lstrNomObj & " con la identificación '" & lstrIdObjeto &
                                    "' no existe"
                    End If
                Else
                    SVacie()
                    If Not IsNothing(aobjIDLlave(0)) Then
                        lstrMens = "El Dato ingresado, '" & lstrDatoIng & "', no es válido"
                    End If
                End If
            Else
                SVacie()
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
                If Not String.IsNullOrEmpty(lstrMens) Then
                    lstrMens &= ". Refresca la Ventana!"
                    SLevanteEventoNot(lstrMens, "", EnuIdMens.EnuAbrir,
                                EnuSeveridadNot.EnuInformacion)
                Else
                    SLevanteEventoNot("", "", EnuIdMens.EnuAbrir,
                                EnuSeveridadNot.EnuOk)
                End If
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Public Sub SVayaAlPrimero()
        GobjPanDat.SControleProcesoObj(True)
        If BlnEsNavegable Then
            If DtbTablaNavegacion.Rows.Count > 0 Then
                MdrwIdObjetoActual = DtbTablaNavegacion.Rows(0)
                SObtegaDataRowObjetoNavegable(FstrFiltro)
            Else
                SObtegaDataRowObjetoNavegable(FstrFiltro(MobjValorUltimaLlave))
                MdrwIdObjetoActual = Nothing
            End If
        Else
            If DtbTablaColeccion.Rows.Count > 0 Then
                DrwRegistroActual = (DtbTablaColeccion.Select)(0)
            End If
        End If
        SLeaValores(True)
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Public Sub SVayaAlAnterior()
        GobjPanDat.SControleProcesoObj(True)
        If DtbTablaNavegacion.Rows.Count > 0 Then
            Dim lentIndiceRegistro As Integer = DtbTablaNavegacion.Rows.IndexOf(MdrwIdObjetoActual)
            If lentIndiceRegistro > 0 Then
                MdrwIdObjetoActual = DtbTablaNavegacion.Rows.Item(lentIndiceRegistro - 1)
                SObtegaDataRowObjetoNavegable(FstrFiltro)
            End If
        Else
            DrwRegistroActual = Nothing
            MdrwIdObjetoActual = Nothing
        End If
        SLeaValores(True)
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Public Sub SVayaAlSiguiente()
        GobjPanDat.SControleProcesoObj(True)
        If DtbTablaNavegacion.Rows.Count > 0 Then
            Dim lentIndiceRegistro As Integer = DtbTablaNavegacion.Rows.IndexOf(MdrwIdObjetoActual)
            If lentIndiceRegistro < DtbTablaNavegacion.Rows.Count - 1 Then
                MdrwIdObjetoActual = DtbTablaNavegacion.Rows.Item(lentIndiceRegistro + 1)
                SObtegaDataRowObjetoNavegable(FstrFiltro)
            Else
                DrwRegistroActual = Nothing
                MdrwIdObjetoActual = Nothing
            End If
        Else
            DrwRegistroActual = Nothing
        End If
        SLeaValores(True)
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Public Sub SVayaAlUltimo()
        GobjPanDat.SControleProcesoObj(True)
        If DtbTablaNavegacion.Rows.Count > 0 Then
            Dim lentIndiceRegistro As Integer = DtbTablaNavegacion.Rows.IndexOf(MdrwIdObjetoActual)
            If lentIndiceRegistro < DtbTablaNavegacion.Rows.Count - 1 Then
                MdrwIdObjetoActual = DtbTablaNavegacion.Rows.Item(DtbTablaNavegacion.Rows.Count - 1)
                SObtegaDataRowObjetoNavegable(FstrFiltro)
            End If
        Else
            SObtegaDataRowObjetoNavegable(FstrFiltro(MobjValorUltimaLlave))
            MdrwIdObjetoActual = Nothing
        End If
        SLeaValores(True)
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Friend Function FstrCamposIndiceTabla() As String(,)
        Dim lshrLimiteSupIndice As Short = FshrLimiteSupIndice()
        Dim lstrCamposIndiceTabla(,) As String = Nothing
        If lshrLimiteSupIndice > -1 Then
            ReDim lstrCamposIndiceTabla(lshrLimiteSupIndice, 1)
            For Each lobjProp As ClsCBPropiedad In ColPropiedades
                With lobjProp
                    If .BlnEsLlave Then
                        If Not MblnEsAutonumerico Then
                            MblnEsAutonumerico = .BlnEsAutonumerico
                        End If
                        lstrCamposIndiceTabla(.BytPosLlave, 0) = lobjProp.StrNombreCampoBD
                        lstrCamposIndiceTabla(.BytPosLlave, 1) = lobjProp.StrOrdenIdice
                    End If
                End With
            Next
        End If
        Return lstrCamposIndiceTabla
    End Function
    Private Function FclmPrimaryKey() As DataColumn()
        Dim lclmPrimaryKey As DataColumn() = Nothing
        Dim lshrLimiteSupIndice As Short = FshrLimiteSupIndice()
        Dim lstrCamposIndiceTabla As String(,) = FstrCamposIndiceTabla()
        If lshrLimiteSupIndice > -1 Then
            ReDim lclmPrimaryKey(lshrLimiteSupIndice)
            For i = 0 To lshrLimiteSupIndice
                lclmPrimaryKey(i) = DtbTablaColeccion.Columns(lstrCamposIndiceTabla(i, 0))
            Next
        End If
        Return lclmPrimaryKey
    End Function
    Private Function FblnHayCambios() As Boolean
        Dim lblnCambioObjeto As Boolean = False
        If Not IsNothing(ColPropiedades) Then
            For Each lobjPropiedad As ClsCBPropiedad In ColPropiedades
                If lobjPropiedad.BlnCambio Then
                    lblnCambioObjeto = True
                End If
            Next
        End If
        Return lblnCambioObjeto
    End Function
    ' Manejo notificaciones
    Protected Sub SLevanteEventoNot(astrMensNot As String, astrMensEx As String,
            aenuIdMens As EnuIdMens, aenuSevNot As EnuSeveridadNot)
        Dim lobjNotiEven As New ClsNotiEventArgs
        lobjNotiEven.SRegistreNotifica(astrMensNot, astrMensEx, aenuIdMens, aenuSevNot)
        RaiseEvent EvnNotifica(Me, lobjNotiEven)
    End Sub
    Protected Friend Sub SLevanteEventoNot(aobjPropiedad As ClsCBPropiedad, astrMensNot As String,
            astrMensEx As String, aenuIdMens As EnuIdMens, aenuSevNot As EnuSeveridadNot)
        Dim lobjNotiEven As New ClsNotiEventArgs
        lobjNotiEven.SRegistreNotifica(astrMensNot, astrMensEx, aenuIdMens, aenuSevNot)
        RaiseEvent EvnNotifica(aobjPropiedad, lobjNotiEven)
    End Sub
    Friend Overridable Function FblnNotificaOk(aenuIdMensNot As EnuIdMens) As Boolean
        Dim lblnOk As Boolean = False
        Select Case aenuIdMensNot
            Case EnuIdMens.EnuAbrir
                lblnOk = BlnExiste
            Case EnuIdMens.EnuNoCreable
                lblnOk = FblnEsCreable()
            Case EnuIdMens.EnuNoModificable
                lblnOk = FblnEsModificable()
        End Select
        Return lblnOk
    End Function
#End Region

#Region "Funciones y procedimientos privados"
    Private Function FstrFiltro() As String
        Dim lstrfiltro As String = String.Empty
        Dim lenuTipoDato As EnuTipoValor
        If Not (IsNothing(DtbTablaNavegacion) OrElse IsNothing(MdrwIdObjetoActual)) Then
            For Each ldclColumna As DataColumn In DtbTablaNavegacion.Columns
                lenuTipoDato = FenuTipoDatoCampoBD(ldclColumna.ColumnName)
                Select Case lenuTipoDato
                    Case EnuTipoValor.enuByte, EnuTipoValor.enuDecimal, EnuTipoValor.enuDouble,
                            EnuTipoValor.enuInteger, EnuTipoValor.enuLong, EnuTipoValor.enuShort,
                            EnuTipoValor.enuSingle, EnuTipoValor.enuUInteger, EnuTipoValor.enuULong,
                            EnuTipoValor.enuUShort
                        lstrfiltro &= ldclColumna.ColumnName & " = " &
                                ClsPanorama.FobjValorCampo(MdrwIdObjetoActual(ldclColumna.ColumnName),
                                        lenuTipoDato)
                    Case EnuTipoValor.enuBoolean
                        lstrfiltro &= ldclColumna.ColumnName & " = " &
                                ClsPanorama.FobjValorCampo(MdrwIdObjetoActual(ldclColumna.ColumnName),
                                        lenuTipoDato)
                    Case EnuTipoValor.enuDate
                        Dim lstrFecha As String = CType(ClsPanorama.FobjValorCampo(MdrwIdObjetoActual(
                                ldclColumna.ColumnName), lenuTipoDato), String)
                        lstrFecha = ClsPanoramaDat.FstrFechaNormalizada(lstrFecha)
                        lstrfiltro &= ldclColumna.ColumnName & " = '" & lstrFecha & "'"
                    Case EnuTipoValor.enuDateTime
                        Dim lstrFecha As String = CType(ClsPanorama.FobjValorCampo(MdrwIdObjetoActual(
                                ldclColumna.ColumnName), lenuTipoDato), String)
                        lstrFecha = ClsPanoramaDat.FstrFechaHoraNormalizada(lstrFecha)
                        lstrfiltro &= ldclColumna.ColumnName & " = '" & lstrFecha & "'"
                    Case EnuTipoValor.enuString
                        lstrfiltro &= ldclColumna.ColumnName & " = '" &
                                ClsPanorama.FobjValorCampo(MdrwIdObjetoActual(ldclColumna.ColumnName),
                                        lenuTipoDato) & "'"
                End Select
                lstrfiltro &= " AND "
            Next
            lstrfiltro = lstrfiltro.Substring(0, lstrfiltro.Length - 5)
        End If
        Return lstrfiltro
    End Function
    Private Function FstrFiltro(aobjIdLlave() As Object) As String
        Dim lstrFiltro As String = String.Empty
        Dim lenuTipoDato As EnuTipoValor
        Dim lstrCamposIndice(,) As String = ColCamposIndexesTablas(HstrNombreTabla)
        Dim lobjValor As Object
        Dim i As Byte = 0
        For Each lstrCampo As String In lstrCamposIndice
            If lstrCampo.ToUpper <> "ASC" AndAlso lstrCampo.ToUpper <> "DESC" Then
                lenuTipoDato = FenuTipoDatoCampoBD(lstrCampo)
                If Not IsNothing(aobjIdLlave) Then
                    lobjValor = ClsPanorama.FobjValorCampo(aobjIdLlave(i), lenuTipoDato)
                    Select Case lenuTipoDato
                        Case EnuTipoValor.enuByte, EnuTipoValor.enuDecimal, EnuTipoValor.enuDouble,
                                EnuTipoValor.enuInteger, EnuTipoValor.enuLong, EnuTipoValor.enuShort,
                                EnuTipoValor.enuSingle, EnuTipoValor.enuUInteger, EnuTipoValor.enuULong,
                                EnuTipoValor.enuUShort, EnuTipoValor.enuBoolean
                            lstrFiltro &= lstrCampo & " = " & lobjValor
                        Case EnuTipoValor.enuDate
                            Dim lstrFecha As String = CType(lobjValor, String)
                            lstrFecha = ClsPanoramaDat.FstrFechaNormalizada(lstrFecha)
                            lstrFiltro &= lstrCampo & " = '" & lstrFecha & "'"
                        Case EnuTipoValor.enuDateTime
                            Dim lstrFecha As String = CType(lobjValor, String)
                            lstrFecha = ClsPanoramaDat.FstrFechaHoraNormalizada(lstrFecha)
                            lstrFiltro &= lstrCampo & " = '" & lstrFecha & "'"
                        Case EnuTipoValor.enuString
                            lstrFiltro &= lstrCampo & " = '" & lobjValor & "'"
                    End Select
                Else
                    lstrFiltro &= lstrCampo & " IS NULL"
                End If
                lstrFiltro &= " AND "
                i += 1
            End If
        Next
        lstrFiltro = lstrFiltro.Substring(0, lstrFiltro.Length - 5)
        Return lstrFiltro
    End Function
    Protected Shared Function FblnValorLlaveEsNull(aobjvalorLlave As Object()) As Boolean
        Dim lblnValorLlaveIsNull = aobjvalorLlave Is Nothing OrElse (aobjvalorLlave.Length = 0)
        If Not lblnValorLlaveIsNull Then
            Dim lblnTodosNull = True
            For i = 0 To aobjvalorLlave.Length - 1
                If Not IsNothing(aobjvalorLlave(i)) Then
                    lblnTodosNull = False
                    Exit For
                End If
            Next
            lblnValorLlaveIsNull = lblnTodosNull
        End If
        Return lblnValorLlaveIsNull
    End Function
    ''' <summary>
    ''' Actualiza el datarow del objeto actual con los cambios hechos a los valores de las propiedades
    ''' del objeto y puebla las colecciones "mcolNombresCamposCambio" y "mcolDatosNuevos" y registra los cambios en 
    ''' el log de la aplicación
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SActualiceDataRow()
        McolNombresCamposCambio.Clear()
        McolDatosNuevos.Clear()
        For Each lobjProp As ClsCBPropiedad In ColPropiedades
            If Not String.IsNullOrEmpty(lobjProp.StrNombreCampoBD) Then
                If lobjProp.BlnCambio Then
                    If IsNothing(lobjProp.ObjValorPro) Then
                        DrwRegistroActual(lobjProp.StrNombreCampoBD) = DBNull.Value
                    Else
                        DrwRegistroActual(lobjProp.StrNombreCampoBD) = lobjProp.ObjValorPro
                    End If
                    McolNombresCamposCambio.Add(lobjProp.StrNombreCampoBD)
                    McolDatosNuevos.Add(lobjProp.ObjValorPro, lobjProp.StrNombreCampoBD)
                    If lobjProp.BlnRegistrarLogCambio Then
                        SRegistreCambioLogApp(lobjProp)
                    End If
                End If
            End If
        Next
    End Sub
    Private Sub SRegistreCambioLogApp(aobjProp As ClsCBPropiedad)
        If Not GenuTipoInstanciamiento = EnuTipoInstanciamiento.enuInstalacion Then
            GobjPanorama.SRegistreCambioLogApp(StrIdObjeto, HstrNombreClase, aobjProp.StrNombre,
                aobjProp.ObjValorOriginal.ToString, aobjProp.ObjValorNuevo.ToString)
        End If
    End Sub
    ''' <summary>
    ''' Establece el valor de del Array mobjValorUltimaLlave() para el objeto actual y 
    ''' puebla la colección mcolDatosRef para permitir referenciar el objeto actual en la BD.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SDetermineReferencias()
        MobjValorUltimaLlave = Nothing
        McolDatosRef.Clear()
        Dim lshrLimiteSupIndice As Short = FshrLimiteSupIndice()
        If lshrLimiteSupIndice > -1 Then
            ReDim MobjValorUltimaLlave(lshrLimiteSupIndice)
            For Each lobjProp As ClsCBPropiedad In ColPropiedades
                If lobjProp.BlnEsLlave Then
                    MobjValorUltimaLlave(lobjProp.BytPosLlave) = lobjProp.ObjValorPro
                End If
            Next
            For i = 0 To lshrLimiteSupIndice
                McolDatosRef.Add(MobjValorUltimaLlave(i))
            Next
        End If
    End Sub
    Private Sub SObtegaDataRowObjetoNavegable(astrFiltro As String)
        Dim lblnNoHayError As Boolean = False
        Using ldstObjeto As New DataSet
            Dim ldtbObjetoActual As DataTable = Nothing
            Try
                GobjPanDat.SdsDataSet(ldstObjeto, HstrNombreTabla, {"*"}, {{"", ""}}, astrFiltro, True,
                        Array.Empty(Of String))
                lblnNoHayError = True
            Catch ex As ArgumentNullException
                Throw
            Catch ex As PanDatException
                Throw
            Catch ex As Exception
                Throw
            Finally
                If lblnNoHayError Then
                    ldtbObjetoActual = ldstObjeto.Tables(0)
                    If ldtbObjetoActual.Rows.Count > 0 Then
                        DrwRegistroActual = ldtbObjetoActual.Rows(0)
                    Else
                        DrwRegistroActual = Nothing
                    End If
                Else
                    DrwRegistroActual = Nothing
                End If
            End Try
        End Using
    End Sub
    ''' <summary>
    ''' Crea un nuevo datarow a partir de dtbObjetoActual y lo puebla con los datos de las propiedades
    ''' del objeto. Además puebla las colecciones "mcolNombresCamposCambio" y "mcolDatosNuevos"
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SPuebleDataRowNuevo()
        McolNombresCamposCambio.Clear()
        McolDatosNuevos.Clear()
        For Each lobjProp As ClsCBPropiedad In ColPropiedades
            If Not String.IsNullOrEmpty(lobjProp.StrNombreCampoBD) Then
                If Not lobjProp.BlnEsAutonumerico Then
                    If Not IsNothing(lobjProp.ObjValorPro) Then
                        DrwRegistroActual(lobjProp.StrNombreCampoBD) = lobjProp.ObjValorPro
                        McolNombresCamposCambio.Add(lobjProp.StrNombreCampoBD)
                        McolDatosNuevos.Add(lobjProp.ObjValorPro, lobjProp.StrNombreCampoBD)
                    End If
                End If
            End If
        Next
    End Sub
    Private Sub SInserteRegistro()
        If FblnEsCreable(MobjValorUltimaLlave) Then
            GobjPanDat.SInserteRegistro(HstrNombreTabla, McolNombresCamposCambio, McolDatosNuevos)
        Else
            Dim lstrMens = "El objeto no fue creado porque ya existe uno igual!"
            SLevanteEventoNot(lstrMens, "", EnuIdMens.EnuNoCreable,
                    EnuSeveridadNot.EnuFalta)
        End If
    End Sub
    Private Sub SActualiceRegistro()
        If McolNombresCamposCambio.Count > 0 Then
            Dim lentRegistrosAfectados As Integer = GobjPanDat.SActualiceRegistro(HstrNombreTabla,
                    McolNombresCamposCambio, McolDatosNuevos, ColNombresCamposRef, McolDatosRef)
            If lentRegistrosAfectados = 0 Then
                Throw New ErrorInesperadoPanLException("No se actualizo el Registro!")
            End If
        End If
    End Sub
    Private Sub STermineActualizacion()
        If MenuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Select Case HenuTipoObjeto
                Case EnuModoInstanciaObjDef.enuNavegable, EnuModoInstanciaObjDef.enuUnico, EnuModoInstanciaObjDef.None
                    SNormaliceEstado(True)
                Case EnuModoInstanciaObjDef.enuDeColeccion
                    SNormaliceEstado(True)
                    If Not MblnEsAutonumerico AndAlso HenuIdClase <>
                            EnuIdClasesPanDef.EnuCenutiliOriCop Then
                        DtbTablaColeccion.Rows.Add(DrwRegistroActual)
                    End If
            End Select
        Else
            SNormaliceEstado(True)
        End If
        SResetPropiedades()
    End Sub
#End Region
End Class