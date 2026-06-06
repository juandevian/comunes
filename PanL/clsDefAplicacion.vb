Friend Class ClsDefAplicacion
#Region "Definiciones"
    Inherits PanL.ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "PanTblDefAplicaciones"
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Definición de Aplicación.
    ''' </summary>
    ''' <param name="aenuModoInstanciaObj">Indica si se instancia como un objeto navegable o como un Objeto único.</param>
    ''' <remarks>Si se instancia como un objeto navegable, se crea un datatable que contiene las columnas de
    ''' la llave con las llaves de todos los objetos y queda a la espera de que se indique que objeto abrir.
    ''' Si se instancia como un objeto único, queda a la espera de recibir el valor de los campos de la llave 
    ''' para abrir dicho objeto. </remarks>
    Public Sub New(aenuModoInstanciaObj As enuModoInstanciaObjDef)
        If aenuModoInstanciaObj = enuModoInstanciaObjDef.enuDeColeccion Then
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        End If
        Dim lstrCamposSelect As String()
        HobjPadre = Nothing
        hblnEsCreable = False
        hblnEsModificable = False
        HblnEsSuprimible = False
        hblnEsAnulable = False
        If aenuModoInstanciaObj = enuModoInstanciaObjDef.enuNavegable Then
            lstrCamposSelect = {objIdAppDefShr.strNombreCampoBD}
        Else
            henuTipoObjeto = enuModoInstanciaObjDef.enuUnico
            lstrCamposSelect = {"*"}
        End If
        hcolTablas.Add(MCSTRNOMBRETABLA)
        hcolCamposSelect.Add(lstrCamposSelect)
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
            Return EnuIdClasesPanDef.enuDefAplicacion
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Def. Aplicación"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjNombreAppStr As New ClsNombreAppStr(Me)
    Friend ReadOnly Property ObjEdicionStr As New ClsEdicionStr(Me)
    Friend ReadOnly Property ObjPrefijoAppStr As New ClsPrefijoAppStr(Me)
    Friend ReadOnly Property ObjPrefijoEdicionStr As New ClsPrefijoEdicionStr(Me)
    Friend ReadOnly Property ObjIdAppDefShr As New ClsIdAppDefShr(Me)
    Friend ReadOnly Property ObjPrefijoTablasStr As New ClsPrefijoTablasStr(Me)
    Friend ReadOnly Property ObjTipoLicenciamientoByt As New ClsTipoLicenciamientoByt(Me)
    Friend ReadOnly Property ObjUsaCarpetasBln As New ClsUsaCarpetasBln(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjEdicionStr)
                HcolPropiedades.Add(ObjIdAppDefShr)
                HcolPropiedades.Add(ObjNombreAppStr)
                HcolPropiedades.Add(ObjPrefijoAppStr)
                HcolPropiedades.Add(ObjPrefijoEdicionStr)
                HcolPropiedades.Add(ObjPrefijoTablasStr)
                HcolPropiedades.Add(ObjTipoLicenciamientoByt)
                HcolPropiedades.Add(ObjUsaCarpetasBln)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras Propiedades"
    Friend ReadOnly Property StrNombreCompleto As String
        Get
            If blnExiste Then
                Return objNombreAppStr.objValorPro & " " & objEdicionStr.objValorPro
            Else
                Return ""
            End If
        End Get
    End Property
    Friend ReadOnly Property StrNombreTipoLicenciamiento As String
        Get
            Return ClsAdministrador.FstrNombreDatoConstantePan(
                    EnuGrupoConstantesPanDef.enuTipoLicenciamiento,
                    ObjTipoLicenciamientoByt.ObjValorPro)
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Friend Overrides Function FblnEsSuprimible() As Boolean
        Return False
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsEdicionStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Edicion"
        HshrLongitud = 20
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "Edicion"
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 3,
                    ShrLongitud, BlnEsRequerido)
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
Friend Class ClsIdAppDefShr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdApp"
        HenuTipoValor = EnuTipoValor.enuUShort
        HstrNombreCampoBd = "IdAplicacion"
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 0
    End Sub
    Public Overrides Sub SValide()
        Dim lobjLlavePrincipal(0) As Object
        Dim lblnEsValido As Boolean
        Dim lobjPadre As ClsDefAplicacion = ObjPadre
        lblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, EnuListaAplicaciones.EnuAdministrador,
                EnuListaAplicaciones.EnuOrionCop, BlnEsRequerido, EnuTipoValor.enuShort)
        If Not BlnLeyendoOrigen Then
            HstrMens = String.Empty
            If lblnEsValido Then
                lobjLlavePrincipal(0) = HobjValorNew
                If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                    If Not lobjPadre.FblnEsCreable(lobjLlavePrincipal) Then
                        HstrMens = "La Id. de la Aplicacion ingresada ya existe!"
                        lblnEsValido = False
                    End If
                ElseIf lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                    If HobjValorNew <> HobjValorPro OrElse (Not lobjPadre.BlnExiste) Then
                        lobjPadre.SAbra(lobjLlavePrincipal)
                    End If
                    If Not lobjPadre.BlnExiste Then
                        HstrMens = "La Id. de la Aplicacion ingresada no existe!"
                        If Not IsNothing(lobjPadre.ObjValorUltimaLlave) Then
                            lobjPadre.SAbra(lobjPadre.ObjValorUltimaLlave)
                        Else
                            lblnEsValido = False
                        End If
                    End If
                End If
            Else
                HstrMens = "La Id. de la Aplicacion ingresada no es válida!"
            End If
            If Not String.IsNullOrEmpty(HstrMens) Then
                SNotifiqueDatInv()
            End If
        End If
        HblnEsValido = lblnEsValido
    End Sub
    Private Sub EPosCambio() Handles Me.EvnPosCambio
        Dim lobjPadre As ClsDefAplicacion = ObjPadre
        With lobjPadre
            .ObjEdicionStr.SValide()
            .ObjTipoLicenciamientoByt.SValide()
            .ObjPrefijoAppStr.SValide()
            .ObjPrefijoEdicionStr.SValide()
        End With
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsNombreAppStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Nombre"
        HshrLongitud = 20
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "Nombre"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 4,
                ShrLongitud, BlnEsRequerido)
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
Friend Class ClsPrefijoAppStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "PrefijoApp"
        HshrLongitud = 3
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "PrefijoAplicacion"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, ShrLongitud,
                ShrLongitud, BlnEsRequerido)
        If lblnEsValido Then
            If Not IsNothing(HobjValorNew) Then
                HobjValorNew = HobjValorNew.ToString.ToUpper
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
Friend Class ClsPrefijoEdicionStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "PrefijoEdicion"
        HshrLongitud = 2
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "PrefijoEdicion"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        If lblnEsValido Then
            If Not IsNothing(HobjValorNew) Then
                HobjValorNew = HobjValorNew.ToString.ToUpper
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
Friend Class ClsPrefijoTablasStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "PrefijoTablas"
        HshrLongitud = 3
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "PrefijoTablas"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoString(HobjValorNew, 3,
                ShrLongitud, BlnEsRequerido)
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
Friend Class ClsTipoLicenciamientoByt
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TipoLicenciamiento"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = "IDTblTipoLicenciamiento"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsDefAplicacion = ObjPadre
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew,
                EnuTipoLicenciamientoDef.enuPorEstacion,
                EnuTipoLicenciamientoDef.enuPorInstancia, HblnEsRequerido)
        If Not lblnEsValido Then
            If HobjValorNew = EnuTipoLicenciamientoDef.None AndAlso
                    (lobjPadre.ObjIdAppDefShr.ObjValorPro = EnuListaAplicaciones.EnuAdministrador) Then
                lblnEsValido = True
            End If
        ElseIf HobjValorNew <> EnuTipoLicenciamientoDef.None AndAlso
                    (lobjPadre.ObjIdAppDefShr.ObjValorPro = EnuListaAplicaciones.EnuAdministrador) Then
            lblnEsValido = False
        End If
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return CByte(HobjValorPro).ToString
        End If
    End Function
End Class
Friend Class ClsUsaCarpetasBln
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "UsaCarpetas"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = "UsaCarpetas"
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsDefAplicacion = ObjPadre
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        If lblnEsValido Then
            If HobjValorNew AndAlso lobjPadre.ObjIdAppDefShr.ObjValorPro =
                    EnuListaAplicaciones.EnuAdministrador Then
                lblnEsValido = False
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
#End Region