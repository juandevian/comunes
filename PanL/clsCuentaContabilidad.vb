Friend Class ClsCuentaContabilidad
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "PanCuentasContabilidad"
    '
    Private ReadOnly MobjPadre As ClsCarpeta = Nothing
#End Region

#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Cuenta de Contabilidad.
    ''' </summary>
    ''' <param name="aenuModoInstanciaObj">Indica si se instancia como un objeto navegable o como un objeto único</param>
    ''' <remarks>Si se instancia como un objeto navegable, se crea un datatable que contiene las columnas de
    ''' la llave con las llaves de todos los objetos y queda a la espera de que se indique que objeto abrir.
    ''' Si se instancia como un objeto único, queda a la espera de recibir el valor de los campos de la llave 
    ''' para abrir dicho objeto. </remarks>
    Friend Sub New(aobjPadre As ClsCarpeta, aenuModoInstanciaObj As EnuModoInstanciaObjDef)
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuDeColeccion Then
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        End If
        Dim lstrCamposSelect As String()
        HobjPadre = aobjPadre
        MobjPadre = aobjPadre
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuNavegable Then
            Dim lstrFiltro = ClsIdCarpetaCuentaShr.SstrNombreCampoBd & " = " &
                    aobjPadre.ObjIdCarpetaShr.ToString()
            HblnEsAnulable = False
            HcolFiltros.Add(lstrFiltro)
            lstrCamposSelect = {ObjIdCarpetaCuentaShr.StrNombreCampoBD, ObjIdCuentaContStr.StrNombreCampoBD}
        Else
            HblnEsCreable = False
            HblnEsModificable = False
            HblnEsSuprimible = False
            HblnEsAnulable = False
            HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
            lstrCamposSelect = {"*"}
        End If
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
    End Sub
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Carpeta al cual pertenece la colección de cuentas de 
    ''' contabilidad de la cual este objeto forma parte</param>
    ''' <param name="adrwCuentaCont">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsCarpeta, adrwCuentaCont As DataRow)
        HobjPadre = aobjPadre
        MobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        '
        DrwRegistroActual = adrwCuentaCont
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
    Public Shared ReadOnly Property SstrNombreTabla As String
        Get
            Return MCSTRNOMBRETABLA
        End Get
    End Property
    Protected Friend Overrides ReadOnly Property HenuIdClase As EnuIdClasesPanDef
        Get
            Return EnuIdClasesPanDef.enuCuentaContabilidad
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Cuenta de Contabilidad"
        End Get
    End Property
    Friend Overrides ReadOnly Property HstrNombreObj As String
        Get
            Return Chr(34) & ObjNombreCuentaStr.ObjValorPro & Chr(34)
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjIdCarpetaCuentaShr As New ClsIdCarpetaCuentaShr(Me)
    Friend ReadOnly Property ObjIdCuentaContStr As New ClsIdCuentaContStr(Me)
    Friend ReadOnly Property ObjNombreCuentaStr As New ClsNombreCuentaStr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjIdCarpetaCuentaShr)
                HcolPropiedades.Add(ObjIdCuentaContStr)
                HcolPropiedades.Add(ObjNombreCuentaStr)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#End Region

#Region "Procedimientos y funciones invalidantes"
    Public Overrides Function FblnEsAnulable() As Boolean
        Return MyBase.FblnEsAnulable()
    End Function
    Public Overrides ReadOnly Property BlnEsCreable As Boolean
        Get
            Return True
        End Get
    End Property
    Friend Overrides Function FblnEsSuprimible() As Boolean
        Dim lstrNombreColumnas As String() = {ObjIdCarpetaCuentaShr.StrNombreCampoBD,
                ObjIdCuentaContStr.StrNombreCampoBD}
        Dim lstrCondiciones As String() = {" = " & ObjIdCarpetaCuentaShr.ToString, " = '" &
                ObjIdCuentaContStr.ToString & "'"}
        Dim lstrTablasExcluir As String() = {SstrNombreTabla}
        Dim lblnEsSuprimible As Boolean = FblnPermitidoSuprimir()
        lblnEsSuprimible = lblnEsSuprimible AndAlso
                ClsPanorama.FblnEsEliminableReg(lstrTablasExcluir, lstrNombreColumnas,
                        lstrCondiciones, True, True)
        If lblnEsSuprimible Then
            lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg(lstrTablasExcluir,
                lstrNombreColumnas, lstrCondiciones, True, False)
        End If
        Return lblnEsSuprimible
    End Function
    Protected Friend Overrides Sub SInicialiceObj()
        ObjIdCarpetaCuentaShr.ObjValorPro = MobjPadre.ObjIdCarpetaShr.ObjValorPro
    End Sub
    Protected Friend Overrides Sub SPrepareParaImportacion()
        MyBase.SPrepareParaImportacion()
        If IsNothing(DrwRegistroActual) Then
            Dim lobjCtaCon As ClsCuentaContabilidad = MobjPadre.ObjCuentaContabilidad("")
            DrwRegistroActual = lobjCtaCon.DrwRegistroActual.Table.NewRow
        Else
            DrwRegistroActual = DrwRegistroActual.Table.NewRow
        End If
    End Sub
    Protected Friend Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        ObjIdCarpetaCuentaShr.ObjValorPro = MobjPadre.ObjIdCarpetaShr.ObjValorPro
        MyBase.SActualice(ablnExigeRequeridos)
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdCuentaContStr.ToString
        End Get
    End Property
#End Region

#Region "Procedimientos del objeto"
    Friend Overrides Function FblnSonValidosDatosOrigen(adtbOrigen As DataTable,
            astrColumnasRelacionadas As String(), ablnReinicie As Boolean,
            ByRef astrMens As String) As Boolean
        If adtbOrigen Is Nothing Then
            Throw New ArgumentNullException(NameOf(adtbOrigen))
        End If
        Dim lstrColFalta = FstrColumnasFalta(adtbOrigen, astrColumnasRelacionadas)
        Dim lstrMens = String.Empty, i = 0, lstrColumnaOrigen As String
        Dim lblnEsValido = String.IsNullOrEmpty(lstrColFalta)
        If Not lblnEsValido Then
            lstrMens = "Las siguientes columnas faltan o tienen nombre errado: '" &
                    lstrColFalta & "'"
        Else
            Dim lshrIdCarpeta As Short, lstrNombre As String, lstrIdCtaCont As String
            For Each ldrwOrigen As DataRow In adtbOrigen.Rows
                i += 1
                lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                        ClsIdCarpetaShr.SstrNombreCampoBd)
                lshrIdCarpeta = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                        EnuTipoValor.EnuShort)
                lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                        ClsNombreCuentaStr.SstrNombreCampoBd)
                lstrNombre = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                        EnuTipoValor.EnuString)
                lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                        ClsIdCuentaContStr.SstrNombreCampoBd)
                lstrIdCtaCont = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                        EnuTipoValor.EnuString)
                If Not (lshrIdCarpeta = 0 AndAlso String.IsNullOrEmpty(lstrNombre) AndAlso
                        String.IsNullOrEmpty(lstrIdCtaCont)) Then
                    lblnEsValido = lshrIdCarpeta > 0 AndAlso lshrIdCarpeta < Short.MaxValue AndAlso
                            lshrIdCarpeta = GshrIdCarpeta
                    If Not lblnEsValido Then
                        lstrMens = "El Dato de la " & Chr(34) & "Carpeta" & Chr(34) &
                                " en la Fila " & i.ToString() & " no es valido!"
                        Exit For
                    End If
                    lblnEsValido = Not String.IsNullOrEmpty(lstrNombre)
                    If Not lblnEsValido Then
                        lstrMens = "El Nombre de la Cuenta en la Fila" & i.ToString() & " no es valido!"
                        Exit For
                    End If
                    lblnEsValido = Not String.IsNullOrEmpty(lstrIdCtaCont)
                    If Not lblnEsValido Then
                        lstrMens = "El Número de la Cuenta en la Fila " & i.ToString() & " no es valido!"
                        Exit For
                    End If
                End If
            Next
        End If
        If Not lblnEsValido Then
            astrMens = lstrMens & " - Las Cuentas no fueron importadas!"
        End If
        Return lblnEsValido
    End Function
    Friend Function FentIndiceCta(adtbCuentas As DataTable, astrNombreCta As String,
            aentInd As Integer) As Integer
        Dim lstrMens = String.Empty, lentIndice = -1
        Dim lstrNomCta = "*" & UCase(astrNombreCta) & "*"
        If adtbCuentas IsNot Nothing AndAlso adtbCuentas.Rows.Count > 0 Then
            Dim lstrFiltro = ClsNombreCuentaStr.SstrNombreCampoBd & " LIKE '" &
                        lstrNomCta & "'"
            Dim ldrwCuentas() As DataRow = adtbCuentas.Select(lstrFiltro)
            If ldrwCuentas.Length > aentInd - 1 Then
                Dim ldrwCuenta As DataRow = ldrwCuentas(aentInd - 1)
                lentIndice = adtbCuentas.Rows.IndexOf(ldrwCuenta)
            ElseIf aentInd > 1 Then
                lstrMens = "No hay más Cuentas que coincidad con este Nombre!"
            Else
                lstrMens = "Una Cuentas Contable con este Nombre no existe!"
            End If
        Else
            lstrMens = "Aún no hay Cuentas Contable creadas!"
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEventoNot(lstrMens, "", 0, EnuSeveridadNot.EnuInformacion)
        End If
        Return lentIndice
    End Function
    Private Shared Function FstrColumnasFalta(adtbOrigen As DataTable,
                astrColumnasRelacionadas As String()) As String
        Dim lstrColumnaFal = String.Empty, lstrNomColumnaOrigen As String
        Dim lstrColumnaDestino As String
        Dim lcolColumnasDestino As New Collection From {
            {ClsIdCarpetaShr.SstrNombreCampoBd, ClsIdCarpetaShr.SstrNombreCampoBd},
            {ClsNombreCuentaStr.SstrNombreCampoBd, ClsNombreCuentaStr.SstrNombreCampoBd},
            {ClsIdCuentaContStr.SstrNombreCampoBd, ClsIdCuentaContStr.SstrNombreCampoBd}
        }
        For Each ldclOri As DataColumn In adtbOrigen.Columns
            lstrNomColumnaOrigen = ldclOri.ColumnName
            lstrColumnaDestino = FstrColumnaDestino(astrColumnasRelacionadas, lstrNomColumnaOrigen)
            If lcolColumnasDestino.Contains(lstrColumnaDestino) Then
                lcolColumnasDestino.Remove(lstrColumnaDestino)
            End If
        Next
        Dim i = 1
        If lcolColumnasDestino.Count > 0 Then
            For Each lstrCol As String In lcolColumnasDestino
                If i = lcolColumnasDestino.Count Then
                    lstrColumnaFal &= lstrCol
                Else
                    lstrColumnaFal &= lstrCol & ", "
                End If
            Next
        End If
        Return lstrColumnaFal
    End Function
#End Region
End Class

#Region "Clases de Propiedad"
Friend Class ClsIdCarpetaCuentaShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCarpeta"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdCarpeta"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 0
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsCuentaContabilidad = ObjPadre
        Dim lobjAbuelo As ClsCarpeta = lobjPadre.ObjPadre
        Dim lblnEsValido As Boolean
        lblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                    BlnEsRequerido, EnuTipoValor)
        If Not BlnLeyendoOrigen Then
            If lblnEsValido Then
                lblnEsValido = (HobjValorNew = lobjAbuelo.ObjIdCarpetaShr.ObjValorPro)
            End If
            If Not lblnEsValido Then
                HstrMens = "La Id. de la Carpeta ingresada no es válida!"
                SNotifiqueDatInv()
            End If
        End If
        HblnEsValido = lblnEsValido
    End Sub
    Private Sub EPosSetValor() Handles Me.EvnPosSetValor
        Dim lobjPadre As ClsCuentaContabilidad = ObjPadre
        If BlnEsValido Then
            If Not IsNothing(lobjPadre.ObjIdCuentaContStr.ObjValorPro) Then
                lobjPadre.ObjIdCuentaContStr.SValide()
            End If
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

Friend Class ClsIdCuentaContStr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsCuentaContabilidad = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCuentaCont"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdCuentaCont"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 1
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1,
                ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            If Not BlnLeyendoOrigen Then
                HstrMens = String.Empty
                HblnEsValido = IsNumeric(HobjValorNew)
                If Not HblnEsValido Then
                    HstrMens = "La Cuenta de Contabilidad debe contener solo Números!"
                    SNotifiqueDatInv()
                    Exit Sub
                End If
                With MobjPadre
                    Dim lblnExisteCtaCon = ClsCarpeta.FblnExisteCuentaCont(HobjValorNew)
                    If .EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                        If .ObjIdCarpetaCuentaShr.BlnEsValido Then
                            HblnEsValido = Not lblnExisteCtaCon
                            If lblnExisteCtaCon Then
                                HstrMens = "La Cuenta de Contabilidad ingresada ya existe!"
                            End If
                        End If
                    Else
                        If Not lblnExisteCtaCon Then
                            HstrMens = "La Cuenta de Contabilidad ingresada no existe!"
                        End If
                        HobjValorNew = HobjValorPro
                    End If
                End With
            End If
        Else
            If (MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso
                        Not String.IsNullOrEmpty(HobjValorNew.ToString)) OrElse
                        MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                HstrMens = "El valor ingresado, '" & HobjValorNew.ToString & "', no es valido!"
            End If
        End If
        If Not String.IsNullOrEmpty(HstrMens) Then
            SNotifiqueDatInv()
        End If
    End Sub
    Public Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(ObjValorPro) Then
            Return HobjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class

Friend Class ClsNombreCuentaStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Nombre"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Nombre"
        HshrLongitud = 100
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew,
                4, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            HobjValorNew = HobjValorNew.ToString.TrimEnd
            HobjValorNew = HobjValorNew.ToString.ToUpper
        Else
            HstrMens = "El Nombre debe tener más de tres Caracteres!"
            SNotifiqueDatInv()
        End If
    End Sub
    Public Shared ReadOnly Property SstrNombreCampoBd As String
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
#End Region