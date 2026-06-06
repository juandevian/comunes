Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Threading
Public MustInherit Class ClsFormInterface
#Region "Definiciones"
#Region "Herencia e interfaz"
    ' Herencia
    Inherits System.Windows.Window
    Implements IWinPanoramaIU

    ' Procedimientos que deben tener las Ventanas que heredan de esta clase.
    Protected MustOverride Sub SLoad() Implements IWinPanoramaIU.SLoad
    Protected MustOverride Sub SInicialiceObjeto() Implements IWinPanoramaIU.SInicialiceObjeto
    Protected MustOverride Sub SInicialiceControles() Implements IWinPanoramaIU.SInicialiceControles

    Protected MustOverride Sub SMuestreDatos() Implements IWinPanoramaIU.SMuestreDatos
    Protected MustOverride Sub SValide() Implements IWinPanoramaIU.SValide
    Protected MustOverride Sub SRegistre() Implements IWinPanoramaIU.SRegistre
    Protected MustOverride Sub SConfigureMenuesPropios() Implements IWinPanoramaIU.SConfigureMenuesPropios
    Protected MustOverride ReadOnly Property StrNombreVentana As String Implements IWinPanoramaIU.StrNombreVentana
    Protected Friend MustOverride ReadOnly Property EnuIdVentana As EnuIdVentanaDef Implements IWinPanoramaIU.EnuIdVentana
#End Region

#Region "Enumeradores"
    Private Enum EnuTipoMensaje As Integer
        enuNormal = 0
        enuMensaje
        enuCamposInsatisfechos
    End Enum
#End Region

#Region "Delegados"
    Private Delegate Sub SdgtActualizaNoti(dp As DependencyProperty, Content As Object)
    Private MdgtLblActualizaNot As SdgtActualizaNoti = Nothing
#End Region

#Region "Variables"
    Public WithEvents HwinBusqueda As WinBusqueda = Nothing
    Private McolControlesLlave As Collection = Nothing
    Private MctlControlDespuesDeLlave As Controls.Control = Nothing
    Private ReadOnly McolControlesRestringidos As New Collection
    Private McolLabels As Collection = Nothing
    Private MgrdGridPrincipal As Grid = Nothing
    Private McolPermisosWin As Collection = Nothing
    Protected WithEvents HtlbMiBarraHerramientas As ToolBar = Nothing
    ' Campos Heredables
    Protected HenuIdVentana As EnuIdVentanaDef = EnuIdVentanaDef.None
    Protected HblnSeEstaCerrando As Boolean = False
    Protected HblnLogOnRegistrado As Boolean = False
    Protected HblnCancelando As Boolean = False
    Protected HblnCargandoForma As Boolean = False
    Protected HblnMostrandoDatos As Boolean = False
    ' Objetos Heredables
    Protected Friend WinPadre As ClsFormInterface = Nothing
    Protected Friend WithEvents ObjObjetoWin As ClsCBObjetoPan = Nothing
    Protected Friend WithEvents ObjHijoObjWin As ClsCBObjetoPan = Nothing
    Protected Friend BlnVentanaAux As Boolean = False
    ' Propiedades
    Protected HbttCrear As Button = Nothing
    Protected HbttModificar As Button = Nothing
    Protected HbttSuprimir As Button = Nothing
    Protected HbttAnular As Button = Nothing
    Protected HbttGuardar As Button = Nothing
    Protected HbttRefrescar As Button = Nothing
    Protected HbttAlPrimero As Button = Nothing
    Protected HbttAlSiguiente As Button = Nothing
    Protected HbttAlAnterior As Button = Nothing
    Protected HbttAlUltimo As Button = Nothing
    Protected HbttBuscar As Button = Nothing
    Protected HbttCerrar As Button = Nothing
    Protected HbttCalculadora As Button = Nothing
    Protected HbttCalendario As Button = Nothing
    Protected HbttMensajes As Button = Nothing
    Protected HbttNotas As Button = Nothing
    Protected HbttAyuda As Button = Nothing
    Protected HbttTercero As Button = Nothing
    Protected HbttImprimir As Button = Nothing
    Protected HbttReportes As Button = Nothing
    Protected HbttCamara As Button = Nothing
    Protected HbttSalir As Button = Nothing
    Protected HbttAceptar As Button = Nothing
    Protected HbttCancelar As Button = Nothing
    Protected HsepAccion As Separator = Nothing
    Protected HmnuAcciones As MenuItem = Nothing
    Protected HmnuHerramientas As MenuItem = Nothing
    Protected HmnuNavegar As MenuItem = Nothing
    Protected HmnuAyuda As MenuItem = Nothing
    Protected HmnuCrear As MenuItem = Nothing
    Protected HmnuModificar As MenuItem = Nothing
    Protected HmnuSuprimir As MenuItem = Nothing
    Protected HmnuAnular As MenuItem = Nothing
    Protected HmnuGuardar As MenuItem = Nothing
    Protected HmnuContenidoAyuda As MenuItem = Nothing
    Protected HmnuRefrescar As MenuItem = Nothing
    Protected HmnuAlPrimero As MenuItem = Nothing
    Protected HmnuAlSiguiente As MenuItem = Nothing
    Protected HmnuAlAnterior As MenuItem = Nothing
    Protected HmnuAlUltimo As MenuItem = Nothing
    Protected HmnuBuscar As MenuItem = Nothing
    Protected HmnuCalendario As MenuItem = Nothing
    Protected HmnuCalculadora As MenuItem = Nothing
    Protected HmnuNotas As MenuItem = Nothing
    Protected HmnuMensajes As MenuItem = Nothing
    Protected HmnuImprimir As MenuItem = Nothing
    Protected HmnuCerrar As MenuItem = Nothing
    Protected HmnuSalir As MenuItem = Nothing
    Protected HmnuMiMenu As Menu = Nothing
    Protected HlblEmpresa As Label = Nothing
    Protected HlblCentroUtil As Label = Nothing
    Protected HlblPeriodo As Label = Nothing
    Protected HlblUsuario As Label = Nothing
    Protected HpnlPanelControl As DockPanel
    Protected HlblMensajes As Label = Nothing
    ' Eventos 
    Private Event EvnNotifica As EventHandler(Of ClsNotiEventArgs)
    'Variables
    Private WithEvents MnuBotonGrande As MenuItem = Nothing
    Private WithEvents MnuBotonMediano As MenuItem = Nothing
    Private WithEvents MnuBotonPequeño As MenuItem = Nothing
    Private MtxtEstado As TextBox = Nothing

    Private MenuElementosToolBar As EnuElementosToolBarDef = EnuElementosToolBarDef.None
    Private MenuElementosAdicionales As EnuElementosAdicionalesDef = EnuElementosAdicionalesDef.None
    Private MentItemsValidar As Integer = 0
    Private MstcValidador As StcValidadorWpf() = Nothing
    Private MblnHabilitadosCtlsLlave As Boolean = False
    Private MblnEsVentanaPrincipal As Boolean = False
    Private MblnEsVentanaParametrizar As Boolean = False
    Private MblnPropietariaEsPrincipal As Boolean = False
    Private MblnVenProceso As Boolean = False
#End Region

#Region "Propiedades autoimplementadas"
    Friend Property EnuOperacionEnWin As EnuOperacionEnVentana = EnuOperacionEnVentana.cenuConsultando
    Friend Property EnuTipoPermisoObjWin As EnuPermisosDef = EnuPermisosDef.None
    Public Property ObjValorLlave() As Object = Nothing
    Friend Property BlnSiempreCreando As Boolean = False
    Friend Property BlnCanceleCierra As Boolean = False
#End Region
#End Region

#Region "Constructores"
    Protected Sub New()

    End Sub
#End Region

#Region "Apariencia de la forma"
    ''' <summary>
    ''' Carga la forma de acuerdo a los valores pasados en los argumentos. Normalmente es llamado desde el
    ''' evento Loaded de la ventana.
    ''' </summary>
    ''' <param name="aenuElementosAdicionales">Especifica los botones adicionales que aparederán en la barra de
    ''' herramientas.</param>
    ''' <param name="aentDimensionItemsValidacion">Indica la cantidad de datos a validar en la ventana.</param>
    ''' <param name="acolControlesLlave">Colección que contiene los controles que muestran los datos que
    ''' conforman la identificación inequivoca del objeto que maneja la forma.</param>
    ''' <param name="actlControlDespuesDeLlave">Es el control siguiente a los controles contenidos en el 
    ''' argumento "acolControlesLlave"</param>
    ''' <remarks></remarks>
    Protected Sub SCargueForma(aenuElementosAdicionales As EnuElementosAdicionalesDef,
            aentDimensionItemsValidacion As Integer, acolControlesLlave As Collection,
            actlControlDespuesDeLlave As Control, ablnVentanaProceso As Boolean)
        GobjPanDat.SControleProcesoObj(True)
        MblnVenProceso = ablnVentanaProceso
        HblnCargandoForma = True
        MblnEsVentanaPrincipal = (EnuIdVentana = EnuIdVentanaDef.EnuMWOrionCop) OrElse
                (EnuIdVentana = EnuIdVentanaDef.EnuMwAdminOrion)
        MblnEsVentanaParametrizar = (EnuIdVentana = EnuIdVentanaDef.enuParametrizacion)
        If Not IsNothing(WinPadre) Then
            MblnPropietariaEsPrincipal = WinPadre.EnuIdVentana =
                    EnuIdVentanaDef.EnuMwAdminOrion OrElse WinPadre.EnuIdVentana =
                            EnuIdVentanaDef.EnuMWOrionCop
        End If
        If WinPadre IsNot Nothing AndAlso EnuIdVentana <> EnuIdVentanaDef.EnuLogOn Then
            WinPadre.Visibility = Visibility.Hidden
        ElseIf MblnEsVentanaPrincipal Then
            Visibility = Visibility.Hidden
        End If
        McolControlesLlave = acolControlesLlave
        MctlControlDespuesDeLlave = actlControlDespuesDeLlave
        MenuElementosAdicionales = aenuElementosAdicionales
        SAsigneMisControles()
        MdgtLblActualizaNot = New SdgtActualizaNoti(AddressOf HlblMensajes.SetValue)
        If Not (MblnEsVentanaPrincipal AndAlso MblnEsVentanaParametrizar) Then
            SInicialiceObjeto()
        End If
        SDetermineElementosTlb()
        SSeleccioneElementosTlb()
        SAsigneBotonesTlb()
        If HmnuMiMenu IsNot Nothing Then
            SDetermineMenuVentana()
        End If
        SRefresqueBotones()
        SDimensioneItemsValidacion(aentDimensionItemsValidacion)
        SInicialiceControles()
        If Not (MblnEsVentanaPrincipal) Then
            SEstablezcaEstado()
        End If
        SConfigureComandos()
        SMuestreDatos()
        SVentanaAuxiliar()
        GobjPanDat.SControleProcesoObj(False)
        HblnCargandoForma = False
    End Sub

    Protected Sub SChequieMenuContextual()
        If Not IsNothing(HtlbMiBarraHerramientas) Then
            Dim lmnuMenuItem As MenuItem = Nothing
            For Each lmnuMenuItem In HtlbMiBarraHerramientas.ContextMenu.Items
                lmnuMenuItem.IsChecked = False
            Next
            Select Case GenuTamanoIcono
                Case EnuTamanoIconos.EnuPequeño
                    For Each lmnuMenuItem In HtlbMiBarraHerramientas.ContextMenu.Items
                        If lmnuMenuItem.Name = "MnuBttPeq" Then
                            lmnuMenuItem.IsChecked = True
                            Exit Select
                        End If
                    Next
                Case EnuTamanoIconos.EnuMediano
                    For Each lmnuMenuItem In HtlbMiBarraHerramientas.ContextMenu.Items
                        If lmnuMenuItem.Name = "MnuBttMed" Then
                            lmnuMenuItem.IsChecked = True
                            Exit Select
                        End If
                    Next
                Case EnuTamanoIconos.EnuGrande
                    For Each lmnuMenuItem In HtlbMiBarraHerramientas.ContextMenu.Items
                        If lmnuMenuItem.Name = "MnuBttGde" Then
                            lmnuMenuItem.IsChecked = True
                            Exit Select
                        End If
                    Next
            End Select
        End If
    End Sub

    Protected Sub SAdicioneControlRestringido(actlControl As Control)
        If Not IsNothing(actlControl) Then
            McolControlesRestringidos.Add(actlControl)
        End If
    End Sub

    Friend Function FstrNombreDoc() As String
        Dim lstrArti = FstrArticulo()
        Dim lstrNombreDoc = " " & ObjObjetoWin.HstrNombreObj
        Dim lstrNombre = lstrArti & ObjObjetoWin.StrNombreClase & lstrNombreDoc
        Return lstrNombre
    End Function

    Private Function FstrArticulo()
        Dim lstrArt = "El "
        If ObjObjetoWin.HenuIdClase = EnuIdClasesPanDef.EnuConsultaSql OrElse
                ObjObjetoWin.HenuIdClase = EnuIdClasesPanDef.EnuInteresMora OrElse
                ObjObjetoWin.HenuIdClase = EnuIdClasesPanDef.EnuCuentaBanco OrElse
                ObjObjetoWin.HenuIdClase = EnuIdClasesPanDef.EnuCuentaContabilidad OrElse
                ObjObjetoWin.HenuIdClase = EnuIdClasesPanDef.EnuFactura OrElse
                ObjObjetoWin.HenuIdClase = EnuIdClasesPanDef.EnuNotaAjusteCuotaAdmin OrElse
                ObjObjetoWin.HenuIdClase = EnuIdClasesPanDef.EnuNotaCr OrElse
                ObjObjetoWin.HenuIdClase = EnuIdClasesPanDef.EnuNotaDevAnt OrElse
                ObjObjetoWin.HenuIdClase = EnuIdClasesPanDef.EnuNotaReversaCr Then
            lstrArt = "La "
        End If
        Return lstrArt
    End Function
#Region "Manejo tamaño botones toolbar"
    ''' <summary>
    ''' Determina el tamaño y el estilo de la barra de herramientas y sus elementos segun la variable
    ''' global GenuTamanoIcono.
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub SRefresqueBotones()
        If Not IsNothing(HtlbMiBarraHerramientas) Then
            Select Case GenuTamanoIcono
                Case EnuTamanoIconos.EnuPequeño
                    SRefresquePeq()
                Case EnuTamanoIconos.EnuMediano
                    SRefresqueMed()
                Case EnuTamanoIconos.EnuGrande
                    SRefresqueGde()
            End Select
            SChequieMenuContextual()
            SHabiliteBotonesTlb()
        End If
    End Sub
    Private Sub SRefresquePeq()
        Dim lrodRowControl As RowDefinition = FindName("rwdControl")
        lrodRowControl.Style = FindResource("RecRdePeq")
        If MblnEsVentanaPrincipal OrElse MblnEsVentanaParametrizar Then
            HtlbMiBarraHerramientas.Style = FindResource("RecTlbPeqPpal")
        Else
            HtlbMiBarraHerramientas.Style = FindResource("RecTlbPeq")
        End If
        For Each lctlControl As Control In HtlbMiBarraHerramientas.Items
            If lctlControl.GetType.Name = "Button" Then
                If MblnEsVentanaPrincipal OrElse MblnEsVentanaParametrizar Then
                    lctlControl.Style = FindResource("RecBttTlbPeqPpal")
                End If
            ElseIf lctlControl.GetType.Name = "Separator" Then
                If MblnEsVentanaPrincipal OrElse MblnEsVentanaParametrizar Then
                    lctlControl.Style = FindResource("RecSepTlbPeqPpal")
                Else
                    lctlControl.Style = FindResource("RecSepTlbPeq")
                End If
            End If
        Next
    End Sub
    Private Sub SRefresqueMed()
        Dim lrodRowControl As RowDefinition = FindName("rwdControl")
        lrodRowControl.Style = FindResource("RecRdeMed")
        If MblnEsVentanaPrincipal OrElse MblnEsVentanaParametrizar Then
            HtlbMiBarraHerramientas.Style = FindResource("RecTlbMedPpal")
        Else
            HtlbMiBarraHerramientas.Style = FindResource("RecTlbMed")
        End If
        For Each lctlControl As Control In HtlbMiBarraHerramientas.Items
            If lctlControl.GetType.Name = "Button" Then
                If MblnEsVentanaPrincipal OrElse MblnEsVentanaParametrizar Then
                    lctlControl.Style = FindResource("RecBttTlbMedPpal")
                End If
            ElseIf lctlControl.GetType.Name = "Separator" Then
                If MblnEsVentanaPrincipal OrElse MblnEsVentanaParametrizar Then
                    lctlControl.Style = FindResource("RecSepTlbMedPpal")
                Else
                    lctlControl.Style = FindResource("RecSepTlbMed")
                End If
            End If
        Next
    End Sub
    Private Sub SRefresqueGde()
        Dim lrodRowControl As RowDefinition = FindName("rwdControl")
        lrodRowControl.Style = FindResource("RecRdeGde")
        If MblnEsVentanaPrincipal OrElse MblnEsVentanaParametrizar Then
            HtlbMiBarraHerramientas.Style = FindResource("RecTlbGdePpal")
        Else
            HtlbMiBarraHerramientas.Style = FindResource("RecTlbGde")
        End If
        For Each lctlControl As Control In HtlbMiBarraHerramientas.Items
            If lctlControl.GetType.Name = "Button" Then
                If MblnEsVentanaPrincipal OrElse MblnEsVentanaParametrizar Then
                    lctlControl.Style = FindResource("RecBttTlbGdePpal")
                Else
                    lctlControl.Style = FindResource("RecBttTlbGdePpal")
                End If
            ElseIf lctlControl.GetType.Name = "Separator" Then
                If MblnEsVentanaPrincipal OrElse MblnEsVentanaParametrizar Then
                    lctlControl.Style = FindResource("RecSepTlbGdePpal")
                Else
                    lctlControl.Style = FindResource("RecSepTlbGde")
                End If
            End If
        Next
    End Sub
    Private Sub EConMenu_Closing(sender As Object, e As EventArgs) Handles HtlbMiBarraHerramientas.ContextMenuClosing
        SRefresqueBotones()
    End Sub
    Private Sub EMenuConGde_Click(sender As Object, e As EventArgs) Handles MnuBotonGrande.Checked
        MnuBotonMediano.IsChecked = False
        MnuBotonPequeño.IsChecked = False
        GenuTamanoIcono = EnuTamanoIconos.EnuGrande
    End Sub
    Private Sub EMenuConMed_Click(sender As Object, e As EventArgs) Handles MnuBotonMediano.Checked
        MnuBotonGrande.IsChecked = False
        MnuBotonPequeño.IsChecked = False
        GenuTamanoIcono = EnuTamanoIconos.EnuMediano
    End Sub
    Private Sub EMenuConPeq_Click(sender As Object, e As EventArgs) Handles MnuBotonPequeño.Checked
        MnuBotonGrande.IsChecked = False
        MnuBotonMediano.IsChecked = False
        GenuTamanoIcono = EnuTamanoIconos.EnuPequeño
    End Sub
#End Region
#End Region

#Region "Propiedades de solo lectura"
    Protected ReadOnly Property HcolLabelsBarraEstado As Collection
        Get
            If IsNothing(McolLabels) Then
                McolLabels = New Collection From {
                    HlblEmpresa,
                    HlblCentroUtil,
                    HlblUsuario,
                    HlblPeriodo
                }
            End If
            Return McolLabels
        End Get
    End Property
#End Region

#Region "Procedimientos"
#Region "Determina elementos ToolBar y Menu"
    ''' <summary>
    ''' Asigna el valor final de la variable local "menuElementosToolBar" de acuerdo al tipo de objeto
    '''  y a los argumentos pasador en "sCargueForma"
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SDetermineElementosTlb()
        If EnuIdVentana <> EnuIdVentanaDef.EnuLogOn AndAlso
                EnuIdVentana <> EnuIdVentanaDef.EnuCambioContraseña AndAlso
                EnuIdVentana <> EnuIdVentanaDef.EnuRegistroClave AndAlso
                EnuIdVentana <> EnuIdVentanaDef.EnuRegistroCenUtil AndAlso
                EnuIdVentana <> EnuIdVentanaDef.EnuAutorizaDscto Then
            SDetermineElementosGrales()
            If Not (MblnEsVentanaPrincipal OrElse MblnEsVentanaParametrizar) Then
                If Not MblnVenProceso Then
                    SDetermineElementosObjeto()
                    SDetermineElementosAdicionales()
                End If
            ElseIf MblnEsVentanaParametrizar Then
                MenuElementosToolBar += EnuElementosToolBarDef.EnuRefrescar
                MenuElementosToolBar += EnuElementosToolBarDef.EnuSepAccion
            End If
        Else
            MenuElementosToolBar = EnuElementosToolBarDef.None
        End If
    End Sub

    Private Sub SDetermineElementosGrales()
        If MblnEsVentanaPrincipal Then
            MenuElementosToolBar += EnuElementosToolBarDef.EnuSalir + EnuElementosToolBarDef.EnuRefrescar +
                    EnuElementosToolBarDef.EnuSepAccion
        Else
            MenuElementosToolBar += EnuElementosToolBarDef.EnuCerrar
        End If
        MenuElementosToolBar += EnuElementosToolBarDef.EnuSepSalir
        MenuElementosToolBar += EnuElementosToolBarDef.EnuCalendario +
                EnuElementosToolBarDef.EnuCalculadora + EnuElementosToolBarDef.EnuNotas +
                EnuElementosToolBarDef.EnuSepHerramientas
        MenuElementosToolBar += EnuElementosToolBarDef.EnuAyuda
    End Sub

    Private Sub SDetermineElementosObjeto()
        If Not IsNothing(ObjObjetoWin) Then
            With ObjObjetoWin
                If .BlnEsCreable Then
                    MenuElementosToolBar += EnuElementosToolBarDef.EnuCrear
                End If
                If .BlnEsModificable Then
                    MenuElementosToolBar += EnuElementosToolBarDef.EnuModificar
                End If
                If .BlnEsAnulable Then
                    MenuElementosToolBar += EnuElementosToolBarDef.EnuAnular
                End If
                If .FblnPermitidoSuprimir Then
                    MenuElementosToolBar += EnuElementosToolBarDef.EnuSuprimir
                End If
                If .BlnEsCreable OrElse .BlnEsModificable Then
                    MenuElementosToolBar += EnuElementosToolBarDef.EnuGuardar
                End If
                MenuElementosToolBar += EnuElementosToolBarDef.EnuRefrescar
                MenuElementosToolBar += EnuElementosToolBarDef.EnuSepAccion
                If .BlnEsNavegable Then
                    SDetermineElementosNavegacion()
                Else
                    If MenuElementosToolBar And EnuElementosToolBarDef.EnuCrear Then
                        MenuElementosToolBar -= EnuElementosToolBarDef.EnuCrear
                    End If
                    If MenuElementosToolBar And EnuElementosToolBarDef.EnuSuprimir Then
                        MenuElementosToolBar -= EnuElementosToolBarDef.EnuSuprimir
                    End If
                End If
                MenuElementosToolBar += EnuElementosToolBarDef.EnuEstado
                MenuElementosToolBar += EnuElementosToolBarDef.EnuSepEstado
            End With
        End If
    End Sub

    Private Sub SDetermineElementosNavegacion()
        MenuElementosToolBar += EnuElementosToolBarDef.EnuAlPrimero + EnuElementosToolBarDef.EnuAlAnterior +
                EnuElementosToolBarDef.EnuAlSiguiente + EnuElementosToolBarDef.EnuAlUltimo +
                EnuElementosToolBarDef.EnuBuscar + EnuElementosToolBarDef.EnuSepNavegar
    End Sub

    Private Sub SDetermineElementosAdicionales()
        Dim lblnHayAdicionales = False
        If MenuElementosAdicionales And EnuElementosAdicionalesDef.EnuBuscar Then
            If Not MenuElementosToolBar And EnuElementosToolBarDef.EnuBuscar Then
                MenuElementosToolBar += EnuElementosToolBarDef.EnuBuscar
                lblnHayAdicionales = True
            End If
        End If
        If MenuElementosAdicionales And EnuElementosAdicionalesDef.EnuCamara Then
            MenuElementosToolBar += EnuElementosToolBarDef.EnuCamara
            lblnHayAdicionales = True
        End If
        If MenuElementosAdicionales And EnuElementosAdicionalesDef.EnuTercero Then
            MenuElementosToolBar += EnuElementosToolBarDef.EnuTercero
            lblnHayAdicionales = True
        End If
        If MenuElementosAdicionales And EnuElementosAdicionalesDef.EnuImprimir Then
            MenuElementosToolBar += EnuElementosToolBarDef.EnuImprimir
            lblnHayAdicionales = True
        End If
        If MenuElementosAdicionales And EnuElementosAdicionalesDef.EnuReportes Then
            MenuElementosToolBar += EnuElementosToolBarDef.EnuReportes
            lblnHayAdicionales = True
        End If
        If lblnHayAdicionales Then
            MenuElementosToolBar += EnuElementosToolBarDef.EnuSepVarios
        End If
    End Sub

    ''' <summary>
    ''' Con base en la variable local "menuElementosToolBar" selecciona los elementos de la ToolBar que
    ''' seran mostrados.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SSeleccioneElementosTlb()
        If Not IsNothing(HtlbMiBarraHerramientas) Then
            If MenuElementosToolBar <> EnuElementosToolBarDef.None Then
                For Each lctlControl As Control In HtlbMiBarraHerramientas.Items
                    If Not CType((CType(MenuElementosToolBar, Integer) And
                            CType(lctlControl.Tag, Integer)), Boolean) Then
                        lctlControl.Visibility = Visibility.Collapsed
                    End If
                Next
            Else
                HtlbMiBarraHerramientas.Visibility = Visibility.Collapsed
            End If
        End If
    End Sub

    ''' <summary>
    ''' Determina y establece el menu de la ventana y sus items.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SDetermineMenuVentana()
        If EnuIdVentana <> EnuIdVentanaDef.EnuLogOn AndAlso
                EnuIdVentana <> EnuIdVentanaDef.enuRegistroClave AndAlso
                EnuIdVentana <> EnuIdVentanaDef.EnuRegistroCenUtil AndAlso
                EnuIdVentana <> EnuIdVentanaDef.EnuCambioContraseña AndAlso
                EnuIdVentana <> EnuIdVentanaDef.EnuAutorizaDscto AndAlso
                EnuIdVentana <> EnuIdVentanaDef.EnuRevisaNovs Then
            SConformeMenu()
            SAsigneMisMenues()
            SSeleccioneElementosMenu()
            SConfigureMenuesPropios()
        End If
    End Sub

    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SConformeMenu()
        If Not (MblnEsVentanaPrincipal OrElse MblnEsVentanaParametrizar OrElse
                HenuIdVentana = EnuIdVentanaDef.EnuRevisaNovs OrElse
                HenuIdVentana = EnuIdVentanaDef.EnuPropietario) Then
            HmnuMiMenu.Items.Add(FindResource("RecMnuAcciones"))
        End If
        If Not IsNothing(ObjObjetoWin) Then
            If ObjObjetoWin.BlnEsNavegable Then
                HmnuMiMenu.Items.Add(FindResource("RecMnuNavegar"))
            End If
        End If
        If HmnuMiMenu IsNot Nothing Then
            HmnuMiMenu.Items.Add(FindResource("RecMnuHerramientas"))
            HmnuMiMenu.Items.Add(FindResource("RecMnuAyuda"))
        End If
    End Sub

    ''' <summary>
    ''' Con base en la variable local "menuElementosToolBar" selecciona los elementos del Menu que
    ''' seran mostrados.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SSeleccioneElementosMenu()
        If Not (MblnEsVentanaPrincipal OrElse MblnEsVentanaParametrizar) Then
            MenuElementosToolBar += EnuElementosToolBarDef.EnuCerrar -
                    EnuElementosToolBarDef.EnuSalir
        End If
        If Not IsNothing(HmnuMiMenu) Then
            If HmnuMiMenu.Items.Contains(HmnuAcciones) Then
                If Not (MblnEsVentanaPrincipal OrElse MblnEsVentanaParametrizar) Then
                    SSeleccioneMenuItems(HmnuAcciones)
                End If
            End If
            If HmnuMiMenu.Items.Contains(HmnuHerramientas) Then
                SSeleccioneMenuItems(HmnuHerramientas)
            End If
            If HmnuMiMenu.Items.Contains(HmnuNavegar) Then
                SSeleccioneMenuItems(HmnuNavegar)
            End If
            If HmnuMiMenu.Items.Contains(HmnuAyuda) Then
                SSeleccioneMenuItems(HmnuAyuda)
            End If
        End If
    End Sub

    Private Sub SSeleccioneMenuItems(amnuMenuItem As MenuItem)
        For Each lobjObjeto As Object In amnuMenuItem.Items
            If TypeOf lobjObjeto Is MenuItem Then
                Dim lmnui As MenuItem = lobjObjeto
                If Not CType((CType(MenuElementosToolBar, Integer) And CType(lmnui.Tag, Integer)),
                        Boolean) Then
                    lmnui.Visibility = Visibility.Collapsed
                Else
                    If lmnui.Items.Count > 0 Then
                        SSeleccioneMenuItems(lmnui)
                    End If
                End If
            End If
        Next
    End Sub
#End Region

#Region "Asigan variables a los controles generales de la ventana"
    ''' <summary>
    ''' Asigna a variables de modulo los siguientes controles generales: ToolBar, Menu, Botones de Toolbar,
    ''' Botones Aceptar y Cancelar. Ademas inicializa el Menu Contextual de la ToolBar.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SAsigneMisControles()
        Dim lblnEncontradoTB = False, lblnEncontradoLB = False
        Dim lblEncontrado As Label
        HpnlPanelControl = FindName("PanelControl")
        If Not IsNothing(HpnlPanelControl) Then
            For Each uieElemento As UIElement In HpnlPanelControl.Children
                If TypeOf uieElemento Is ToolBar Then
                    HtlbMiBarraHerramientas = CType(uieElemento, ToolBar)
                    lblnEncontradoTB = True
                ElseIf TypeOf uieElemento Is Label Then
                    lblEncontrado = CType(uieElemento, Label)
                    If lblEncontrado.Name = "lblNotifica" Then
                        HlblMensajes = lblEncontrado
                        lblnEncontradoLB = True
                    End If
                End If
                If lblnEncontradoTB AndAlso lblnEncontradoLB Then
                    Exit For
                End If
            Next
            HmnuMiMenu = FindName("MenuVen")
        End If
        MgrdGridPrincipal = FindName("grdPpal")
        If Not (MblnEsVentanaPrincipal OrElse MblnEsVentanaParametrizar) Then
            SAsigneBtotonesAceCan()
        End If
        SAsignePieVentana()
        SChequieMenuContextual()
    End Sub

    Private Sub SAsigneBotonesTlb()
        If Not IsNothing(HtlbMiBarraHerramientas) Then
            For Each lobjObjeto As Object In HtlbMiBarraHerramientas.Items
                If Not FblnAsignoBoton(lobjObjeto) Then
                    Select Case True
                        Case TypeOf lobjObjeto Is TextBox
                            MtxtEstado = lobjObjeto
                        Case TypeOf lobjObjeto Is Button
                            Select Case lobjObjeto.Name
                                Case "bttNotas"
                                    HbttNotas = lobjObjeto
                                Case "bttMensajes"
                                    HbttMensajes = lobjObjeto
                                Case "bttAlPrimero"
                                    HbttAlPrimero = lobjObjeto
                                Case "bttAlSiguiente"
                                    HbttAlSiguiente = lobjObjeto
                                Case "bttAlAnterior"
                                    HbttAlAnterior = lobjObjeto
                                Case "bttAlUltimo"
                                    HbttAlUltimo = lobjObjeto
                                Case "bttBuscar"
                                    HbttBuscar = lobjObjeto
                                Case "bttTercero"
                                    HbttTercero = lobjObjeto
                                Case "bttImprimir"
                                    HbttImprimir = lobjObjeto
                                Case "bttReportes"
                                    HbttReportes = lobjObjeto
                                Case "bttCamara"
                                    HbttCamara = lobjObjeto
                                Case "bttAyuda"
                                    HbttAyuda = lobjObjeto
                            End Select
                        Case TypeOf lobjObjeto Is Separator
                            If lobjObjeto.Name = "sepAccion" Then
                                HsepAccion = lobjObjeto
                            End If
                    End Select
                End If
            Next
        End If
    End Sub

    Private Function FblnAsignoBoton(aobjObjeto As Object) As Boolean
        Dim lblnAsigno = False
        If TypeOf aobjObjeto Is Button Then
            lblnAsigno = True
            Select Case aobjObjeto.Name
                Case "bttSalir"
                    HbttSalir = aobjObjeto
                Case "bttCerrar"
                    HbttCerrar = aobjObjeto
                Case "bttCrear"
                    HbttCrear = aobjObjeto
                Case "bttModificar"
                    HbttModificar = aobjObjeto
                Case "bttSuprimir"
                    HbttSuprimir = aobjObjeto
                Case "bttAnular"
                    HbttAnular = aobjObjeto
                Case "bttGuardar"
                    HbttGuardar = aobjObjeto
                Case "bttRefrescar"
                    HbttRefrescar = aobjObjeto
                Case "bttCalendario"
                    HbttCalendario = aobjObjeto
                Case "bttCalculadora"
                    HbttCalculadora = aobjObjeto
                Case Else
                    lblnAsigno = False
            End Select
        End If
        Return lblnAsigno
    End Function

    Private Sub SAsigneBtotonesAceCan()
        Dim lcnvBotones As Canvas
        Dim lcnvAceCan As Canvas = Nothing
        Dim lgrdPieVentana As Grid = FindName("grdPieVentana")
        If Not IsNothing(lgrdPieVentana) Then
            lcnvBotones = lgrdPieVentana.FindName("cnvBotones")
            If Not IsNothing(lcnvBotones) Then
                For Each lobjObjeto As Object In lcnvBotones.Children
                    If TypeOf lobjObjeto Is Canvas AndAlso lobjObjeto.Name = "cnvAceptaCancela" Then
                        lcnvAceCan = lobjObjeto
                    End If
                Next
                If Not IsNothing(lcnvAceCan) Then
                    For Each lctlObjeto As Object In lcnvAceCan.Children
                        If TypeOf lctlObjeto Is Button Then
                            If lctlObjeto.Name = "bttAceptar" Then
                                HbttAceptar = lctlObjeto
                                HbttAceptar.ToolTip = "En Modo 'Consultando' este botón cierra la Ventana;" & vbCrLf &
                                        "En Modo 'Creando' o 'Modicando' acepta y graba los Cambios."
                            ElseIf lctlObjeto.Name = "bttCancelar" Then
                                HbttCancelar = lctlObjeto
                                HbttCancelar.ToolTip = "En Modo 'Consultando' este botón cierra " &
                                        "la Ventana;" & vbCrLf & "En Modo 'Creando' o 'Modicando' " &
                                        " cancela los Cambios " & vbCrLf & "y cambia la Ventana " &
                                        "al Modo 'Consultando'."
                            End If
                        End If
                    Next
                End If
            End If
        End If
    End Sub

    Private Sub SAsignePieVentana()
        Dim lgrdPieVentana As Grid = FindName("grdPieVentana")
        If Not MblnEsVentanaPrincipal Then
            If EnuIdVentana = EnuIdVentanaDef.EnuLogOn OrElse
                    EnuIdVentana = EnuIdVentanaDef.EnuCambioContraseña Then
                SAsignePieVentanaLog(lgrdPieVentana)
            Else
                SAsignePieVentanaGral(lgrdPieVentana)
            End If
        ElseIf EnuIdVentana = EnuIdVentanaDef.EnuMwAdminOrion Then
            SAsignePieVentanaGral(lgrdPieVentana)
        End If
    End Sub

    Private Sub SAsignePieVentanaLog(agrdPieVentana As Grid)
        If Not IsNothing(agrdPieVentana) Then
            Dim ldcpPiePagina As DockPanel
            For Each lobjObjeto As Object In agrdPieVentana.Children
                If TypeOf lobjObjeto Is DockPanel AndAlso lobjObjeto.Name = "dcpMensajesLog" Then
                    ldcpPiePagina = lobjObjeto
                    HlblMensajes = ldcpPiePagina.Children(0)
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub SAsignePieVentanaGral(agrdPieVentana As Grid)
        If Not IsNothing(agrdPieVentana) Then
            Dim ldcpPiePagina As DockPanel = Nothing
            For Each lobjObjeto As Object In agrdPieVentana.Children
                If TypeOf lobjObjeto Is DockPanel AndAlso lobjObjeto.Name = "dcpPiePaginaGral" Then
                    ldcpPiePagina = lobjObjeto
                    Exit For
                End If
            Next
            If IsNothing(ldcpPiePagina) Then
                For Each lobjObjeto As Object In agrdPieVentana.Children
                    If TypeOf lobjObjeto Is DockPanel AndAlso lobjObjeto.Name = "dcpMensajesLog" Then
                        ldcpPiePagina = lobjObjeto
                        HlblMensajes = ldcpPiePagina.Children(0)
                        Exit For
                    End If
                Next
            Else
                Dim lsbrGeneral As StatusBar
                lsbrGeneral = ldcpPiePagina.Children(0)
                If ldcpPiePagina.Children.Count > 1 Then
                    HlblMensajes = ldcpPiePagina.Children(1)
                End If
                If Not IsNothing(lsbrGeneral) Then
                    For Each lobjObjeto As Object In lsbrGeneral.Items
                        If TypeOf lobjObjeto Is StatusBarItem Then
                            Dim lsbiStatusBarIitem As StatusBarItem = lobjObjeto
                            Select Case lsbiStatusBarIitem.Name
                                Case "sbiEmpresa"
                                    HlblEmpresa = lsbiStatusBarIitem.Content
                                Case "sbiCentroUtil"
                                    HlblCentroUtil = lsbiStatusBarIitem.Content
                                Case "sbiUsuario"
                                    HlblUsuario = lsbiStatusBarIitem.Content
                                Case "sbiPeriodo"
                                    HlblPeriodo = lsbiStatusBarIitem.Content
                            End Select
                        End If
                    Next
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Asigna variables de modulo a los menues generales de la ventana
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SAsigneMisMenues()
        For Each lobjObjeto As Object In HmnuMiMenu.Items
            If TypeOf lobjObjeto Is MenuItem Then
                Select Case lobjObjeto.Name
                    Case "MnuAcciones"
                        If Not (MblnEsVentanaPrincipal OrElse MblnEsVentanaParametrizar) Then
                            HmnuAcciones = lobjObjeto
                            SAsigneMenuAcciones()
                        End If
                    Case "MnuHerramientas"
                        HmnuHerramientas = lobjObjeto
                        SAsigneMenuHerramientas()
                    Case "MnuNavegar"
                        HmnuNavegar = lobjObjeto
                        SAsigneMenuNavegar()
                    Case "MnuAyuda"
                        HmnuAyuda = lobjObjeto
                        SAsigneMenuAyuda()
                End Select
            End If
        Next
        Dim lcmnToolBarMenuCon As ContextMenu = HtlbMiBarraHerramientas.ContextMenu
        For Each lmnuiMenuItem As MenuItem In lcmnToolBarMenuCon.Items
            Select Case lmnuiMenuItem.Name
                Case "MnuBttGde"
                    MnuBotonGrande = lmnuiMenuItem
                Case "MnuBttMed"
                    MnuBotonMediano = lmnuiMenuItem
                Case "MnuBttPeq"
                    MnuBotonPequeño = lmnuiMenuItem
            End Select
        Next
    End Sub

    Private Sub SAsigneMenuAcciones()
        For Each lobjObjeto As Object In HmnuAcciones.Items
            If TypeOf lobjObjeto Is MenuItem Then
                Dim lmnu As MenuItem = lobjObjeto
                Select Case lmnu.Name
                    Case "MnuCrear"
                        HmnuCrear = lobjObjeto
                    Case "MnuModificar"
                        HmnuModificar = lobjObjeto
                    Case "MnuSuprimir"
                        HmnuSuprimir = lobjObjeto
                    Case "MnuAnular"
                        HmnuAnular = lobjObjeto
                    Case "MnuGuardar"
                        HmnuGuardar = lobjObjeto
                    Case "MnuRefrescar"
                        HmnuRefrescar = lobjObjeto
                    Case "MnuCerrar"
                        HmnuCerrar = lobjObjeto
                End Select
            End If
        Next
    End Sub

    Private Sub SAsigneMenuHerramientas()
        HmnuHerramientas.Name = "MnuHerramientas"
        For Each lobjObjeto As Object In HmnuHerramientas.Items
            If TypeOf lobjObjeto Is MenuItem Then
                Dim lmnu As MenuItem = lobjObjeto
                Select Case lmnu.Name
                    Case "MnuCalendario"
                        HmnuCalendario = lobjObjeto
                    Case "MnuCalculadora"
                        HmnuCalculadora = lobjObjeto
                    Case "MnuNotas"
                        HmnuNotas = lobjObjeto
                    Case "MnuMensajes"
                        HmnuMensajes = lobjObjeto
                End Select
            End If
        Next
    End Sub

    Private Sub SAsigneMenuNavegar()
        For Each lobjObjeto As Object In HmnuNavegar.Items
            If TypeOf lobjObjeto Is MenuItem Then
                Dim lmnu As MenuItem = lobjObjeto
                Select Case lmnu.Name
                    Case "MnuAlPrimero"
                        HmnuAlPrimero = lobjObjeto
                    Case "MnuAlAnterior"
                        HmnuAlAnterior = lobjObjeto
                    Case "MnuAlSiguiente"
                        HmnuAlSiguiente = lobjObjeto
                    Case "MnuAlUltimo"
                        HmnuAlUltimo = lobjObjeto
                    Case "MnuBuscar"
                        HmnuBuscar = lobjObjeto
                End Select
            End If
        Next
    End Sub

    Private Sub SAsigneMenuAyuda()
        For Each lobjObjeto As Object In HmnuAyuda.Items
            If TypeOf lobjObjeto Is MenuItem Then
                Dim lmnu As MenuItem = lobjObjeto
                Select Case lmnu.Name
                    Case "MnuContenidoAyuda"
                        HmnuContenidoAyuda = lobjObjeto
                End Select
            End If
        Next
    End Sub
#End Region

#Region "Establece el estado de los controles"
    ' Establece el estado de los controles y botones de acuerdo al estado de la ventana y su comportamiento
    ''' <summary>
    ''' Dimensiona la matriz "mstcValidador" de acuerdo a la cantidad de controles a validar. Este dato
    ''' es pasado como un argumento en "sCargueForma".
    ''' </summary>
    ''' <param name="aentItemsValidar"></param>
    ''' <remarks></remarks>
    Protected Sub SDimensioneItemsValidacion(aentItemsValidar As Integer)
        MentItemsValidar = aentItemsValidar
        If MentItemsValidar > 0 Then
            ReDim MstcValidador(MentItemsValidar - 1)
        End If
        SInicialiceValido()
    End Sub

    Protected Sub SInicialiceValido()
        For i = 0 To MentItemsValidar - 1
            StcValidValido(i) = True
        Next
    End Sub

    Private Sub SEstablezcaEstado()
        If IsNothing(WinPadre) OrElse MblnPropietariaEsPrincipal Then
            If IsNothing(McolControlesLlave) OrElse McolControlesLlave.Count = 0 Then
                SPermisos(False)
            Else
                If Not IsNothing(ObjValorLlave) Then
                    For i As Byte = 1 To McolControlesLlave.Count
                        McolControlesLlave(i).text = ObjValorLlave(i - 1)
                    Next
                End If
                SPermisos(ObjObjetoWin.FblnEsCreable(ObjValorLlave))
            End If
        Else
            SPermisos(ObjObjetoWin.FblnEsCreable(ObjValorLlave))
        End If
    End Sub

    Private Sub SPermisos(ablnCreable As Boolean)
        If EnuIdVentana <> EnuIdVentanaDef.EnuLogOn Then
            SHabiliteMenues()
        End If
        If Not (MblnEsVentanaPrincipal OrElse MblnEsVentanaParametrizar) Then
            Select Case EnuOperacionEnWin
                Case EnuOperacionEnVentana.CenuConsultando
                    SHabiliteWin(False)
                Case EnuOperacionEnVentana.CenuCreando
                    SHabiliteWin(True)
                    EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando
                    If IsNothing(ObjValorLlave) OrElse ablnCreable Then
                        SCrearClic()
                    End If
                Case EnuOperacionEnVentana.CenuModificando
                    SHabiliteWin(True)
                    EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando
                    If (Not IsNothing(ObjValorLlave)) OrElse
                                    (Not IsNothing(ObjObjetoWin)) Then
                        If Not ablnCreable Then
                            SModificarClic()
                        Else
                            If IsNothing(ObjValorLlave) AndAlso (Not IsNothing(ObjObjetoWin)) Then
                                SModificarClic()
                            End If
                        End If
                    End If
            End Select
        End If
    End Sub

    ''' <summary>
    ''' Muestra el estado de la ventana en el control "txtEstado" de la barra de herramientas
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SMuestreEstado()
        If Not MblnVenProceso AndAlso HmnuMiMenu IsNot Nothing Then
            MtxtEstado.Visibility = Visibility.Visible
            Select Case EnuOperacionEnWin
                Case EnuOperacionEnVentana.CenuCreando
                    MtxtEstado.Background = Brushes.Yellow
                    MtxtEstado.Text = My.Resources.EstCreando
                Case EnuOperacionEnVentana.CenuModificando
                    MtxtEstado.Background = Brushes.Yellow
                    MtxtEstado.Text = My.Resources.EstModificando
                Case Else
                    MtxtEstado.Background = Brushes.Aquamarine
                    MtxtEstado.Text = My.Resources.EstConsultando
            End Select
        End If
    End Sub

    ''' <summary>
    ''' Habilita los botones de la ToolBar y los Menues con base en el comportamiento establecido en el 
    ''' objeto de la ventana (objObjetoWin), en el estado de la ventana y en los permisos del usuario actual.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overridable Sub SHabiliteBotonesTlb()
        Dim lblnHabilite As Boolean
        If Not IsNothing(HtlbMiBarraHerramientas) Then
            SHabiliteBotonesAccion()
            If Not IsNothing(HbttAnular) Then
                lblnHabilite = (HbttAnular.Visibility = Visibility.Visible) AndAlso
                        (Not ObjObjetoWin.FblnEstaVacioOrigenDatos) AndAlso
                         (EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando) AndAlso
                         CType(EnuTipoPermisoObjWin And EnuPermisosDef.enuAnular, Boolean) AndAlso
                         ObjObjetoWin.FblnEsAnulable
                SHabiliteBotonTlb(lblnHabilite, HbttAnular)
                SHabiliteMenuItem(lblnHabilite, HmnuAnular)
            End If
            If Not IsNothing(HbttRefrescar) Then
                lblnHabilite = (EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando)
                SHabiliteBotonTlb(lblnHabilite, HbttRefrescar)
                SHabiliteMenuItem(lblnHabilite, HmnuRefrescar)
            End If
            If Not IsNothing(HbttImprimir) Then
                lblnHabilite = (EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando) AndAlso
                        CType(EnuTipoPermisoObjWin And EnuPermisosDef.enuImprimir, Boolean)
                If lblnHabilite Then
                    If Not IsNothing(ObjObjetoWin) Then
                        lblnHabilite = (ObjObjetoWin.EnuTipoObjeto = EnuModoInstanciaObjDef.enuUnico) OrElse
                                Not ObjObjetoWin.FblnEstaVacioOrigenDatos
                    Else
                        lblnHabilite = False
                    End If
                End If
                SHabiliteBotonTlb(lblnHabilite, HbttImprimir)
            End If
            If HbttCerrar.Visibility = Visibility.Visible Then
                HbttCerrar.IsEnabled = True
            End If
        End If
        SHabiliteBotonesNavegar()
    End Sub

    Private Sub SHabiliteBotonesAccion()
        Dim lblnHabilite As Boolean
        If Not IsNothing(HbttCrear) Then
            lblnHabilite = (HbttCrear.Visibility = Visibility.Visible) AndAlso
                    CType(EnuTipoPermisoObjWin And EnuPermisosDef.enuCrear, Boolean) AndAlso
                    EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando
            SHabiliteBotonTlb(lblnHabilite, HbttCrear)
            SHabiliteMenuItem(lblnHabilite, HmnuCrear)
        End If
        If Not IsNothing(HbttModificar) Then
            lblnHabilite = (HbttModificar.Visibility = Visibility.Visible) AndAlso
                    (Not ObjObjetoWin.FblnEstaVacioOrigenDatos) AndAlso
                    ObjObjetoWin.BlnExiste AndAlso
                    (EnuTipoPermisoObjWin And EnuPermisosDef.enuModificar) AndAlso
                    ObjObjetoWin.FblnEsModificable AndAlso
                    EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando
            SHabiliteBotonTlb(lblnHabilite, HbttModificar)
            SHabiliteMenuItem(lblnHabilite, HmnuModificar)
        End If
        If Not IsNothing(HbttSuprimir) Then
            lblnHabilite = (HbttSuprimir.Visibility = Visibility.Visible) AndAlso
                    (Not ObjObjetoWin.FblnEstaVacioOrigenDatos) AndAlso
                     (EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando) AndAlso
                     CType(EnuTipoPermisoObjWin And EnuPermisosDef.enuSuprimir, Boolean)
            SHabiliteBotonTlb(lblnHabilite, HbttSuprimir)
            SHabiliteMenuItem(lblnHabilite, HmnuSuprimir)
        End If
    End Sub

    Private Sub SHabiliteBotonesNavegar()
        If Not IsNothing(HtlbMiBarraHerramientas) Then
            Dim lblnNavegar As Boolean = Not IsNothing(ObjObjetoWin)
            If lblnNavegar Then lblnNavegar = ObjObjetoWin.BlnEsNavegable
            If lblnNavegar Then
                If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando OrElse
                         ObjObjetoWin.FblnEstaVacioOrigenDatos OrElse
                         ObjObjetoWin.EnuTipoObjeto = EnuModoInstanciaObjDef.enuUnico Then
                    lblnNavegar = False
                End If
                If lblnNavegar Then
                    SHabiliteBotonTlb(Not ObjObjetoWin.FblnEsElPrimerRegistro, HbttAlPrimero)
                    SHabiliteMenuItem(Not ObjObjetoWin.FblnEsElPrimerRegistro, HmnuAlPrimero)
                    SHabiliteBotonTlb(Not ObjObjetoWin.FblnEsElPrimerRegistro, HbttAlAnterior)
                    SHabiliteMenuItem(Not ObjObjetoWin.FblnEsElPrimerRegistro, HmnuAlAnterior)
                    SHabiliteBotonTlb(Not ObjObjetoWin.FblnEsElUltimoRegistro, HbttAlSiguiente)
                    SHabiliteMenuItem(Not ObjObjetoWin.FblnEsElUltimoRegistro, HmnuAlSiguiente)
                    SHabiliteBotonTlb(Not ObjObjetoWin.FblnEsElUltimoRegistro, HbttAlUltimo)
                    SHabiliteMenuItem(Not ObjObjetoWin.FblnEsElUltimoRegistro, HmnuAlUltimo)
                    SHabiliteBotonTlb(True, HbttBuscar)
                    SHabiliteMenuItem(True, HmnuBuscar)
                Else
                    SHabiliteBotonTlb(False, HbttAlPrimero)
                    SHabiliteMenuItem(False, HmnuAlPrimero)
                    SHabiliteBotonTlb(False, HbttAlAnterior)
                    SHabiliteMenuItem(False, HmnuAlAnterior)
                    SHabiliteBotonTlb(False, HbttAlSiguiente)
                    SHabiliteMenuItem(False, HmnuAlSiguiente)
                    SHabiliteBotonTlb(False, HbttAlUltimo)
                    SHabiliteMenuItem(False, HmnuAlUltimo)
                    SHabiliteBotonTlb(False, HbttBuscar)
                    SHabiliteMenuItem(False, HmnuBuscar)
                End If
            End If
        End If
    End Sub

    Protected Sub SHabiliteBotonTlb(ablnHabilita As Boolean, abttBoton As FrameworkElement)
        If abttBoton IsNot Nothing Then
            If ablnHabilita Then
                abttBoton.Style = FindResource("RecBttHabilitado")
            Else
                abttBoton.Style = FindResource("RecBttDesHabilitado")
            End If
        End If
    End Sub

    Protected Overridable Sub SHabiliteMenues()
        If GobjPanorama.ObjUsuarioActual.ObjIdUsuarioStr.ObjValorPro <> GCSTRUSUARIOU Then
            If IsNothing(McolPermisosWin) Then
                McolPermisosWin = GobjPanorama.ObjUsuarioActual.FcolPermisosAccionesFrm(EnuIdVentana)
            End If
        End If
        If Not IsNothing(HmnuMiMenu) Then
            Dim lmnuMenuItem As MenuItem
            For Each lobjObjeto As Object In HmnuMiMenu.Items
                If TypeOf lobjObjeto Is MenuItem Then
                    lmnuMenuItem = lobjObjeto
                    SHabiliteMenuPan(lmnuMenuItem)
                End If
            Next
        End If
    End Sub

    Protected Sub SHabiliteMenuPan(amnuMenuItem As MenuItem)
        If TypeOf amnuMenuItem Is MenuItemPan Then
            Dim lmnuMenuItemPan As MenuItemPan = amnuMenuItem
            Dim lblnHabilite = True
            If GobjPanorama.ObjUsuarioActual.ObjIdUsuarioStr.ObjValorPro <> GCSTRUSUARIOU Then
                Dim lentIdAccion = lmnuMenuItemPan.EntIdAccion
                lblnHabilite = FblnHabilitarMenuPan(lentIdAccion)
            End If
            If lblnHabilite Then
                SHabiliteMenuItemPan(True, lmnuMenuItemPan)
                If lmnuMenuItemPan.Items.Count > 0 Then
                    For Each lobjObjeto As Object In lmnuMenuItemPan.Items
                        If TypeOf lobjObjeto Is MenuItem Then
                            Dim lmnuMenuItem As MenuItem = lobjObjeto
                            SHabiliteMenuPan(lmnuMenuItem)
                        End If
                    Next
                End If
            Else
                SHabiliteMenuItemPan(False, lmnuMenuItemPan)
            End If
        ElseIf TypeOf amnuMenuItem Is MenuItem Then
            If amnuMenuItem.Items.Count > 0 Then
                For Each lobjObjeto As Object In amnuMenuItem.Items
                    If TypeOf lobjObjeto Is MenuItem Then
                        Dim lmnuMenuItem As MenuItem = lobjObjeto
                        SHabiliteMenuPan(lmnuMenuItem)
                    End If
                Next
            End If
        End If
    End Sub

    Protected Shared Function FblnHabilitarMenuPan(aenuIDVentana As EnuIdVentanaDef,
            aentIdAccion As Integer) As Boolean
        Dim lblnHabilitar As Boolean = True
        Dim lcolPermisosWin As Collection
        If GstrIdUsuario <> GCSTRUSUARIOU Then
            lcolPermisosWin = GobjPanorama.ObjUsuarioActual.FcolPermisosAccionesFrm(aenuIDVentana)
            For Each ldrwPermiso In lcolPermisosWin
                If ClsPanorama.FobjValorCampo(ldrwPermiso("IdAccion"), EnuTipoValor.enuInteger) =
                        aentIdAccion Then
                    lblnHabilitar = ClsPanorama.FobjValorCampo(ldrwPermiso("Permitida"),
                            EnuTipoValor.enuBoolean)
                    Exit For
                End If
            Next
        End If
        Return lblnHabilitar
    End Function

    Protected Function FblnHabilitarMenuPan(aentIdAccion As Integer) As Boolean
        Dim lblnHabilitar As Boolean = True
        If GstrIdUsuario <> GCSTRUSUARIOU Then
            If IsNothing(McolPermisosWin) Then
                McolPermisosWin = GobjPanorama.ObjUsuarioActual.FcolPermisosAccionesFrm(EnuIdVentana)
            End If
            For Each ldrwPermiso In McolPermisosWin
                If ClsPanorama.FobjValorCampo(ldrwPermiso("IdAccion"), EnuTipoValor.EnuInteger) =
                        aentIdAccion Then
                    lblnHabilitar = ClsPanorama.FobjValorCampo(ldrwPermiso("Permitida"),
                            EnuTipoValor.EnuBoolean)
                    Exit For
                End If
            Next
        End If
        Return lblnHabilitar
    End Function

    Protected Shared Sub SHabiliteMenuItem(ablnHabilita As Boolean, amnuiMenuItem As MenuItem)
        If amnuiMenuItem IsNot Nothing Then
            If ablnHabilita Then
                amnuiMenuItem.IsEnabled = True
                amnuiMenuItem.Opacity = 1
            Else
                amnuiMenuItem.IsEnabled = False
                amnuiMenuItem.Opacity = 0.5
            End If
        End If
    End Sub

    Protected Sub SHabiliteMenu(ablnHabilita As Boolean, amnuiMenuItem As MenuItem)
        If amnuiMenuItem IsNot Nothing Then
            If ablnHabilita Then
                amnuiMenuItem.Style = FindResource("RecMnuItemPriHab")
            Else
                amnuiMenuItem.Style = FindResource("RecMnuItemPriDes")
            End If
        End If
    End Sub

    Protected Shared Sub SHabiliteMenuItemPan(ablnHabilita As Boolean, amnuiMenuItem As MenuItem)
        If amnuiMenuItem IsNot Nothing Then
            If ablnHabilita Then
                amnuiMenuItem.IsEnabled = True
                amnuiMenuItem.Opacity = 1
            Else
                amnuiMenuItem.IsEnabled = False
                amnuiMenuItem.Opacity = 0.5
            End If
        End If
    End Sub

    Protected Sub SHabiliteMenuItemPan(ablnHabilita As Boolean,
            amnuiMenuItemPan As MenuItemPan)
        If amnuiMenuItemPan.BlnEsPrimario Then
            SHabiliteMenu(ablnHabilita, amnuiMenuItemPan)
            Exit Sub
        End If
        If amnuiMenuItemPan IsNot Nothing Then
            If ablnHabilita Then
                amnuiMenuItemPan.IsEnabled = True
                amnuiMenuItemPan.Opacity = 1
            Else
                amnuiMenuItemPan.IsEnabled = False
                amnuiMenuItemPan.Opacity = 0.5
            End If
        End If
    End Sub
#End Region

#Region "Habilita o deshabilita los controles"
    ' Habilita o deshabilita los controles de la ventana dependiendo de su estado
    ' y otras consideraciones"
    Protected Sub SHabiliteWin(ablnHabilita As Boolean)
        MblnHabilitadosCtlsLlave = False
        SHabiliteControles(ablnHabilita, MgrdGridPrincipal)
        If Not MblnVenProceso Then
            SMuestreEstado()
        End If
    End Sub

    Protected Sub SVentanaAuxiliar()
        If BlnVentanaAux Then
            For Each lctrLlave As Control In McolControlesLlave
                lctrLlave.Style = FindResource("RecCtlNoHabilitado")
            Next
            Dim lstrMens = "VENTANA AUXILIAR - " & StrNombreVentana
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub

    Private Sub SHabiliteControles(ablnHabilita As Boolean, aobjObjeto As Object)
        If TypeOf aobjObjeto Is Control Then
            If FblnEsControlHabilitable(aobjObjeto) Then
                SHabiliteControl(ablnHabilita, aobjObjeto)
            Else
                If TypeOf aobjObjeto Is DependencyObject Then
                    For Each lobjHijo As Object In LogicalTreeHelper.GetChildren(aobjObjeto)
                        SHabiliteControles(ablnHabilita, lobjHijo)
                    Next
                End If
            End If
        Else
            If TypeOf aobjObjeto Is DependencyObject Then
                For Each lobjHijo As Object In LogicalTreeHelper.GetChildren(aobjObjeto)
                    If Not (TypeOf lobjHijo Is DockPanel AndAlso lobjHijo.Name = "PanelControl") Then
                        SHabiliteControles(ablnHabilita, lobjHijo)
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub SHabiliteControl(ablnHabilita As Boolean, actlControl As Control)
        Dim lctlControl = actlControl
        Dim lblnEsControlLlave = False
        Dim lblnControlRestringido = False
        If Not (MblnHabilitadosCtlsLlave OrElse IsNothing(McolControlesLlave)) Then
            lblnEsControlLlave = FblnHabilitoControlLlave(lctlControl)
        End If
        If Not lblnEsControlLlave Then
            If Not IsNothing(McolControlesRestringidos) Then
                For i = 1 To McolControlesRestringidos.Count
                    If lctlControl.Equals(McolControlesRestringidos(i)) Then
                        lblnControlRestringido = True
                        Exit For
                    End If
                Next i
            End If
            If Not lblnControlRestringido Then
                If TypeOf actlControl Is Button Then
                    Select Case actlControl.Name
                        Case "bttSearch", "bttOperation"
                            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
                                actlControl.Style = FindResource("RecBttAceDesha")
                            Else
                                actlControl.Style = FindResource("RecBttAceCan")
                            End If
                        Case Else
                            If Not (lctlControl.Equals(HbttAceptar) OrElse
                                    lctlControl.Equals(HbttCancelar)) Then
                                If ablnHabilita Then
                                    actlControl.Style = FindResource("RecCtlHabilitado")
                                Else
                                    actlControl.Style = FindResource("RecCtlNoHabilitado")
                                End If
                            End If
                    End Select
                Else
                    If ablnHabilita Then
                        actlControl.Style = FindResource("RecCtlHabilitado")
                    Else
                        actlControl.Style = FindResource("RecCtlNoHabilitado")
                    End If
                End If
            End If
        Else
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuModificando Then
                actlControl.Style = FindResource("RecCtlNoHabilitado")
            Else
                actlControl.Style = FindResource("RecCtlHabilitado")
            End If
        End If
    End Sub

    Private Function FblnHabilitoControlLlave(actlControl As Control) As Boolean
        Dim lblnHabCtlLlave = False
        Static lentCantControlesLlave As Integer
        If Not IsNothing(McolControlesLlave) Then
            For i = 1 To McolControlesLlave.Count
                If actlControl.Equals(McolControlesLlave(i)) Then
                    lblnHabCtlLlave = True
                    Select Case EnuOperacionEnWin
                        Case EnuOperacionEnVentana.CenuConsultando
                            If ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                                actlControl.Style = FindResource("RecCtlNoHabilitado")
                            Else
                                actlControl.Style = FindResource("RecCtlHabilitado")
                                If i = 1 Then
                                    actlControl.Focus()
                                End If
                            End If
                        Case EnuOperacionEnVentana.CenuCreando
                            If IsNothing(ObjValorLlave) Then
                                actlControl.Style = FindResource("RecCtlHabilitado")
                            Else
                                actlControl.Style = FindResource("RecCtlNoHabilitado")
                            End If
                        Case EnuOperacionEnVentana.CenuModificando
                            actlControl.Style = FindResource("RecCtlNoHabilitado")
                    End Select
                    lentCantControlesLlave += 1
                    If lentCantControlesLlave = McolControlesLlave.Count Then
                        MblnHabilitadosCtlsLlave = True
                        lentCantControlesLlave = 0
                    End If
                    Exit For
                End If
            Next
        End If
        Return lblnHabCtlLlave
    End Function

    Private Shared Function FblnEsControlHabilitable(actlControl As Control) As Boolean
        Dim lblnEsCtrlHab = False
        If TypeOf actlControl Is TextBox OrElse TypeOf actlControl Is ComboBox OrElse
                TypeOf actlControl Is DatePicker OrElse TypeOf actlControl Is CheckBox OrElse
                TypeOf actlControl Is RadioButton OrElse TypeOf actlControl Is Button OrElse
                TypeOf actlControl Is DataGrid OrElse TypeOf actlControl Is PasswordBox OrElse
                TypeOf actlControl Is ListBox Then
            lblnEsCtrlHab = True
        End If
        Return lblnEsCtrlHab
    End Function
#End Region

#Region "Validacion"
    Protected Property StcValidaControl(aintIndice As Integer) As Control
        Get
            Return MstcValidador(aintIndice).CtlControl
        End Get
        Set(value As Control)
            MstcValidador(aintIndice).CtlControl = value
        End Set
    End Property

    Protected Property StcValidValido(aintIndice As Integer) As Boolean
        Get
            Return MstcValidador(aintIndice).BlnValido
        End Get
        Set(value As Boolean)
            MstcValidador(aintIndice).BlnValido = value
        End Set
    End Property

    Protected Function FblnEstanTodosBien(ablnNotifique As Boolean) As Boolean
        Dim lblnValido As Boolean = True
        Dim i As Integer
        If MentItemsValidar > 0 Then
            For i = MstcValidador.GetLowerBound(0) To MstcValidador.GetUpperBound(0)
                If MstcValidador(i).BlnValido Then
                    If TypeOf MstcValidador(i).CtlControl Is Label Then
                        MstcValidador(i).CtlControl.Style = FindResource("RecCtlValido")
                    ElseIf TypeOf MstcValidador(i).CtlControl Is CheckBox Then
                        MstcValidador(i).CtlControl.Foreground = Brushes.Black
                        MstcValidador(i).CtlControl.FontWeight = FontWeights.Normal
                    ElseIf TypeOf MstcValidador(i).CtlControl Is GroupBox Then
                        MstcValidador(i).CtlControl.Foreground = Brushes.Black
                        MstcValidador(i).CtlControl.FontWeight = FontWeights.Normal
                    ElseIf TypeOf MstcValidador(i).CtlControl Is RadioButton Then
                        MstcValidador(i).CtlControl.Foreground = Brushes.Black
                        MstcValidador(i).CtlControl.FontWeight = FontWeights.Normal
                    End If
                Else
                    If TypeOf MstcValidador(i).CtlControl Is Label Then
                        MstcValidador(i).CtlControl.Style = FindResource("RecLblInvalido")
                    ElseIf TypeOf MstcValidador(i).CtlControl Is CheckBox Then
                        MstcValidador(i).CtlControl.Foreground = Brushes.Red
                        MstcValidador(i).CtlControl.FontWeight = FontWeights.Bold
                    ElseIf TypeOf MstcValidador(i).CtlControl Is GroupBox Then
                        MstcValidador(i).CtlControl.Foreground = Brushes.Red
                        MstcValidador(i).CtlControl.FontWeight = FontWeights.Bold
                    ElseIf TypeOf MstcValidador(i).CtlControl Is RadioButton Then
                        MstcValidador(i).CtlControl.Foreground = Brushes.Red
                        MstcValidador(i).CtlControl.FontWeight = FontWeights.Bold
                    End If
                    lblnValido = False
                End If
            Next i
        End If
        If Not IsNothing(HtlbMiBarraHerramientas) Then
            Dim lblnHabilitar = False
            Select Case EnuOperacionEnWin
                Case EnuOperacionEnVentana.CenuConsultando
                    lblnHabilitar = False
                Case EnuOperacionEnVentana.CenuCreando, EnuOperacionEnVentana.CenuModificando
                    lblnHabilitar = lblnValido
            End Select
            SHabiliteBotonTlb(lblnHabilitar, HbttGuardar)
            SHabiliteMenuItem(lblnHabilitar, HmnuGuardar)
        End If
        If Not IsNothing(HbttAceptar) Then
            If lblnValido Then
                HbttAceptar.Style = FindResource("RecBttAceCan")
                If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando OrElse
                        EnuIdVentana = EnuIdVentanaDef.EnuAutorizaDscto OrElse
                        EnuIdVentana = EnuIdVentanaDef.EnuLogOn Then
                    If ablnNotifique Then
                        SLevanteEveNoti("", "", 0, EnuSeveridadNot.EnuOk)
                    End If
                End If
            Else
                HbttAceptar.Style = FindResource("RecBttAceDesha")
            End If
        End If
        If ablnNotifique AndAlso Not (HblnCargandoForma OrElse lblnValido) AndAlso
                ObjObjetoWin IsNot Nothing Then
            SLevanteEveNoti(My.Resources.CamposSinSatisfacer, "", 0,
                    EnuSeveridadNot.EnuCamInsatis)
        End If
        Return lblnValido
    End Function

    Protected Function FblnEstanTodosBien() As Boolean
        Dim lblnValido As Boolean = True
        Dim i As Integer
        If MentItemsValidar > 0 Then
            For i = MstcValidador.GetLowerBound(0) To MstcValidador.GetUpperBound(0)
                If MstcValidador(i).BlnValido Then
                    If TypeOf MstcValidador(i).CtlControl Is Label Then
                        MstcValidador(i).CtlControl.Style = FindResource("RecCtlValido")
                    ElseIf TypeOf MstcValidador(i).CtlControl Is CheckBox Then
                        MstcValidador(i).CtlControl.Foreground = Brushes.Black
                        MstcValidador(i).CtlControl.FontWeight = FontWeights.Normal
                    ElseIf TypeOf MstcValidador(i).CtlControl Is GroupBox Then
                        MstcValidador(i).CtlControl.Foreground = Brushes.Black
                        MstcValidador(i).CtlControl.FontWeight = FontWeights.Normal
                    ElseIf TypeOf MstcValidador(i).CtlControl Is RadioButton Then
                        MstcValidador(i).CtlControl.Foreground = Brushes.Black
                        MstcValidador(i).CtlControl.FontWeight = FontWeights.Normal
                    End If
                Else
                    If TypeOf MstcValidador(i).CtlControl Is Label Then
                        MstcValidador(i).CtlControl.Style = FindResource("RecLblInvalido")
                    ElseIf TypeOf MstcValidador(i).CtlControl Is CheckBox Then
                        MstcValidador(i).CtlControl.Foreground = Brushes.Red
                        MstcValidador(i).CtlControl.FontWeight = FontWeights.Bold
                    ElseIf TypeOf MstcValidador(i).CtlControl Is GroupBox Then
                        MstcValidador(i).CtlControl.Foreground = Brushes.Red
                        MstcValidador(i).CtlControl.FontWeight = FontWeights.Bold
                    ElseIf TypeOf MstcValidador(i).CtlControl Is RadioButton Then
                        MstcValidador(i).CtlControl.Foreground = Brushes.Red
                        MstcValidador(i).CtlControl.FontWeight = FontWeights.Bold
                    End If
                    lblnValido = False
                End If
            Next i
        End If
        If Not IsNothing(HtlbMiBarraHerramientas) Then
            Dim lblnHabilitar = False
            Select Case EnuOperacionEnWin
                Case EnuOperacionEnVentana.CenuConsultando
                    lblnHabilitar = False
                Case EnuOperacionEnVentana.CenuCreando, EnuOperacionEnVentana.CenuModificando
                    lblnHabilitar = lblnValido
            End Select
            SHabiliteBotonTlb(lblnHabilitar, HbttGuardar)
            SHabiliteMenuItem(lblnHabilitar, HmnuGuardar)
        End If
        If Not IsNothing(HbttAceptar) Then
            If lblnValido Then
                If Not (ObjObjetoWin.FblnEstaVacioOrigenDatos AndAlso EnuOperacionEnWin =
                        EnuOperacionEnVentana.CenuConsultando) OrElse
                        HenuIdVentana = EnuIdVentanaDef.EnuRevPrefacturas Then
                    HbttAceptar.Style = FindResource("RecBttAceCan")
                    SLevanteEveNoti("", "", 0, EnuSeveridadNot.EnuOk)
                End If
            Else
                HbttAceptar.Style = FindResource("RecBttAceDesha")
            End If
        End If
        If Not (HblnCargandoForma OrElse lblnValido) AndAlso ObjObjetoWin IsNot Nothing Then
            If Not (ObjObjetoWin.FblnEstaVacioOrigenDatos AndAlso EnuOperacionEnWin =
                    EnuOperacionEnVentana.CenuConsultando) OrElse
                    HenuIdVentana = EnuIdVentanaDef.EnuLogOn Then
                SLevanteEveNoti(My.Resources.CamposSinSatisfacer, "", 0,
                    EnuSeveridadNot.EnuCamInsatis)
            End If
        End If
        Return lblnValido
    End Function

    Protected Overridable Function FblnNotificaOk(aenuIdMensNot As EnuIdMens) As Boolean
        Return True
    End Function
#End Region
#End Region

#Region "Comandos"
    Private Sub SConfigureComandos()
        SConfigureCmdCrear()
        SConfigureCmdModificar()
        SConfigureCmdGuardar()
        SConfigureCmdSuprimir()
        SConfigureCmdAnular()
        SConfigureCmdRefrescar()
        SConfigureCmdAlPrimero()
        SConfigureCmdAlAnterior()
        SConfigureCmdAlSiguiente()
        SConfigureCmdAlUltimo()
        SConfigureCmdBuscar()
        SConfigureCmdCerrarWin()
        SConfigureCmdAceptar()
        SConfigureCmdCancelar()
        SConfigureCmdCalendario()
        SConfigureCmdCalculadora()
        SConfigureCmdNotas()
        SConfigureCmdMensajes()
        SConfigureCmdImprimir()
        SConfigureCmdAyuda()
        If MblnEsVentanaPrincipal Then
            SConfigureCmdSalirApp()
        End If
    End Sub

#Region "Comando Crear" 'Ok
    Private Sub SConfigureCmdCrear()
        If HbttCrear IsNot Nothing AndAlso (HbttCrear.Visibility = Visibility.Visible) Then
            Dim cmdCrear As New CommandBinding(ClsComandos.RucCrear, AddressOf CmdCrear_Executed,
                                                AddressOf CmdCrear_CanExecute)
            CommandBindings.Add(cmdCrear)
            HbttCrear.Command = ClsComandos.RucCrear
            HmnuCrear.Command = ClsComandos.RucCrear
        End If
    End Sub
    Private Sub CmdCrear_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SCrearClic()
    End Sub
    Private Sub CmdCrear_CanExecute(sender As System.Object, e As System.Windows.Input.CanExecuteRoutedEventArgs)
        Dim lblnHabilite = False
        If Not IsNothing(HbttCrear) Then
            lblnHabilite = (HbttCrear.Visibility = Visibility.Visible) AndAlso
                    CType(EnuTipoPermisoObjWin And EnuPermisosDef.enuCrear, Boolean)
        End If
        e.CanExecute = lblnHabilite
    End Sub
#End Region

#Region "Comando Modificar" 'Ok
    Private Sub SConfigureCmdModificar()
        If Not IsNothing(HbttModificar) AndAlso (HbttModificar.Visibility = Visibility.Visible) Then
            Dim cmdModificar As New CommandBinding(ClsComandos.RucModificar, AddressOf CmdModificar_Executed,
                                                        AddressOf CmdModificar_CanExecute)
            CommandBindings.Add(cmdModificar)
            HbttModificar.Command = ClsComandos.RucModificar
            HmnuModificar.Command = ClsComandos.RucModificar
        End If
    End Sub
    Private Sub CmdModificar_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SModificarClic()
    End Sub
    Private Sub CmdModificar_CanExecute(sender As System.Object, e As System.Windows.Input.CanExecuteRoutedEventArgs)
        Dim lblnHabilite = False
        If Not IsNothing(HbttModificar) Then
            If EnuIdVentana = EnuIdVentanaDef.EnuImportar OrElse
                    EnuIdVentana = EnuIdVentanaDef.EnuUbicaciones OrElse
                    EnuIdVentana = EnuIdVentanaDef.EnuOpcionesBK OrElse
                    EnuIdVentana = EnuIdVentanaDef.EnuCuentasContabilidad Then
                lblnHabilite = True
            Else
                lblnHabilite = (HbttModificar.Visibility = Visibility.Visible) AndAlso
                        ObjObjetoWin IsNot Nothing AndAlso
                        (Not ObjObjetoWin.FblnEstaVacioOrigenDatos) AndAlso ObjObjetoWin.BlnExiste AndAlso
                        CType(EnuTipoPermisoObjWin And EnuPermisosDef.enuModificar, Boolean) AndAlso
                        ObjObjetoWin.FblnEsModificable()
            End If
        End If
        e.CanExecute = lblnHabilite
    End Sub
#End Region

#Region "Comando Guardar" 'Ok
    Private Sub SConfigureCmdGuardar()
        If Not IsNothing(HbttGuardar) AndAlso (HbttGuardar.Visibility = Visibility.Visible) Then
            Dim cmdGuardar As New CommandBinding(ClsComandos.RucGuardar,
                    AddressOf CmdGuardar_Executed, AddressOf CmdGuardar_CanExecute)
            CommandBindings.Add(cmdGuardar)
            HbttGuardar.Command = ClsComandos.RucGuardar
            If Not IsNothing(HmnuGuardar) Then
                HmnuGuardar.Command = ClsComandos.RucGuardar
            End If
        End If
    End Sub
    Private Sub CmdGuardar_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SGuardarClic()
    End Sub
    Private Sub CmdGuardar_CanExecute(sender As System.Object, e As System.Windows.Input.CanExecuteRoutedEventArgs)
        Dim lblnHabilite = FblnEstanTodosBien(False)
        If Not IsNothing(HbttGuardar) Then
            If EnuIdVentana = EnuIdVentanaDef.EnuUbicaciones Then
                Dim lobjUbicacion As Ubicacion.ClsUbicacion = ObjObjetoWin
                lblnHabilite = lobjUbicacion.EnuTipoAccion <> EnuTipoAccionDef.None
            Else
                lblnHabilite = (HbttGuardar.Visibility = Visibility.Visible) AndAlso
                        Not IsNothing(ObjObjetoWin)
            End If
        End If
        e.CanExecute = lblnHabilite
    End Sub
#End Region

#Region "Comando Suprimir" 'Ok
    Private Sub SConfigureCmdSuprimir()
        If Not IsNothing(HbttSuprimir) AndAlso (HbttSuprimir.Visibility = Visibility.Visible) Then
            Dim cmdSuprimir As New CommandBinding(ClsComandos.RucSuprimir, AddressOf CmdSuprimir_Executed,
                                                  AddressOf CmdSuprimir_CanExecute)
            CommandBindings.Add(cmdSuprimir)
            HbttSuprimir.Command = ClsComandos.RucSuprimir
            HmnuSuprimir.Command = ClsComandos.RucSuprimir
        End If
    End Sub
    Private Sub CmdSuprimir_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SSuprimirClic()
    End Sub
    Private Sub CmdSuprimir_CanExecute(sender As System.Object, e As System.Windows.Input.CanExecuteRoutedEventArgs)
        Dim lblnHabilite = False
        If Not IsNothing(HbttSuprimir) Then
            lblnHabilite = (HbttSuprimir.Visibility = Visibility.Visible) AndAlso
                    (Not ObjObjetoWin.FblnEstaVacioOrigenDatos) AndAlso
                    (EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando) AndAlso
                    CType(EnuTipoPermisoObjWin And EnuPermisosDef.enuSuprimir, Boolean)
        End If
        e.CanExecute = lblnHabilite
    End Sub
#End Region

#Region "Comando Anular" 'Ok
    Private Sub SConfigureCmdAnular()
        If Not IsNothing(HbttAnular) AndAlso (HbttAnular.Visibility = Visibility.Visible) Then
            Dim cmdAnular As New CommandBinding(ClsComandos.RucAnular, AddressOf CmdAnular_Executed,
                                                        AddressOf CmdAnular_CanExecute)
            CommandBindings.Add(cmdAnular)
            HbttAnular.Command = ClsComandos.RucAnular
            HmnuAnular.Command = ClsComandos.RucAnular
        End If
    End Sub

    Private Sub CmdAnular_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SAnularClic()
    End Sub

    Private Sub CmdAnular_CanExecute(sender As System.Object, e As System.Windows.Input.CanExecuteRoutedEventArgs)
        Dim lblnHabilite = False
        If Not IsNothing(HbttAnular) Then
            lblnHabilite = (HbttAnular.Visibility = Visibility.Visible) AndAlso
                    (Not ObjObjetoWin.FblnEstaVacioOrigenDatos) AndAlso
                     (EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando) AndAlso
                     CType(EnuTipoPermisoObjWin And EnuPermisosDef.enuAnular, Boolean)
        End If
        e.CanExecute = lblnHabilite
    End Sub
#End Region
#Region "Comando Refrescar"
    Private Sub SConfigureCmdRefrescar()
        If Not IsNothing(HbttRefrescar) AndAlso (HbttRefrescar.Visibility = Visibility.Visible) Then
            Dim cmdRefrescar As New CommandBinding(ClsComandos.RucRefrescar, AddressOf CmdRefrescar_Executed,
                                                     AddressOf CmdRefrescar_CanExecute)
            CommandBindings.Add(cmdRefrescar)
            HbttRefrescar.Command = ClsComandos.RucRefrescar
            HmnuRefrescar.Command = ClsComandos.RucRefrescar
        End If
    End Sub
    Private Sub CmdRefrescar_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SRefrescarClic()
    End Sub
    Private Sub CmdRefrescar_CanExecute(sender As System.Object, e As System.Windows.Input.CanExecuteRoutedEventArgs)
        Dim lblnHabilite = False
        If Not IsNothing(HbttRefrescar) Then
            lblnHabilite = EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando
        End If
        e.CanExecute = lblnHabilite
    End Sub
#End Region
#Region "Comando AlPrimero"
    Private Sub SConfigureCmdAlPrimero()
        If Not IsNothing(HbttAlPrimero) AndAlso (HbttAlPrimero.Visibility = Visibility.Visible) Then
            Dim cmdAlPrimero As New CommandBinding(ClsComandos.RucAlPrimero, AddressOf CmdAlPrimero_Executed,
                                                       AddressOf CmdAlPrimero_CanExecute)
            CommandBindings.Add(cmdAlPrimero)
            HbttAlPrimero.Command = ClsComandos.RucAlPrimero
            If Not IsNothing(HmnuAlPrimero) Then
                HmnuAlPrimero.Command = ClsComandos.RucAlPrimero
            End If
        End If
    End Sub
    Private Sub CmdAlPrimero_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SAlPrimeroClic()
    End Sub
    Private Sub CmdAlPrimero_CanExecute(sender As System.Object, e As System.Windows.Input.CanExecuteRoutedEventArgs)
        Dim lblnHabilitar As Boolean = ObjObjetoWin.BlnEsNavegable AndAlso Not IsNothing(HbttAlPrimero) AndAlso
                HbttAlPrimero.Visibility = Visibility.Visible
        If lblnHabilitar Then
            lblnHabilitar = Not ObjObjetoWin.FblnEsElPrimerRegistro AndAlso
                    EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando
        End If
        e.CanExecute = lblnHabilitar
    End Sub
#End Region
#Region "Comando AlAnterior"
    Private Sub SConfigureCmdAlAnterior()
        If Not IsNothing(HbttAlAnterior) AndAlso (HbttAlAnterior.Visibility = Visibility.Visible) Then
            Dim cmdAlAnterior As New CommandBinding(ClsComandos.RucAlAnterior, AddressOf CmdAlAnterior_Executed,
                                                    AddressOf CmdAlAnterior_CanExecute)
            CommandBindings.Add(cmdAlAnterior)
            HbttAlAnterior.Command = ClsComandos.RucAlAnterior
            If Not IsNothing(HmnuAlAnterior) Then
                HmnuAlAnterior.Command = ClsComandos.RucAlAnterior
            End If
        End If
    End Sub

    Private Sub CmdAlAnterior_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SAlAnteriorClic()
    End Sub

    Private Sub CmdAlAnterior_CanExecute(sender As System.Object, e As System.Windows.Input.CanExecuteRoutedEventArgs)
        Dim lblnHabilitar As Boolean = ObjObjetoWin.BlnEsNavegable AndAlso Not IsNothing(HbttAlAnterior) AndAlso
                HbttAlAnterior.Visibility = Visibility.Visible
        If lblnHabilitar Then
            lblnHabilitar = Not ObjObjetoWin.FblnEsElPrimerRegistro AndAlso
                    EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando
        End If
        e.CanExecute = lblnHabilitar
    End Sub
#End Region
#Region "Comando AlSiguiente"
    Private Sub SConfigureCmdAlSiguiente()
        If Not IsNothing(HbttAlSiguiente) AndAlso (HbttAlSiguiente.Visibility = Visibility.Visible) Then
            Dim cmdAlSiguiente As New CommandBinding(ClsComandos.RucAlSiguiente, AddressOf CmdAlSiguiente_Executed,
                                                   AddressOf CmdAlSiguiente_CanExecute)
            CommandBindings.Add(cmdAlSiguiente)
            HbttAlSiguiente.Command = ClsComandos.RucAlSiguiente
            If Not IsNothing(HmnuAlSiguiente) Then
                HmnuAlSiguiente.Command = ClsComandos.RucAlSiguiente
            End If
        End If
    End Sub
    Private Sub CmdAlSiguiente_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SAlSiguienteClic()
    End Sub
    Private Sub CmdAlSiguiente_CanExecute(sender As System.Object, e As System.Windows.Input.CanExecuteRoutedEventArgs)
        Dim lblnHabilitar As Boolean = ObjObjetoWin.BlnEsNavegable AndAlso Not IsNothing(HbttAlSiguiente) AndAlso
                HbttAlSiguiente.Visibility = Visibility.Visible
        If lblnHabilitar Then
            lblnHabilitar = Not ObjObjetoWin.FblnEsElUltimoRegistro AndAlso
                    EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando
        End If
        e.CanExecute = lblnHabilitar
    End Sub
#End Region
#Region "Comando Al Ultimo"
    Private Sub SConfigureCmdAlUltimo()
        If Not IsNothing(HbttAlUltimo) AndAlso (HbttAlUltimo.Visibility = Visibility.Visible) Then
            Dim cmdAlUltimo As New CommandBinding(ClsComandos.RucAlUltimo, AddressOf CmdAlUltimo_Executed,
                                                      AddressOf CmdAlUltimo_CanExecute)
            CommandBindings.Add(cmdAlUltimo)
            HbttAlUltimo.Command = ClsComandos.RucAlUltimo
            If Not IsNothing(HmnuAlUltimo) Then
                HmnuAlUltimo.Command = ClsComandos.RucAlUltimo
            End If
        End If
    End Sub

    Private Sub CmdAlUltimo_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SAlUltimoClic()
    End Sub

    Private Sub CmdAlUltimo_CanExecute(sender As System.Object, e As System.Windows.Input.CanExecuteRoutedEventArgs)
        Dim lblnHabilitar As Boolean = ObjObjetoWin.BlnEsNavegable AndAlso Not IsNothing(HbttAlUltimo) AndAlso
                HbttAlUltimo.Visibility = Visibility.Visible
        If lblnHabilitar Then
            lblnHabilitar = Not ObjObjetoWin.FblnEsElUltimoRegistro AndAlso
                    EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando
        End If
        e.CanExecute = lblnHabilitar
    End Sub
#End Region
#Region "Comando Buscar"
    Private Sub SConfigureCmdBuscar()
        If Not IsNothing(HbttBuscar) AndAlso (HbttBuscar.Visibility = Visibility.Visible) Then
            Dim cmdBuscar As New CommandBinding(ClsComandos.RucBuscar, AddressOf CmdBuscar_Executed,
                                                    AddressOf CmdBuscar_CanExecute)
            CommandBindings.Add(cmdBuscar)
            HbttBuscar.Command = ClsComandos.RucBuscar
            If Not IsNothing(HmnuBuscar) Then
                HmnuBuscar.Command = ClsComandos.RucBuscar
            End If
        End If
    End Sub

    Private Sub CmdBuscar_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SBuscarClic()
    End Sub

    Private Sub CmdBuscar_CanExecute(sender As System.Object, e As System.Windows.Input.CanExecuteRoutedEventArgs)
        Dim lblnNavegar As Boolean = ObjObjetoWin.BlnEsNavegable
        If lblnNavegar Then
            If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
                lblnNavegar = False
            End If
        End If
        e.CanExecute = lblnNavegar
    End Sub
#End Region
#Region "Comando Salir"
    Private Sub SConfigureCmdSalirApp()
        If Not IsNothing(HbttSalir) AndAlso (HbttSalir.Visibility = Visibility.Visible) Then
            Dim cmdSalirApp As New CommandBinding(ClsComandos.RucSalir, AddressOf CmdSalirApp_Executed,
                        AddressOf CmdSalirApp_CanExecute)
            CommandBindings.Add(cmdSalirApp)
            HbttSalir.Command = ClsComandos.RucSalir
            HmnuSalir.Command = ClsComandos.RucSalir
        End If
    End Sub
    Private Sub CmdSalirApp_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SSalirClic()
    End Sub
    Private Sub CmdSalirApp_CanExecute(sender As System.Object, e As System.Windows.Input.CanExecuteRoutedEventArgs)
        Dim lblHabilitar = (MblnEsVentanaPrincipal) AndAlso
                HbttSalir.Visibility = Visibility.Visible
        e.CanExecute = lblHabilitar
    End Sub
#End Region
#Region "Comando Cerrar"
    Private Sub SConfigureCmdCerrarWin()
        If Not IsNothing(HbttCerrar) AndAlso (HbttCerrar.Visibility = Visibility.Visible) Then
            Dim cmdCerrar As New CommandBinding(ClsComandos.RucCerrar, AddressOf CmdCerrar_Executed,
                                                    AddressOf CmdCerrar_CanExecute)
            CommandBindings.Add(cmdCerrar)
            HbttCerrar.Command = ClsComandos.RucCerrar
            HmnuCerrar.Command = ClsComandos.RucCerrar
        End If
    End Sub

    Private Sub CmdCerrar_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SCerrarClic()
    End Sub

    Private Sub CmdCerrar_CanExecute(sender As System.Object, e As System.Windows.Input.CanExecuteRoutedEventArgs)
        Dim lblHabilitar = EnuIdVentana <> EnuIdVentanaDef.EnuMWOrionCop AndAlso
                    EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando
        e.CanExecute = lblHabilitar
    End Sub
#End Region
#Region "Comando Aceptar"
    Private Sub SConfigureCmdAceptar()
        If Not IsNothing(HbttAceptar) AndAlso (HbttAceptar.Visibility = Visibility.Visible) Then
            Dim cmdAceptar As New CommandBinding(ClsComandos.RucAceptar, AddressOf CmdAceptar_Executed,
                                                            AddressOf CmdAceptar_CanExecute)
            CommandBindings.Add(cmdAceptar)
            HbttAceptar.Command = ClsComandos.RucAceptar
        End If
    End Sub
    Private Sub CmdAceptar_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SAceptarClic()
    End Sub
    Private Sub CmdAceptar_CanExecute(sender As System.Object, e As System.Windows.Input.CanExecuteRoutedEventArgs)
        Dim lblnHabilitar = Not IsNothing(HbttCancelar) AndAlso FblnEstanTodosBien(False)
        e.CanExecute = lblnHabilitar
    End Sub
#End Region
#Region "Comando Cancelar"
    Private Sub SConfigureCmdCancelar()
        If Not IsNothing(HbttCancelar) AndAlso (HbttCancelar.Visibility = Visibility.Visible) Then
            Dim cmdCancelar As New CommandBinding(ClsComandos.RucCancelar, AddressOf CmdCancelar_Executed,
                                                            AddressOf CmdCancelar_CanExecute)
            CommandBindings.Add(cmdCancelar)
            HbttCancelar.Command = ClsComandos.RucCancelar
        End If
    End Sub
    Private Sub CmdCancelar_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SCancelarClic()
    End Sub
    Private Sub CmdCancelar_CanExecute(sender As System.Object, e As System.Windows.Input.CanExecuteRoutedEventArgs)
        e.CanExecute = Not IsNothing(HbttCancelar)
    End Sub
#End Region
#Region "Comando Calendario"
    Private Sub SConfigureCmdCalendario()
        If Not IsNothing(HbttCalendario) Then
            Dim cmdCalendario As New CommandBinding(ClsComandos.RucCalendario, AddressOf CmdCalendario_Executed,
                                                AddressOf CmdCalendario_CanExecute)
            CommandBindings.Add(cmdCalendario)
            HbttCalendario.Command = ClsComandos.RucCalendario
            HmnuCalendario.Command = ClsComandos.RucCalendario
        End If
    End Sub
    Private Sub CmdCalendario_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SCalendarioClic()
    End Sub
    Private Sub CmdCalendario_CanExecute(sender As System.Object, e As System.Windows.Input.
            CanExecuteRoutedEventArgs)
        e.CanExecute = Not IsNothing(HbttCalendario)
    End Sub
#End Region
#Region "Comando Calculadora"
    Private Sub SConfigureCmdCalculadora()
        If Not IsNothing(HbttCalculadora) Then
            Dim cmdCalculadora As New CommandBinding(ClsComandos.RucCalculadora, AddressOf CmdCalculadora_Executed,
                                                AddressOf CmdCalculadora_CanExecute)
            CommandBindings.Add(cmdCalculadora)
            HbttCalculadora.Command = ClsComandos.RucCalculadora
            HmnuCalculadora.Command = ClsComandos.RucCalculadora
        End If
    End Sub
    Private Sub CmdCalculadora_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SCalculadoraClic()
    End Sub
    Private Sub CmdCalculadora_CanExecute(sender As System.Object, e As System.Windows.Input.CanExecuteRoutedEventArgs)
        e.CanExecute = Not IsNothing(HbttCalculadora)
    End Sub
#End Region
#Region "Comando Notas"
    Private Sub SConfigureCmdNotas()
        If Not IsNothing(HbttNotas) Then
            Dim cmdNotas As New CommandBinding(ClsComandos.RucNotas, AddressOf CmdNotas_Executed,
                                                AddressOf CmdNotas_CanExecute)
            CommandBindings.Add(cmdNotas)
            HbttNotas.Command = ClsComandos.RucNotas
            HmnuNotas.Command = ClsComandos.RucNotas
        End If
    End Sub

    Private Sub CmdNotas_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SNotasClic()
    End Sub

    Private Sub CmdNotas_CanExecute(sender As System.Object, e As System.Windows.Input.CanExecuteRoutedEventArgs)
        e.CanExecute = Not IsNothing(HbttNotas)
    End Sub
#End Region
#Region "Comando Mensajes"
    Private Sub SConfigureCmdMensajes()
        If Not IsNothing(HbttMensajes) Then
            Dim cmdMensajes As New CommandBinding(ClsComandos.RucMensajes, AddressOf CmdMensajes_Executed,
                                                AddressOf CmdMensajes_CanExecute)
            CommandBindings.Add(cmdMensajes)
            HbttMensajes.Command = ClsComandos.RucMensajes
            HmnuMensajes.Command = ClsComandos.RucMensajes
        End If
    End Sub

    Private Sub CmdMensajes_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SMensajesClic()
    End Sub

    Private Sub CmdMensajes_CanExecute(sender As System.Object, e As System.Windows.Input.CanExecuteRoutedEventArgs)
        e.CanExecute = Not IsNothing(HbttMensajes)
    End Sub
#End Region
#Region "Comando Imprimir"
    Private Sub SConfigureCmdImprimir()
        If Not IsNothing(HbttImprimir) Then
            Dim cmdImprimir As New CommandBinding(ClsComandos.RucImprimir, AddressOf CmdImprimir_Executed,
                                                AddressOf CmdImprimir_CanExecute)
            CommandBindings.Add(cmdImprimir)
            HbttImprimir.Command = ClsComandos.RucImprimir
            If Not IsNothing(HmnuImprimir) Then
                HmnuImprimir.Command = ClsComandos.RucImprimir
            End If
        End If
    End Sub

    Private Sub CmdImprimir_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SImprimirClic()
    End Sub

    Private Sub CmdImprimir_CanExecute(sender As System.Object, e As System.Windows.Input.CanExecuteRoutedEventArgs)
        e.CanExecute = Not (IsNothing(HbttImprimir) OrElse IsNothing(HmnuImprimir))
    End Sub
#End Region
#Region "Comando Ayuda"
    Private Sub SConfigureCmdAyuda()
        If Not IsNothing(HbttAyuda) Then
            Dim cmdAyuda As New CommandBinding(ClsComandos.RucAyuda, AddressOf CmdAyuda_Executed,
                                                AddressOf CmdAyuda_CanExecute)
            CommandBindings.Add(cmdAyuda)
            HbttAyuda.Command = ClsComandos.RucAyuda
            If Not IsNothing(HmnuContenidoAyuda) Then
                HmnuContenidoAyuda.Command = ClsComandos.RucAyuda
            End If
        End If
    End Sub

    Private Sub CmdAyuda_Executed(sender As System.Object, e As System.Windows.Input.ExecutedRoutedEventArgs)
        SAyudaClic()
    End Sub

    Private Sub CmdAyuda_CanExecute(sender As System.Object, e As System.Windows.Input.CanExecuteRoutedEventArgs)
        e.CanExecute = Not IsNothing(HbttAyuda)
    End Sub
#End Region
#End Region

#Region "Operaciones"
#Region "Menu Acciones"
    Protected Sub SCrearClic()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
                If EnuTipoPermisoObjWin And EnuPermisosDef.EnuCrear Then
                    SCree()
                End If
            Else
                SCancelarClic()
            End If
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanLException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanLException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                GobjPanDat.SControleProcesoObj(False)
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Protected Friend Overridable Sub SCree()
        SLevanteEveOk()
        GobjPanDat.SControleProcesoObj(True)
        If ObjObjetoWin.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            ObjObjetoWin.SCreeObj(ObjValorLlave)
        ElseIf ObjObjetoWin.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
            Throw New ErrorInesperadoPanLException("Estado no esperado del Objeto")
        End If
        EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando
        SMuestreDatos()
        SFrmAdicione()
        If (McolControlesLlave IsNot Nothing) AndAlso (Not IsNothing(ObjValorLlave)) Then
            For i As Byte = 1 To McolControlesLlave.Count
                McolControlesLlave(i).Text = ObjValorLlave(i - 1)
            Next
        End If
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Protected Sub SModificarClic()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
                If CType(EnuTipoPermisoObjWin And EnuPermisosDef.enuModificar, Boolean) Then
                    SModifique()
                End If
            Else
                If HbttCancelar IsNot Nothing Then
                    SCancelarClic()
                End If
            End If
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanLException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanLException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                GobjPanDat.SControleProcesoObj(False)
            Else
                EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando
                ObjObjetoWin.SNormaliceEstado(True)
                GobjPanDat.SControleProcesoObj(False, True)
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
    Protected Friend Overridable Sub SModifique()
        If ObjObjetoWin.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            EnuOperacionEnWin = EnuOperacionEnVentana.CenuModificando
            ObjObjetoWin.SModifique()
            If Not IsNothing(HbttCancelar) Then
                HbttCancelar.Content = My.Resources.Cancelar
            End If
            SHabiliteWin(True)
            If Not IsNothing(MctlControlDespuesDeLlave) Then
                If MctlControlDespuesDeLlave.IsEnabled Then
                    MctlControlDespuesDeLlave.Focus()
                End If
            End If
            SRegistre()
            SValide()
        Else
            Throw New ErrorInesperadoPanLException("Estado inesperado del objeto!")
        End If
    End Sub
    Protected Sub SSuprimirClic()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            SSuprima()
            lblnNoHayError = True
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ArgumentNullException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                GobjPanDat.SControleProcesoObj(False)
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuError)
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Protected Overridable Sub SSuprima()
        Dim lblnSuprimio = False
        Dim lstrMens = String.Empty
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            If EnuTipoPermisoObjWin And EnuPermisosDef.enuSuprimir Then
                If ObjObjetoWin.FblnEsSuprimible() Then
                    If MsgBox("Esta seguro de suprimir el presente " & ObjObjetoWin.StrNombreClase & "?",
                            MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Supresión") = MsgBoxResult.Yes Then
                        lstrMens = FstrNombreDoc()
                        lblnSuprimio = ObjObjetoWin.FblnSuprimio()
                        If lblnSuprimio Then
                            ObjObjetoWin.SVayaAlPrimero()
                            SMuestreDatos()
                        Else
                            SRefrescarClic()
                        End If
                    End If
                    If lblnSuprimio Then
                        If lstrMens.StartsWith("El") Then
                            lstrMens &= " fue suprimido exitosamente!"
                        Else
                            lstrMens &= " fue suprimida exitosamente!"
                        End If
                    Else
                        If lstrMens.StartsWith("El") Then
                            lstrMens &= " no fue suprimido!"
                        Else
                            lstrMens &= " no fue suprimida!"
                        End If
                    End If
                Else
                    lstrMens = "No es posible eliminar, por cuanto tiene información ligada!"
                End If
            Else
                lstrMens = "El Usuario actual no tiene Permisos para esta Acción!"
            End If
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Protected Sub SAnularClic()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Dim lblnAnulo As Boolean = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
                If CType(EnuTipoPermisoObjWin And EnuPermisosDef.enuAnular, Boolean) Then
                    lblnAnulo = SAnule()
                End If
            End If
            lblnNoHayError = True
        Catch ex As PanLException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
                GobjPanDat.SControleProcesoObj(False)
                If lblnAnulo Then
                    SRefrescarClic()
                End If
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Protected Overridable Function SAnule() As Boolean
        Dim lblnAnulo = False, lstrMens = String.Empty
        With ObjObjetoWin
            If CType(EnuTipoPermisoObjWin And EnuPermisosDef.EnuAnular, Boolean) Then
                Dim lobjCentroUtil As ClsCentroUtilidad =
                                GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual
                Dim lblnPuede As Boolean = lobjCentroUtil.ObjEstadoContratoByt.ObjValorPro <>
                                EnuEstadoContrato.EnuSuspendido
                If Not lblnPuede Then
                    lstrMens = "No es posible la anulación. El contrtrato está en estado suspendido!"
                End If
                If lblnPuede Then
                    If .FblnEsAnulable Then
                        If MsgBox("Desea Anular el presente Documento?", MsgBoxStyle.YesNo,
                              "Confirmar Anulación") = MsgBoxResult.Yes Then
                            .SAnule()
                            lblnAnulo = True
                            SFinaliceOperacion()
                            lstrMens = "El presente objeto fue Anulado!"
                        End If
                    Else
                        lstrMens = "El(La)  " & ObjObjetoWin.StrNombreClase &
                            " no puede ser anulado(a)!"
                    End If
                End If
            End If
        End With
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
        Return lblnAnulo
    End Function
    Protected Sub SGuardarClic()
        SGuarde()
#If DES = 1 Then
        Debug.Print(GobjPanDat.FentProceso.ToString & " Fin SGuardarClic")
#End If
    End Sub
    Protected Overridable Sub SGuarde()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Dim lblnGuardo As Boolean
        Try
            GobjPanDat.SControleProcesoObj(True)
            If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
                SRegistre()
                SValide()
                lblnGuardo = FblnGravo()
                If lblnGuardo Then
                    If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
                        lstrMens = FstrNombreDoc()
                        SFinaliceOperacion()
                        If lstrMens.StartsWith("El") Then
                            lstrMens &= " fue creado exitosamente!"
                        Else
                            lstrMens &= " fue creada exitosamente!"
                        End If
                        If BlnSiempreCreando Then
                            SCrearClic()
                        End If
                    Else
                        SFinaliceOperacion()
                    End If
                End If
            End If
            lblnNoHayError = True
        Catch ex As PanLException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ArgumentNullException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                GobjPanDat.SControleProcesoObj(False)
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
                SFinaliceOperacion()
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
    Protected Friend Overridable Function FblnGravo() As Boolean
        Dim lblnGravo = False
        If FblnEstanTodosBien() Then
            ObjObjetoWin.SActualice(True)
            lblnGravo = True
        End If
        Return lblnGravo
    End Function
    Protected Sub SRefrescarClic()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            Try
                GobjPanDat.SControleProcesoObj(True)
                SRefresqueWin()
                lblnNoHayError = True
            Catch ex As PanLException
                lstrMens &= ex.Message
                lstrMensEx = ex.ToString
            Catch ex As PanDatException
                lstrMens &= ex.Message
                lstrMensEx = ex.ToString
            Catch ex As Exception
                lstrMens &= ex.Message
                lstrMensEx = ex.ToString
            Finally
                If lblnNoHayError Then
                    GobjPanDat.SControleProcesoObj(False)
                Else
                    GobjPanDat.SAborteTransaccion()
                    GobjPanDat.SControleProcesoObj(False, True)
                    SFinaliceOperacion()
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
            End Try
        End If
    End Sub
    ''' <summary>
    ''' Refresca el objeto de la forma leyendolo de nuevo de la base de datos y muestra las
    ''' propeidades del objeto refrescado.
    ''' </summary>
    ''' <remarks>La operación de la ventana debe estar en Consultando!</remarks>
    Protected Friend Overridable Sub SRefresqueWin()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            If Not IsNothing(ObjObjetoWin) Then
                ObjObjetoWin.SRefresqueObj()
            End If
            SMuestreDatos()
            If ObjObjetoWin IsNot Nothing AndAlso ObjObjetoWin.BlnExiste Then
                SHabiliteMenues()
            End If
        End If
    End Sub
    Protected Overridable Sub SCerrarClic()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            If Not IsNothing(ObjObjetoWin) Then
                If GblnOK Then
                    If ObjObjetoWin.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                        lstrMens = "Estado del objeto inapropiado al cerrar la forma " &
                            StrNombreVentana & "!"
                        GblnOK = False
                    End If
                End If
            End If
            lblnNoHayError = True
        Catch ex As PanLException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuError)
                End If
                GobjPanDat.SControleProcesoObj(False)
                HblnSeEstaCerrando = True
                Close()
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
                SCerrarClic()
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
    Protected Sub SSalirClic()
        Close()
    End Sub
#End Region

#Region "Menu -> Herramientas"
    Protected Shared Sub SCalculadoraClic()
        Dim lstrTray As String
        lstrTray = FstrTrayecCalculadoraExe()
        If Not String.IsNullOrEmpty(lstrTray) Then
            Interaction.Shell(lstrTray, AppWinStyle.NormalFocus)
        End If
    End Sub
    Protected Shared Sub SCalendarioClic()
        Dim lfrmCalenda = New FrmCalendario
        lfrmCalenda.Show()
    End Sub
    Protected Shared Sub SNotasClic()
        Dim lstrTray As String
        If GenuIdAplicacion = EnuListaAplicaciones.EnuAdministrador Then
            lstrTray = FstrTrayecBlockNotasExe() & " " & GstrTrayDat & "NotasAdminOrion.txt"
        Else
            lstrTray = FstrTrayecBlockNotasExe() & " " & GstrTrayDatos & "\NotasOrionPlus.txt"
        End If
        If Not String.IsNullOrEmpty(lstrTray) Then
            Interaction.Shell(lstrTray, AppWinStyle.NormalFocus)
        End If
    End Sub
    Protected Shared Sub SMensajesClic()
        '
    End Sub
    Protected Sub SAyudaClic()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Dim lstrTrayAyuda = GstrTrayAppDat & "OriF1.pdf"
        Dim lpsiProceso As New ProcessStartInfo With {
            .FileName = lstrTrayAyuda
        }
        Try
            Process.Start(lpsiProceso)
            lblnNoHayError = True
        Catch ex As PanLException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Finally
            If Not lblnNoHayError Then
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
#End Region

#Region "Menu Navegar"
    Protected Sub SAlPrimeroClic()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            SNormaliceEstado()
            ObjObjetoWin.SVayaAlPrimero()
            SActualiceVentana()
            lblnNoHayError = True
        Catch ex As PanLException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ArgumentNullException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Finally
            If Not lblnNoHayError Then
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
    Protected Sub SAlAnteriorClic()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            SNormaliceEstado()
            ObjObjetoWin.SVayaAlAnterior()
            SActualiceVentana()
            lblnNoHayError = True
        Catch ex As PanLException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ArgumentNullException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Finally
            If Not lblnNoHayError Then
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
    Protected Sub SAlSiguienteClic()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            SNormaliceEstado()
            ObjObjetoWin.SVayaAlSiguiente()
            SActualiceVentana()
            lblnNoHayError = True
        Catch ex As PanLException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ArgumentNullException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Finally
            If Not lblnNoHayError Then
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
    Protected Sub SAlUltimoClic()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            SNormaliceEstado()
            ObjObjetoWin.SVayaAlUltimo()
            SActualiceVentana()
            lblnNoHayError = True
        Catch ex As PanLException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ArgumentNullException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Finally
            If Not lblnNoHayError Then
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
    Protected Overridable Sub SNavegueObj()
        '
    End Sub
    Private Sub SNormaliceEstado()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuModificando Then
            SRegistre()
            SFinaliceOperacion()
            SModificarClic()
        End If
    End Sub
    Private Sub SActualiceVentana()
        SNavegueObj()
        SRefresqueWin()
    End Sub
#End Region

#Region "Buscar"
    Protected Sub SBuscarClic()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            SBuscar()
            lblnNoHayError = True
        Catch ex As PanLException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Finally
            If Not lblnNoHayError Then
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
    Protected Overridable Sub SBuscar()
        Cursor = Input.Cursors.Wait
        If IsNothing(HwinBusqueda) Then
            HwinBusqueda = New WinBusqueda With {
                .WinPadre = Me
            }
        End If
        If FblnDefinioBusqueda() Then
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuModificando Then
                SRegistre()
                SModificarClic()
            End If
            HwinBusqueda.ShowDialog()
        End If
        HwinBusqueda = Nothing
        Cursor = Input.Cursors.Arrow
    End Sub
    Protected Overridable Sub SBuscar(astrTituloVentana As String, ablnOcultarBttCancelar As Boolean)
        Cursor = Input.Cursors.Wait
        If IsNothing(HwinBusqueda) Then
            HwinBusqueda = New WinBusqueda With {
                .WinPadre = Me
            }
            If Not String.IsNullOrEmpty(astrTituloVentana) Then
                HwinBusqueda.Title = astrTituloVentana
                HwinBusqueda.lblTitulo.Content = astrTituloVentana
            End If
            If ablnOcultarBttCancelar Then
                HwinBusqueda.SOculteBotonCancelar()
            End If
        End If
        If FblnDefinioBusqueda() Then
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuModificando Then
                SRegistre()
                SModificarClic()
            End If
            HwinBusqueda.ShowDialog()
        End If
        HwinBusqueda = Nothing
        Cursor = Input.Cursors.Arrow
    End Sub
    Protected Friend Overridable Function FblnDefinioBusqueda() As Boolean
        Return False
    End Function
    Friend Property StrResultadoBusqueda As String = String.Empty
    Friend Property StrResutadosBusqueda As String()
    Friend Property BlnBusquedaOk As Boolean = False
#End Region

#Region "Menu Otros"
    Protected Sub SImprimirClic()
        SImprima()
    End Sub

    Protected Overridable Sub SImprima()
        '
    End Sub
#End Region

#Region "Botones Aceptar y Cancelar"
    Protected Sub SAceptarClic()
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            SGuardarClic()
        Else
            If EnuIdVentana = EnuIdVentanaDef.EnuLogOn OrElse EnuIdVentana = EnuIdVentanaDef.EnuAutorizaDscto Then
                SGuardarClic()
                SCerrarClic()
            Else
                SCerrarClic()
            End If
        End If
    End Sub
    Protected Sub SCancelarClic()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            SCancele()
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanLException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanLException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                GobjPanDat.SControleProcesoObj(False)
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Protected Friend Overridable Sub SCancele()
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            SFinaliceOperacion()
            SRefresqueWin()
            If ObjObjetoWin IsNot Nothing AndAlso Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                SLevanteEveNoti("", "", 0, EnuSeveridadNot.EnuOk)
            End If
        Else
            SCerrarClic()
        End If
    End Sub
#End Region

#Region "Acciones complementarias"
    Protected Sub SFrmAdicione()
        If Not IsNothing(HbttCancelar) Then
            HbttCancelar.Content = My.Resources.Cancelar
        End If
        SHabiliteWin(True)
        SValide()
        If McolControlesLlave IsNot Nothing Then
            If McolControlesLlave.Count > 0 Then
                If McolControlesLlave(1).IsEnabled Then
                    McolControlesLlave(1).Focus()
                ElseIf MctlControlDespuesDeLlave IsNot Nothing Then
                    If MctlControlDespuesDeLlave.IsEnabled Then
                        MctlControlDespuesDeLlave.Focus()
                    End If
                End If
            ElseIf MctlControlDespuesDeLlave IsNot Nothing Then
                If MctlControlDespuesDeLlave.IsEnabled Then
                    MctlControlDespuesDeLlave.Focus()
                End If
            End If
        Else
            If Not IsNothing(MctlControlDespuesDeLlave) Then
                MctlControlDespuesDeLlave.Focus()
            End If
        End If
    End Sub
    Protected Friend Overridable Sub SFinaliceOperacion()
        Dim lstrMens As String = "Los datos del objeto " & ObjObjetoWin.StrNombreClase &
                " han cambiado!" & vbCrLf & "Desea guardar los cambios?"
        Dim lblnCreanado = EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            With ObjObjetoWin
                If .EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuConsultando Then
                    If FblnEstanTodosBien() AndAlso .BlnTengoCambios Then
                        If MsgBox(lstrMens, vbYesNo, "Aceptar Cambios") = vbYes Then
                            .SActualice(True)
                            If lblnCreanado Then
                                If BlnSiempreCreando Then
                                    lstrMens = FstrNombreDoc()
                                    If lstrMens.StartsWith("El") Then
                                        lstrMens &= " fue creado exitosamente!"
                                    Else
                                        lstrMens &= " fue creada exitosamente!"
                                    End If
                                    SEstablezcaWinConsultando()
                                    SCrearClic()
                                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                                Else
                                    .SNormaliceEstado(True)
                                    SEstablezcaWinConsultando()
                                End If
                            End If
                        Else
                            .SNormaliceEstado(True)
                            SEstablezcaWinConsultando()
                            If BlnCanceleCierra Then
                                SCerrarClic()
                            End If
                        End If
                    Else
                        .SNormaliceEstado(True)
                        SEstablezcaWinConsultando()
                        If BlnCanceleCierra Then
                            SCerrarClic()
                        End If
                    End If
                Else
                    .SNormaliceEstado(True)
                    SEstablezcaWinConsultando()
                End If
            End With
        End If
    End Sub
    Protected Friend Overridable Sub SEstablezcaWinConsultando()
        Dim lblnRefresque = (EnuOperacionEnWin <> EnuOperacionEnVentana.CenuCreando)
        EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando
        If Not IsNothing(HbttCancelar) Then
            HbttCancelar.Content = My.Resources.CerrarBtn
        End If
        SHabiliteWin(False)
        If EnuIdVentana <> EnuIdVentanaDef.EnuVerificaInt Then
            If lblnRefresque Then
                SRefrescarClic()
            End If
        End If
        If McolControlesLlave IsNot Nothing Then
            If McolControlesLlave.Count > 0 Then
                For Each lctrlControlLlave As Control In McolControlesLlave
                    lctrlControlLlave.IsEnabled = True
                Next
                If McolControlesLlave(1).IsEnabled Then
                    McolControlesLlave(1).Focus()
                End If
            End If
        End If
    End Sub
#End Region
#End Region

#Region "Procesos de Notificaciones por eventos"
    ' Definiciones
    Private MobjSenderProp As ClsCBPropiedad = Nothing
    Private MobjSenderObjPan As ClsCBObjetoPan = Nothing
    Private MobjNotiEveArg As ClsNotiEventArgs = Nothing
    Private MobjE As ClsNotiEventArgs = Nothing
    Private ReadOnly MblnNotiSonoras As Boolean = GobjAdministrador.BlnNotificacionSonora
    Protected Friend Sub SLevanteEveNoti(astrMensaje As String, astrMensajeEx As String, aentNroMens As Integer,
            aenuSevMens As EnuSeveridadNot)
        Dim lobjNotiEven As New ClsNotiEventArgs
        If aenuSevMens = EnuSeveridadNot.EnuError OrElse aenuSevMens = EnuSeveridadNot.EnuExcep Then
            If WinPadre IsNot Nothing AndAlso EnuIdVentana <> EnuIdVentanaDef.EnuLogOn Then
                GblnOK = False
                SCerrarClic()
                WinPadre.SLevanteEveNoti(astrMensaje, astrMensajeEx, aentNroMens, aenuSevMens)
            Else
                lobjNotiEven.SRegistreNotifica(astrMensaje, astrMensajeEx, aentNroMens, aenuSevMens)
                RaiseEvent EvnNotifica(Me, lobjNotiEven)
            End If
        Else
            lobjNotiEven.SRegistreNotifica(astrMensaje, astrMensajeEx, aentNroMens, aenuSevMens)
            RaiseEvent EvnNotifica(Me, lobjNotiEven)
        End If
    End Sub
    Protected Friend Sub SLevanteEveOk()
        Dim lobjNotiEven As New ClsNotiEventArgs
        lobjNotiEven.SRegistreNotifica(String.Empty, String.Empty, 0, EnuSeveridadNot.EnuOk)
        RaiseEvent EvnNotifica(Nothing, lobjNotiEven)
    End Sub
    ' Método que captura los eventos de notificación
    Private Sub Evn_Notifica(aobjSender As Object, e As ClsNotiEventArgs) _
            Handles ObjObjetoWin.EvnNotifica, ObjHijoObjWin.EvnNotifica, Me.EvnNotifica
        MobjE = e
        Static lblnBloquear As Boolean = False
        If MobjE.EnuSevNotifica = EnuSeveridadNot.EnuError OrElse
                MobjE.EnuSevNotifica = EnuSeveridadNot.EnuExcep Then
            MobjNotiEveArg = MobjE
            SProceseNoti()
            lblnBloquear = True
        End If
        If Not lblnBloquear Then
            If aobjSender Is Nothing AndAlso e.EnuSevNotifica = EnuSeveridadNot.EnuOk Then
                MobjSenderObjPan = Nothing
                MobjSenderProp = Nothing
                MobjNotiEveArg = MobjE
            End If
            If MobjNotiEveArg IsNot Nothing AndAlso
                    MobjE.EnuSevNotifica >= MobjNotiEveArg.EnuSevNotifica Then
                MobjNotiEveArg = MobjE
                If TypeOf aobjSender Is ClsCBObjetoPan Then
                    MobjSenderObjPan = aobjSender
                    MobjSenderProp = Nothing
                ElseIf TypeOf aobjSender Is ClsCBPropiedad Then
                    MobjSenderObjPan = Nothing
                    MobjSenderProp = aobjSender
                Else
                    MobjSenderObjPan = Nothing
                    MobjSenderProp = Nothing
                End If
            Else
                If TypeOf aobjSender Is ClsCBPropiedad Then
                    SProceseNotiProp(CType(aobjSender, ClsCBPropiedad))
                ElseIf TypeOf aobjSender Is ClsCBObjetoPan Then
                    SProceseNotiObjPan(aobjSender)
                ElseIf TypeOf aobjSender Is ClsFormInterface Then
                    SProceseNotiFI()
                End If
            End If
            If EnuIdVentana = EnuIdVentanaDef.EnuEFac AndAlso TypeOf aobjSender IsNot
                         ClsFormInterface Then
                Exit Sub
            End If
            SProceseNoti()
        End If
    End Sub
    Private Sub SProceseNotiProp(aobjSender As ClsCBPropiedad)
        If MobjSenderObjPan Is Nothing Then
            If MobjSenderProp Is Nothing Then
                MobjSenderProp = aobjSender
                MobjNotiEveArg = MobjE
            Else
                If MobjNotiEveArg.EnuSevNotifica = EnuSeveridadNot.EnuFalta Then
                    If MobjSenderProp.FblnNotiInfoOk(MobjNotiEveArg.EnuIdMensNot) Then
                        MobjSenderProp = aobjSender
                        MobjNotiEveArg = MobjE
                    End If
                ElseIf MobjNotiEveArg.EnuSevNotifica = EnuSeveridadNot.EnuDatoInvalido Then
                    If MobjSenderProp.BlnEsValido Then
                        MobjSenderProp = aobjSender
                        MobjNotiEveArg = MobjE
                    End If
                Else
                    MobjSenderProp = aobjSender
                    MobjNotiEveArg = MobjE
                End If
            End If
        Else
            If MobjE.EnuSevNotifica = EnuSeveridadNot.EnuDatoInvalido Then
                MobjSenderObjPan = Nothing
                MobjSenderProp = aobjSender
                MobjNotiEveArg = MobjE
            Else
                If MobjNotiEveArg.EnuIdMensNot > 0 Then
                    If MobjSenderObjPan.FblnNotificaOk(MobjNotiEveArg.EnuIdMensNot) Then
                        MobjSenderObjPan = Nothing
                        MobjSenderProp = aobjSender
                        MobjNotiEveArg = MobjE
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub SProceseNotiObjPan(aobjSender As ClsCBObjetoPan)
        If MobjSenderProp Is Nothing Then
            If MobjSenderObjPan Is Nothing Then
                MobjSenderObjPan = aobjSender
                MobjNotiEveArg = MobjE
            Else
                If MobjNotiEveArg.EnuIdMensNot > EnuIdMens.None Then
                    If MobjSenderObjPan.FblnNotificaOk(MobjNotiEveArg.EnuIdMensNot) Then
                        MobjSenderObjPan = aobjSender
                        MobjNotiEveArg = MobjE
                    End If
                Else
                    MobjSenderObjPan = aobjSender
                    MobjNotiEveArg = MobjE
                End If
            End If
        Else
            If MobjSenderProp.ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                If MobjSenderProp.BlnEsValido Then
                    MobjSenderProp = Nothing
                    MobjSenderObjPan = aobjSender
                    MobjNotiEveArg = MobjE
                End If
            Else
                MobjSenderProp = Nothing
                MobjSenderObjPan = aobjSender
                MobjNotiEveArg = MobjE
            End If
        End If
    End Sub
    Private Sub SProceseNotiFI()
        Dim lblnProcesar As Boolean = MobjSenderProp Is Nothing AndAlso MobjSenderObjPan Is Nothing
        If Not lblnProcesar Then
            If MobjSenderProp IsNot Nothing Then
                If MobjSenderProp.ObjPadre.EnuEstadoActualizacion <>
                        EnuEstadoObjetoDef.enuConsultando OrElse
                        HenuIdVentana = EnuIdVentanaDef.EnuLogOn Then
                    If MobjNotiEveArg.EnuSevNotifica = EnuSeveridadNot.EnuDatoInvalido Then
                        lblnProcesar = MobjSenderProp.BlnEsValido
                    ElseIf MobjNotiEveArg.EnuSevNotifica = EnuSeveridadNot.EnuFalta Then
                        lblnProcesar = MobjSenderProp.FblnNotiInfoOk(MobjE.EnuIdMensNot)
                    End If
                Else
                    lblnProcesar = True
                End If
            ElseIf MobjSenderObjPan IsNot Nothing AndAlso MobjNotiEveArg.EnuIdMensNot > 0 Then
                lblnProcesar = MobjSenderObjPan.FblnNotificaOk(MobjNotiEveArg.EnuIdMensNot)
            End If
        Else
            If MobjNotiEveArg IsNot Nothing Then
                lblnProcesar = MobjE.EnuSevNotifica >= MobjNotiEveArg.EnuSevNotifica
                If Not lblnProcesar Then
                    lblnProcesar = (MobjE.EnuSevNotifica = EnuSeveridadNot.EnuOk OrElse
                            MobjE.EnuSevNotifica = EnuSeveridadNot.EnuCamInsatis) AndAlso
                            MobjNotiEveArg.EnuSevNotifica <= EnuSeveridadNot.EnuInformacion
                End If
            Else
                MobjNotiEveArg = MobjE
            End If
        End If
        If lblnProcesar Then
            MobjSenderObjPan = Nothing
            MobjSenderProp = Nothing
            MobjNotiEveArg = MobjE
        End If
    End Sub
    Private Sub SProceseNoti()
        Select Case MobjNotiEveArg.EnuSevNotifica
            Case EnuSeveridadNot.EnuOk
                SLimpieNotificaciones()
                MobjNotiEveArg = MobjE
                MobjSenderProp = Nothing
                MobjSenderObjPan = Nothing
            Case EnuSeveridadNot.EnuInformacion, EnuSeveridadNot.EnuFalta
                SMuestreInformacion()
            Case EnuSeveridadNot.EnuAdvertencia
                SMuestreAdveretencia()
            Case EnuSeveridadNot.EnuCamInsatis
                SMuestreCamInsatisfechos()
                MobjSenderProp = Nothing
                MobjSenderObjPan = Nothing
            Case EnuSeveridadNot.EnuDatoInvalido
                SMuestreDatoInvalido()
            Case EnuSeveridadNot.EnuExcep
                SMuestreEx()
            Case EnuSeveridadNot.EnuError
                SMuestreErr()
        End Select
        If MobjNotiEveArg.EnuSevNotifica <> EnuSeveridadNot.EnuOk Then
            SNotifiqueSon()
        End If
    End Sub
    Private Sub SNotifiqueSon()
        Static lstrMens As String = ""
        If lstrMens <> MobjNotiEveArg.StrMensaje Then
            lstrMens = MobjNotiEveArg.StrMensaje
            If GblnNotiSonoras Then
                Dim j = 0
                Select Case MobjNotiEveArg.EnuSevNotifica
                    Case EnuSeveridadNot.EnuOk
                '
                    Case EnuSeveridadNot.EnuInformacion, EnuSeveridadNot.EnuFalta
                        j = 1
                    Case EnuSeveridadNot.EnuAdvertencia
                        j = 2
                    Case EnuSeveridadNot.EnuCamInsatis
                '
                    Case EnuSeveridadNot.EnuDatoInvalido
                        j = 1
                    Case EnuSeveridadNot.EnuExcep
                        j = 3
                    Case EnuSeveridadNot.EnuError
                        j = 3
                End Select
                If j > 0 Then
                    Dim i = 0
                    Dim ldblSegIni As Double = DateAndTime.Timer
                    Do Until i = j
                        i += 1
                        Console.Beep()
                        While DateAndTime.Timer - ldblSegIni < 0.5
                        End While
                    Loop
                End If
            End If
        End If
    End Sub
    Private Sub SLimpieNotificaciones()
        If Not (HlblMensajes Is Nothing OrElse HblnSeEstaCerrando) Then
            Dim lstrMens = StrNombreVentana
            Dim NoUsado = Dispatcher.Invoke(MdgtLblActualizaNot, DispatcherPriority.ContextIdle,
                    New Object() {ContentProperty, lstrMens})
            SColorieLblMensaje_New()
        End If
    End Sub
    Private Sub SMuestreInformacion()
        Dim lstrMens As String
        If MobjNotiEveArg.EnuSevNotifica = EnuSeveridadNot.EnuInformacion Then
            lstrMens = My.Resources.Informacion
        Else
            lstrMens = My.Resources.Falta
        End If
        lstrMens &= MobjNotiEveArg.StrMensaje
        Dim NoUsado = Dispatcher.Invoke(MdgtLblActualizaNot, DispatcherPriority.ContextIdle,
                    New Object() {ContentProperty, lstrMens})
        SColorieLblMensaje_New()
    End Sub
    Private Sub SMuestreAdveretencia()
        Dim lstrMens = String.Empty
        If Not String.IsNullOrEmpty(MobjNotiEveArg.StrMensaje) Then
            lstrMens = My.Resources.Advertencia & MobjNotiEveArg.StrMensaje
        End If
        Dim NoUsado = Dispatcher.Invoke(MdgtLblActualizaNot, DispatcherPriority.ContextIdle,
                    New Object() {ContentProperty, lstrMens})
        SColorieLblMensaje_New()
    End Sub
    Private Sub SMuestreCamInsatisfechos()
        If Not (HlblMensajes Is Nothing OrElse HblnSeEstaCerrando) Then
            Dim lstrMens = My.Resources.CamposSinSatisfacer
            Dim NoUsado = Dispatcher.Invoke(MdgtLblActualizaNot, DispatcherPriority.ContextIdle,
                    New Object() {ContentProperty, lstrMens})
            SColorieLblMensaje_New()
        End If
    End Sub
    Private Sub SMuestreDatoInvalido()
        Dim lstrMens = My.Resources.DatInv & MobjNotiEveArg.StrMensaje
        Dim NoUsado = Dispatcher.Invoke(MdgtLblActualizaNot, DispatcherPriority.ContextIdle,
                    New Object() {ContentProperty, lstrMens})
        SColorieLblMensaje_New()
    End Sub
    Private Sub SMuestreEx()
        Dim lstrMens = My.Resources.Excep & MobjNotiEveArg.StrMensaje
        lstrMens &= " Por favor reinicie la aplicación. Si se repite el error por favor informe a soporte!"
        If HlblMensajes IsNot Nothing Then
            Dim NoUsado = Dispatcher.Invoke(MdgtLblActualizaNot, DispatcherPriority.ContextIdle,
                    New Object() {ContentProperty, lstrMens})
            SColorieLblMensaje_New()
        End If
        ClsPanorama.SEscribaArchivoError(MobjNotiEveArg.StrMensajeEx)
        Dim Unused = MsgBox("Se ha presentado un error!. Por favor reinicie el programa." & vbCrLf &
                "Si el problema persiste, informe a soporte!",
                MsgBoxStyle.OkOnly, "Excepción")
        EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando
        SCancele()
    End Sub
    Private Sub SMuestreErr()
        Dim lstrMens = My.Resources.ErrorStr & " " & MobjNotiEveArg.StrMensaje
        lstrMens &= " Por favor reinicie la aplicación y si se repite el error informe a soporte!"
        If HlblMensajes IsNot Nothing Then
            Dim NoUsado = Dispatcher.Invoke(MdgtLblActualizaNot, DispatcherPriority.ContextIdle,
                    New Object() {ContentProperty, lstrMens})
            SColorieLblMensaje_New()
        Else
            MsgBox(lstrMens, vbOKOnly, "Error")
        End If
    End Sub
    Private Sub SColorieLblMensaje_New()
        If Not IsNothing(HlblMensajes) Then
            Dim lenuSevNotifica = MobjNotiEveArg.EnuSevNotifica
            HlblMensajes.FontWeight = FontWeights.Bold
            If lenuSevNotifica = EnuSeveridadNot.EnuOk Then
                HlblMensajes.Background = Brushes.Transparent
            Else
                HlblMensajes.Background = Brushes.White
            End If
            Select Case lenuSevNotifica
                Case EnuSeveridadNot.EnuOk
                    If MblnEsVentanaPrincipal Then
                        HlblMensajes.Foreground = Brushes.Gray
                    Else
                        HlblMensajes.Foreground = Brushes.White
                    End If
                Case EnuSeveridadNot.EnuDatoInvalido
                    HlblMensajes.Foreground = Brushes.Blue
                Case EnuSeveridadNot.EnuAdvertencia
                    If String.IsNullOrEmpty(MobjNotiEveArg.StrMensaje) Then
                        HlblMensajes.Background = Brushes.Transparent
                    Else
                        HlblMensajes.Foreground = Brushes.CornflowerBlue
                    End If
                Case EnuSeveridadNot.EnuError, EnuSeveridadNot.EnuExcep
                    HlblMensajes.Foreground = Brushes.DarkRed
                    HlblMensajes.Background = Brushes.LightGray
                Case EnuSeveridadNot.EnuCamInsatis
                    HlblMensajes.Foreground = Brushes.Red
                Case EnuSeveridadNot.EnuFalta
                    HlblMensajes.Foreground = Brushes.DarkRed
                Case Else
                    HlblMensajes.Foreground = Brushes.Blue
            End Select
        End If
    End Sub
#End Region

#Region "Eventos en la Ventana"
    Protected Friend Sub EwinLoaded(sender As Object, e As EventArgs) Handles Me.Loaded
        SLoad()
    End Sub

    Protected Overridable Sub EwinClosed(sender As Object, e As EventArgs) Handles Me.Closed
        If EnuIdVentana = EnuIdVentanaDef.enuRegistroClave Then
            If GblnOK Then
                If Not FblnEstanTodosBien(False) Then
                    GblnOK = False
                End If
            End If
        End If
        If WinPadre IsNot Nothing Then
            If WinPadre.Visibility <> Visibility.Visible Then
                WinPadre.Visibility = Visibility.Visible
                If EnuIdVentana <> EnuIdVentanaDef.EnuLogOn Then
                    WinPadre.SRefrescarClic()
                End If
            End If
        End If
    End Sub
#End Region
End Class