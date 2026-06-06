Public Class WinBusqueda
#Region "Definiciones"
    Implements IDisposable
#Region "Enumeradores"
    Private Enum EnuTBTipoBusqueda As Integer
        enuTBSencilla = 0
        enuTBCompuesta
    End Enum
#End Region
    ' Variables de modulo
    Private ReadOnly McolTablasPri As New Collection
    Private ReadOnly McolCamposTablaPri As New Collection
    Private ReadOnly McolTablasSec As New Collection
    Private ReadOnly McolCamposTablaSec As New Collection
    Private ReadOnly McolCamposPriRel As New Collection
    Private ReadOnly McolCamposSecRel As New Collection
    Private ReadOnly McolCamposBusqueda As New Collection
    Private ReadOnly McolRetornar2Col As New Collection
    Private ReadOnly McolFiltros As New Collection
    Private ReadOnly McolCamposRetornar As New Collection
    Private ReadOnly McolTipo As New Collection
    Private MstrCampoIndice As String = String.Empty
    Private MdtbBusqueda As Data.DataTable = Nothing
    Private MdvwBusqueda As DataView = Nothing
    Private MbttAceptar As Controls.Button = Nothing
    Private MbttCancelar As Controls.Button = Nothing
    Private MblnRetornar2Col As Boolean = False
    ' Propiedades autoimplementadas
    Friend Property WinPadre As ClsFormInterface = Nothing
#End Region
#Region "Procedimientos y funciones"
    Private Sub SAsigneBotones()
        Dim lcnvPieWin As Canvas = FindName("cnvBotones")
        Dim lcnvBotones As Canvas = Nothing
        If Not IsNothing(lcnvPieWin) Then
            For Each lobjObjeto As Object In lcnvPieWin.Children
                If TypeOf lobjObjeto Is Canvas Then
                    If lobjObjeto.Name = "cnvAceptaCancela" Then
                        lcnvBotones = lobjObjeto
                        Exit For
                    End If
                End If
            Next
            For Each lobjObjeto As Object In lcnvBotones.Children
                If TypeOf lobjObjeto Is Controls.Button Then
                    Dim lbttBoton As Controls.Button = lobjObjeto
                    If lbttBoton.Name = "bttAceptar" Then
                        MbttAceptar = lbttBoton
                        MbttAceptar.TabIndex = 3
                    ElseIf lbttBoton.Name = "bttCancelar" Then
                        MbttCancelar = lbttBoton
                        MbttCancelar.Content = "_Cancelar"
                        MbttCancelar.TabIndex = 4
                    End If
                End If
            Next
        End If
    End Sub
    ''' <summary>
    ''' Recibe los elementos necesarios para que la forma "frmBusqueda" funcione adecuadamente cuando la busqueda
    ''' contiene campos de dos tablas como por ejemplo el "IdCliente" de la tabla "TraLicencias" y el nombre del
    ''' Cliente de la tabla "TraClientes"
    ''' </summary>
    ''' <param name="astrNombreBusqueda">Es el nombre que identifica la busqueda.</param>
    ''' <param name="astrNombreTablaPri">Nombre de la tabla primaria.</param>
    ''' <param name="astrNombreTablaSec">Nombre de la tabla secundaria.</param>
    ''' <param name="astrCamposTablaPri">Array con los nombres de los campos de la tabla primaria que van a ser
    ''' mostrados.</param>
    ''' <param name="astrCamposTablaSec">Array con los nombres de los campos de la tabla secundaria que van a ser
    ''' mostrados.</param>
    ''' <param name="astrCamposPriRel">Array con los nombres de los campos de la tabla primaria que se
    ''' relacionan con los campos de la tabala secundaria.</param>
    ''' <param name="astrCamposSecRel">Array con los nombres de los campos de la tabla secundaria que se
    ''' relacionan con los campos de la tabala primaria.</param>
    ''' <param name="astrCampoBusqueda">Nombre del campo que contiene los datos a buscar. Este campo se debe 
    ''' calificar con "P." si pertenece a la tabla primaria o "S." si pertenece a la tabla secundaria.</param>
    ''' <param name="astrCampoRetornar">Nombre del campo que contiene el valor que retorna la busqueda 
    ''' y que debe identificar el objeto buscado</param>
    ''' <param name="astrfiltro">Una expresion de cadena que filtra los registros mostrados</param>
    ''' <remarks>si el campo de busqueda pertenece a la tabla primaria se el debe anteponer "P." y si 
    ''' pertenece a la tabla secundaria se le debe anteponer "S."</remarks>
    Public Sub SDefinaBusqueda(astrNombreBusqueda As String, astrNombreTablaPri As String,
            astrNombreTablaSec As String, astrCamposTablaPri() As String,
            astrCamposTablaSec() As String, astrCamposPriRel() As String,
            astrCamposSecRel() As String, astrCampoBusqueda As String,
            astrCampoRetornar As String, astrfiltro As String,
            ablnCampoRetornarTablaPri As Boolean)
        cboNombreBusqueda.Items.Add(astrNombreBusqueda)
        Dim lblnTablaPriContCampoRetornar = astrCamposTablaPri.Contains(astrCampoRetornar)
        If Not lblnTablaPriContCampoRetornar Then
            If astrCamposTablaPri(0).StartsWith("DISTINCT") Then
                lblnTablaPriContCampoRetornar = astrCamposTablaPri(0).Substring(8).Trim =
                    astrCampoRetornar
            End If
        End If
        If Not (lblnTablaPriContCampoRetornar OrElse
                astrCamposTablaSec.Contains(astrCampoRetornar)) Then
            ReDim Preserve astrCamposTablaPri(astrCamposTablaPri.Count)
            astrCamposTablaPri(astrCamposTablaPri.Count - 1) = astrCampoRetornar
        End If
        If ablnCampoRetornarTablaPri Then
            ReDim Preserve astrCamposTablaPri(astrCamposTablaPri.Count)
            astrCamposTablaPri(astrCamposTablaPri.Count - 1) = astrCampoRetornar & " AS CampoRetornar"
        Else
            ReDim Preserve astrCamposTablaSec(astrCamposTablaSec.Count)
            astrCamposTablaSec(astrCamposTablaSec.Count - 1) = astrCampoRetornar & " AS CampoRetornar"
        End If
        McolTablasPri.Add(astrNombreTablaPri, astrNombreBusqueda)
        McolCamposTablaPri.Add(astrCamposTablaPri, astrNombreBusqueda)
        McolTablasSec.Add(astrNombreTablaSec, astrNombreBusqueda)
        McolCamposPriRel.Add(astrCamposPriRel, astrNombreBusqueda)
        McolCamposSecRel.Add(astrCamposSecRel, astrNombreBusqueda)
        McolCamposTablaSec.Add(astrCamposTablaSec, astrNombreBusqueda)
        McolCamposBusqueda.Add(astrCampoBusqueda, astrNombreBusqueda)
        McolCamposRetornar.Add(astrCampoRetornar, astrNombreBusqueda)
        McolRetornar2Col.Add(False, astrNombreBusqueda)
        McolFiltros.Add(astrfiltro, astrNombreBusqueda)
        McolTipo.Add(EnuTBTipoBusqueda.enuTBCompuesta, astrNombreBusqueda)
    End Sub
    ''' <summary>
    ''' Recibe los elementos necesarios para que la forma "frmBusqueda" funcione adecuadamente cuando la busqueda
    ''' contiene campos de dos tablas como por ejemplo el "IdCliente" de la tabla "TraLicencias" y el nombre del
    ''' Cliente de la tabla "TraClientes" y y debe devolver los valores de dos campos que identifican 
    ''' el objeto buscado. 
    ''' </summary>
    ''' <param name="astrNombreBusqueda">Es el nombre que identifica la busqueda.</param>
    ''' <param name="astrNombreTablaPri">Nombre de la tabla primaria.</param>
    ''' <param name="astrNombreTablaSec">Nombre de la tabla secundaria.</param>
    ''' <param name="astrCamposTablaPri">Array con los nombres de los campos de la tabla primaria que van a ser
    ''' mostrados.</param>
    ''' <param name="astrCamposTablaSec">Array con los nombres de los campos de la tabla secundaria que van a ser
    ''' mostrados.</param>
    ''' <param name="astrCamposPriRel">Array con los nombres de los campos de la tabla primaria que se
    ''' relacionan con los campos de la tabala secundaria.</param>
    ''' <param name="astrCamposSecRel">Array con los nombres de los campos de la tabla secundaria que se
    ''' relacionan con los campos de la tabala primaria.</param>
    ''' <param name="astrCampoBusqueda">Nombre del campo que contiene los datos a buscar. Este campo se debe 
    ''' calificar con "P." si pertenece a la tabla primaria o "S." si pertenece a la tabla secundaria.</param>
    ''' <param name="astrCamposRetornar">Nombre de los campo que contienen los valores que retorna la busqueda 
    ''' y que debe identificar el objeto buscado</param>
    ''' <param name="astrfiltro">Una expresion de cadena que filtra los registros mostrados</param>
    ''' <remarks>si el campo de busqueda pertenece a la tabla primaria se el debe anteponer "P." y si 
    ''' pertenece a la tabla secundaria se le debe anteponer "S."</remarks>
    Public Sub SDefinaBusqueda(astrNombreBusqueda As String, astrNombreTablaPri As String,
            astrNombreTablaSec As String, astrCamposTablaPri() As String,
            astrCamposTablaSec() As String, astrCamposPriRel() As String,
            astrCamposSecRel() As String, astrCampoBusqueda As String,
            astrCamposRetornar() As String, astrfiltro As String,
            ablnCamposRetornarTablaPri As Boolean)
        If astrCamposRetornar Is Nothing Then
            Throw New ArgumentNullException(NameOf(astrCamposRetornar))
        End If
        cboNombreBusqueda.Items.Add(astrNombreBusqueda)
        If Not (astrCamposTablaPri.Contains(astrCamposRetornar(0)) OrElse
                astrCamposTablaSec.Contains(astrCamposRetornar(0))) Then
            ReDim Preserve astrCamposTablaPri(astrCamposTablaPri.Count)
            astrCamposTablaPri(astrCamposTablaPri.Count - 1) = astrCamposRetornar(0)
        End If
        If Not (astrCamposTablaPri.Contains(astrCamposRetornar(1)) OrElse
                astrCamposTablaSec.Contains(astrCamposRetornar(1))) Then
            ReDim Preserve astrCamposTablaPri(astrCamposTablaPri.Count)
            astrCamposTablaPri(astrCamposTablaPri.Count - 1) = astrCamposRetornar(1)
        End If
        If ablnCamposRetornarTablaPri Then
            ReDim Preserve astrCamposTablaPri(astrCamposTablaPri.Count + 1)
            astrCamposTablaPri(astrCamposTablaPri.Count - 2) = astrCamposRetornar(0) & " AS CampoRetornar"
            astrCamposTablaPri(astrCamposTablaPri.Count - 1) = astrCamposRetornar(1) & " AS CampoRetornar1"
        Else
            ReDim Preserve astrCamposTablaSec(astrCamposTablaSec.Count + 1)
            astrCamposTablaSec(astrCamposTablaSec.Count - 2) = astrCamposRetornar(0) & " AS CampoRetornar"
            astrCamposTablaSec(astrCamposTablaSec.Count - 1) = astrCamposRetornar(1) & " AS CampoRetornar1"
        End If
        McolTablasPri.Add(astrNombreTablaPri, astrNombreBusqueda)
        McolCamposTablaPri.Add(astrCamposTablaPri, astrNombreBusqueda)
        McolTablasSec.Add(astrNombreTablaSec, astrNombreBusqueda)
        McolCamposPriRel.Add(astrCamposPriRel, astrNombreBusqueda)
        McolCamposSecRel.Add(astrCamposSecRel, astrNombreBusqueda)
        McolCamposTablaSec.Add(astrCamposTablaSec, astrNombreBusqueda)
        McolCamposBusqueda.Add(astrCampoBusqueda, astrNombreBusqueda)
        McolCamposRetornar.Add(astrCamposRetornar, astrNombreBusqueda)
        McolRetornar2Col.Add(True, astrNombreBusqueda)
        McolFiltros.Add(astrfiltro, astrNombreBusqueda)
        McolTipo.Add(EnuTBTipoBusqueda.enuTBCompuesta, astrNombreBusqueda)
    End Sub
    ''' <summary>
    ''' Recibe los elementos necesarios para que la forma "frmBusqueda" funcione adecuadamente cuando la busqueda
    ''' contiene campos de una tabla.
    ''' </summary>
    ''' <param name="astrNombreBusqueda">Es el nomnre que identifica la busqueda.</param>
    ''' <param name="astrNombreTabla">Nombre de la tabla en la cual se busca</param>
    ''' <param name="astrCamposMostrar">Array con los nombres de la tabla primaria que van a ser mostrados</param>
    ''' <param name="astrCampoBusqueda">Nombre del campo que contiene los datos a buscar.</param>
    ''' <param name="astrCampoRetornar">Nombre del campo que contiene el valor que retorna la busqueda 
    ''' y que debe identificar el objeto buscado</param>
    ''' <param name="astrfiltro">Una expresion de cadena que filtra los registros mostrados</param>
    ''' <remarks></remarks>
    Public Sub SDefinaBusqueda(astrNombreBusqueda As String, astrNombreTabla As String,
            astrCamposMostrar() As String, astrCampoBusqueda As String,
            astrCampoRetornar As String, astrfiltro As String)
        cboNombreBusqueda.Items.Add(astrNombreBusqueda)
        If Not astrCamposMostrar.Contains(astrCampoRetornar) Then
            ReDim Preserve astrCamposMostrar(astrCamposMostrar.Count)
            astrCamposMostrar(astrCamposMostrar.Count - 1) = astrCampoRetornar
        End If
        ReDim Preserve astrCamposMostrar(astrCamposMostrar.Count)
        astrCamposMostrar(astrCamposMostrar.Count - 1) = astrCampoRetornar & " AS CampoRetornar"
        McolTablasPri.Add(astrNombreTabla, astrNombreBusqueda)
        McolCamposTablaPri.Add(astrCamposMostrar, astrNombreBusqueda)
        McolCamposBusqueda.Add(astrCampoBusqueda, astrNombreBusqueda)
        McolCamposRetornar.Add(astrCampoRetornar, astrNombreBusqueda)
        McolRetornar2Col.Add(False, astrNombreBusqueda)
        McolFiltros.Add(astrfiltro, astrNombreBusqueda)
        McolTipo.Add(EnuTBTipoBusqueda.enuTBSencilla, astrNombreBusqueda)
    End Sub
    ''' <summary>
    ''' Recibe los elementos necesarios para que la forma "frmBusqueda" funcione adecuadamente cuando la busqueda
    ''' contiene campos de una tabla y debe devolver los valores de dos campos que identifican el objeto buscado.
    ''' </summary>
    ''' <param name="astrNombreBusqueda">Es el nomnre que identifica la busqueda.</param>
    ''' <param name="astrNombreTabla">Nombre de la tabla en la cual se busca</param>
    ''' <param name="astrCamposMostrar">Array con los nombres de la tabla primaria que van a ser mostrados</param>
    ''' <param name="astrCampoBusqueda">Nombre del campo que contiene los datos a buscar.</param>
    ''' <param name="astrCamposRetornar">Nombre de los campos que contiene los valores que retorna la busqueda 
    ''' y que debe identificar el objeto buscado</param>
    ''' <param name="astrfiltro">Una expresion de cadena que filtra los registros mostrados</param>
    ''' <remarks></remarks>
    Public Sub SDefinaBusqueda(astrNombreBusqueda As String, astrNombreTabla As String,
            astrCamposMostrar() As String, astrCampoBusqueda As String,
            astrCamposRetornar As String(), astrFiltro As String)
        If astrCamposRetornar Is Nothing Then
            Throw New ArgumentNullException(NameOf(astrCamposRetornar))
        End If
        cboNombreBusqueda.Items.Add(astrNombreBusqueda)
        If Not astrCamposMostrar.Contains(astrCamposRetornar(0)) Then
            ReDim Preserve astrCamposMostrar(astrCamposMostrar.Count)
            astrCamposMostrar(astrCamposMostrar.Count - 1) = astrCamposRetornar(0)
        End If
        If Not astrCamposMostrar.Contains(astrCamposRetornar(1)) Then
            ReDim Preserve astrCamposMostrar(astrCamposMostrar.Count)
            astrCamposMostrar(astrCamposMostrar.Count - 1) = astrCamposRetornar(1)
        End If
        ReDim Preserve astrCamposMostrar(astrCamposMostrar.Count + 1)
        astrCamposMostrar(astrCamposMostrar.Count - 1) = astrCamposRetornar(0) & " AS CampoRetornar"
        astrCamposMostrar(astrCamposMostrar.Count) = astrCamposRetornar(1) & " AS CampoRetornar1"
        McolTablasPri.Add(astrNombreTabla, astrNombreBusqueda)
        McolCamposTablaPri.Add(astrCamposMostrar, astrNombreBusqueda)
        McolCamposBusqueda.Add(astrCampoBusqueda, astrNombreBusqueda)
        McolCamposRetornar.Add(astrCamposRetornar, astrNombreBusqueda)
        McolRetornar2Col.Add(True, astrNombreBusqueda)
        McolFiltros.Add(astrFiltro, astrNombreBusqueda)
        McolTipo.Add(EnuTBTipoBusqueda.enuTBSencilla, astrNombreBusqueda)
    End Sub
    Private Sub SRetorneBusqueda()
        Dim ldgvDataGrid As DataRowView
        ldgvDataGrid = dgrBusqueda.SelectedItem
        If Not IsNothing(ldgvDataGrid) Then
            WinPadre.BlnBusquedaOk = True
            If MblnRetornar2Col Then
                ReDim WinPadre.StrResutadosBusqueda(1)
                WinPadre.StrResutadosBusqueda(0) = ldgvDataGrid("CampoRetornar")
                WinPadre.StrResutadosBusqueda(1) = ldgvDataGrid("CampoRetornar1")
            Else
                WinPadre.StrResultadoBusqueda = ldgvDataGrid("CampoRetornar")
            End If
            Close()
        Else
            SMuestreMensaje("No se ha seleccionado Fila alguna!", EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Friend Sub SOculteBotonCancelar()
        Canvas.SetRight(MbttAceptar, Canvas.GetRight(MbttCancelar))
        MbttCancelar.Visibility = Visibility.Collapsed
    End Sub
    Private Sub SMuestreMensaje(astrMens As String, aenuSevNot As EnuSeveridadNot)
        Dim lstrMens = String.Empty
        If Not String.IsNullOrEmpty(astrMens) Then
            Select Case aenuSevNot
                Case EnuSeveridadMen.enuAdvertencia ' Advertencia
                    lstrMens = My.Resources.Advertencia
                Case EnuSeveridadMen.enuInformacion ' Informacion
                    lstrMens = My.Resources.Informacion
                Case EnuSeveridadMen.enuError ' Error
                    lstrMens = My.Resources.ErrorStr
            End Select
            lstrMens &= astrMens
            SColorieLblMensaje(aenuSevNot)
        Else
            lstrMens = String.Empty
            lblMensajesBus.Content = "BUSQUEDA"
            SColorieLblMensaje(EnuSeveridadNot.EnuOk)
        End If
        If lblMensajesBus IsNot Nothing Then
            lblMensajesBus.Content = lstrMens
        End If
    End Sub
    Private Sub SColorieLblMensaje(aenuSevNot As EnuSeveridadNot)
        If Not IsNothing(lblMensajesBus) Then
            lblMensajesBus.FontWeight = FontWeights.Bold
            If aenuSevNot = EnuSeveridadNot.EnuOk Then
                lblMensajesBus.Background = System.Windows.Media.Brushes.Transparent
            Else
                lblMensajesBus.Background = System.Windows.Media.Brushes.White
            End If
            Select Case aenuSevNot
                Case EnuSeveridadNot.EnuOk
                    lblMensajesBus.Foreground = System.Windows.Media.Brushes.White
                Case EnuSeveridadNot.EnuDatoInvalido
                    lblMensajesBus.Foreground = System.Windows.Media.Brushes.Blue
                Case EnuSeveridadNot.EnuAdvertencia
                    lblMensajesBus.Foreground = System.Windows.Media.Brushes.CornflowerBlue
                Case EnuSeveridadNot.EnuError, EnuSeveridadNot.EnuExcep
                    lblMensajesBus.Foreground = System.Windows.Media.Brushes.DarkRed
                    lblMensajesBus.Background = System.Windows.Media.Brushes.LightGray
                Case EnuSeveridadNot.EnuCamInsatis
                    lblMensajesBus.Foreground = System.Windows.Media.Brushes.Red
                Case EnuSeveridadNot.EnuFalta
                    lblMensajesBus.Foreground = System.Windows.Media.Brushes.DarkRed
                Case Else
                    lblMensajesBus.Foreground = System.Windows.Media.Brushes.Blue
            End Select
        End If
    End Sub
#End Region
#Region "Implementa IDisposable"
    Protected Overridable Overloads Sub Dispose(disposing As Boolean)
        If disposing Then
            MdvwBusqueda.Dispose()
            MdtbBusqueda.Dispose()
        End If
    End Sub
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region
#Region "Eventos de la ventana y los controles"
    Public Sub New()
        InitializeComponent()
        SAsigneBotones()
    End Sub
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        Top = 10
        Left = 700
        cboNombreBusqueda.SelectedIndex = 0
        If cboNombreBusqueda.Items.Count > 1 Then
            cboNombreBusqueda.Focus()
        Else
            txtBuscar.Focus()
        End If
        lblMensajesBus.Content = "BUSQUEDA"
    End Sub
    Private Sub Cbo_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles _
            cboNombreBusqueda.SelectionChanged
        Dim lstrMens = String.Empty, lblnNoHayError = False
        If cboNombreBusqueda.SelectedIndex >= 0 Then
            Dim lstrNombreBusqueda As String = cboNombreBusqueda.SelectedItem.ToString
            Dim lenuTipo As EnuTBTipoBusqueda = McolTipo(lstrNombreBusqueda)
            Dim lstrNombreTablaPri As String = McolTablasPri(lstrNombreBusqueda)
            Dim lstrCamposTablaPri As String() = McolCamposTablaPri(lstrNombreBusqueda)
            MblnRetornar2Col = McolRetornar2Col(lstrNombreBusqueda)
            Dim lstrNombreTablaSec As String
            Dim lstrCamposTablaSec As String()
            Dim lstrCamposRelPri As String()
            Dim lstrCamposRelSec As String()
            Dim lstrFiltro As String = McolFiltros(lstrNombreBusqueda)
            MstrCampoIndice = McolCamposBusqueda(lstrNombreBusqueda)
            Dim lstrIndice As String(,) = {{"", ""}}
            Dim lstrCamposGrupo As String() = Array.Empty(Of String)
            Try
                If lenuTipo = EnuTBTipoBusqueda.enuTBSencilla Then
                    MdtbBusqueda = ClsPanorama.FdtbDataTable(lstrNombreTablaPri,
                            lstrCamposTablaPri, lstrIndice, lstrFiltro, False, lstrCamposGrupo)
                Else
                    lstrNombreTablaSec = McolTablasSec(lstrNombreBusqueda)
                    lstrCamposTablaSec = McolCamposTablaSec(lstrNombreBusqueda)
                    lstrCamposRelPri = McolCamposPriRel(lstrNombreBusqueda)
                    lstrCamposRelSec = McolCamposSecRel(lstrNombreBusqueda)
                    MdtbBusqueda = ClsPanorama.FdtbDataTable(lstrNombreTablaPri, lstrCamposTablaPri,
                            lstrNombreTablaSec, lstrCamposTablaSec, lstrCamposRelPri, lstrCamposRelSec,
                            {{"", ""}}, lstrFiltro, Array.Empty(Of String), True)
                End If
                MdvwBusqueda = New DataView(MdtbBusqueda) With {
                    .Sort = MstrCampoIndice & " ASC"
                }
                grdBusqueda.DataContext = MdvwBusqueda
                lblnNoHayError = True
            Catch ex As ProveedorBdPanException
                lstrMens &= ex.Message
            Catch ex As PanDatException
                lstrMens &= ex.Message
            Catch ex As ArgumentNullException
                lstrMens &= ex.Message
            Catch ex As Exception
                lstrMens &= ex.Message
            Finally
                If Not lblnNoHayError Then
                    WinPadre.BlnBusquedaOk = False
                    WinPadre.StrResultadoBusqueda = String.Empty
                    Dim Nousado = MsgBox(lstrMens, vbOKOnly, "Error")
                    Close()
                Else
                    txtBuscar.Text = String.Empty
                    If cboNombreBusqueda.Items.Count = 1 Then
                        txtBuscar.Focus()
                    End If
                End If
            End Try
        End If
    End Sub
    Private Sub Txt_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtBuscar.TextChanged
        Dim lstrFiltro = MstrCampoIndice & " LIKE '*" & txtBuscar.Text & "*'"
        MdvwBusqueda.RowFilter = lstrFiltro
    End Sub
    Private Sub Dgr_AutoGeneratedColumns(sender As Object, e As EventArgs) Handles dgrBusqueda.AutoGeneratedColumns
        For Each ldgcGridCol As DataGridColumn In dgrBusqueda.Columns
            If ldgcGridCol.Header = "CampoRetornar" OrElse ldgcGridCol.Header = "CampoRetornar1" Then
                ldgcGridCol.Visibility = Visibility.Collapsed
            End If
        Next
    End Sub
    Private Sub Ctrl_KeyDown(sender As Object, e As Input.KeyEventArgs) Handles txtBuscar.KeyDown
        If e.Key = Key.Enter Then
            If TypeOf sender Is Controls.TextBox Then
                SRetorneBusqueda()
            End If
        End If
    End Sub
    Private Sub DgrBusqueda_KeyUp(sender As Object, e As Input.KeyEventArgs) Handles dgrBusqueda.KeyUp
        If e.Key = Key.Return Then
            SRetorneBusqueda()
        End If
    End Sub
    Private Sub Dgr_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles _
            dgrBusqueda.MouseRightButtonUp
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If lelmElemento.Name = "dgrBusqueda" Then
            SRetorneBusqueda()
        End If
    End Sub
    Private Sub Dgr_MouseDc(sender As Object, e As MouseButtonEventArgs) Handles _
            dgrBusqueda.MouseDoubleClick
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If lelmElemento.Name = "dgrBusqueda" Then
            SRetorneBusqueda()
        End If
    End Sub
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Controls.Button Then
            Select Case True
                Case lelmElemento.Equals(MbttAceptar)
                    SRetorneBusqueda()
                Case lelmElemento.Equals(MbttCancelar)
                    WinPadre.BlnBusquedaOk = False
                    WinPadre.StrResultadoBusqueda = String.Empty
                    Close()
            End Select
        End If
    End Sub
    Private Sub DgrBusqueda_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles dgrBusqueda.SelectionChanged
        If dgrBusqueda.SelectedItem IsNot Nothing Then
            SMuestreMensaje("", EnuSeveridadNot.EnuOk)
        End If
    End Sub
#End Region
End Class
