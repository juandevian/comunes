Imports System.Windows
Imports System.Windows.Controls
Module MdefOriWin
#Region "Constantes globales"
    Friend Const GCDBLNODOGRANDE As Double = 35.0
    Friend Const GCDBLNODOMEDIANO As Double = 30.0
    Friend Const GCDBLNODOPEQUENO As Double = 25.0
    Friend Const GCDBLFUENTENODOGRANDE As Double = 14.0
    Friend Const GCDBLFUENTENODOMEDIANO As Double = 12.0
    Friend Const GCDBLFUENTENODOPEQUENO As Double = 10.0
#End Region
#Region "Procedimientos"
    ''' <summary>
    ''' Procedimiento que llena el ComboBox "acboComboBox" con los datos contenidos en la columna
    ''' cuyo nombre es "Dato" del array de DataRows "adrwDataRow"
    ''' el cual debe estar basado en la tabla "PanTblConstantes"
    ''' </summary>
    ''' <param name="adrwDataRow">Array de DataRows que contiene los datos para poblar el combobox</param>
    ''' <param name="acboComboBox">ComboBox que sera poblado con los datos del DatRow()</param>
    ''' <remarks></remarks>
    Friend Sub SPuebleComboBox(adrwDataRow As DataRow(), acboComboBox As ComboBox)
        acboComboBox.Items.Clear()
        If Not IsNothing(adrwDataRow) AndAlso adrwDataRow.Length > 0 Then
            For i As Integer = 0 To adrwDataRow.Length - 1
                acboComboBox.Items.Add(adrwDataRow(i)("Dato"))
            Next
        End If
        acboComboBox.SelectedIndex = 0
    End Sub
    ''' <summary>
    ''' Ordena el DataGrid por columna pasada en el argumento "adgrColumna" y en la forma pasada en el
    ''' argumento "alsdSortDireccion"
    ''' </summary>
    ''' <param name="adgrGrid">DataGrid a ordenar</param>
    ''' <param name="adgcColumna">Columna por la cual se ordena el DataGrid "adgrGrid"</param>
    ''' <param name="alsdSortDireccion">Dirección de ordenamiento.</param>
    ''' <remarks></remarks>
    Friend Sub SOrdeneDataGrid(adgrGrid As DataGrid, adgcColumna As DataGridColumn,
                astrNombreColumna As String, alsdSortDireccion As ListSortDirection)
        adgrGrid.Items.SortDescriptions.Clear()
        adgrGrid.Items.SortDescriptions.Add(New SortDescription(astrNombreColumna, alsdSortDireccion))
        SApliqueSortDirection(adgrGrid, adgcColumna, alsdSortDireccion)
        adgrGrid.Items.Refresh()
    End Sub
    Private Sub SApliqueSortDirection(adgrGrid As DataGrid, adgcColumna As DataGridColumn,
            alsdListaSortDireccion As ListSortDirection)
        For Each ldgcColumna As DataGridColumn In adgrGrid.Columns
            ldgcColumna.SortDirection = Nothing
        Next
        adgcColumna.SortDirection = alsdListaSortDireccion
    End Sub
    Friend Sub SPuebleBarraEstadoAdmin(acolLabels As Collection)
        Dim llblCuatro As Label = acolLabels(4)
        acolLabels(1).Content = "Usuario actual: " & GstrIdUsuario
        acolLabels(1).ToolTip = acolLabels(1).Content
        acolLabels(2).Content = "Estación: " & GstrOrigenActual
        acolLabels(2).ToolTip = acolLabels(2).Content
        acolLabels(3).Content = "Fecha y hora: " & Now.ToString()
        acolLabels(3).ToolTip = acolLabels(3).Content
        llblCuatro.Visibility = Visibility.Collapsed
    End Sub
#End Region
#Region "Funciones"
    Public Function FstrTrayecCalculadoraExe() As String
        Dim dir As String
        dir = Environment.SystemDirectory
        If Not String.IsNullOrEmpty(dir & "\CALC.EXE") Then
            Return dir & "\CALC.EXE"
        Else
            Return ""
        End If
    End Function
    Public Function FstrTrayecBlockNotasExe() As String
        Dim dir As String
        dir = Environment.SystemDirectory
        If Not String.IsNullOrEmpty(dir & "\NOTEPAD.EXE") Then
            Return dir & "\NOTEPAD.EXE"
        Else
            Return ""
        End If
    End Function
    Friend Function FmnuiMenuItem(astrNombre As String, astrheader As String,
            astrEstilo As String) As Controls.MenuItem
        Dim lmnuiMenuItem As New Controls.MenuItem With {
            .Name = astrNombre,
            .Header = astrheader,
            .HorizontalContentAlignment = HorizontalAlignment.Center,
            .VerticalContentAlignment = VerticalAlignment.Center,
        .Style = Application.Current.FindResource(astrEstilo)
        }
        Return lmnuiMenuItem
    End Function
    Friend Function FmnuiMenuItem(astrNombre As String, astrheader As String,
             astrEstilo As String, astrShortcut As String) As Controls.MenuItem
        Dim lmnuiMenuItem As New Controls.MenuItem With {
            .Name = astrNombre,
            .Header = astrheader,
            .InputGestureText = astrShortcut,
            .HorizontalContentAlignment = HorizontalAlignment.Center,
            .VerticalContentAlignment = VerticalAlignment.Center,
            .Style = Application.Current.FindResource(astrEstilo)
        }
        Return lmnuiMenuItem
    End Function
    Friend Function FmnuiMenuItemPan(astrNombre As String, astrheader As String,
            aentIdAccion As Integer, astrShortcut As String) As MenuItemPan
        Dim lmnuiMenuItemPan As New MenuItemPan
        With lmnuiMenuItemPan
            .Name = astrNombre
            .Header = astrheader
            .Foreground = Brushes.Black
            .HorizontalContentAlignment = HorizontalAlignment.Center
            .VerticalContentAlignment = VerticalAlignment.Center
            .EntIdAccion = aentIdAccion
            .InputGestureText = astrShortcut
            .IsTabStop = False
        End With
        Return lmnuiMenuItemPan
    End Function
    Friend Function FmnuiMenuItemPan(astrNombre As String, astrheader As String,
            aentIdAccion As Integer, astrShortcut As String, ablnPrimario As Boolean) As MenuItemPan
        Dim lmnuiMenuItemPan As New MenuItemPan
        With lmnuiMenuItemPan
            If ablnPrimario Then
                .Style = Application.Current.FindResource("RecMnuItemPri")
            Else
                .Style = Application.Current.FindResource("RecMnuItemSec")
            End If
            .Name = astrNombre
            .Header = astrheader
            .Foreground = Brushes.White
            .HorizontalContentAlignment = HorizontalAlignment.Center
            .VerticalContentAlignment = VerticalAlignment.Center
            .EntIdAccion = aentIdAccion
            .InputGestureText = astrShortcut
            .IsTabStop = False
            .BlnEsPrimario = ablnPrimario
        End With
        Return lmnuiMenuItemPan
    End Function
    Friend Function FmnuiMenuItemPan(astrNombre As String, astrheader As String,
            aentIdAccion As Integer, astrShortcut As String, astrNombreImg As String) As MenuItemPan
        Dim lmnuiMenuItemPan As New MenuItemPan
        With lmnuiMenuItemPan
            .Name = astrNombre
            .Header = astrheader
            .Foreground = Brushes.Black
            .HorizontalContentAlignment = HorizontalAlignment.Center
            .VerticalContentAlignment = VerticalAlignment.Center
            .EntIdAccion = aentIdAccion
            .InputGestureText = astrShortcut
            .IsTabStop = False
            If Not String.IsNullOrEmpty(astrNombreImg) Then
                .Icon = New System.Windows.Controls.Image() With {
                    .Source = New BitmapImage(New Uri("RecImagenes/" & astrNombreImg, UriKind.Relative))
                }
            End If
        End With
        Return lmnuiMenuItemPan
    End Function
    Friend Function FtviTviPan(astrEtiqueta As String, astrTrayectoriaImagen As String) As TreeViewItem
        Dim ldblTamañoLetra As Double
        ' Objeto que será devuelto
        Dim ltviPan As New TreeViewItem With {
            .Margin = New Thickness(1, 1, 0, 1)
        }
        ' Crea un Stack Panel
        Dim lstpTvi As New StackPanel With {
            .Orientation = Orientation.Horizontal,
            .Height = 25
        }
        ldblTamañoLetra = GCDBLFUENTENODOMEDIANO
        ' Crea Imagen
        Dim limgTvi As Image = Nothing
        If Not String.IsNullOrEmpty(astrTrayectoriaImagen) Then
            Dim luriImgTvi As New Uri("pack://application:,,,/" & astrTrayectoriaImagen)
            limgTvi = New Image With {
                .Source = New BitmapImage(luriImgTvi)
            }
        End If
        ' Etiqueta
        Dim lblTvi As New Controls.Label With {
            .Content = astrEtiqueta,
            .FontSize = ldblTamañoLetra
        }
        ' Poner elementos en stack
        If limgTvi IsNot Nothing Then
            lstpTvi.Children.Add(limgTvi)
        End If
        lstpTvi.Children.Add(lblTvi)
        ' Asignar Stack a Header
        ltviPan.Header = lstpTvi
        ltviPan.IsTabStop = True
        Return ltviPan
    End Function
    Friend Function FtviTviPan(astrEtiqueta As String, astrTrayectoriaImagen As String,
                               adblAltura As Double) As TreeViewItem
        ' Objeto que será devuelto
        Dim ltviPan As New TreeViewItem
        Try
            ltviPan.Padding = New Thickness(0, 1, 0, 1)
            ltviPan.Height = adblAltura
            ltviPan.Padding = New Thickness(0)
            ltviPan.VerticalContentAlignment = VerticalAlignment.Center
            ' Crea un Stack Panel
            Dim lstpTvi As New StackPanel With {
                .Orientation = Controls.Orientation.Horizontal,
                .VerticalAlignment = VerticalAlignment.Center,
                .Height = adblAltura,
                .Margin = New Thickness(0)
            }
            ' Crea Imagen
            Dim limgTvi As New Controls.Image
            Dim luriImgTvi As New Uri("pack://application:,,,/" & astrTrayectoriaImagen)
            limgTvi.Source = New BitmapImage(luriImgTvi)
            limgTvi.Margin = New Thickness(1)
            ' Etiqueta
            Dim lblTvi As New Controls.Label With {
                .Content = astrEtiqueta,
                .VerticalContentAlignment = VerticalAlignment.Center,
                .FontSize = GCDBLFUENTENODOMEDIANO,
                .Padding = New Thickness(2, 0, 2, 0)
            }
            ' Poner elementos en stack
            lstpTvi.Children.Add(limgTvi)
            lstpTvi.Children.Add(lblTvi)
            ' Asignar Stack a Header
            ltviPan.Header = lstpTvi
        Catch ex As Exception
            Throw
        End Try
        Return ltviPan
    End Function
    Friend Function FtviTviPanCheck(astrEtiqueta As String, astrTag As String) As TreeViewItem
        Dim ldblTamañoLetra As Double
        ' Objeto que será devuelto
        Dim ltviPan As New TreeViewItem With {
            .Tag = astrTag
        }
        ' Crea un Stack Panel
        Dim lstpTvi As New StackPanel With {
            .Orientation = Orientation.Horizontal,
            .Height = 20
        }
        ' CheckBox
        Dim lchkTvi As New CheckBox With {
                .IsChecked = False,
                .Padding = New Thickness(0),
                .VerticalAlignment = VerticalAlignment.Center,
                .Tag = astrTag,
                .IsEnabled = False
            }
        ldblTamañoLetra = GCDBLFUENTENODOGRANDE
        ' Etiqueta
        Dim lblTvi As New Label With {
            .Content = Space(3) & astrEtiqueta,
            .FontSize = ldblTamañoLetra,
            .Padding = New Thickness(0),
            .VerticalAlignment = VerticalAlignment.Center
        }
        ' Poner elementos en stack
        lstpTvi.Children.Add(lchkTvi)
        lstpTvi.Children.Add(lblTvi)
        ' Asignar Stack a Header
        ltviPan.Header = lstpTvi
        Return ltviPan
    End Function
    Friend Sub SNombreTvi(atviNodo As TreeViewItem, astrEtiqueta As String)
        Dim ldblTamañoLetra = GCDBLFUENTENODOMEDIANO
        ' Etiqueta
        Dim lblTvi As New Controls.Label With {
            .Padding = New Thickness(0, 1, 0, 1),
            .Content = astrEtiqueta,
            .FontSize = ldblTamañoLetra
        }
        atviNodo.Header = lblTvi
    End Sub
    Friend Sub SMarqueTodos(atviNodo As TreeViewItem, ablnChequiado As Boolean)
        For Each ltviHijo As TreeViewItem In atviNodo.Items
            If ltviHijo.Items.Count > 0 Then
                SMarqueTodos(ltviHijo, ablnChequiado)
            End If
            SMarqueNodo(ltviHijo, ablnChequiado)
        Next
    End Sub
    Friend Sub SHabiliteNodo(atviNodo As TreeViewItem, ablnHabiltar As Boolean)
        If atviNodo.IsEnabled <> ablnHabiltar Then
            atviNodo.IsEnabled = ablnHabiltar
        End If
        SCambieHabilChkTvi(atviNodo, ablnHabiltar)
        For Each ltviNodo As TreeViewItem In atviNodo.Items
            atviNodo.IsEnabled = ablnHabiltar
            SCambieHabilChkTvi(ltviNodo, ablnHabiltar)
            If ltviNodo.Items.Count > 0 Then
                SHabiliteNodo(ltviNodo, ablnHabiltar)
            End If
        Next
    End Sub
    Friend Sub SMarqueNodo(atviNodo As TreeViewItem, ablnChequeado As Boolean)
        For Each lobjHijo As Object In atviNodo.Header.Children
            If TypeOf lobjHijo Is CheckBox Then
                Dim lchkNodo As CheckBox = lobjHijo
                lchkNodo.IsChecked = ablnChequeado
            End If
        Next
    End Sub
    Friend Sub SCambieHabilChkTvi(atviNodo As TreeViewItem, ablnHabilite As Boolean)
        For Each lobjHijo As Object In atviNodo.Header.Children
            If TypeOf lobjHijo Is CheckBox Then
                Dim lchkNodo As CheckBox = lobjHijo
                lchkNodo.IsEnabled = ablnHabilite
            End If
        Next
    End Sub
    Friend Function FblnNodoChequiado(atviNodo As TreeViewItem) As Boolean
        Dim lblnNodoCheck As Boolean = False
        Dim lchkNodo As CheckBox
        For Each lobjHijo As Object In atviNodo.Header.Children
            If TypeOf lobjHijo Is CheckBox Then
                lchkNodo = lobjHijo
                lblnNodoCheck = lchkNodo.IsChecked
                Exit For
            End If
        Next
        Return lblnNodoCheck
    End Function
#End Region
End Module
