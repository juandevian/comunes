Imports System.Drawing
Friend Class ClsImagen
#Region "Definiciones"
    ' Herencia
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "PanImagenes"
#End Region

#Region "Constructores"
    Public Sub New(aobjPadre As Object, adrwImagen As DataRow)
        If adrwImagen Is Nothing Then
            Throw New ArgumentNullException(NameOf(adrwImagen))
        End If
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        '
        DrwRegistroActual = adrwImagen
        DtbTablaColeccion = adrwImagen.Table
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
            Return EnuIdClasesPanDef.enuImagen
        End Get
    End Property

    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Imagen"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjIdImagenDbl As New ClsIdImagenDbl(Me)
    Friend ReadOnly Property ObjIdCategoriaByt As New ClsIdCategoriaByt(Me)
    Friend ReadOnly Property ObjOrdinalImagenShr As New ClsOrdinalImagenShr(Me)
    Friend ReadOnly Property ObjPropiedadImagenImg As New ClsPropiedadImagenImg(Me)
    Friend ReadOnly Property ObjFechaDtm As New ClsFechaDtm(Me)
    Friend ReadOnly Property ObjDescripcionStr As New ClsDescripcionStr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjIdImagenDbl, ObjIdImagenDbl.StrNombre)
                HcolPropiedades.Add(ObjIdCategoriaByt, ObjIdCategoriaByt.StrNombre)
                HcolPropiedades.Add(ObjOrdinalImagenShr, ObjOrdinalImagenShr.StrNombre)
                HcolPropiedades.Add(ObjPropiedadImagenImg, ObjPropiedadImagenImg.StrNombre)
                HcolPropiedades.Add(ObjFechaDtm, ObjFechaDtm.StrNombre)
                HcolPropiedades.Add(ObjDescripcionStr, ObjDescripcionStr.StrNombre)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras Propiedades"
    Friend ReadOnly Property StrNombreEnuCategoriaImgen As String
        Get
            Return ClsAdministrador.FstrNombreDatoConstantePan(EnuGrupoConstantesPanDef.enuCategoriaImagen,
                    ObjIdCategoriaByt.ObjValorPro)
        End Get
    End Property
#End Region
#End Region

#Region "Procedimientos y funciones invalidantes"
    Protected Friend Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            GobjPanDat.SControleProcesoObj(True)
            Try
                SNumereObj()
                MyBase.SActualice(ablnExigeRequeridos)
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
        Else
            MyBase.SActualice(ablnExigeRequeridos)
        End If
    End Sub
    Friend Overrides Function FblnEsSuprimible() As Boolean
        Dim lblnEsSuprimible = FblnPermitidoSuprimir()
        If lblnEsSuprimible Then
            Dim lshrOrdinal As Short = FshrOrdinalActual()
            lblnEsSuprimible = (lshrOrdinal = ObjOrdinalImagenShr.ObjValorPro)
        End If
        Return lblnEsSuprimible
    End Function
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdImagenDbl.ToString
        End Get
    End Property
#End Region

#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Dim lshrOrdinal As Short = FshrOrdinalActual() + 1
            ObjOrdinalImagenShr.ObjValorPro = lshrOrdinal
        End If
    End Sub

    Private Function FshrOrdinalActual() As Short
        Dim lstrFiltro As String = "IdImagen = '" & ObjIdImagenDbl.ObjValorPro &
                "' And IdTblCategoria = " & ObjIdCategoriaByt.ObjValorPro
        Dim lshrOrdinal As Short = 0
        Dim lobjOrdinal As Object = GobjPanDat.FobjMaxValorCampo("PanImagenes", "Ordinal",
                lstrFiltro)
        If Not (IsNothing(lobjOrdinal) OrElse IsDBNull(lobjOrdinal)) Then
            If IsNumeric(lobjOrdinal) Then
                lshrOrdinal = CType(lobjOrdinal, Short)
            Else
                lshrOrdinal = 0
            End If
        End If
        Return lshrOrdinal
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsIdImagenDbl
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As clsCBObjetoPan)
        MyBase.New(aobjPadre)
        hstrNombre = "IdImagen"
        henuTipoValor = EnuTipoValor.enuDouble
        hstrNombreCampoBd = "IdImagen"
        hblnEsRequerido = True
        hblnEsLlave = True
        hbytPosicionLlave = 0
    End Sub
    Public Overrides Sub SValide()
        hblnEsValido = (clsPanorama.fblnEsValidoNumero(hobjValorNew, GCDBLMINTERC, GCDBLMAXTERC, blnEsRequerido))
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(hobjValorPro) Then
            Return ""
        Else
            Return hobjValorPro.ToString
        End If
    End Function

End Class

Friend Class ClsIdCategoriaByt
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdCategoria"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = "IdTblCategoria"
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 1
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = (ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuCategoriaImagenDef.enuFotoTercero,
                EnuCategoriaImagenDef.enuFirmas, HblnEsRequerido))
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        HobjValorNew = EnuCategoriaImagenDef.None
        HobjValorPro = HobjValorNew
        HblnEsValido = False
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsOrdinalImagenShr
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
        HblnEsValido = (ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue, BlnEsRequerido))
        If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            HblnEsValido = True
        End If
    End Sub

    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function

End Class

Friend Class ClsPropiedadImagenImg
    Inherits ClsCBPropiedad
    Private MbytImagenGuardada As Byte() = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Imagen"
        HenuTipoValor = EnuTipoValor.enuImagen
        HstrNombreCampoBd = "Imagen"
    End Sub
    Private Sub EPreSetValor(aobjSender As Object, e As ClsPanEventArgs) _
            Handles Me.EvnPreSetValor
        If IsNothing(HobjValorNew) Then
            HblnEsValido = (False)
        Else
            If BlnLeyendoOrigen Then
                If HobjValorNew.GetType.Name = "Byte[]" Then
                    HobjValorPro = HobjValorNew
                    HblnEsValido = (True)
                ElseIf HobjValorNew.GetType.Name = "Bitmap" Then
                    HobjValorPro = HobjValorNew
                    SConviertaAByte()
                    SAsigneCambio(True, True)
                Else
                    HobjValorPro = Nothing
                    HblnEsValido = (False)
                End If
            Else
                If HobjValorNew.GetType.Name = "Bitmap" Then
                    HobjValorPro = HobjValorNew
                    SConviertaAByte()
                    ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                    HblnEsValido = (True)
                    SAsigneCambio(True, True)
                Else
                    HblnEsValido = (False)
                End If
            End If
        End If
        e.BlnCancele = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = False
        If BlnLeyendoOrigen Then
            If IsArray(HobjValorNew) Then
                If HobjValorNew.GetType.Name = "Byte" Then
                    lblnEsValido = True
                End If
            End If
        End If
        HblnEsValido = lblnEsValido
    End Sub
    Friend Function FimgImagenGuardada() As Image
        If Not IsNothing(BytImagenGuardada) Then
            Using lmstImagenBinaria As New MemoryStream(BytImagenGuardada)
                Dim limgImagen As Image = Image.FromStream(lmstImagenBinaria)
                Return limgImagen
            End Using
        Else
            Return Nothing
        End If
    End Function
    Friend Function FmstImagenBinaria() As MemoryStream
        Dim lmstImagenBinaria As MemoryStream = Nothing
        If Not IsNothing(BytImagenGuardada) Then
            lmstImagenBinaria = New MemoryStream(BytImagenGuardada)
        End If
        Return lmstImagenBinaria
    End Function
    Friend ReadOnly Property BytImagenGuardada As Byte()
        Get
            If IsNothing(MbytImagenGuardada) Then
                MbytImagenGuardada = HobjValorPro
            End If
            Return MbytImagenGuardada
        End Get
    End Property
    Private Sub SConviertaAByte()
        If HobjValorPro IsNot Nothing Then
            Using lmstImagenBinaria As New MemoryStream
                HobjValorPro.Save(lmstImagenBinaria, Imaging.ImageFormat.Jpeg)
                HobjValorPro = lmstImagenBinaria.GetBuffer
                HblnEsValido = (True)
            End Using
        Else
            HblnEsValido = (False)
        End If
    End Sub
    Public Overrides Function ToString() As String
        Return "Propiedad Imagen"
    End Function
End Class

Friend Class ClsFechaDtm
    Inherits ClsCBPropiedad

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Fecha"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = "Fecha"
        HblnEsRequerido = True
    End Sub

    Public Overrides Sub SValide()
        HblnEsValido = (ClsPanorama.FblnEsValidoFecha(HobjValorNew, DateAdd(DateInterval.Year, -1,
                Date.Today), Date.Today, BlnEsRequerido))
    End Sub

    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class

Friend Class ClsDescripcionStr
    Inherits ClsCBPropiedad

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Descripcion"
        HshrLongitud = 50
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "Descripcion"
    End Sub

    Public Overrides Sub SValide()
        HblnEsValido = (ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud, BlnEsRequerido))
    End Sub

    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
#End Region
