Imports System.Drawing
Friend Class ClsTercero
#Region "Definiciones"
    Inherits PanL.ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "PanTerceros"
    Private MobjFoto As ClsImagen = Nothing
    'Variables
    Private MintUltimaFoto As Integer = 0
#End Region

#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Tercero.
    ''' </summary>
    ''' <param name="aenuModoInstanciaObj">Indica si se instancia como un objeto navegable o como un Objeto único.</param>
    ''' <remarks>Si se instancia como un objeto navegable, se crea un datatable que contiene las columnas de
    ''' la llave con las llaves de todos los objetos y queda a la espera de que se indique que objeto abrir.
    ''' Si se instancia como un objeto único, queda a la espera de recibir el valor de los campos de la llave 
    ''' para abrir dicho objeto. </remarks>
    Public Sub New(aenuModoInstanciaObj As EnuModoInstanciaObjDef)
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuDeColeccion Then
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        End If
        Dim lstrCamposSelect As String()
        HobjPadre = Nothing
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuNavegable Then
            HblnEsAnulable = False
            lstrCamposSelect = {ObjIdTerceroDbl.StrNombreCampoBD}
        Else
            HblnEsCreable = False
            HblnEsSuprimible = False
            HblnEsAnulable = False
            HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
            lstrCamposSelect = {"*"}
        End If
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
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
            Return EnuIdClasesPanDef.EnuTercero
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Tercero"
        End Get
    End Property
    Friend Overrides ReadOnly Property HstrNombreObj As String
        Get
            Return Chr(34) & ObjIdTerceroDbl.ToString & Chr(34)
        End Get
    End Property
#End Region

#Region "Propiedades Prop"
    Friend ReadOnly Property ObjApellidoPrimeroStr As New ClsApellidoPrimeroStr(Me)
    Friend ReadOnly Property ObjApellidoSegundoStr As New ClsApellidoSegundoStr(Me)
    Friend ReadOnly Property ObjCodigoPostalStr As New ClsCodigoPostalStr(Me)
    Friend ReadOnly Property ObjDireccionUnoStr As New ClsDireccionUnoStr(Me)
    Friend ReadOnly Property ObjDireccionDosStr As New ClsDireccionDosStr(Me)
    Friend ReadOnly Property ObjPaginaWebStr As New ClsPaginaWebStr(Me)
    Friend ReadOnly Property ObjEmailStr As New ClsEmailStr(Me)
    Friend ReadOnly Property ObjFechaModificacionDtm As New ClsFechaModificacionDtm(Me)
    Friend ReadOnly Property ObjCiudadDirShr As New ClsCiudadDirShr(Me)
    Friend ReadOnly Property ObjDepartamentoDirByt As New ClsDepartamentoDirByt(Me)
    Friend ReadOnly Property ObjPaisDirStr As New ClsPaisDirStr(Me)
    Friend ReadOnly Property ObjTipoDocIdentidadByt As New ClsTipoDocIdentidadByt(Me)
    Friend ReadOnly Property ObjIdTerceroDbl As New ClsIdTerceroDbl(Me)
    Friend ReadOnly Property ObjIdUsuarioCreoStr As New ClsIdUsuarioCreoStr(Me)
    Friend ReadOnly Property ObjIdUsuarioModificoStr As New ClsIdUsuarioModificoStr(Me)
    Friend ReadOnly Property ObjNombrePrimeroStr As New ClsNombrePrimeroStr(Me)
    Friend ReadOnly Property ObjNombreSegundoStr As New ClsNombreSegundoStr(Me)
    Friend ReadOnly Property ObjRazonSocialStr As New ClsRazonSocialStr(Me)
    Friend ReadOnly Property ObjTelefonoUnoStr As New ClsTelefonoUnoStr(Me)
    Friend ReadOnly Property ObjTelefonoDosStr As New ClsTelefonoDosStr(Me)
    Friend ReadOnly Property ObjTipoTerceroByt As New ClsTipoTerceroByt(Me)
    Friend ReadOnly Property ObjCelularStr As New ClsCelularStr(Me)
    Friend ReadOnly Property ObjCelular2Str As New ClsCelular2Str(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjApellidoPrimeroStr)
                HcolPropiedades.Add(ObjApellidoSegundoStr)
                HcolPropiedades.Add(ObjCodigoPostalStr)
                HcolPropiedades.Add(ObjDireccionUnoStr)
                HcolPropiedades.Add(ObjDireccionDosStr)
                HcolPropiedades.Add(ObjPaginaWebStr)
                HcolPropiedades.Add(ObjEmailStr)
                HcolPropiedades.Add(ObjFechaCreacionDtm)
                HcolPropiedades.Add(ObjFechaModificacionDtm)
                HcolPropiedades.Add(ObjPaisDirStr)
                HcolPropiedades.Add(ObjDepartamentoDirByt)
                HcolPropiedades.Add(ObjCiudadDirShr)
                HcolPropiedades.Add(ObjTipoDocIdentidadByt)
                HcolPropiedades.Add(ObjIdTerceroDbl)
                HcolPropiedades.Add(ObjIdUsuarioCreoStr)
                HcolPropiedades.Add(ObjIdUsuarioModificoStr)
                HcolPropiedades.Add(ObjNombrePrimeroStr)
                HcolPropiedades.Add(ObjNombreSegundoStr)
                HcolPropiedades.Add(ObjRazonSocialStr)
                HcolPropiedades.Add(ObjTelefonoUnoStr)
                HcolPropiedades.Add(ObjTelefonoDosStr)
                HcolPropiedades.Add(ObjCelularStr)
                HcolPropiedades.Add(ObjCelular2Str)
                HcolPropiedades.Add(ObjTipoTerceroByt)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region

#Region "Otras Propiedades"
    Friend ReadOnly Property ImgUltimaImagen As ClsImagen
        Get
            Dim limgUltIma As ClsImagen = Nothing
            Dim lcolImagenes = FcolImagenes(EnuCategoriaImagenDef.enuFotoTercero)
            If lcolImagenes.Count > 0 Then
                limgUltIma = lcolImagenes(lcolImagenes.Count)
            End If
            Return limgUltIma
        End Get
    End Property
    Friend ReadOnly Property StrTelefono As String
        Get
            Dim lstrTel = String.Empty
            If Not String.IsNullOrEmpty(ObjCelularStr.ToString()) Then
                lstrTel = ObjCelularStr.ToString
            ElseIf Not String.IsNullOrEmpty(ObjCelular2Str.ToString()) Then
                lstrTel = ObjCelular2Str.ToString()
            ElseIf Not String.IsNullOrEmpty(ObjTelefonoUnoStr.ToString()) Then
                lstrTel = ObjTelefonoUnoStr.ToString()
            ElseIf Not String.IsNullOrEmpty(ObjTelefonoDosStr.ToString()) Then
                lstrTel = ObjTelefonoDosStr.ToString()
            End If
            Return lstrTel
        End Get
    End Property
#End Region
#End Region

#Region "Procedimientos y funciones invalidantes"
    Protected Friend Overrides Sub SVacie()
        MyBase.SVacie()
        MintUltimaFoto = 0
    End Sub

    Protected Friend Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        If EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            GobjPanDat.SControleProcesoObj(True)
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                ObjFechaModificacionDtm.ObjValorPro = Date.Now
                ObjFechaCreacionDtm.ObjValorPro = Date.Now
                ObjIdUsuarioCreoStr.ObjValorPro = GstrIdUsuario
                ObjIdUsuarioModificoStr.ObjValorPro = GstrIdUsuario
            ElseIf EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                ObjFechaModificacionDtm.ObjValorPro = Date.Now
                ObjIdUsuarioModificoStr.ObjValorPro = GstrIdUsuario
            End If
            Dim lblnNoHayError = False
            Try
                GobjPanDat.SInicialiceTransaccion()
                If FblnCambioNombre() Then
                    SActualiceCliente()
                End If
                MyBase.SActualice(ablnExigeRequeridos)
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
                    GobjPanDat.SControleProcesoObj(False, True)
                Else
                    GobjPanDat.SConfirmeTransaccion()
                    GobjPanDat.SControleProcesoObj(False)
                End If
            End Try
        End If
    End Sub

    Protected Friend Overrides Sub SCreeObj(aobjValorLlave() As Object)
        MyBase.SCreeObj(aobjValorLlave)
        If Not IsNothing(ObjValorLlave) Then
            Dim ldblIdTercero = ObjValorLlave(0)
            ObjIdTerceroDbl.ObjValorPro = ldblIdTercero
        End If
    End Sub

    Friend Overrides Function FblnEsSuprimible() As Boolean
        Dim lblnEsSuprimible As Boolean = BlnEsSuprimible
        If lblnEsSuprimible Then
            ' La condicion no debe incluir el nombre de de la columna
            Dim lstrCondicion As String = " = " & ObjIdTerceroDbl.ObjValorPro
            lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg({SstrNombreTabla},
                    ObjIdTerceroDbl.StrNombreCampoBD, lstrCondicion, True, False)
            If lblnEsSuprimible Then
                lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg({SstrNombreTabla},
                        ObjIdTerceroDbl.StrNombreCampoBD, lstrCondicion, True, True)
            End If
        End If
        Return lblnEsSuprimible
    End Function

    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdTerceroDbl.ToString
        End Get
    End Property

    Friend Overrides Function FblnSonValidosDatosOrigen(adtbOrigen As DataTable,
            astrColumnasRelacionadas As String(), ablnReinicie As Boolean,
            ByRef astrMens As String) As Boolean
        Dim lblnEsValido = False, i = 0, lstrColumnaOrigen As String
        Dim lstrEMail As String, ldblIdTercero As Double, lbytIdTipoTer As Byte, lbytIdTipoId As Byte
        For Each ldrwOrigen As DataRow In adtbOrigen.Rows
            i += 1
            lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                    ClsIdTerceroDbl.SstrNombreCampoBd)
            If Not IsDBNull(ldrwOrigen(lstrColumnaOrigen)) Then
                lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                        ClsEmailStr.SstrNombreCampoBd)
                lstrEMail = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                        EnuTipoValor.EnuString)
                lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                        ClsIdTerceroDbl.SstrNombreCampoBd)
                ldblIdTercero = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                        EnuTipoValor.EnuDouble)
                lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                        ClsTipoTerceroByt.SstrNombreCampoBd)
                lbytIdTipoTer = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                        EnuTipoValor.EnuByte)
                lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                        ClsTipoDocIdentidadByt.SstrNombreCampoBd)
                lbytIdTipoId = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                        EnuTipoValor.EnuByte)
                lblnEsValido = ClsPanorama.FblnEsValidoNumero(ldblIdTercero, 1, Double.MaxValue,
                        True, EnuTipoValor.EnuDouble)
                If Not lblnEsValido Then
                    astrMens = "La Identificación del Tercero en el Registro " & i.ToString &
                            " no es válido!"
                    Exit For
                End If
                lblnEsValido = ClsPanorama.FblnEsValidoString(lstrEMail, 5, 100, False)
                If Not String.IsNullOrEmpty(lstrEMail) Then
                    lblnEsValido = ClsPanorama.FblnEsValidoEMail(lstrEMail)
                End If
                If Not lblnEsValido Then
                    astrMens = "El Email en el Registro " & i.ToString & " no es válido!"
                    Exit For
                End If
                lblnEsValido = ClsPanorama.FblnEsValidoNumero(lbytIdTipoTer, 1, 2, True,
                        EnuTipoValor.EnuByte)
                If Not lblnEsValido Then
                    astrMens = "El Tipo de Tercero en el Registro " & i.ToString &
                            " no es válido!"
                    Exit For
                End If
                lblnEsValido = ClsPanorama.FblnEsValidoNumero(lbytIdTipoId, 1, 7, True,
                        EnuTipoValor.EnuByte)
                If Not lblnEsValido Then
                    astrMens = "El Tipo de Identificación en el Registro " & i.ToString & " no es válido!"
                    Exit For
                End If
            Else
                lblnEsValido = False
                astrMens = "La Identificación del Tercero en el Registro " & i.ToString &
                        " no es válida!"
            End If
        Next
        Return lblnEsValido
    End Function
#End Region

#Region "Procedimientos del objeto"
    Friend Function FstrNombreCompleto() As String
        Dim lstrNomComp As String
        If ObjRazonSocialStr.ObjValorPro <> "" Then
            lstrNomComp = ObjRazonSocialStr.ObjValorPro
        Else
            lstrNomComp = ObjNombrePrimeroStr.ObjValorPro
            If ObjNombreSegundoStr.ObjValorPro <> "" Then
                lstrNomComp &= " " & ObjNombreSegundoStr.ObjValorPro
            End If
            lstrNomComp &= " " & ObjApellidoPrimeroStr.ObjValorPro
            If ObjApellidoSegundoStr.ObjValorPro <> "" Then
                lstrNomComp &= " " & ObjApellidoSegundoStr.ObjValorPro
            End If
        End If
        Return lstrNomComp
    End Function

    Private Function FblnCambioNombre() As Boolean
        Dim lblnCambio = ObjNombrePrimeroStr.BlnCambio OrElse ObjNombreSegundoStr.BlnCambio OrElse
                    ObjApellidoPrimeroStr.BlnCambio OrElse ObjApellidoSegundoStr.BlnCambio OrElse
                    ObjRazonSocialStr.BlnCambio
        Return lblnCambio
    End Function

    Private Sub SActualiceCliente()
        Dim lobjApp = ClsAdministrador.FobjAppActual
        If lobjApp.ObjIdAppShr.ObjValorPro = EnuListaAplicaciones.EnuOrionCop Then
            Dim lstrExpSql = "UPDATE OriClientes SET NombreCompleto = '" &
                        FstrNombreCompleto() & "' WHERE IdTerceroCliente = " &
                        ObjIdTerceroDbl.ObjValorPro
            GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
            lstrExpSql = "UPDATE OriPropietarios SET NombreCompleto = '" &
                        FstrNombreCompleto() & "' WHERE IdTerceroPropietario = " &
                        ObjIdTerceroDbl.ObjValorPro
            GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        End If
    End Sub

    ''' <summary>
    ''' Cambia las Identidad de un tercero en todas las tablas donde exista un campo cuyo nombre empiece
    ''' por "IdTercero"
    ''' </summary>
    ''' <param name="adblIdTerceroActual">Valor de la Identidad actual del Tercero</param>
    ''' <param name="adblIdTerceroNueva">Valor de la nueva Identidad para el Tercero</param>
    ''' <remarks></remarks>
    Shared Sub SCambieIdTercero(adblIdTerceroActual As Double, adblIdTerceroNueva As Double)
        If adblIdTerceroActual = 0 OrElse adblIdTerceroNueva = 0 Then
            Dim lstrNombreArgumentoNulo As String
            If adblIdTerceroActual = 0 Then
                lstrNombreArgumentoNulo = "adblIdTerceroActual"
            Else
                lstrNombreArgumentoNulo = "adblIdTerceroNueva"
            End If
            Throw New ArgumentNullException(lstrNombreArgumentoNulo)
        End If
        Dim lstrNombeColumnaComo As String = "IdTercero"
        Dim lstrNombreTablasAActualizar As String() = GobjPanDat.FstrNombreTablasContienenColumna(lstrNombeColumnaComo)
        Dim lcolCamposACambiar As New Collection
        Dim lcolDatosNuevos As New Collection
        Dim lcolCamposRef As Collection
        Dim lcoldatosRef As New Collection
        GobjPanDat.SControleProcesoObj(True)
        GobjPanDat.SInicialiceTransaccion()
        For Each lstrNombreTabla As String In lstrNombreTablasAActualizar
            Dim lstrNombreCampos As String() = GobjPanDat.FstrNombresColumnasEnTabla(lstrNombreTabla,
                    lstrNombeColumnaComo)
            For Each lstrNombreCampo As String In lstrNombreCampos
                lcolCamposACambiar.Clear()
                lcolDatosNuevos.Clear()
                lcoldatosRef.Clear()
                lcolCamposACambiar.Add(lstrNombreCampo)
                lcolCamposRef = lcolCamposACambiar
                lcolDatosNuevos.Add(adblIdTerceroNueva.ToString)
                lcoldatosRef.Add(adblIdTerceroActual.ToString)
                GobjPanDat.SActualiceRegistro(lstrNombreTabla, lcolCamposACambiar,
                                 lcolDatosNuevos, lcolCamposRef, lcoldatosRef)
            Next
        Next
        GobjPanDat.SConfirmeTransaccion()
        GobjPanDat.SControleProcesoObj(False)
    End Sub

    Private Function FdtbImagenes(aenuCategoriaImagen As EnuCategoriaImagenDef) As DataTable
        Dim lstrFiltro As String = "IdImagen = '" & ObjIdTerceroDbl.ObjValorPro &
                    "' And IdTblCategoria = " & aenuCategoriaImagen
        Dim ldtbImagenes = ClsPanorama.FdtbDataTable("PanImagenes", {"*"},
                        {{"IdImagen", "ASC"}, {"IdTblCategoria", "ASC"}, {"Ordinal", "ASC"}},
                        lstrFiltro)
        Return ldtbImagenes
    End Function
#Region "Manejo de Imagenes"
    Friend Function FcolImagenes(aenuCategoriaImagen As EnuCategoriaImagenDef) As Collection
        Dim lcolImagenes As New Collection
        Dim ldtbImagenes = FdtbImagenes(aenuCategoriaImagen)
        For Each ldrwImagen As DataRow In ldtbImagenes.Select
            Dim lobjImagen As New ClsImagen(Me, ldrwImagen) With {
                .EnuPermisosObj = EnuPermisosObj
            }
            lobjImagen.SLeaValores(True)
            lcolImagenes.Add(lobjImagen)
        Next
        MintUltimaFoto = lcolImagenes.Count
        ClsPanorama.SActualiceCol(lcolImagenes)
        Return lcolImagenes
    End Function

    Friend Sub SAdicioneImagen(aimgImagen As Image, aenuCategImagen As EnuCategoriaImagenDef)
        If EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuConsultando Then
            If Not IsNothing(aimgImagen) AndAlso ObjIdTerceroDbl.BlnEsValido Then
                Dim lstrDescripcion = String.Empty
                Dim ldtbImagenes = FdtbImagenes(aenuCategImagen)
                Dim ldrwNuevaImagen As DataRow = ldtbImagenes.NewRow
                Dim lobjNuevaImagen As New ClsImagen(Me, ldrwNuevaImagen) With {
                        .EnuPermisosObj = Me.EnuPermisosObj
                    }
                lobjNuevaImagen.SCreeObj(Nothing)
                Select Case aenuCategImagen
                    Case EnuCategoriaImagenDef.EnuFotoTercero
                        lstrDescripcion = "Fotografia del tercero "
                    Case EnuCategoriaImagenDef.EnuFirmas
                        lstrDescripcion = "Firma del tercero "
                    Case EnuCategoriaImagenDef.EnuDocumentos
                        lstrDescripcion = "Documentos del tercero "
                    Case EnuCategoriaImagenDef.EnuFotoExtintor
                        lstrDescripcion = "Extintor del tercero "
                    Case EnuCategoriaImagenDef.EnuMedicas
                        lstrDescripcion = "Medicas del tercero "
                End Select
                With lobjNuevaImagen
                    .ObjIdImagenDbl.ObjValorPro = ObjIdTerceroDbl.ObjValorPro
                    .ObjIdCategoriaByt.ObjValorPro = aenuCategImagen
                    .ObjPropiedadImagenImg.BlnLeyendoOrigen = True
                    .ObjPropiedadImagenImg.ObjValorPro = aimgImagen
                    .ObjFechaDtm.ObjValorPro = Date.Today
                    .ObjDescripcionStr.ObjValorPro = lstrDescripcion &
                            ObjIdTerceroDbl.ToString
                End With
                lobjNuevaImagen.SActualice(True)
                If aenuCategImagen = EnuCategoriaImagenDef.EnuFotoTercero Then
                    Dim lcolImagenes = FcolImagenes(EnuCategoriaImagenDef.EnuFotoTercero)
                    MintUltimaFoto = lcolImagenes.Count
                End If
            End If
        End If
    End Sub

    Friend Function FmstFotoAnterior() As MemoryStream
        Dim lmstImagen As MemoryStream = Nothing
        Dim lcolImagenes As Collection = FcolImagenes(EnuCategoriaImagenDef.EnuFotoTercero)
        If lcolImagenes.Count > 0 Then
            MobjFoto = lcolImagenes(MintUltimaFoto)
            lmstImagen = MobjFoto.ObjPropiedadImagenImg.FmstImagenBinaria
            MintUltimaFoto -= 1
            If MintUltimaFoto = 0 Then
                MintUltimaFoto = lcolImagenes.Count
            End If
        End If
        Return lmstImagen
    End Function

    Friend Function FblnSuprimioFoto() As Boolean
        Dim lblnSuprimio = False
        If MobjFoto.FblnPermitidoSuprimir() Then
            lblnSuprimio = MobjFoto.FblnSuprimio()
            Dim lcolImagenes = FcolImagenes(EnuCategoriaImagenDef.EnuFotoTercero)
            If lblnSuprimio Then
                MintUltimaFoto = lcolImagenes.Count
            Else
                MintUltimaFoto += 1
                If MintUltimaFoto > lcolImagenes.Count Then
                    MintUltimaFoto = 1
                End If
            End If
        End If
        Return lblnSuprimio
    End Function

    Friend ReadOnly Property DtmFechaFoto As Date
        Get
            Dim ldtmfecha As Date = Nothing
            If Not IsNothing(MobjFoto) Then
                ldtmfecha = MobjFoto.ObjFechaDtm.ObjValorPro
            End If
            Return ldtmfecha
        End Get
    End Property
    ' Firmas
    Friend Function FmstFirma() As MemoryStream
        Dim lmstImagen As MemoryStream = Nothing
        Dim lcolImagenes As Collection = FcolImagenes(EnuCategoriaImagenDef.enuFirmas)
        If lcolImagenes.Count > 0 Then
            Dim lobjImagen As ClsImagen = lcolImagenes(1)
            lmstImagen = lobjImagen.ObjPropiedadImagenImg.FmstImagenBinaria
        End If
        Return lmstImagen
    End Function

    Friend Function FblnSuprimioFirma() As Boolean
        Dim lblnSuprimio = False
        Dim lcolImagenes = FcolImagenes(EnuCategoriaImagenDef.EnuFirmas)
        If lcolImagenes.Count > 0 Then
            Dim lobjFirma As ClsImagen = lcolImagenes(1)
            If lobjFirma.FblnPermitidoSuprimir Then
                lblnSuprimio = lobjFirma.FblnSuprimio()
            End If
        End If
        Return lblnSuprimio
    End Function

    Friend Function FobjFirma() As ClsImagen
        Dim lobjFirma As ClsImagen = Nothing
        Dim ldtbFirma = FdtbImagenes(EnuCategoriaImagenDef.EnuFirmas)
        If ldtbFirma.Rows.Count > 0 Then
            Dim ldrwFirma = ldtbFirma.Rows(0)
            lobjFirma = New ClsImagen(Me, ldrwFirma)
            lobjFirma.SLeaValores(True)
        End If
        Return lobjFirma
    End Function

    Friend Function FbytFirma() As Byte()
        Dim lbytFirma As Byte() = Nothing
        Dim lobjFirma = FobjFirma()
        If Not IsNothing(lobjFirma) Then
            lbytFirma = lobjFirma.ObjPropiedadImagenImg.BytImagenGuardada
        End If
        Return lbytFirma
    End Function
#End Region
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsApellidoPrimeroStr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsTercero = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "Apellido1"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "ApellidoPrimero"
        HshrLongitud = 20
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = True
        HblnEsRequerido = MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro <> EnuTipoDocIdDef.enuNit AndAlso
                    MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro <> EnuTipoDocIdDef.enuNuip
        If BlnEsRequerido Then
            HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 2, ShrLongitud, BlnEsRequerido)
        ElseIf Not IsNothing(HobjValorNew) Then
            If Not String.IsNullOrEmpty(HobjValorNew) Then
                HblnEsValido = False
            End If
        End If
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If Not IsNothing(HobjValorNew) And HblnEsValido Then
                HobjValorNew = FstrNombreTercero(HobjValorNew)
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsApellidoSegundoStr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsTercero = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "Apellido2"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "ApellidoSegundo"
        HshrLongitud = 20
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = True
        If MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro <> EnuTipoDocIdDef.enuNit AndAlso
                    MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro <> EnuTipoDocIdDef.enuNuip Then
            HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        ElseIf Not IsNothing(HobjValorNew) Then
            If Not String.IsNullOrEmpty(HobjValorNew) Then
                HblnEsValido = False
            End If
        End If
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If Not IsNothing(HobjValorNew) And HblnEsValido Then
                HobjValorNew = FstrNombreTercero(HobjValorNew)
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsCelularStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Celular"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Celular"
        HenuTipoValor = EnuTipoValor.EnuString
        HshrLongitud = 16
        HStrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        If HobjValorNew IsNot Nothing Then
            HobjValorNew = HobjValorNew.ToString().Replace(" ", "")
        End If
        HblnEsRequerido = If(HobjValorNew <> String.Empty, True, False)
        HblnEsValido = (ClsPanorama.FblnEsValidoStringNumerico(HobjValorNew, 10, ShrLongitud,
                BlnEsRequerido))
        If Not HblnEsValido Then
            HstrMens = "El campo Teléfono solo admite números. Debe tener al menos 10 " &
                    "caracteres o puede dejarlo vacío.!"
            SNotifiqueDatInv()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsCelular2Str
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Celular2"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Celular2"
        HenuTipoValor = EnuTipoValor.EnuString
        HshrLongitud = 16
        HStrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        If HobjValorNew IsNot Nothing Then
            HobjValorNew = HobjValorNew.ToString().Replace(" ", "")
        End If
        HblnEsRequerido = If(HobjValorNew <> String.Empty, True, False)
        HblnEsValido = (ClsPanorama.FblnEsValidoStringNumerico(HobjValorNew, 10, ShrLongitud,
                BlnEsRequerido))
        If Not HblnEsValido Then
            HstrMens = "El campo Teléfono solo admite números. Debe tener al menos 10 " &
                    "caracteres o puede dejarlo vacío.!"
            SNotifiqueDatInv()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsCodigoPostalStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "CodigoPostal"
    Private ReadOnly MobjPadre As ClsTercero = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Código Postal"
        HshrLongitud = 6
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoStringNumerico(HobjValorNew, ShrLongitud,
                ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            HblnEsValido = MobjPadre.ObjCiudadDirShr.BlnEsValido
        End If
        If HblnEsValido Then
            If MobjPadre.ObjCiudadDirShr.ObjValorPro = 0 Then
                HblnEsValido = HobjValorNew = String.Empty
            Else
                HblnEsValido = Not String.IsNullOrEmpty(HobjValorNew.ToString())
            End If
        End If
    End Sub
    Private Sub EPosCambioVlr(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosCambio
        MobjPadre.ObjDireccionUnoStr.SValide()
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsDireccionUnoStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Direccion1"
    Private ReadOnly MobjPadre As ClsTercero = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "DireccionUno"
        HshrLongitud = 80
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "Direccion1"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            HblnEsValido = MobjPadre.ObjCiudadDirShr.BlnEsValido
        End If
        If HblnEsValido Then
            If String.IsNullOrEmpty(MobjPadre.ObjCodigoPostalStr.ToString()) Then
                HblnEsValido = HobjValorNew = String.Empty
            Else
                HblnEsValido = Not String.IsNullOrWhiteSpace(HobjValorNew) AndAlso
                        Not String.IsNullOrEmpty(HobjValorNew)
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsDireccionDosStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Direccion2"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "DireccionDos"
        HshrLongitud = 80
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "Direccion2"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = (ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud, BlnEsRequerido))
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsEmailStr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsTercero = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "Email"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Email"
        HshrLongitud = 100
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 5, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            HobjValorNew = HobjValorNew.ToString.Trim
            If Not String.IsNullOrEmpty(HobjValorNew) Then
                HblnEsValido = ClsPanorama.FblnEsValidoEMail(HobjValorNew)
            End If
            If Not HblnEsValido Then
                HstrMens = "El Email del Tercero no es valido!"
                SNotifiqueDatInv()
            Else
                If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuConsultando Then
                    If Not IsNothing(HobjValorNew) And HblnEsValido Then
                        HobjValorNew = HobjValorNew.ToString.Trim
                    End If
                End If
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsPaginaWebStr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsTercero = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "PaginaWeb"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Pagina Web"
        HshrLongitud = 50
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 7, HshrLongitud, BlnEsRequerido)
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If Not IsNothing(HobjValorNew) And HblnEsValido Then
                HobjValorNew = HobjValorNew.ToString.Trim
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsFechaExpedicionDocDtm
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "FechaExpedicionDoc"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = "FechaExpedicionDoc"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoFecha(HobjValorNew, GCDTMFECHANULA, Date.Today, BlnEsRequerido)
        HblnEsValido = lblnEsValido
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        HobjValorPro = GCDTMFECHANULA
        HobjValorNew = HobjValorPro
        HblnEsValido = False
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return Format(GCDTMFECHANULA, GCSTRFMTFECHASIMPLE)
        Else
            Return Format(HobjValorPro, GCSTRFMTFECHASIMPLE)
        End If
    End Function
End Class

Friend Class ClsFechaModificacionDtm
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "FechaModificacion"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = "FechaModificacion"
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim ldtmFechaMin As Date = DateAdd(DateInterval.Minute, -1, Date.Now)
        Dim ldtmFechaMax As Date = DateAdd(DateInterval.Minute, 1, Date.Now)
        Dim lblnEsValido As Boolean = True
        Dim lobjPadre As ClsTercero = ObjPadre
        If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            lblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax, BlnEsRequerido)
        End If
        HblnEsValido = lblnEsValido
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        HobjValorPro = Date.Now
        HobjValorNew = HobjValorPro
        HblnEsValido = False
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return Format(GCDTMFECHANULA, GCSTRFMTFECHASIMPLE)
        Else
            Return Format(HobjValorPro, GCSTRFMTFECHASIMPLE)
        End If
    End Function
End Class

Friend Class ClsCiudadDirShr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsTercero = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "CiudadDir"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = "IdCiudadDir"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCSHRMINCIUD, GCSHRMAXCIUD,
                BlnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            HblnEsValido = MobjPadre.ObjDepartamentoDirByt.BlnEsValido
        End If
        If HblnEsValido Then
            If String.IsNullOrEmpty(MobjPadre.ObjPaisDirStr.ToString()) Then
                HblnEsValido = HobjValorNew = 0
            Else
                HblnEsValido = HobjValorNew > 0
            End If
        End If
    End Sub
    Private Sub EPosCambioVlr(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosCambio
        MobjPadre.ObjCodigoPostalStr.SValide()
        MobjPadre.ObjDireccionUnoStr.SValide()
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsCiudadExpShr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CiudadExp"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = "IdCiudadExp"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = (ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCSHRMINCIUD, GCSHRMAXCIUD, BlnEsRequerido, EnuTipoValor))
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsDepartamentoDirByt
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsTercero = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "DepartamentoDir"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = "IdDepartamentoDir"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCBYTMINDPTO, GCBYTMAXDPTO,
                BlnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            If String.IsNullOrEmpty(MobjPadre.ObjPaisDirStr.ToString()) Then
                HblnEsValido = HobjValorNew = 0
            Else
                HblnEsValido = HobjValorNew > 0
            End If
        End If
    End Sub
    Private Sub EPosCambioVlr(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosCambio
        MobjPadre.ObjCiudadDirShr.SValide()
    End Sub
    Public Overrides Function ToString() As String
        If HobjValorPro = 0 Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsDepartamentoExpByt
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "DepartamentoExp"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = "IdDepartamentoExp"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = (ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCBYTMINDPTO, GCBYTMAXDPTO, BlnEsRequerido, EnuTipoValor))
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsPaisDirStr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsTercero = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = ObjPadre
        HstrNombre = "PaisDir"
        HenuTipoValor = EnuTipoValor.EnuString
        HshrLongitud = 2
        HstrNombreCampoBd = "IdPaisDir"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, HshrLongitud, BlnEsRequerido)
    End Sub
    Private Sub EPosCambioVlr(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosCambio
        MobjPadre.ObjDepartamentoDirByt.SValide()
        MobjPadre.ObjCiudadDirShr.SValide()
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsPaisExpStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "PaisExp"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 2
        HstrNombreCampoBd = "IdPaisExp"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = (ClsPanorama.FblnEsValidoString(HobjValorNew, 0, HshrLongitud, BlnEsRequerido))
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return String.Empty
        Else
            Return HobjValorPro.ToString()
        End If
    End Function
End Class

Friend Class ClsTipoDocIdentidadByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTblTipoId"
    Private ReadOnly MobjPadre As ClsTercero = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "TipoDocIdentidad"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = (ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuTipoDocIdDef.enuCedulaCiudadania,
                EnuTipoDocIdDef.enuPEP, HblnEsRequerido))
    End Sub
    Private Sub EPosSetValor() Handles Me.EvnPosSetValor
        If HblnEsValido Then
            MobjPadre.ObjRazonSocialStr.SValide()
            MobjPadre.ObjApellidoPrimeroStr.SValide()
            MobjPadre.ObjApellidoSegundoStr.SValide()
            MobjPadre.ObjNombrePrimeroStr.SValide()
            MobjPadre.ObjNombreSegundoStr.SValide()
            MobjPadre.ObjPaginaWebStr.SValide()
            If HobjValorPro <> EnuTipoDocIdDef.enuNit AndAlso
                         HobjValorPro <> EnuTipoDocIdDef.enuNuip Then
                MobjPadre.ObjIdTerceroDbl.SbyDigitoVerificacion = -1
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        Dim lstrTipopDoc = String.Empty
        If Not IsNothing(HobjValorPro) Then
            Dim lenuTipoDoc As EnuTipoDocIdDef = CType(HobjValorPro, Integer)
            Select Case lenuTipoDoc
                Case EnuTipoDocIdDef.enuCedulaCiudadania
                    lstrTipopDoc = "C.C."
                Case EnuTipoDocIdDef.enuNit
                    lstrTipopDoc = "NIT"
                Case EnuTipoDocIdDef.enuCedulaExtranjeria
                    lstrTipopDoc = "Cédula de Extranjería"
                Case EnuTipoDocIdDef.enuNuip
                    lstrTipopDoc = "NUIP"
                Case EnuTipoDocIdDef.enuPasaporte
                    lstrTipopDoc = "Pasaporte"
                Case EnuTipoDocIdDef.enuRegistroCivil
                    lstrTipopDoc = "Registro Civil"
                Case EnuTipoDocIdDef.enuTarjetaIdentidad
                    lstrTipopDoc = "T.I."
            End Select
        End If
        Return lstrTipopDoc
    End Function
End Class

Friend Class ClsTipoTerceroByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTblTipoTercero"
    Private ReadOnly MobjPadre As ClsTercero = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "TipoTercero"
        HenuTipoValor = EnuTipoValor.EnuByte
        HStrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuTipoTerceroDef.enuPersonaNatural,
                EnuTipoTerceroDef.enuPersonaJuridica, HblnEsRequerido)
        If lblnEsValido Then
            If HobjValorNew = EnuTipoTerceroDef.enuPersonaNatural AndAlso
                    (MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro = EnuTipoDocIdDef.enuNit OrElse
                     MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro = EnuTipoDocIdDef.enuNuip) Then
                lblnEsValido = False
            ElseIf HobjValorNew = EnuTipoTerceroDef.enuPersonaJuridica AndAlso
                    (MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro <> EnuTipoDocIdDef.enuNit AndAlso
                    MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro <> EnuTipoDocIdDef.enuNuip) Then
                lblnEsValido = False
            End If
        End If
        HblnEsValido = lblnEsValido
    End Sub
    Private Sub ClsTipoTerceroByt_evnPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If Not (e.BlnVaciandoObjeto OrElse IsNothing(MobjPadre.ObjIdTerceroDbl.ObjValorPro)) Then
            MobjPadre.ObjIdTerceroDbl.SValide()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsIdTerceroDbl
    Inherits ClsCBPropiedad
    Private MsbyDigitoVerificacion As SByte = -1
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTercero"
    Private ReadOnly MobjPadre As ClsTercero = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdTercero"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 0
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MsbyDigitoVerificacion = -1
        HblnEsValido = False
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC,
                    GCDBLMAXTERC, BlnEsRequerido)
        If HblnEsValido Then
            HblnEsValido = HobjValorNew <> GCDBLTERCERONULO
            If Not HblnEsValido Then
                HstrMens = "Este número de identificación está reservado por el programa!"
            End If
        End If
        If HblnEsValido Then
            Dim lobjLlavePrincipal() As Object
            If Not BlnLeyendoOrigen Then
                lobjLlavePrincipal = {HobjValorNew}
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                    If MobjPadre.FblnExisteLlave(lobjLlavePrincipal) Then
                        HstrMens = "El Tercero con el Id. " & HobjValorNew.ToString & " ya existe!"
                        HblnEsValido = False
                    End If
                ElseIf MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                    HblnEsValido = MobjPadre.FblnExisteLlave(lobjLlavePrincipal)
                    If HblnEsValido Then
                        If HobjValorNew <> HobjValorPro Then
                            MobjPadre.SAbra(lobjLlavePrincipal)
                        End If
                        If MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro = EnuTipoDocIdDef.EnuNit OrElse
                                MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro = EnuTipoDocIdDef.EnuNuip Then
                            MsbyDigitoVerificacion = FsbyDigitoVarificacion()
                        End If
                    Else
                        HstrMens = "El Tercero con el Id. " & HobjValorNew.ToString & " no existe!"
                    End If
                Else
                    HblnEsValido = (HobjValorOriginal = HobjValorNew)
                    If Not HblnEsValido Then
                        HstrMens = "No es permitido cambiar la Identidad del Tercero!"
                    End If
                End If
                If HblnEsValido Then
                    If MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro = EnuTipoDocIdDef.enuNit OrElse
                            MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro = EnuTipoDocIdDef.enuNuip Then
                        If MsbyDigitoVerificacion <> FsbyDigitoVarificacion() Then
                            HblnEsValido = False
                        End If
                    End If
                End If
            End If
        Else
            If Not MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando AndAlso
                    HobjValorNew <> GCDBLTERCERONULO Then
                HstrMens = "La Id. del Tercero Ingresada no es valida!"
            End If
        End If
        If Not HblnEsValido AndAlso Not String.IsNullOrEmpty(HstrMens) Then
            SNotifiqueDatInv()
        End If
    End Sub
    Public Property SbyDigitoVerificacion As Object
        Get
            If Not (MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro = EnuTipoDocIdDef.enuNit OrElse
                            MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro = EnuTipoDocIdDef.enuNuip) Then
                MsbyDigitoVerificacion = -1
            Else
                MsbyDigitoVerificacion = -2
            End If
            If MsbyDigitoVerificacion = -2 Then
                MsbyDigitoVerificacion = FsbyDigitoVarificacion()
            End If
            Return MsbyDigitoVerificacion
        End Get
        Set(value As Object)
            Dim lblnEsValido As Boolean
            If MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro = EnuTipoDocIdDef.enuNit OrElse
                    MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro = EnuTipoDocIdDef.enuNuip Then
                lblnEsValido = ClsPanorama.FblnEsValidoNumero(value, 0, 9, True,
                        EnuTipoValor.enuSByte)
                If lblnEsValido Then
                    If value = FsbyDigitoVarificacion() Then
                        lblnEsValido = True
                        MsbyDigitoVerificacion = value
                    Else
                        lblnEsValido = False
                    End If
                End If
                HblnEsValido = lblnEsValido
            Else
                If Not IsNumeric(value) Then
                    If String.IsNullOrEmpty(value) Then
                        MsbyDigitoVerificacion = -1
                    Else
                        HblnEsValido = (False)
                    End If
                ElseIf value <> -1 Then
                    HblnEsValido = (False)
                Else
                    MsbyDigitoVerificacion = value
                End If
            End If
        End Set
    End Property
    Private Function FsbyDigitoVarificacion() As SByte
        Dim i As Integer
        Dim llngAcum As Long = 0
        Dim lsbyDigitoVer As SByte
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        For i = 0 To HobjValorNew.ToString.Length - 1
            llngAcum += CInt(HobjValorNew.ToString.Substring(i, 1)) *
                    Choose(Len(HobjValorNew.ToString) - i, 3, 7, 13, 17, 19, 23, 29, 37, 41, 43, 47, 53, 59, 67, 71)
        Next i
        Select Case llngAcum Mod 11
            Case 0
                lsbyDigitoVer = 0
            Case 1
                lsbyDigitoVer = 1
            Case Else
                lsbyDigitoVer = 11 - (llngAcum Mod 11)
        End Select
        Return lsbyDigitoVer
    End Function
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsIdUsuarioCreoStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdUsuarioCreo"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "IdUsuarioCreo"
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If String.IsNullOrEmpty(HobjValorNew) Then
            HobjValorNew = GstrIdUsuario
        End If
        HblnEsValido = (ClsPanorama.FblnEsValidoString(HobjValorNew, 2, ShrLongitud, BlnEsRequerido))
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsIdUsuarioModificoStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdUsuarioModifico"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "IdUsuarioModifico"
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If String.IsNullOrEmpty(HobjValorNew) Then
            HobjValorNew = GstrIdUsuario
        End If
        HblnEsValido = (ClsPanorama.FblnEsValidoString(HobjValorNew, 2, ShrLongitud, BlnEsRequerido))
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsNombrePrimeroStr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsTercero = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "Nombre1"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "NombrePrimero"
        HshrLongitud = 20
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = True
        If MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro <> EnuTipoDocIdDef.enuNit AndAlso
                    MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro <> EnuTipoDocIdDef.enuNuip Then
            HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 2, ShrLongitud, BlnEsRequerido)
        ElseIf Not IsNothing(HobjValorNew) Then
            If Not String.IsNullOrEmpty(HobjValorNew) Then
                HblnEsValido = False
            End If
        End If
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If Not IsNothing(HobjValorNew) And HblnEsValido Then
                HobjValorNew = FstrNombreTercero(HobjValorNew)
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsNombreSegundoStr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsTercero = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "Nombre2"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "NombreSegundo"
        HshrLongitud = 20
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = True
        If MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro <> EnuTipoDocIdDef.enuNit AndAlso
                    MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro <> EnuTipoDocIdDef.enuNuip Then
            HblnEsValido = (ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido))
        ElseIf Not IsNothing(HobjValorNew) Then
            If Not String.IsNullOrEmpty(HobjValorNew) Then
                HblnEsValido = False
            End If
        End If
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If Not IsNothing(HobjValorNew) And HblnEsValido Then
                HobjValorNew = FstrNombreTercero(HobjValorNew)
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsRazonSocialStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "RazonSocial"
    Private ReadOnly MobjPadre As ClsTercero = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "RazonSocial"
        HshrLongitud = 100
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = True
        If MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro = EnuTipoDocIdDef.enuNit OrElse
                    MobjPadre.ObjTipoDocIdentidadByt.ObjValorPro = EnuTipoDocIdDef.enuNuip Then
            HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        ElseIf Not IsNothing(HobjValorNew) Then
            If Not String.IsNullOrEmpty(HobjValorNew) Then
                HblnEsValido = False
            End If
        End If
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If Not IsNothing(HobjValorNew) And HblnEsValido Then
                HobjValorNew = FstrNombreTercero(HobjValorNew)
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If Not e.BlnVaciandoObjeto Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                MobjPadre.ObjApellidoPrimeroStr.SValide()
                MobjPadre.ObjNombrePrimeroStr.SValide()
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsTelefonoUnoStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Telefono1"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TelefonoUno"
        HenuTipoValor = EnuTipoValor.EnuString
        HshrLongitud = 16
        HStrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HobjValorNew = HobjValorNew.ToString().Replace(" ", "")
        HblnEsRequerido = If(HobjValorNew <> String.Empty, True, False)
        HblnEsValido = (ClsPanorama.FblnEsValidoStringNumerico(HobjValorNew, 10, ShrLongitud,
                BlnEsRequerido))
        If Not HblnEsValido Then
            HstrMens = "El campo Teléfono solo admite números. Debe tener al menos 10 " &
                    "caracteres o puede dejarlo vacío.!"
            SNotifiqueDatInv()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsTelefonoDosStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Telefono2"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TelefonoDos"
        HenuTipoValor = EnuTipoValor.EnuString
        HshrLongitud = 16
        HStrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HobjValorNew = HobjValorNew.ToString().Replace(" ", "")
        HblnEsRequerido = If(HobjValorNew <> String.Empty, True, False)
        HblnEsValido = (ClsPanorama.FblnEsValidoStringNumerico(HobjValorNew, 10, ShrLongitud,
                BlnEsRequerido))
        If Not HblnEsValido Then
            HstrMens = "El campo Teléfono solo admite números. Debe tener al menos 10 " &
                    "caracteres o puede dejarlo vacío.!"
            SNotifiqueDatInv()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
#End Region