Namespace ActualizaBd
    Friend Class ClsPoneAlDiaMySql
        Inherits clsCBPoneAlDiaBD
#Region "Definiciones"
        Private MstrCharacterSetBD As String = String.Empty
        Private MstrCollationNameBD As String = String.Empty
        Private MstrCollationNameTabla As String = String.Empty
        Private MdtbDatabasesBD As DataTable = Nothing
        Private MdtbTablasBD As DataTable = Nothing
        Private MdtbColumnasBD As DataTable = Nothing
        Private MdtbIndicesBD As DataTable = Nothing
        Private MdtbColumnasIndicesBD As DataTable = Nothing
#End Region
#Region "Constructores"
        Public Sub New()
            If Not gobjPanDat.blnRegistrado Then
                Throw New System.InvalidOperationException("El módulo no esta registrado")
            End If
        End Sub
#End Region
#Region "Estructura BD a partir conexión"
        ''' <summary>
        ''' Crea la estructura de la base de datos a partir de la conexion a MySql
        ''' </summary>
        ''' <remarks></remarks>
        ''' 
        Protected Overrides Sub SGenereEstructuraActualBD(astrPrefijoTablas As String)
            Dim lobjTabla As ClsTabla
            SVerifiqueEstadoBD()
            If HenuEstadoBD <> EnuEstadoBaseDatos.enuNoExiste Then
                Dim lstrNomBD = StrNombreBD
                SEstablezcaObjetosSchemaBd()
                Dim lstrFiltroTablas As String = "TABLE_NAME LIKE '" & astrPrefijoTablas & "*'"
                Dim ldrwBaseDatos As DataRow = MdtbDatabasesBD.Select("database_name = '" &
                        lstrNomBD & "'")(0)
                MstrCharacterSetBD = ldrwBaseDatos("DEFAULT_CHARACTER_SET_NAME")
                MstrCollationNameBD = ldrwBaseDatos("DEFAULT_COLLATION_NAME")
                If IsNothing(HobjEstructuraBD) Then
                    HobjEstructuraBD = New ClsEstructuraBD
                End If
                HobjBaseDatosBD = HobjEstructuraBD.ObjBaseDatos
                HobjBaseDatosBD.SAsignaPropiedadesBD(HstrNombreBDEnXml, MstrCharacterSetBD,
                    MstrCollationNameBD, HentVersionBD)
                Dim ldrwTablas As DataRow() = MdtbTablasBD.Select(lstrFiltroTablas)
                If ldrwTablas.Length = 0 Then
                    HenuEstadoBD = EnuEstadoBaseDatos.enuDespoblada
                End If
                For Each ldrwTabla As DataRow In ldrwTablas
                    MstrCollationNameTabla = ldrwTabla("TABLE_COLLATION")
                    If MstrCollationNameTabla = MstrCollationNameBD Then
                        MstrCollationNameTabla = String.Empty
                    End If
                    lobjTabla = HobjBaseDatosBD.FobjNuevaTabla(ldrwTabla("TABLE_NAME"),
                        MstrCollationNameTabla)
                    SAdicioneColumnas(lobjTabla)
                    SAdicioneIndices(lobjTabla)
                    If lobjTabla.StrNombre.ToUpper.Contains("TBL") Then
                        SAdicioneRegistros(lobjTabla)
                    End If
                Next
                'If Not clsActualizaBD.fblnHayError() Then
                '    hobjBaseDatosBD.sAdicioneRelaciones(aBDBD)
                'End If
            End If
        End Sub
        Friend Function FobjBaseDatosDB(astrPrefTablas As String) As ClsBaseDatos
            SGenereEstructuraActualBD(astrPrefTablas)
            If String.IsNullOrEmpty(HobjBaseDatosBD.StrNombreBD) Then
                Dim lstrNomBd = ""
                If astrPrefTablas = My.Resources.PrefPan Then
                    lstrNomBd = StrNombreBD
                ElseIf astrPrefTablas = My.Resources.PrefOri Then
                    lstrNomBd = My.Resources.NomBdOri
                End If
                HobjBaseDatosBD.SAsigneNombreBD(lstrNomBd)
            End If
            Return HobjBaseDatosBD
        End Function
        Private Sub SVerifiqueEstadoBD()
            If Not GobjPanDat.BlnExisteBdPanorama Then
                HenuEstadoBD = EnuEstadoBaseDatos.enuNoExiste
                HentVersionBD = 0
            Else
                Dim lentIdVersonBD = GobjPanDat.FentVersionBD(GshrIdAplicacion)
                HentVersionBD = lentIdVersonBD
            End If
        End Sub
        Private Sub SEstablezcaObjetosSchemaBd()
            Dim lstrFiltro(3) As String, lblnNoHayError = False
            lstrFiltro(1) = StrNombreBD.ToLower
            With GobjPanDat
                Try
                    .SControleProcesoObj(True)
                    .SAbraConexionBd()
                    MdtbDatabasesBD = .CnnConexionBd_App.GetSchema("DataBases", lstrFiltro)
                    MdtbTablasBD = .CnnConexionBd_App.GetSchema("Tables", lstrFiltro)
                    MdtbColumnasBD = .CnnConexionBd_App.GetSchema("Columns", lstrFiltro)
                    MdtbIndicesBD = .CnnConexionBd_App.GetSchema("Indexes", lstrFiltro)
                    MdtbColumnasIndicesBD = .CnnConexionBd_App.GetSchema("IndexColumns", lstrFiltro)
                    lblnNoHayError = True
                Catch ex As Exception
                    Throw
                Finally
                    If lblnNoHayError Then
                        .SControleProcesoObj(False)
                    Else
                        .SControleProcesoObj(False, True)
                    End If
                End Try
            End With
        End Sub
        Private Sub SAdicioneColumnas(ByRef aobjTabla As ClsTabla)
            Dim lblnRequerido As Boolean
            Dim lstrTipoColumna As String
            Dim lblnAutoNumerico As Boolean
            Dim lstrCharacterSetCol As String
            Dim lstrCollationNameCol As String
            Dim lstrValorPorDefecto As String = String.Empty
            Dim lstrLongitud As String
            Dim lstrFiltro As String = ("TABLE_NAME = '" & aobjTabla.StrNombre & "'")
            Dim ldrwColumnas() As DataRow = MdtbColumnasBD.Select(lstrFiltro)
            For Each ldrwColumna As DataRow In ldrwColumnas
                If Not IsDBNull(ldrwColumna("EXTRA")) AndAlso ldrwColumna("EXTRA").ToString = "auto_increment" Then
                    lblnAutoNumerico = True
                Else
                    lblnAutoNumerico = False
                End If
                If ldrwColumna("IS_NULLABLE").ToUpper = "NO" Then
                    lblnRequerido = True
                Else
                    lblnRequerido = False
                End If
                lstrTipoColumna = FstrTipoDatosScriptSchema(ldrwColumna("COLUMN_TYPE"))
                lstrLongitud = FstrLongitudStringScriptSchema(ldrwColumna("COLUMN_TYPE"))
                If Not IsDBNull(ldrwColumna("CHARACTER_SET_NAME")) Then
                    lstrCharacterSetCol = ldrwColumna("CHARACTER_SET_NAME")
                    lstrCollationNameCol = ldrwColumna("COLLATION_NAME")
                    If Not String.IsNullOrEmpty(MstrCollationNameTabla) Then
                        If MstrCollationNameTabla = lstrCollationNameCol Then
                            lstrCharacterSetCol = String.Empty
                            lstrCollationNameCol = String.Empty
                        End If
                    Else
                        If MstrCharacterSetBD = lstrCharacterSetCol Then
                            lstrCharacterSetCol = String.Empty
                        End If
                        If MstrCollationNameBD = lstrCollationNameCol Then
                            lstrCollationNameCol = String.Empty
                        End If
                    End If
                Else
                    lstrCharacterSetCol = String.Empty
                    lstrCollationNameCol = String.Empty
                End If
                If Not IsDBNull(ldrwColumna("COLUMN_DEFAULT")) Then
                    If ldrwColumna("COLUMN_DEFAULT") <> "''" Then
                        lstrValorPorDefecto = ldrwColumna("COLUMN_DEFAULT")
                        If lstrValorPorDefecto.Contains("'") Then
                            lstrValorPorDefecto = lstrValorPorDefecto.Replace("'", "")
                        End If
                        If ClsCBPoneAlDiaBD.FblnValorEsValorProDefecto(lstrValorPorDefecto) Then
                            lstrValorPorDefecto = String.Empty
                        End If
                    End If
                End If
                aobjTabla.SAdicioneColumna(ldrwColumna("COLUMN_NAME"), lblnRequerido, lstrTipoColumna, lblnAutoNumerico,
                                           lstrLongitud,
                        lstrCharacterSetCol, lstrCollationNameCol, lstrValorPorDefecto, ldrwColumna("COLUMN_COMMENT"))
            Next
        End Sub
        Private Sub SAdicioneIndices(ByRef aobjTabla As ClsTabla)
            Dim lstrFiltro As String = "TABLE_NAME = '" & aobjTabla.StrNombre & "'"
            Dim ldrwIndices() As DataRow = MdtbIndicesBD.Select(lstrFiltro)
            Dim lstrNombreIndice As String
            Dim lblnEsPrincipal As Boolean
            Dim lblnEsUnico As Boolean
            Dim lobjIndice As ClsIndice = Nothing
            For Each ldrwIndice As DataRow In ldrwIndices
                If ldrwIndice("INDEX_NAME") = "PRIMARY" Then
                    lstrNombreIndice = "PK_" & aobjTabla.StrNombre
                Else
                    lstrNombreIndice = ldrwIndice("INDEX_NAME")
                End If
                lblnEsPrincipal = ldrwIndice("PRIMARY")
                lblnEsUnico = ldrwIndice("UNIQUE")
                lobjIndice = aobjTabla.FobjNuevoIndice(lstrNombreIndice, lblnEsPrincipal, lblnEsUnico)
                SAdicioneColumnasIndice(lobjIndice)
            Next
        End Sub
        Private Sub SAdicioneColumnasIndice(ByRef aobjIndice As ClsIndice)
            Dim lstrPropiedades(1) As String
            Dim lstrNombreIndice As String
            Dim lstrNombreColumnaIndice As String
            Dim lstrOrdenIndice As String
            Dim lstrFiltro As String
            If aobjIndice.StrNombre.StartsWith("PK_") Then
                lstrNombreIndice = "PRIMARY"
            Else
                lstrNombreIndice = aobjIndice.StrNombre
            End If
            lstrFiltro = "INDEX_NAME = '" & lstrNombreIndice & "' AND TABLE_NAME = '" & aobjIndice.ObjPadre.StrNombre & "'"
            Dim ldrwColumnasIndices() As DataRow = MdtbColumnasIndicesBD.Select(lstrFiltro)
            For Each ldrwColumnaIndice As DataRow In ldrwColumnasIndices
                lstrNombreColumnaIndice = ldrwColumnaIndice("COLUMN_NAME")
                If Not aobjIndice.ColColumnasIndice.Contains(lstrNombreColumnaIndice) Then
                    lstrPropiedades(0) = "Nombre&" & ldrwColumnaIndice("COLUMN_NAME").ToUpper
                    If ldrwColumnaIndice("SORT_ORDER").ToUpper = "A" Then
                        lstrOrdenIndice = "ASC"
                    Else
                        lstrOrdenIndice = "DESC"
                    End If
                    aobjIndice.SAdicioneColumnaIndice(lstrNombreColumnaIndice, lstrOrdenIndice)
                Else
                    Throw New ArchivoXmlPanException("La Columna de indice '" & ldrwColumnaIndice("COLUMN_NAME") &
                            "' del indice '" & vbCrLf & lstrNombreIndice & "' ya existe")
                End If
            Next
        End Sub
        Private Shared Sub SAdicioneRegistros(ByRef aobjTabla As ClsTabla)
            Dim lstrSql As String = "SELECT * FROM " & aobjTabla.StrNombre
            Dim lstrNombreColumna() As String = Nothing
            Dim lstrValorCampo() As String = Nothing
            Dim lstrPropiedades() As String = Nothing
            Dim lobjRegistro As ClsRegistro = Nothing
            Dim lstrKey As String = String.Empty
            Dim i As Integer = 0, lblnNoHayError = False
            Try
                GobjPanDat.SControleProcesoObj(True)
                GobjPanDat.SAbraConexionBd()
                Dim ldrRegistros As MySqlDataReader = ClsPanoramaDat.FdrDataReader(GobjPanDat.CnnConexionBd_App, lstrSql)
                With ldrRegistros
                    ReDim lstrNombreColumna(.FieldCount - 1)
                    ReDim lstrValorCampo(.FieldCount - 1)
                    ReDim lstrPropiedades(.FieldCount - 1)
                    For i = 0 To .FieldCount - 1
                        lstrNombreColumna(i) = .GetName(i)
                    Next
                    Do While .Read
                        lstrKey = String.Empty
                        For i = 0 To .FieldCount - 1
                            lstrValorCampo(i) = CType(.GetValue(i), String)
                            lstrPropiedades(i) = lstrNombreColumna(i) & "&" & lstrValorCampo(i)
                            If IsNumeric(lstrValorCampo(i)) Then
                                lstrValorCampo(i) = Val(lstrValorCampo(i))
                            End If
                            lstrKey &= lstrValorCampo(i)
                        Next
                        lobjRegistro = New ClsRegistro(aobjTabla)
                        lobjRegistro.SAsignePropRegistro(lstrPropiedades)
                        aobjTabla.ColRegistros.Add(lobjRegistro, lstrKey)
                    Loop
                End With
                ldrRegistros.Close()
                lblnNoHayError = True
            Catch ex As MySqlException
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
        End Sub
#End Region
#Region "Acciones de Actualizacion - Procedimientos invalidantes"
        Protected Overrides Sub SCreeBaseDatos()
            Dim lstrNomBD = StrNombreBD
            If lstrNomBD.ToUpper <> HobjBaseDatosXml.StrNombreBD.ToUpper Then
                Dim lstrMens = "El nombre de la Base de Datos definido en el Script no es valido!"
                Throw New ArchivoXmlPanException(lstrMens)
            End If
            Dim lstrPropiedades As String()
            Dim lstrsql As String = "CREATE DATABASE " & lstrNomBD
            If Not String.IsNullOrEmpty(HobjBaseDatosXml.StrCharacterSet) Then
                lstrsql &= " DEFAULT CHARACTER SET " & HobjBaseDatosXml.StrCharacterSet
            End If
            If Not String.IsNullOrEmpty(HobjBaseDatosXml.StrCollationName) Then
                lstrsql &= " DEFAULT COLLATE " & HobjBaseDatosXml.StrCollationName
            End If
            GobjPanDat.SEjecuteSentenciaSql(lstrsql)
            lstrPropiedades = {"Nombre&" & HobjBaseDatosXml.StrNombreBD,
                    "CharacterSet&" & HobjBaseDatosXml.StrCharacterSet,
                    "CollationName&" & HobjBaseDatosXml.StrCollationName,
                    "Version&" & HentVersionBD}
            HobjBaseDatosBD = HobjEstructuraBD.ObjBaseDatos
            HobjBaseDatosBD.SAsignaPropiedadesBD(lstrPropiedades)
            lstrsql = "USE " & HobjBaseDatosXml.StrNombreBD
            GobjPanDat.SEjecuteSentenciaSql(lstrsql)
        End Sub
        Protected Overrides Sub SActualiceBD()
            Dim lstrSql As String = "ALTER DATABASE " & HobjBaseDatosXml.StrNombreBD
            lstrSql &= " CHARACTER SET " & HobjBaseDatosXml.StrCharacterSet
            lstrSql &= " COLLATE " & HobjBaseDatosXml.StrCollationName
            GobjPanDat.SEjecuteSentenciaSql(lstrSql)
            HobjBaseDatosBD.SConviertaCharacterSet(HobjBaseDatosXml.StrCharacterSet,
                        HobjBaseDatosXml.StrCollationName)
        End Sub
        Protected Overrides Sub SVersioneBD(ablnActualiceTamano As Boolean)
            Dim lcolNombreColumnas As New Collection
            Dim lcolDatos As New Collection
            Dim lcolTipoDatos As New Collection
            Dim lcolColumnaIndice As New Collection
            Dim lcolDatosRef As New Collection
            Dim lblnInstalandoApp As Boolean = (GenuTipoInstanciamiento = EnuTipoInstanciamiento.enuInstalacion)
            Dim lstrTablaVer As String
            If GshrIdAplicacion = 999 Then
                lstrTablaVer = "TraVersiones"
            Else
                lstrTablaVer = "PanVersiones"
            End If
            If HenuEstadoBD = EnuEstadoBaseDatos.enuDespoblada OrElse
                    HenuEstadoBD = EnuEstadoBaseDatos.enuNoExiste OrElse lblnInstalandoApp Then
                lcolNombreColumnas.Add("IdAplicacion")
                lcolNombreColumnas.Add("Version")
                lcolNombreColumnas.Add("Tamano")
                lcolDatos.Add(GshrIdAplicacion.ToString)
                lcolDatos.Add(HentVerBDEnXml)
                lcolDatos.Add(0)
                lcolTipoDatos.Add("INTEGER")
                lcolTipoDatos.Add("LONG")
                lcolTipoDatos.Add("DOUBLE")
                SInserteRegistro(lstrTablaVer, lcolNombreColumnas,
                        lcolDatos, lcolTipoDatos)
            Else
                lcolNombreColumnas.Add("Version")
                lcolDatos.Add(HentVerBDEnXml, "Version")
                lcolColumnaIndice.Add("IdAplicacion")
                lcolDatosRef.Add(GshrIdAplicacion, "IdAplicacion")
                GobjPanDat.SActualiceRegistro(lstrTablaVer, lcolNombreColumnas, lcolDatos,
                        lcolColumnaIndice, lcolDatosRef)
            End If
        End Sub
        Protected Overrides Sub SCreeTabla()
            Dim lstrSql As String = FstrConstruyaExpSqlCreeTabla(HobjTablaXml)
            GobjPanDat.SEjecuteSentenciaSql(lstrSql)
            HobjBaseDatosBD.SAdicioneTabla(HobjTablaXml)
            FblnExisteTablaEnBD(HobjTablaXml.StrNombre)
        End Sub
        Protected Overrides Sub SRenombreTabla(astrNombreTablaOri As String,
                    astrNombreTablaDes As String)
            Dim lstrSql As String = "ALTER TABLE " & astrNombreTablaOri & " RENAME TO " & astrNombreTablaDes
            GobjPanDat.SEjecuteSentenciaSql(lstrSql)
            HobjBaseDatosBD.SRenombreTabla(astrNombreTablaOri, astrNombreTablaDes)
        End Sub
        Protected Overrides Sub SVinculeTablas(astrNombreTabla As String,
                astrBaseDatosExterna As String)

        End Sub
        Protected Overrides Sub SVerifiqueCollationTabla(astrCoolationNameOriginalBD As String)
            Dim lstrNombreTabla As String = HobjTablaXml.StrNombre
            Dim lstrCharacterSet As String
            Dim lstrCollationNameTblXml As String
            Dim lstrCollationNameTblBD As String
            Dim lblnColNamheredadaXml As Boolean = False
            Dim lblnColNamheredadaBD As Boolean = False
            Dim lblnActualizarBD As Boolean = False
            Dim lblnActualizarEstructuraBD As Boolean
            If String.IsNullOrEmpty(HobjTablaXml.StrCollationName) Then
                lstrCollationNameTblXml = HobjTablaXml.ObjPadre.StrCollationName
                lblnColNamheredadaXml = True
            Else
                lstrCollationNameTblXml = HobjTablaXml.StrCollationName
            End If
            If String.IsNullOrEmpty(HobjTablaBD.StrCollationName) Then
                lstrCollationNameTblBD = astrCoolationNameOriginalBD
                lblnColNamheredadaBD = True
            Else
                lstrCollationNameTblBD = HobjTablaBD.StrCollationName
                If lstrCollationNameTblBD = astrCoolationNameOriginalBD Then
                    lblnColNamheredadaBD = True
                End If
            End If
            If lblnColNamheredadaBD Then
                If lblnColNamheredadaXml Then
                    lblnActualizarBD = True
                Else
                    If lstrCollationNameTblBD <> lstrCollationNameTblXml Then
                        lblnActualizarBD = True
                    End If
                End If
            Else
                If lstrCollationNameTblBD <> lstrCollationNameTblXml Then
                    lblnActualizarBD = True
                End If
            End If
            lblnActualizarEstructuraBD = (lblnColNamheredadaXml AndAlso (Not lblnColNamheredadaBD)) OrElse
                    (Not lblnColNamheredadaXml)
            If lblnActualizarBD Then
                lstrCharacterSet = lstrCollationNameTblXml.Split("_")(0)
                Dim lstrSql As String = "ALTER TABLE " & lstrNombreTabla
                lstrSql &= " CONVERT TO CHARACTER SET " & lstrCharacterSet
                lstrSql &= " COLLATE " & lstrCollationNameTblXml
                GobjPanDat.SEjecuteSentenciaSql(lstrSql)
            End If
            If lblnActualizarEstructuraBD Then
                If Not lblnColNamheredadaXml Then
                    HobjTablaBD.SConviertaCollationName(lstrCollationNameTblXml)
                Else
                    HobjTablaBD.SConviertaCollationName("")
                End If
            End If
        End Sub
        Protected Overrides Function FblnCambioCollationColumna(aobjColumnaXML As ClsColumna) As Boolean
            If aobjColumnaXML Is Nothing Then
                Throw New ArgumentNullException(NameOf(aobjColumnaXML))
            End If
            Dim lstrCollationNameTblXml As String = aobjColumnaXML.ObjPadre.StrCollationName
            Dim lstrCharacterSetColXml As String = aobjColumnaXML.StrCharacterSet
            Dim lstrCollationNameColXml As String = aobjColumnaXML.StrCollationName
            If String.IsNullOrEmpty(lstrCollationNameTblXml) Then
                lstrCollationNameTblXml = aobjColumnaXML.ObjPadre.ObjPadre.StrCollationName
            End If
            If String.IsNullOrEmpty(lstrCollationNameColXml) Then
                lstrCollationNameColXml = lstrCollationNameTblXml
            End If
            If String.IsNullOrEmpty(lstrCharacterSetColXml) Then
                lstrCharacterSetColXml = lstrCollationNameColXml.Split("_")(0)
            End If
            If lstrCollationNameColXml <> lstrCollationNameTblXml OrElse
                    lstrCharacterSetColXml <> lstrCollationNameTblXml.Split("_")(0) Then
                Return True
            Else
                Return False
            End If
        End Function
        Protected Overrides Function FblnDifierenInd(aobjIndiceXml As ClsIndice) As Boolean
            If aobjIndiceXml Is Nothing Then
                Throw New ArgumentNullException(NameOf(aobjIndiceXml))
            End If
            Dim lblnDifieren As Boolean = False
            Dim lblnEsPrincipal As Boolean = aobjIndiceXml.BlnPrincipal
            Dim lblnEsUnico As Boolean = aobjIndiceXml.BlnUnico
            Dim lcolColumnasIndice As Collection = aobjIndiceXml.ColColumnasIndice
            Dim lobjColumnaIndice As ClsColumnaIndice
            Dim lobjColumnaIndiceBD As ClsColumnaIndice
            Dim lblnExiste As Boolean
            If HobjIndiceBD.BlnPrincipal <> lblnEsPrincipal Then
                lblnDifieren = True
            End If
            If Not lblnDifieren AndAlso HobjIndiceBD.BlnUnico <> lblnEsUnico Then
                lblnDifieren = True
            End If
            If Not lblnDifieren AndAlso HobjIndiceBD.ColColumnasIndice.Count <> lcolColumnasIndice.Count Then
                lblnDifieren = True
            End If
            If Not lblnDifieren Then
                For Each lobjColumnaIndice In lcolColumnasIndice
                    If Not HobjIndiceBD.ColColumnasIndice.Contains(lobjColumnaIndice.StrNombre) Then
                        lblnDifieren = True
                    End If
                Next
            End If
            If Not lblnDifieren Then
                For Each lobjColumnaIndice In lcolColumnasIndice
                    lblnExiste = False
                    For i As UShort = 1 To HobjIndiceBD.ColColumnasIndice.Count
                        lobjColumnaIndiceBD = HobjIndiceBD.ColColumnasIndice(i)
                        If lobjColumnaIndiceBD.StrNombre.ToUpper = lobjColumnaIndice.StrNombre.ToUpper Then
                            lblnExiste = True
                            If lobjColumnaIndiceBD.BlnAscendente <> lobjColumnaIndice.BlnAscendente Then
                                lblnDifieren = True
                            End If
                            Exit For
                        End If
                    Next i
                    If Not lblnExiste Then
                        lblnDifieren = True
                    End If
                Next
            End If
            Return lblnDifieren
        End Function
        Protected Overrides Sub SCreeIndice(aobjIndiceXml As ClsIndice)
            If aobjIndiceXml Is Nothing Then
                Throw New ArgumentNullException(NameOf(aobjIndiceXml))
            End If
            Dim lcolColumnasIndice As Collection = aobjIndiceXml.ColColumnasIndice
            Dim lobjColumnaIndice As ClsColumnaIndice
            Dim lstrNombreIndice As String = aobjIndiceXml.StrNombre
            Dim lblnEsPrimario As Boolean = aobjIndiceXml.BlnPrincipal
            Dim lblnEsUnico As Boolean = aobjIndiceXml.BlnUnico
            Dim lstrNombreColumna As String
            Dim lstrModo As String
            For i = 1 To lcolColumnasIndice.Count
                lobjColumnaIndice = lcolColumnasIndice.Item(i)
                If Not FblnExisteColumnaEnTablaBD(lobjColumnaIndice.StrNombre) Then
                    Throw New ArchivoXmlPanException("La Columna '" & lobjColumnaIndice.StrNombre &
                            "' esta definido en el Indice " & " pero no existe en la Tabla")
                End If
            Next
            Dim lstrSql As String = "ALTER TABLE " & aobjIndiceXml.ObjPadre.StrNombre & " ADD "
            If lblnEsPrimario Then
                lstrSql &= "PRIMARY KEY" & " ("
            ElseIf lblnEsUnico Then
                lstrSql &= "UNIQUE INDEX " & lstrNombreIndice & " ("
            Else
                lstrSql &= "INDEX " & lstrNombreIndice & " ("
            End If
            For i = 1 To lcolColumnasIndice.Count
                lobjColumnaIndice = lcolColumnasIndice.Item(i)
                If lobjColumnaIndice.BlnAscendente Then
                    lstrModo = "ASC"
                Else
                    lstrModo = "DESC"
                End If
                lstrNombreColumna = lobjColumnaIndice.StrNombre
                lstrSql &= lstrNombreColumna & " " & lstrModo & ", "
            Next
            lstrSql = lstrSql.Substring(0, Len(lstrSql) - 2) & ")"
            GobjPanDat.SEjecuteSentenciaSql(lstrSql)
            HobjTablaBD.SAdicioneIndice(aobjIndiceXml)
            FblnExisteIndiceEnBD(lstrNombreIndice)
        End Sub
        Protected Overrides Sub SElimineIndice(astrNombreTabla As String, astrNombreIndice As String)
            If astrNombreTabla Is Nothing Then
                Throw New ArgumentNullException(NameOf(astrNombreTabla))
            End If
            If astrNombreIndice Is Nothing Then
                Throw New ArgumentNullException(NameOf(astrNombreIndice))
            End If
            Dim lstrNombreIndice As String = astrNombreIndice
            Dim lstrSql As String
            If lstrNombreIndice.ToUpper.StartsWith("PK_") OrElse lstrNombreIndice = "PRIMARY" Then
                lstrSql = "ALTER TABLE " & astrNombreTabla & " DROP PRIMARY KEY"
            Else
                lstrSql = "DROP INDEX " & lstrNombreIndice & " ON " & astrNombreTabla
            End If
            If FblnExisteTablaEnBD(astrNombreTabla) Then
                If FblnExisteIndiceEnBD(astrNombreIndice) Then
                    GobjPanDat.SEjecuteSentenciaSql(lstrSql)
                    HobjTablaBD.SRemuevaIndice(lstrNombreIndice)
                End If
            End If
        End Sub
        Protected Overrides Sub SEscribaRegistros()
            Dim lobjRegistro As ClsRegistro
            For Each lobjRegistro In HobjTablaXml.ColRegistros
                SInserteRegistro(HobjTablaXml.StrNombre,
                        lobjRegistro.ColNombresColumna, lobjRegistro.ColDatos, lobjRegistro.ColTipoDatos)
            Next
        End Sub
        Protected Overrides Sub SCreeColumna(aobjColumnaXml As ClsColumna)
            If aobjColumnaXml Is Nothing Then
                Throw New ArgumentNullException(NameOf(aobjColumnaXml))
            End If
            Dim lstrNombreColumnaXml As String = aobjColumnaXml.StrNombre
            Dim lstrSql As String
            If Not FblnExisteColumnaEnTablaBD(lstrNombreColumnaXml) Then
                lstrSql = FstrConstruyaExpSqlCreeColumna(aobjColumnaXml)
                GobjPanDat.SEjecuteSentenciaSql(lstrSql)
            End If
            If Not FblnExisteColumnaEnTablaBD(lstrNombreColumnaXml) Then
                HobjTablaBD.SAdicioneColumna(aobjColumnaXml)
            End If
        End Sub
        Protected Overrides Sub SRenombreColumna(astrNombreTabla As String,
            astrNombreColumnaOriginal As String, astrNombreColumnaNuevo As String)
            Dim lstrSql As String
            If FblnExisteTablaEnBD(astrNombreTabla) Then
                SElimineIndicePKdeColumna(astrNombreColumnaOriginal)
                If Not FblnExisteColumnaEnTablaBD(astrNombreColumnaNuevo) Then
                    If FblnExisteColumnaEnTablaBD(astrNombreColumnaOriginal) Then
                        lstrSql = FstrConstruyaExpSqlRenombrarColumna(HobjColumnaBD, astrNombreColumnaNuevo)
                        GobjPanDat.SEjecuteSentenciaSql(lstrSql)
                        HobjTablaBD.SRenombreColumna(astrNombreColumnaOriginal, astrNombreColumnaNuevo)
                    End If
                End If
            End If
        End Sub
        Protected Overrides Sub SCambieColumna(aobjColumnaXml As ClsColumna)
            If aobjColumnaXml Is Nothing Then
                Throw New ArgumentNullException(NameOf(aobjColumnaXml))
            End If
            Dim lstrSql As String
            lstrSql = FstrConstruyaExpSqlCambiarColumna(aobjColumnaXml)
            GobjPanDat.SEjecuteSentenciaSql(lstrSql)
            HobjTablaBD.SCambieColumna(HobjColumnaBD, aobjColumnaXml)
            If Not FblnExisteColumnaEnTablaBD(aobjColumnaXml.StrNombre) Then
                Throw New PanDatException("La columna no pertenece a la tabla!")
            End If
        End Sub
        Protected Overrides Sub SReemplaseNullPorCero()
            With HobjColumnaBD
                Dim lstrSql As String = "UPDATE " & .ObjPadre.StrNombre & " SET " & .StrNombre & " = '0' WHERE " &
                        .StrNombre & " IS NULL"
                GobjPanDat.SEjecuteSentenciaSql(lstrSql)
            End With
        End Sub
        Protected Overrides Sub SCreeRegistro(aobjRegistroXml As ClsRegistro)
            If aobjRegistroXml Is Nothing Then
                Throw New ArgumentNullException(NameOf(aobjRegistroXml))
            End If
            Dim lstrNombreTabla As String = aobjRegistroXml.ObjPadre.StrNombre
            Dim lcolNombresColumnas As Collection = aobjRegistroXml.ColNombresColumna
            Dim lcolDatos As Collection = aobjRegistroXml.ColDatos
            Dim lcolTipoDatos As Collection = aobjRegistroXml.ColTipoDatos
            Dim lobjRegistroBD As ClsRegistro
            Dim lstrKey As String = String.Empty
            SInserteRegistro(lstrNombreTabla, lcolNombresColumnas, lcolDatos, lcolTipoDatos)
            lobjRegistroBD = aobjRegistroXml
            For Each lobjDato As Object In lobjRegistroBD.ColDatos
                If IsNumeric(lobjDato) Then
                    lobjDato = Val(lobjDato)
                End If
                lstrKey &= lobjDato
            Next
            HobjTablaBD.ColRegistros.Add(lobjRegistroBD, lstrKey)
        End Sub
        Protected Overrides Sub SActualiceReg(aobjRegistroXml As ClsRegistro,
                astrNombreColumnaCambio As String)
            If aobjRegistroXml Is Nothing Then
                Throw New ArgumentNullException(NameOf(aobjRegistroXml))
            End If
            If astrNombreColumnaCambio Is Nothing Then
                Throw New ArgumentNullException(NameOf(astrNombreColumnaCambio))
            End If
            Dim lstrNombreTabla As String = aobjRegistroXml.ObjPadre.StrNombre
            Dim lcolNombresColumnasCambio As New Collection
            Dim lcolDatosNuevos As New Collection
            Dim lcolNombreColumnasRef As New Collection
            Dim lcolDatosRef As New Collection
            lcolNombresColumnasCambio.Add(astrNombreColumnaCambio)
            lcolDatosNuevos.Add(aobjRegistroXml.ColDatos(astrNombreColumnaCambio), astrNombreColumnaCambio)
            For Each lobjColumnaIndice As ClsColumnaIndice In aobjRegistroXml.ObjPadre.FcolColumnasRef
                lcolNombreColumnasRef.Add(lobjColumnaIndice.StrNombre)
            Next
            For Each lstrNombreColumnaIndice As String In lcolNombreColumnasRef
                lcolDatosRef.Add(aobjRegistroXml.ColDatos(lstrNombreColumnaIndice))
            Next
            GobjPanDat.SActualiceRegistro(lstrNombreTabla, lcolNombresColumnasCambio,
                    lcolDatosNuevos, lcolNombreColumnasRef, lcolDatosRef)
        End Sub
        Private Shared Sub SInserteRegistro(astrNombreTabla As String, acolNombresColumnas As Collection,
                            acolDatos As Collection, acolTipoDatos As Collection)
            If acolNombresColumnas.Count <> acolDatos.Count Then
                Throw New ArgumentoInvalidoPanException("Los Parametros pasados son incoherentes.")
            End If
            Dim lstrSql As String = FstrConstruyaExpSqlInsertarReg(astrNombreTabla,
                    acolNombresColumnas, acolDatos, acolTipoDatos)
            GobjPanDat.SEjecuteSentenciaSql(lstrSql)
        End Sub
        Protected Overrides Sub SElimineReg(aobjRegistro As ClsRegistro, astrKey As String)
            If aobjRegistro Is Nothing Then
                Throw New ArgumentNullException(NameOf(aobjRegistro))
            End If
            If aobjRegistro Is Nothing OrElse astrKey Is Nothing Then
                Throw New ArgumentNullException(NameOf(astrKey))
            End If
            Dim lstrNombreTabla = aobjRegistro.ObjPadre.StrNombre
            Dim lcolNombresColumnasRef As New Collection
            Dim lcolDatosRef As New Collection
            For Each lobjColumnaIndice As ClsColumnaIndice In aobjRegistro.ObjPadre.FcolColumnasRef
                lcolNombresColumnasRef.Add(lobjColumnaIndice.StrNombre)
            Next
            For Each lstrNombreColumnaIndice As String In lcolNombresColumnasRef
                lcolDatosRef.Add(aobjRegistro.ColDatos(lstrNombreColumnaIndice))
            Next
            Try
                GobjPanDat.SElimineRegistro(lstrNombreTabla, lcolNombresColumnasRef, lcolDatosRef)
                HobjTablaBD.ColRegistros.Remove(astrKey)
            Catch ex As ProveedorBdPanException
                Throw
            Catch ex As ArgumentOutOfRangeException
                Throw
            Catch ex As Exception
                Throw
            End Try
        End Sub
#End Region
    End Class
End Namespace