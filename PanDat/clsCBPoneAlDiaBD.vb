Namespace ActualizaBd
    Friend MustInherit Class ClsCBPoneAlDiaBD
#Region "Definiciones"
        Protected MustOverride Sub SElimineIndice(astrNombreTabla As String, astrNombreIndice As String)
        Protected MustOverride Sub SCreeIndice(aobjIndiceXml As ClsIndice)
        Protected MustOverride Sub SCreeBaseDatos()
        Protected MustOverride Sub SGenereEstructuraActualBD(astrPrefijoTablas As String)
        Protected MustOverride Sub SVersioneBD(ablnActualiceTamano As Boolean)
        Protected MustOverride Sub SActualiceBD()
        Protected MustOverride Sub SCreeTabla()
        Protected MustOverride Sub SVerifiqueCollationTabla(astrCoolationNameOriginalBD As String)
        Protected MustOverride Function FblnCambioCollationColumna(aobjColumnaXML As ClsColumna) As Boolean
        Protected MustOverride Sub SRenombreTabla(astrNombreTablaOri As String,
                astrNombreTablaDes As String)
        Protected MustOverride Sub SVinculeTablas(astrNombreTabla As String,
                astrBaseDatosExterna As String)
        Protected MustOverride Function FblnDifierenInd(aobjIndiceXml As ClsIndice) As Boolean
        '
        Private McolTablasSobranEnBd As Collection = Nothing
        Private MblnCambioCharacterSetBD As Boolean = False
        Private MstrColNamBDOri As String = String.Empty
        Private MstrArchivoScriptXml As String = String.Empty
#End Region
#Region "Propiedades heredables"
        Protected Property HenuEstadoBD As EnuEstadoBaseDatos = EnuEstadoBaseDatos.None
        Protected Property HobjBaseDatosXml As ClsBaseDatos = Nothing
        Protected Property HobjEstructuraBD As ClsEstructuraBD = Nothing
        Protected Property HentVerBDEnXml As Integer = 0
        Protected Property HentVersionBD As Integer = 0
        Protected Property HstrNombreBDEnXml As String = String.Empty
        Protected Property HobjBaseDatosBD As ClsBaseDatos = Nothing
        Protected Property HobjTablaXml As ClsTabla = Nothing
        Protected Property HobjTablaBD As ClsTabla = Nothing
        Protected Property HobjColumnaBD As ClsColumna = Nothing
        Protected Property HobjIndiceBD As ClsIndice = Nothing
#End Region
        ''' <summary>
        ''' Procedimiento llamado desde la que hereda. Inicia todo el proceso de actualización de la base
        ''' de datos a partir de la estructura contenida en el archivo xml respectivo.
        ''' </summary>
        ''' <remarks></remarks>
        Protected Friend Sub SActualiceBaseDatos(astrArchivoXml As String)
            Dim lstrMens As String = String.Empty
            HobjEstructuraBD = New ClsEstructuraBD
            MstrArchivoScriptXml = astrArchivoXml
            SGenereEstructuraXml()
            If GentVerBDEnProg <> HentVerBDEnXml Then
                lstrMens = "El archivo Xml que contiene la estructura de la base de datos," & vbCrLf &
                        "no coincide con el esperado por el Programa." & vbCrLf &
                        "Por favor informe al ingeniero de Soporte."
            Else
                Dim lstrPrefTabla As String
                lstrPrefTabla = HstrNombreBDEnXml.Substring(0, 3)
                SGenereEstructuraActualBD(lstrPrefTabla)
                Select Case HenuEstadoBD
                    Case EnuEstadoBaseDatos.enuNoExiste
                        SCreeBaseDatos()
                        SSincroniceBD()
                    Case EnuEstadoBaseDatos.enuDespoblada
                        SSincroniceBD()
                    Case EnuEstadoBaseDatos.None
                        If GentVerBDEnProg < HentVersionBD Then
                            lstrMens = "El programa esta desactualizado con relación a la base de datos." & vbCrLf &
                                    "La aplicación no puede continuar."
                        ElseIf GentVerBDEnProg > HentVersionBD Then
                            HenuEstadoBD = EnuEstadoBaseDatos.enuDesActualizada
                            SSincroniceBD()
                        Else
                            HenuEstadoBD = EnuEstadoBaseDatos.enuActualizada
                        End If
                End Select
            End If
            If Not String.IsNullOrEmpty(lstrMens) Then
                Throw New ArgumentoInvalidoPanException(lstrMens)
            End If
        End Sub
        Private Sub SGenereEstructuraXml()
            Dim lobjEstructuraXML As New ClsEstructuraBD
            lobjEstructuraXML.SCreaEstructuraXml(MstrArchivoScriptXml)
            HobjBaseDatosXml = lobjEstructuraXML.ObjBaseDatos
            HentVerBDEnXml = HobjBaseDatosXml.EntVersion
            HstrNombreBDEnXml = HobjBaseDatosXml.StrNombreBD
        End Sub
        Private Sub SSincroniceBD()
            Dim lblnEjecutarComandos As Boolean = Not (HenuEstadoBD =
                    EnuEstadoBaseDatos.enuNoExiste OrElse
                    HenuEstadoBD = EnuEstadoBaseDatos.enuDespoblada) AndAlso
                    Not IsNothing(HobjBaseDatosXml.ColComandos)
            If HenuEstadoBD <> EnuEstadoBaseDatos.enuNoExiste Then
                Dim lstrNomBD = StrNombreBD
                If lstrNomBD = HobjBaseDatosXml.StrNombreBD Then
                    If HobjBaseDatosBD.StrCharacterSet <> HobjBaseDatosXml.StrCharacterSet OrElse
                            HobjBaseDatosBD.StrCollationName <> HobjBaseDatosXml.StrCollationName Then
                        MstrColNamBDOri = HobjBaseDatosBD.StrCollationName
                        MblnCambioCharacterSetBD = True
                        SActualiceBD()
                    End If
                End If
            End If
            ' Ejecuto Comandos de inicio en BD
            If lblnEjecutarComandos Then
                If HobjBaseDatosXml.ColComandos.Count > 0 Then
                    SEjecuteComandosXml(EnuSecuenciaAccion.enuInicio, True)
                End If
            End If
            SVerifiqueTablas() 'Ok
            ' Ejecuto Comandos de Fin en BD
            If lblnEjecutarComandos Then
                If HobjBaseDatosXml.ColComandos.Count > 0 Then
                    If FblnHayComandos(True, False) Then
                        SEjecuteComandosXml(EnuSecuenciaAccion.enuFin, True)
                    End If
                End If
            End If
            SVersioneBD(True)
        End Sub
        Private Function FblnHayComandos(ablnBD As Boolean, ablnInicio As Boolean) As Boolean
            Dim lcolComandos As Collection
            Dim lobjComando As ClsComando
            If ablnBD Then
                lcolComandos = HobjBaseDatosXml.ColComandos
            Else
                lcolComandos = HobjTablaXml.ColComandos
            End If
            For Each lobjComando In lcolComandos
                If ablnInicio AndAlso lobjComando.EnuSecuencia = EnuSecuenciaAccion.enuInicio Then
                    Return True
                ElseIf (Not ablnInicio) AndAlso lobjComando.EnuSecuencia = EnuSecuenciaAccion.enuFin Then
                    Return True
                End If
            Next
            Return False
        End Function
        Protected Friend Sub SEjecuteComandosXml(aenuSecuencia As EnuSecuenciaAccion,
                ablnBD As Boolean)
            If GenuTipoInstanciamiento <> EnuTipoInstanciamiento.enuInstalacion Then
                Dim lcolComandos As Collection
                Dim lobjComando As ClsComando
                If ablnBD Then
                    lcolComandos = HobjBaseDatosXml.ColComandos
                Else
                    lcolComandos = HobjTablaXml.ColComandos
                End If
                For Each lobjComando In lcolComandos
                    If lobjComando.EnuSecuencia = aenuSecuencia Then
                        Select Case lobjComando.EnuTipoComando
                            Case EnuTipoComando.enuPropio
                                SEjecuteComandoPropio(lobjComando, ablnBD)
                            Case EnuTipoComando.enuInstruccionSQL
                                SEjecuteComandoSqlXml(lobjComando, ablnBD)
                            Case EnuTipoComando.enuLlamadaProc
                                SEjecuteComandoProc(lobjComando)
                        End Select
                    End If
                Next
            End If
        End Sub
        ''' <summary>
        ''' Ejecuta un comando "Tipo Propio" contenido en la estructura Xml
        ''' </summary>
        ''' <param name="aobjComando"></param>
        ''' <param name="ablnBD"></param>
        ''' <remarks></remarks>
        Private Sub SEjecuteComandoPropio(aobjComando As ClsComando,
                ablnBD As Boolean)
            If (aobjComando.StrCondicion IsNot Nothing) AndAlso
                    Not String.IsNullOrEmpty(aobjComando.StrCondicion) Then
                If Not FblnCumpleCondicion(aobjComando.StrCondicion) Then
                    Exit Sub
                End If
            End If
            Select Case aobjComando.StrAccion
                Case "RBD" ' "ReleaBaseDatos": Generar de nuevo la estructura de la base de datos
                    SReleaEstructuraBd()
                Case "ETB" ' "EliminaTabla": Elimina Tabla
                    SElimineTabla(aobjComando)
                Case "RTB" ' "RenombraTabla": Renombra Tabla
                    SRenombreTabla(aobjComando)
                Case "ACM"  ' "AdicionaColumna": Adiciona Columna
                    SAdicioneColumna(aobjComando)
                Case "CCM" 'Crea una columna en una tabla a partir de una columna de otra tabla
                    SCopieColumna(aobjComando)
                Case "ECM" ' "EliminaColumna": Elimina Columna
                    SElimineColumna(aobjComando, ablnBD)
                Case "ETB" ' "EliminaTabla": Elimina Tabla
                    SElimineColumna(aobjComando, ablnBD)
                Case "RCM" ' "RenombraColumna": Renombra Columna
                    SRenombreColumna(aobjComando, ablnBD)
                Case "EIN" ' "EliminaIndice": Elimina Indice
                    SElimineIndice(aobjComando)
                Case "VTB"  ' "VerificaTabla": Verifica tabla
                    SVerifiqueTabla(aobjComando)
                Case Else
                    Throw New ArchivoXmlPanException("El comando " & aobjComando.StrAccion & " no existe en Panorama")
            End Select
        End Sub
#Region "Ejecucion de Comandos en Archivo XML"
        Protected Sub SReleaEstructuraBd() 'RBD
            Dim lstrPrefTablas = ClsPanoramaDat.FstrPrefijoTablas(GshrIdAplicacion)
            HobjBaseDatosBD = Nothing
            SGenereEstructuraActualBD(lstrPrefTablas)
        End Sub
        Private Sub SElimineTabla(aobjComando As ClsComando) 'ETB
            If FblnExisteTablaEnBD(aobjComando.StrParametros) Then
                If HobjBaseDatosXml.ColTablas.Contains(aobjComando.StrParametros) Then
                    Throw New ArchivoXmlPanException("La tabla '" & aobjComando.StrParametros &
                            "' no puede ser borrada" & vbCrLf & "por cuanto aún esta incluida en el archivo xml.")
                Else
                    SElimineTabla(aobjComando.StrParametros)
                End If
            End If
        End Sub
        Private Sub SRenombreTabla(aobjComando As ClsComando) 'RTB
            Dim lstrNombreTabla As String = aobjComando.StrParametros.Split(",")(0)
            Dim lstrNuevoNombre As String = aobjComando.StrParametros.Split(",")(1)
            If FblnExisteTablaEnBD(lstrNombreTabla) Then
                If HobjBaseDatosXml.ColTablas.Contains(lstrNombreTabla) Then
                    Throw New ArchivoXmlPanException("La tabla '" & lstrNombreTabla & "' no puede ser borrada" &
                            vbCrLf & "por cuanto aún esta incluida en el archivo xml.")
                Else
                    SRenombreTabla(lstrNombreTabla, lstrNuevoNombre)
                End If
            End If
        End Sub
        Private Sub SAdicioneColumna(aobjComando As ClsComando) 'ACM
            Dim lstrNombreColumna As String = aobjComando.StrParametros
            If Not HobjTablaXml.BlnVinculada Then
                SCreeColumna(HobjTablaXml.ColColumnas(lstrNombreColumna))
            End If
        End Sub
        Private Sub SCopieColumna(aobjComando As ClsComando) 'CCM
            Dim lstrPartes = aobjComando.StrParametros.Split(",")
            Dim lstrNomTablaOrigen = lstrPartes(0)
            Dim lstrNomTablaDest = lstrPartes(1)
            Dim lstrNombreColumna = lstrPartes(2)
            Dim lobjTablaOrigen As ClsTabla = HobjBaseDatosBD.ColTablas(lstrNomTablaOrigen)
            Dim lobjTablaDestino As ClsTabla = HobjBaseDatosBD.ColTablas(lstrNomTablaDest)
            Dim lobjColumna As ClsColumna = lobjTablaOrigen.ColColumnas(lstrNombreColumna)
            Dim lobjColumnaDes = FobjColumnaClonada(lobjColumna, lobjTablaDestino)
            If FblnExisteTablaEnBD(lstrNomTablaDest) Then
                SCreeColumna(lobjColumnaDes)
            End If
        End Sub
        Private Shared Function FobjColumnaClonada(aobjColumna As ClsColumna,
                aobjTablaDestino As ClsTabla) As ClsColumna
            Dim lobjNuevaColumna = New ClsColumna(aobjTablaDestino)
            ' Recontuir parametros nueva columna desde columna origen
            Dim lstrNombre = String.Empty, lstrRequerido = String.Empty
            Dim lstrTipoDato = String.Empty, lstrLongitud = String.Empty
            Dim lstrAutonumerico = String.Empty, lstrComentario = String.Empty
            With aobjColumna
                ' Nombre
                lstrNombre = "Nombre&" & .StrNombre
                ' Requerido
                lstrRequerido = "Requerido&" & IIf(.BlnRequerido, "S", "N")
                ' Tipo Dato
                lstrTipoDato = "TipoDato&" & .StrTipoDatos
                ' Longitud
                lstrLongitud = "Longitud&"
                If String.IsNullOrEmpty(.StrLongitud) Then
                    lstrLongitud = "Longitud&" & 0
                Else
                    lstrLongitud = "Longitud&" & CType(.StrLongitud, Integer)
                End If
                ' AutoNumerico
                lstrAutonumerico = "AutoNumerico&" & IIf(.BlnAutoNumerico, "S", "N")
                'Comentario
                lstrComentario = "Comentario&" & .StrComentario
            End With
            Dim lstrPropiedades = {lstrNombre, lstrRequerido, lstrTipoDato, lstrLongitud, lstrAutonumerico,
                                  lstrComentario}
            lobjNuevaColumna.SAsignePropColumna(lstrPropiedades)
            Return lobjNuevaColumna
        End Function
        Private Sub SElimineColumna(aobjComando As ClsComando, ablnBD As Boolean) 'ECM
            Dim lstrNombreColumna As String = aobjComando.StrParametros
            If ablnBD Then
                For Each lobjTabla In HobjBaseDatosXml.ColTablas
                    If Not lobjTabla.blnVinculada Then
                        SElimineColumna(lobjTabla.strNombre, lstrNombreColumna)
                    End If
                Next
            Else
                If Not HobjTablaXml.BlnVinculada Then
                    SElimineColumna(HobjTablaXml.StrNombre, lstrNombreColumna)
                End If
            End If
        End Sub
        Private Sub SRenombreColumna(aobjComando As ClsComando, ablnBD As Boolean) 'RCM
            Dim lstrNombreColumna As String = aobjComando.StrParametros.Split(",")(0)
            Dim lstrNuevoNombre As String = aobjComando.StrParametros.Split(",")(1)
            If ablnBD Then
                For Each lobjTabla In HobjBaseDatosBD.ColTablas
                    If Not lobjTabla.blnVinculada Then
                        If lobjTabla.colColumnas.Contains(lstrNombreColumna) Then
                            SRenombreColumna(lobjTabla.strNombre,
                                    lstrNombreColumna, lstrNuevoNombre)
                        End If
                    End If
                Next
            Else
                If Not HobjTablaXml.BlnVinculada Then
                    If FblnExisteColumnaEnTablaBD(lstrNombreColumna) Then
                        SRenombreColumna(HobjTablaXml.StrNombre,
                                lstrNombreColumna, lstrNuevoNombre)
                    End If
                End If
            End If
        End Sub
        Private Sub SElimineIndice(aobjComando As ClsComando) 'EIN
            If Not HobjTablaXml.BlnVinculada Then
                SElimineIndice(HobjTablaXml.StrNombre, aobjComando.StrParametros)
            End If
        End Sub
        Private Sub SVerifiqueTabla(aobjComando As ClsComando) 'VTB
            Dim lstrTabla As String = aobjComando.StrParametros
            If HobjBaseDatosXml.ColTablas.Contains(lstrTabla) Then
                HobjTablaXml = HobjBaseDatosXml.ColTablas(lstrTabla)
                If FblnExisteTablaEnBD(lstrTabla) Then
                    SVerifiqueTabla()
                Else
                    SCreeTabla()
                End If
            End If
        End Sub
        Private Sub SEjecuteComandoProc(ByRef aobjComando As ClsComando)
            Select Case aobjComando.StrParametros
                Case "SActuBD_V130"
                    If FblnCumpleCondicion(aobjComando.StrCondicion) Then
                        CallByName(Me, "sActuBD_V130", CallType.Method)
                    End If
                Case "SActuBD_V136"
                    If FblnCumpleCondicion(aobjComando.StrCondicion) Then
                        CallByName(Me, "sActuBD_V136", CallType.Method)
                    End If
                Case "SActuBD_V244"
                    If FblnCumpleCondicion(aobjComando.StrCondicion) Then
                        CallByName(Me, "sActuBD_V244", CallType.Method)
                    End If
                Case Else
                    Throw New ErrorInesperadoPanDatException("Parametro en Comando Proc no esperado!")
            End Select
        End Sub
#End Region
        ''' <summary>
        ''' Ejecuta un comando "Tipo Sql" contenido en la estructura Xml
        ''' </summary>
        ''' <param name="aobjComando"></param>
        ''' <remarks></remarks>
        Private Sub SEjecuteComandoSqlXml(aobjComando As ClsComando,
                Optional ablnDB As Boolean = False)
            Dim lstrSql As String
            If FblnCumpleCondicion(aobjComando.StrCondicion) Then
                If aobjComando.StrAccion = "FSH" Then
                    If ablnDB Then
                        lstrSql = "FLUSH TABLES"
                    Else
                        lstrSql = "FLUSH TABLE " & HobjTablaXml.StrNombre
                    End If
                    GobjPanDat.SEjecuteSentenciaSql(lstrSql)
                End If
                If Not String.IsNullOrEmpty(aobjComando.StrParametros) Then
                    lstrSql = aobjComando.StrParametros
                    GobjPanDat.SEjecuteSentenciaSql(lstrSql)
                End If
            End If
        End Sub
        Protected Sub SVerifiqueTablas()
            Dim lblnEsVinculada As Boolean
            Dim lstrBaseDatosVin As String
            Dim i As Integer = 0
            STablasSobranEnBD()
            If Not IsNothing(McolTablasSobranEnBd) AndAlso McolTablasSobranEnBd.Count > 0 Then
                For Each lstrTabla As String In McolTablasSobranEnBd
                    SElimineTabla(lstrTabla)
                Next
            End If
            For i = 1 To HobjBaseDatosXml.ColTablas.Count
                HobjTablaXml = HobjBaseDatosXml.ColTablas.Item(i)
                lblnEsVinculada = HobjTablaXml.BlnVinculada
                lstrBaseDatosVin = HobjTablaXml.StrBDVinculada
                If FblnExisteTablaEnBD(HobjTablaXml.StrNombre) Then
                    If Not lblnEsVinculada Then
                        SVerifiqueTabla()
                    End If
                Else
                    If Not lblnEsVinculada Then
                        ' Se crea la tabla
                        SCreeTabla()
                        If HobjTablaXml.ColRegistros.Count > 0 Then
                            SEscribaRegistros()
                        End If
                        If HobjTablaXml.ColComandos.Count > 0 Then
                            SEjecuteComandosXml(EnuSecuenciaAccion.enuFin, False)
                        End If
                    Else
                        ' Manejo de tabla vinculada
                        SVinculeTablas(HobjTablaXml.StrNombre, lstrBaseDatosVin)
                    End If
                End If
            Next
        End Sub
        Private Sub STablasSobranEnBD()
            If IsNothing(McolTablasSobranEnBd) Then
                McolTablasSobranEnBd = New Collection
            Else
                McolTablasSobranEnBd.Clear()
            End If
            For Each lobjTablaBD As ClsTabla In HobjBaseDatosBD.ColTablas
                If Not HobjBaseDatosXml.ColTablas.Contains(lobjTablaBD.StrNombre) Then
                    McolTablasSobranEnBd.Add(lobjTablaBD.StrNombre)
                End If
            Next
        End Sub
        Private Sub SElimineTabla(astrNombreTabla As String)
            Dim lstrSql As String = FstrConstruyaExpSqlElimineTabla(astrNombreTabla)
            If FblnExisteTablaEnBD(astrNombreTabla) Then
                GobjPanDat.SEjecuteSentenciaSql(lstrSql)
                HobjBaseDatosBD.SRemuevaTabla(astrNombreTabla)
            End If
        End Sub
        Protected Overridable Sub SEscribaRegistros()
            '
        End Sub
        ''' <summary>
        ''' Verifica que la definición de la tabla en el archivo XML corresponda a la tabla en la BD, llevando a cabo los cambios
        ''' necesarios para que sean iguales. 
        ''' </summary>
        ''' <remarks>Si el nombre de la tabla contiene el string "tbl (no case sensitive)" verifica que los registros coincidan;
        ''' si no coinciden, hace los cambios necesarios para que coincidan.</remarks>
        Private Sub SVerifiqueTabla()
            Dim lstrNombreTablaXml As String = HobjTablaXml.StrNombre
            Dim lcolColumnasXml As Collection = HobjTablaXml.ColColumnas
            Dim lcolIndicesXml As Collection = HobjTablaXml.ColIndices
            Dim lcolRegistrosXml As Collection = HobjTablaXml.ColRegistros
            Dim lobjColumnaXml As ClsColumna
            Dim lcolIndicesBD As Collection = HobjTablaBD.ColIndices
            Dim lobjIndiceXml As ClsIndice
            Dim lobjColumnaBD As ClsColumna
            Dim lblnEjecutarComandos As Boolean = Not (HenuEstadoBD =
                    EnuEstadoBaseDatos.enuNoExiste OrElse
                    HenuEstadoBD = EnuEstadoBaseDatos.enuDespoblada) AndAlso
                    Not IsNothing(HobjTablaXml.ColComandos)
            If lblnEjecutarComandos Then
                If HobjTablaXml.ColComandos.Count > 0 Then
                    SEjecuteComandosXml(EnuSecuenciaAccion.enuInicio, False)
                End If
            End If
            If MblnCambioCharacterSetBD Then
                SVerifiqueCollationTabla(MstrColNamBDOri)
            End If
            ' Verifico Columnas
            For Each lobjColumnaXml In lcolColumnasXml
                If Not FblnExisteColumnaEnTablaBD(lobjColumnaXml.StrNombre) Then
                    SCreeColumna(lobjColumnaXml)
                End If
            Next
            For Each lobjColumnaBD In HobjTablaBD.ColColumnas
                HobjColumnaBD = lobjColumnaBD
                If Not lcolColumnasXml.Contains(lobjColumnaBD.StrNombre) Then
                    SElimineColumna(lstrNombreTablaXml, lobjColumnaBD.StrNombre) ' OK
                End If
            Next
            For i = 1 To lcolColumnasXml.Count
                lobjColumnaXml = lcolColumnasXml.Item(i)
                SVerifiqueColumna(lobjColumnaXml) ' OK
            Next
            ' Verifico Indices: elimino indices no deben existir
            For Each lobjIndiceBD As ClsIndice In lcolIndicesBD
                Dim lstrNombreIndice As String
                lstrNombreIndice = lobjIndiceBD.StrNombre
                If Not lcolIndicesXml.Contains(lstrNombreIndice) Then
                    SElimineIndice(lobjIndiceBD.ObjPadre.StrNombre, lobjIndiceBD.StrNombre)
                End If
            Next
            ' Verifico Indices: Creo indices no existen
            For i = 1 To lcolIndicesXml.Count
                lobjIndiceXml = lcolIndicesXml.Item(i)
                If Not FblnExisteIndiceEnBD(lobjIndiceXml.StrNombre) Then
                    SCreeIndice(lobjIndiceXml)
                End If
            Next
            For i = 1 To lcolIndicesXml.Count
                lobjIndiceXml = lcolIndicesXml.Item(i)
                SVerifiqueIndice(lobjIndiceXml)
            Next
            If lstrNombreTablaXml.ToUpper.Contains("TBL") AndAlso lcolRegistrosXml.Count > 0 Then
                SVerifiqueRegistros()
            End If
            If lblnEjecutarComandos Then
                If HobjTablaXml.ColComandos.Count > 0 Then
                    SEjecuteComandosXml(EnuSecuenciaAccion.enuFin, False)
                End If
            End If
        End Sub
        Protected Overridable Sub SCreeColumna(aobjColumnaXml As ClsColumna)
            '
        End Sub
        Private Sub SElimineColumna(astrNombreTabla As String, astrNombreColumna As String)
            Dim lstrSql As String = "ALTER TABLE " & astrNombreTabla & " DROP COLUMN " &
                    astrNombreColumna & ";"
            Dim lstrNombreIndice As String = String.Empty
            If FblnExisteTablaEnBD(astrNombreTabla) Then
                Do While FblnExisteColumnaEnIndiceBD(astrNombreColumna, lstrNombreIndice)
                    SElimineIndice(astrNombreTabla, lstrNombreIndice)
                Loop
                GobjPanDat.SEjecuteSentenciaSql(lstrSql)
                HobjTablaBD.SRemuevaColumna(astrNombreColumna)
            End If
        End Sub
        Private Sub SVerifiqueColumna(aobjColumnaXml As ClsColumna)
            Dim lstrNombreColumna As String = aobjColumnaXml.StrNombre
            Dim lstrValorDefectoXml As String = String.Empty
            Dim lstrValorDefectoBD As String = String.Empty
            Dim lblnCambiarColumna As Boolean = False
            If Not FblnExisteColumnaEnTablaBD(lstrNombreColumna) Then
                SCreeColumna(aobjColumnaXml)
            Else
                If Not IsNothing(aobjColumnaXml.ObjValorDefault) Then
                    If Not FblnValorEsValorProDefecto(aobjColumnaXml.ObjValorDefault.ToString) Then
                        lstrValorDefectoXml = aobjColumnaXml.ObjValorDefault.ToString
                    End If
                End If
                If Not IsNothing(HobjColumnaBD.ObjValorDefault) Then
                    lstrValorDefectoBD = HobjColumnaBD.ObjValorDefault.ToString
                End If
                If aobjColumnaXml.StrTipoDatos.ToUpper.Contains("STRING") Then
                    lblnCambiarColumna = FblnCambioCollationColumna(aobjColumnaXml)
                End If
                If Not lblnCambiarColumna Then
                    If aobjColumnaXml.StrTipoDatos.ToUpper <> HobjColumnaBD.StrTipoDatos.ToUpper OrElse
                            aobjColumnaXml.BlnRequerido <> HobjColumnaBD.BlnRequerido OrElse
                            aobjColumnaXml.StrComentario.ToUpper <> HobjColumnaBD.StrComentario.ToUpper OrElse
                            lstrValorDefectoXml <> lstrValorDefectoBD OrElse aobjColumnaXml.StrLongitud <>
                            HobjColumnaBD.StrLongitud Then
                        lblnCambiarColumna = True
                    End If
                End If
                If lblnCambiarColumna Then
                    If aobjColumnaXml.FblnEsNumerico AndAlso HobjColumnaBD.FblnEsCadena Then
                        SReemplaseNullPorCero()
                    End If
                    SCambieColumna(aobjColumnaXml)
                End If
            End If
        End Sub
        Protected Overridable Sub SReemplaseNullPorCero()
            '
        End Sub
        Protected Overridable Sub SRenombreColumna(astrNombreTabla As String,
                astrNombreColumnaOriginal As String, astrNombreColumnaNuevo As String)
            '
        End Sub
        ''' <summary>
        ''' Verifica si el nombre de la columna pasada en el argumento forma parte del indice Principal; si esto 
        ''' sucede elimina el indice
        ''' </summary>
        ''' <param name="astrNombreColumna">Nombre de la columna a verificar</param>
        ''' <remarks></remarks>
        Protected Sub SElimineIndicePKdeColumna(astrNombreColumna)
            Dim lblnSalir As Boolean = False
            If Not IsNothing(HobjTablaBD) Then
                For Each lobjIndice As ClsIndice In HobjTablaBD.ColIndices
                    If lobjIndice.StrNombre.Substring(0, 2).ToUpper = "PK" Then
                        Dim lobjColInd As ClsColumnaIndice = Nothing
                        For i As Byte = 1 To lobjIndice.ColColumnasIndice.Count
                            lobjColInd = lobjIndice.ColColumnasIndice(i)
                            If lobjColInd.StrNombre.ToUpper = astrNombreColumna.ToUpper Then
                                SElimineIndice(HobjTablaBD.StrNombre, lobjIndice.StrNombre)
                                lblnSalir = True
                            End If
                        Next
                        If lblnSalir Then
                            Exit For
                        End If
                    End If
                Next
            End If
        End Sub
        Protected Overridable Sub SCambieColumna(aobjColumnaXml As ClsColumna)
            '
        End Sub
        Private Sub SVerifiqueIndice(aobjIndiceXml As ClsIndice)
            If FblnExisteIndiceEnBD(aobjIndiceXml.StrNombre) Then
                If FblnEsIndiceDiferente(aobjIndiceXml) Then
                    SElimineIndice(aobjIndiceXml.ObjPadre.StrNombre, aobjIndiceXml.StrNombre)
                    SCreeIndice(aobjIndiceXml)
                End If
            Else
                SCreeIndice(aobjIndiceXml)
            End If
        End Sub
        ''' <summary>
        ''' Verifica si existe una tabla en el catalogo que representa la BD física
        ''' </summary>
        ''' <param name="astrNombreTabla">Nombre de la tabla a buscar</param>
        ''' <returns>Boolean "True" si existe "False" si no existe</returns>
        ''' <remarks></remarks>
        Protected Function FblnExisteTablaEnBD(astrNombreTabla As String) As Boolean
            If HobjBaseDatosBD.ColTablas.Contains(astrNombreTabla) Then
                HobjTablaBD = HobjBaseDatosBD.ColTablas(astrNombreTabla)
                Return True
            Else
                Return False
            End If
        End Function
        ' Verifica que exsista un Columna en la tabla hobjTablaBD de la base de datos. 
        ' Si existe asigna la variable heredable hobjColumnaBD
        Protected Function FblnExisteColumnaEnTablaBD(astrNombreColumna As String) As Boolean
            If HobjTablaBD.ColColumnas.Contains(astrNombreColumna) Then
                HobjColumnaBD = HobjTablaBD.ColColumnas(astrNombreColumna)
                Return True
            Else
                Return False
            End If
        End Function
        Protected Function FblnExisteIndiceEnBD(astrNombreIndice As String) As Boolean
            Dim lblnExiste As Boolean = False
            If HobjTablaBD.ColIndices.Contains(astrNombreIndice) Then
                lblnExiste = True
                HobjIndiceBD = HobjTablaBD.ColIndices(astrNombreIndice)
            Else
                HobjIndiceBD = Nothing
            End If
            Return lblnExiste
        End Function
        Private Function FblnExisteColumnaEnIndiceBD(astrNombreColumna As String,
                ByRef astrNombreIndice As String) As Boolean
            For Each lobjIndice As ClsIndice In HobjTablaBD.ColIndices
                If lobjIndice.ColColumnasIndice.Contains(astrNombreColumna) Then
                    astrNombreIndice = lobjIndice.StrNombre
                    Return True
                End If
            Next
            Return False
        End Function
        Private Function FblnEsIndiceDiferente(aobjIndiceXml As ClsIndice) As Boolean
            Dim lblnDifieren As Boolean = FblnDifierenInd(aobjIndiceXml)
            Return lblnDifieren
        End Function
        Private Sub SVerifiqueRegistros()
            Dim lcolRegistrosXml As Collection = HobjTablaXml.ColRegistros
            Dim lcolRegistrosBD As Collection
            Dim lobjRegistroXml As ClsRegistro
            Dim lobjRegistroBD As ClsRegistro
            Dim lobjDato As Object
            Dim i As UShort = 0
            Dim lstrKey As String
            ' suprimir los registros que existen en la base de datos pero no existen en el archivo Xml
            lcolRegistrosBD = HobjTablaBD.ColRegistros
            For Each lobjRegistroBD In lcolRegistrosBD
                lstrKey = String.Empty
                For i = 1 To lobjRegistroBD.ColDatos.Count
                    lobjDato = lobjRegistroBD.ColDatos(i)
                    If IsNumeric(lobjDato) Then
                        lobjDato = CType(Val(lobjDato), Double)
                    End If
                    lstrKey &= lobjDato
                Next
                If Not lcolRegistrosXml.Contains(lstrKey) Then
                    SElimineReg(lobjRegistroBD, lstrKey)
                End If
            Next
            For Each lobjRegistroXml In lcolRegistrosXml
                lstrKey = String.Empty
                For i = 1 To lobjRegistroXml.ColDatos.Count
                    lobjDato = lobjRegistroXml.ColDatos(i)
                    If IsNumeric(lobjDato) Then
                        lobjDato = Val(lobjDato)
                    End If
                    lstrKey &= lobjDato
                Next
                If Not lcolRegistrosBD.Contains(lstrKey) Then
                    SCreeRegistro(lobjRegistroXml)
                End If
            Next
            For Each lobjRegistroXml In lcolRegistrosXml
                lstrKey = String.Empty
                For i = 1 To lobjRegistroXml.ColDatos.Count
                    lobjDato = lobjRegistroXml.ColDatos(i)
                    If IsNumeric(lobjDato) Then
                        lobjDato = Val(lobjDato)
                    End If
                    lstrKey &= lobjDato
                Next
                lobjRegistroBD = lcolRegistrosBD(lstrKey)
                SVerifiqueRegistro(lobjRegistroXml, lobjRegistroBD)
            Next
        End Sub
        Private Sub SVerifiqueRegistro(ByRef aobjRegistroXml As ClsRegistro,
                ByRef aobjRegistroBD As ClsRegistro)
            For Each lstrNombreColumna As String In aobjRegistroXml.ColNombresColumna
                If aobjRegistroXml.ColDatos(lstrNombreColumna) <> aobjRegistroBD.ColDatos(lstrNombreColumna) Then
                    SActualiceReg(aobjRegistroXml, lstrNombreColumna)
                    aobjRegistroBD.SAsigneValorCampo(lstrNombreColumna, aobjRegistroXml.ColDatos(lstrNombreColumna))
                End If
            Next
        End Sub
        Protected Overridable Sub SCreeRegistro(aobjRegistroXml As ClsRegistro)
            '
        End Sub
        Protected Overridable Sub SActualiceReg(aobjRegistroXml As ClsRegistro,
                astrNombreColumnaCambio As String)
            '
        End Sub
        Protected Overridable Sub SElimineReg(aobjRegistro As ClsRegistro, astrKey As String)
            '
        End Sub
        Private Function FblnCumpleCondicion(astrConicion As String) As Boolean
            Dim lstrVersion As String
            If astrConicion.StartsWith("VBD") Then
                If astrConicion.Contains("<") Then
                    lstrVersion = astrConicion.Substring(astrConicion.IndexOf("<") + 1)
                    If IsNumeric(lstrVersion) Then
                        If HentVersionBD < CType(lstrVersion, Integer) Then
                            Return True
                        End If
                    End If
                End If
            ElseIf String.IsNullOrEmpty(astrConicion) Then
                Return True
            End If
            Return False
        End Function
        ''' <summary>
        ''' Indica si la cadena "astrValor" representa un valor por defecto de los usados por defecto
        ''' </summary>
        ''' <param name="astrValor">Valor a comparar con los Vlores por defecto usados por defecto</param>
        ''' <returns>Buleano</returns>
        ''' <remarks></remarks>
        Protected Shared Function FblnValorEsValorProDefecto(astrValor As String) As Boolean
            Dim lstrValoresPorDefecto As String() = {"0", "0.00", "*", "False", "1900-01-01", "1900-01-01 00:00:00"}
            Return lstrValoresPorDefecto.Contains(astrValor)
        End Function
#Region "Procedimientos y Funciones ejecutados desde Comando del Script de BD"
        ' Cada Factura de estado se actualiza con el agrupador de servicios de la factura viva
        Friend Shared Sub SCreeDocumentosCentrUtil(ashrIdCarpeta As Short, ashrIdCentroUtil As Short)
            Dim lcolNombreCampos As New Collection
            Dim lcolDatos As New Collection
            Dim ldrwCentroUtilOrion = FdtbCentroUtilOrionCop(ashrIdCarpeta, ashrIdCentroUtil).Rows(0)
            lcolNombreCampos.Add("IdCarpeta")
            lcolNombreCampos.Add("IdCentroUtil")
            lcolNombreCampos.Add("IdDocumento")
            lcolNombreCampos.Add("NombreDocumento")
            lcolNombreCampos.Add("NumeracionInicialDoc")
            lcolNombreCampos.Add("PrefijoDocumento")
            lcolNombreCampos.Add("TipoDocumentoCont")
            For i = 1 To 8
                lcolDatos.Add(ashrIdCarpeta, "IdCarpeta")
                lcolDatos.Add(ashrIdCentroUtil, "IdCentroUtil")
                lcolDatos.Add(i, "IdDocumento")
                lcolDatos.Add(FstrNombreDoc(i), "NombreDocumento")
                lcolDatos.Add(FentIdDocumentoIni(ldrwCentroUtilOrion, i), "NumeracionInicialDoc")
                lcolDatos.Add(FstrPrefijoDoc(ldrwCentroUtilOrion, i), "PrefijoDocumento")
                lcolDatos.Add("''", "TipoDocumentoCont")
                GobjPanDat.SInserteRegistro("OriDocumentosContables",
                        lcolNombreCampos, lcolDatos)
                lcolDatos.Clear()
            Next i
        End Sub
        Private Shared Function FstrNombreDoc(ashrIdDocumento As Short) As String
            Dim lstrNombreDoc = String.Empty
            Select Case ashrIdDocumento
                Case 1
                    lstrNombreDoc = "FACTURA DE VENTA"
                Case 2
                    lstrNombreDoc = "RECIBO DE CAJA"
                Case 3
                    lstrNombreDoc = "NOTA APLICACION ANTICIPO"
                Case 4
                    lstrNombreDoc = "NOTA DEBITO INTERESES DE MORA"
                Case 5
                    lstrNombreDoc = "NOTA CREDITO"
                Case 6
                    lstrNombreDoc = "NOTA REINTEGRO ANTICIPO"
                Case 7
                    lstrNombreDoc = "NOTA REVERSION CREDITO"
                Case 8
                    lstrNombreDoc = "COMPROBANTE INTERFAZ CONTABLE"
            End Select
            Return lstrNombreDoc
        End Function
        Private Shared Function FentIdDocumentoIni(adrwCentroUtilOrion As DataRow,
                    ashrIdDocumento As Short) As Integer
            Dim lentIdDoc = 0
            Select Case ashrIdDocumento
                Case 1
                    lentIdDoc = CType(adrwCentroUtilOrion("NumeracionInicialFAC"), Integer)
                Case 2
                    lentIdDoc = CType(adrwCentroUtilOrion("NumeracionInicialREC"), Integer)
                Case 3
                    lentIdDoc = CType(adrwCentroUtilOrion("NumeracionInicialNCO"), Integer)
                Case 4
                    lentIdDoc = CType(adrwCentroUtilOrion("NumeracionInicialNDB"), Integer)
                Case 5
                    lentIdDoc = CType(adrwCentroUtilOrion("NumeracionInicialNCR"), Integer)
                Case 6
                    lentIdDoc = 0
                Case 7
                    lentIdDoc = 0
                Case 8
                    lentIdDoc = 0
            End Select
            Return lentIdDoc
        End Function
        Private Shared Function FstrPrefijoDoc(adrwCentroUtilOrion As DataRow,
                    ashrIdDocumento As Short) As String
            Dim lstrPrefDoc = String.Empty
            Select Case ashrIdDocumento
                Case 1
                    lstrPrefDoc = CType(adrwCentroUtilOrion("PrefijoFactura"), String)
                Case 2
                    lstrPrefDoc = CType(adrwCentroUtilOrion("PrefijoRecibo"), String)
                Case 3
                    lstrPrefDoc = CType(adrwCentroUtilOrion("PrefijoNotaCon"), String)
                Case 4
                    lstrPrefDoc = CType(adrwCentroUtilOrion("PrefijoNotaDb"), String)
                Case 5
                    lstrPrefDoc = CType(adrwCentroUtilOrion("PrefijoNotaCr"), String)
                Case 6
                    lstrPrefDoc = CType(adrwCentroUtilOrion("PrefijoNotaDevAnt"), String)
                Case 7
                    lstrPrefDoc = CType(adrwCentroUtilOrion("PrefijoNotaReversaRC"), String)
                Case 8
                    lstrPrefDoc = String.Empty
            End Select
            Return lstrPrefDoc
        End Function
        Private Shared Function FdtbCentroUtilOrionCop(ashrIdCarpeta As Short,
                    ashridCentroUtil As Short) As DataTable
            Dim ldtbCentroUtilOrionCop As DataTable = Nothing
            Dim lcolNombreTablas As New Collection
            Dim lcolCamposTabla As New Collection
            Dim lcolFiltros As New Collection
            lcolNombreTablas.Add("OriCentrosUtilidadOriCop")
            lcolCamposTabla.Add({"NumeracionInicialFAC", "NumeracionInicialNCO", "NumeracionInicialNCR",
                                 "NumeracionInicialNDB", "NumeracionInicialREC", "PrefijoFactura",
                                 "PrefijoNotaCon", "PrefijoNotaCr", "PrefijoNotaDb",
                                 "PrefijoNotaDevAnt", "PrefijoNotaReversaRC", "PrefijoRecibo"})
            lcolFiltros.Add("IdCarpeta = " & ashrIdCarpeta & " AND IdCentroUtil = " & ashridCentroUtil)
            Using ldsCentrosUtilOrion As New DataSet
                GobjPanDat.SdsDataSet(ldsCentrosUtilOrion, lcolNombreTablas, lcolCamposTabla,
                        Nothing, lcolFiltros)
                ldtbCentroUtilOrionCop = ldsCentrosUtilOrion.Tables(0)
            End Using
            Return ldtbCentroUtilOrionCop
        End Function
#Region "Actualiza a la Versión 130 de la base de datos"
        Public Shared Sub SActuBD_V130()
            'Obetivo: pasar las cuentas descuentos de capital e intereses de mora Cr 
            ' de la Copropiedad a cada uno de los servicios
            Dim lblnOk = False
            Try
                GobjPanDat.SControleProcesoObj(True)
                GobjPanDat.SInicialiceTransaccion()
                ' Cagar tabla de Centros de Utilidad
                Dim lstrCamposSelect() = {"IdCarpeta", "IdCentroUtil"}
                Dim lstrIndice(,) = {{"IdCarpeta", "ASC"}, {"IdCentroUtil", "ASC"}}
                Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect("PanCentrosUtilidad", lstrCamposSelect,
                        lstrIndice, "", Array.Empty(Of String))
                Dim ldtbCentroUtil As DataTable = Nothing
                Using ldsCentrosUtil As New DataSet
                    GobjPanDat.SdsDataSet(ldsCentrosUtil, lstrSql)
                    ldtbCentroUtil = ldsCentrosUtil.Tables(0)
                End Using
                ' Recorro todos los centros de utilidad
                Dim lshrIdCarpta = 0S, lshrIdCentroUtil = 0S
                For Each ldrwCenUtil As DataRow In ldtbCentroUtil.Rows
                    lshrIdCarpta = CType(ldrwCenUtil("IdCarpeta"), Short)
                    lshrIdCentroUtil = CType(ldrwCenUtil("IdCentroUtil"), Short)
                    ' Actualizo la table Servicios
                    Dim lstrCtaDsctoCap = FstrCtaDsctoCap(lshrIdCarpta, lshrIdCentroUtil)
                    Dim lstrCtaMoraCr = FstrCtaMoraCr(lshrIdCarpta, lshrIdCentroUtil)
                    SAcualiceServicios(lshrIdCarpta, lshrIdCentroUtil, lstrCtaDsctoCap, lstrCtaMoraCr)
                Next
                lblnOk = True
            Catch ex As Exception
                Throw
            Finally
                If Not lblnOk Then
                    GobjPanDat.SAborteTransaccion()
                    GobjPanDat.SControleProcesoObj(False, True)
                Else
                    GobjPanDat.SConfirmeTransaccion()
                    GobjPanDat.SControleProcesoObj(False)
                End If
            End Try
        End Sub
        Private Shared Function FstrCtaDsctoCap(ashrIdCarpeta As Short,
                ashrIdCentroUtil As Short) As String
            Dim lstrTabla = "OriCentrosUtilidadOriCop"
            Dim lstrCamposSelect = {"IdCtaDctoCapital"}
            Dim lstrFiltro = "IdCarpeta = " & ashrIdCarpeta.ToString +
                    " AND IdCentroUtil = " & ashrIdCentroUtil.ToString
            Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCamposSelect,
                    {{"", ""}}, lstrFiltro, Array.Empty(Of String))
            Dim ldtbRes As DataTable = Nothing
            Using ldsRes As New DataSet
                GobjPanDat.SdsDataSet(ldsRes, lstrSql)
                ldtbRes = ldsRes.Tables(0)
            End Using
            Dim lstrIdCta = ldtbRes.Rows(0)(0)
            Return lstrIdCta
        End Function
        Private Shared Function FstrCtaMoraCr(ashrIdCarpeta As Short,
                ashrIdCentroUtil As Short) As String
            Dim lstrTabla = "OriCentrosUtilidadOriCop"
            Dim lstrCamposSelect = {"IdCtaMoraCr"}
            Dim lstrFiltro = "IdCarpeta = " & ashrIdCarpeta.ToString +
                    " AND IdCentroUtil = " & ashrIdCentroUtil.ToString
            Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCamposSelect,
                    {{"", ""}}, lstrFiltro, Array.Empty(Of String))
            Dim ldtbRes As DataTable = Nothing
            Using ldsRes As New DataSet
                GobjPanDat.SdsDataSet(ldsRes, lstrSql)
                ldtbRes = ldsRes.Tables(0)
            End Using
            Dim lstrIdCta = ldtbRes.Rows(0)(0)
            Return lstrIdCta
        End Function
        Private Shared Sub SAcualiceServicios(ashrIdCarpeta As Short,
                        ashrIdCentroUtil As Short, astrIdCtaDscto As String,
                        astrIdCtaMoraCr As String)
            Dim lstrTabla = "Servicios"
            Dim lcolNomCampos As New Collection
            Dim lcolDatosNuevos As New Collection
            Dim lcolCamposRef As New Collection
            Dim lcolDatosRef As New Collection
            lcolNomCampos.Add("IdCtaDctoCapital")
            lcolNomCampos.Add("IdCtaMoraCr")
            lcolDatosNuevos.Add(astrIdCtaDscto, "IdCtaDctoCapital")
            lcolDatosNuevos.Add(astrIdCtaMoraCr, "IdCtaMoraCr")
            lcolCamposRef.Add("IdCarpeta")
            lcolCamposRef.Add("IdCentroUtil")
            lcolDatosRef.Add(ashrIdCarpeta)
            lcolDatosRef.Add(ashrIdCentroUtil)
            GobjPanDat.SActualiceRegistro(lstrTabla, lcolNomCampos, lcolDatosNuevos,
                    lcolCamposRef, lcolDatosRef)

        End Sub
#End Region
#Region "Actualiza a la Versión 130 de la base de datos"
        Public Shared Sub SActuBD_V136()
            'Obetivo: pasar eMail, Celular, Telefono1 y Telefono2 del Cliente al Tercero 23/08/2018
            Dim lblnOk = False
            Try
                GobjPanDat.SControleProcesoObj(True)
                GobjPanDat.SInicialiceTransaccion()
                SAcualiceDatosInfo()
                lblnOk = True
            Catch ex As Exception
                Throw
            Finally
                If Not lblnOk Then
                    GobjPanDat.SAborteTransaccion()
                    GobjPanDat.SControleProcesoObj(False, True)
                Else
                    GobjPanDat.SConfirmeTransaccion()
                    GobjPanDat.SControleProcesoObj(False)
                End If
            End Try
        End Sub
        Private Shared Sub SAcualiceDatosInfo()
            Dim lstrSql = "SELECT idTerceroCliente, Celular, eMail, Telefono1, Telefono2 FROM Oriclientes " &
                    "ORDER BY IdTerceroCliente"
            Dim ldtbDatosClie As DataTable = Nothing
            Using ldsCentrosUtil As New DataSet
                GobjPanDat.SdsDataSet(ldsCentrosUtil, lstrSql)
                ldtbDatosClie = ldsCentrosUtil.Tables(0)
            End Using
            Dim lstrCel As String, lstrEmail As String, lstrTel1 As String
            Dim lstrTel2 As String, ldblIdTercero As Double
            For Each ldrwInfo As DataRow In ldtbDatosClie.Rows
                ldblIdTercero = CType(ldrwInfo("idTerceroCliente"), Double)
                If Not IsDBNull(ldrwInfo("Celular")) Then
                    lstrCel = CType(ldrwInfo("Celular"), String)
                    SActuInfoTer(1, lstrCel, ldblIdTercero)
                End If
                If Not IsDBNull(ldrwInfo("eMail")) Then
                    lstrEmail = CType(ldrwInfo("eMail"), String)
                    SActuInfoTer(2, lstrEmail, ldblIdTercero)
                End If
                If Not IsDBNull(ldrwInfo("Telefono1")) Then
                    lstrTel1 = CType(ldrwInfo("Telefono1"), String)
                    SActuInfoTer(3, lstrTel1, ldblIdTercero)
                End If
                If Not IsDBNull(ldrwInfo("Telefono2")) Then
                    lstrTel2 = CType(ldrwInfo("Telefono2"), String)
                    SActuInfoTer(4, lstrTel2, ldblIdTercero)
                End If
            Next
        End Sub
        Private Shared Sub SActuInfoTer(aentInf As Integer, astrInf As String,
                adblIdTercero As Double)
            Dim lstrSql = "UPDATE PanTerceros SET "
            If Not String.IsNullOrEmpty(astrInf) Then
                Select Case aentInf
                    Case 1
                        lstrSql += "Celular = '" & astrInf & "'"
                    Case 2
                        lstrSql += "eMail = '" & astrInf & "'"
                    Case 3
                        lstrSql += "Telefono1 = '" & astrInf & "'"
                    Case 4
                        lstrSql += "Telefono2 = '" & astrInf & "'"
                End Select
                lstrSql += " WHERE IdTercero = " & adblIdTercero.ToString
                GobjPanDat.SEjecuteSentenciaSql(lstrSql)
            End If
        End Sub
#End Region
#Region "Actualiza la bd a la versión 244"
        Public Sub SActuBD_V244()
            Dim lblnNoHayError = False
            Try
                GobjPanDat.SControleProcesoObj(True)
                GobjPanDat.SInicialiceTransaccion()
                Dim lstrExpSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect("OriCentrosUtilidadOriCop",
                        {"IdCarpeta", "IdCentroUtil", "CausaIntAlRecCaja", "CausaIntCierreMes",
                        "CausaIntUltimoDia"}, {{String.Empty, String.Empty}}, String.Empty, {})
                Dim ldtbRes As DataTable = Nothing
                Using ldsRes As New DataSet
                    GobjPanDat.SdsDataSet(ldsRes, lstrExpSql)
                    ldtbRes = ldsRes.Tables(0)
                End Using
                Dim lshrIdCarpeta As Short, lshrIdCenUtil As Short, lblnCausaAlRC As Boolean,
                        lblnCausaFinMes As Boolean, lblnCausaUltimoDia As Boolean,
                        lbytIdModoCausaMora As Byte
                For Each ldrwRes As DataRow In ldtbRes.Rows
                    lshrIdCarpeta = ldrwRes("IdCarpeta")
                    lshrIdCenUtil = ldrwRes("IdCentroUtil")
                    lblnCausaAlRC = ldrwRes("CausaIntAlRecCaja")
                    lblnCausaFinMes = ldrwRes("CausaIntCierreMes")
                    lblnCausaUltimoDia = ldrwRes("CausaIntUltimoDia")
                    Select Case True
                        Case lblnCausaFinMes
                            lbytIdModoCausaMora = 1
                        Case lblnCausaUltimoDia
                            lbytIdModoCausaMora = 2
                        Case lblnCausaAlRC
                            lbytIdModoCausaMora = 3
                        Case Else
                            lbytIdModoCausaMora = 0
                    End Select
                    lstrExpSql = "UPDATE OriCentrosUtilidadOriCop SET IdModoCausaMora = " &
                            lbytIdModoCausaMora & " WHERE IdCarpeta = " & lshrIdCarpeta &
                            " AND IdCentroUtil = " & lshrIdCenUtil
                    GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
                Next
                lblnNoHayError = True
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
        End Sub
#End Region
#End Region
    End Class
End Namespace