Friend Module MActualizaBD
#Region "Definiciones"
    ' Variables
    Friend StrNombreBD As String = String.Empty
#End Region

#Region "Propiedades"
    '
#End Region

#Region "Procedimientos"
    ''' <summary>
    ''' Actualiza las tablas de la aplicacion de acuerdo a las variables globales definidas en 
    ''' "mActualizaBDPub.vb" y asignadas en "clsPanoramaDat.vb" de acuerdo a los parámetros pasados 
    ''' por la aplicación.
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub SActualiceBD(astrNombreArchivoXml As String)
        If astrNombreArchivoXml Is Nothing Then
            Throw New ArgumentNullException(NameOf(astrNombreArchivoXml))
        End If
        Dim lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            Select Case GenuProveedorBD
                Case EnuProveedorBD.enuSQLServer
                    'mobjPoneAlDiaBD = New clsPoneAlDiaSqlServer(Me, lstrArchScript)
                Case EnuProveedorBD.enuOracle
                    'mobjPoneAlDiaBD = New clsPoneAlDiaOracle(Me, astrNombreArchivoXml)
                Case EnuProveedorBD.enuMySql
                    Dim lobjPoneAlDiaBD = New ClsPoneAlDiaMySql()
                    lobjPoneAlDiaBD.SActualiceBaseDatos(astrNombreArchivoXml)
            End Select
            lblnNoHayError = True
        Catch ex As PanDatException
            Throw
        Catch ex As MySql.Data.MySqlClient.MySqlException
            Throw
        Finally
            If lblnNoHayError Then
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub

    Friend Sub SCopieColumna(astrNombreTabla As String, astrNombresColumnaOrigen As String,
                astrNombresColumnaDestino As String)
        Dim lstrsql As String = FstrConstruyaExpSqlCopiarColumna(astrNombreTabla,
                    astrNombresColumnaOrigen, astrNombresColumnaDestino)
        GobjPanDat.SEjecuteSentenciaSql(lstrsql)
    End Sub

    Friend Function FobjBaseDatosDB(astrPrefijoTablas As String) As ClsBaseDatos
        Dim lobjPoneAlDiaMySql As New ClsPoneAlDiaMySql()
        Return lobjPoneAlDiaMySql.FobjBaseDatosDB(astrPrefijoTablas)
    End Function
#End Region

#Region "Funciones"
#Region "Funciones relacionadas con los objetos instanciados a partir de la clase clsEstructuraBD"
    '''<summary>
    ''' Devuelve el string que representa un tipo de datos para ser usado en las instrucciones Sql, 
    ''' a partir de las propiedades de la columna en el script.
    ''' </summary>
    ''' <param name="aobjColumna">Objeto columna que contiene todas las propiedades de la columna</param>
    ''' <returns>String para ser usado en expresión Sql</returns>
    ''' <remarks></remarks>
    Friend Function FstrTipoDatoColumna(aobjColumna As ClsColumna) As String
        Select Case GenuProveedorBD
            Case EnuProveedorBD.enuMySql
                Return FstrTipoDatoColumnaMySql(aobjColumna)
            Case EnuProveedorBD.enuOracle
            Case EnuProveedorBD.enuSQLServer

        End Select
        Return Nothing
    End Function

    ''' <summary>
    ''' Devuelve un string con el valor normalizado del valor de un campo contenido en el script Xml para ser 
    ''' incluido en una expresión SQL.
    ''' </summary>
    ''' <param name="aobjValor">El valor a ser normalizado</param>
    ''' <param name="astrTipoDato">El tipo de dato de la columna contenido en el script Xml para el campo dado.</param>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Friend Function FstrValorCampo(aobjValor As Object, astrTipoDato As String)
        Dim lstrValor = FstrValor(aobjValor, astrTipoDato)
        If String.IsNullOrEmpty(lstrValor) Then
            Select Case astrTipoDato.ToUpper
                Case "DATE"
                    lstrValor = aobjValor.ToString
                    lstrValor = ClsPanoramaDat.FstrFechaNormalizada(lstrValor)
                    lstrValor = "'" & lstrValor & "', "
                Case "DATETIME"
                    lstrValor = aobjValor.ToString
                    lstrValor = ClsPanoramaDat.FstrFechaHoraNormalizada(lstrValor)
                    lstrValor = "'" & lstrValor & "', "
                Case "BOOLEAN"
                    lstrValor = CType(aobjValor, Boolean) & ", "
                Case "STRING", "VSTRING"
                    lstrValor = "'" & aobjValor.ToString & "', "
                Case "OBJECT"
                    '
            End Select
        End If
        Return lstrValor
    End Function

    Private Function FstrValor(aobjValor As Object, astrTipoDato As String) As String
        Dim lstrTipoDat = astrTipoDato.ToUpper
        Dim lstrValor = String.Empty
        If lstrTipoDat = "BYTE" OrElse lstrTipoDat = "SHORT" OrElse lstrTipoDat = "INTEGER" OrElse
                    lstrTipoDat = "LONG" OrElse lstrTipoDat = "UBYTE" OrElse lstrTipoDat = "USHORT" OrElse
                    lstrTipoDat = "UINTEGER" OrElse lstrTipoDat = "ULONG" OrElse lstrTipoDat = "CURRENCY" OrElse
                    lstrTipoDat = "SINGLE" OrElse lstrTipoDat = "DOUBLE" Then
            If String.IsNullOrEmpty(aobjValor.ToString) Then
                lstrValor = "0, "
            Else
                lstrValor = aobjValor.ToString & ", "
            End If
        End If
        Return lstrValor
    End Function

    Friend Function FstrTipoDatoColumnaMySql(aobjColumna As ClsColumna) As String
        If aobjColumna Is Nothing Then
            Throw New ArgumentNullException(NameOf(aobjColumna))
        End If
        Dim lstrMensErr As String = String.Empty
        Dim lushTamano As UShort = 0
        Dim lbytDecimales As Byte = 0, lstrTipoDat = String.Empty
        With aobjColumna
            If Not String.IsNullOrEmpty(.StrLongitud) Then
                If .StrLongitud.Contains(",") Then
                    Dim lstrpar() As String = .StrLongitud.Split(",")
                    lushTamano = CType(lstrpar(0), UShort)
                    lbytDecimales = CType(lstrpar(1), Byte)
                Else
                    lushTamano = CType(.StrLongitud, UShort)
                End If
            End If
            lstrTipoDat = FstrTipoDatoMySql(aobjColumna, lushTamano, lbytDecimales)
            If String.IsNullOrEmpty(lstrTipoDat) Then
                lstrTipoDat = FstrTipoDatoMySql_1(aobjColumna)
                If String.IsNullOrEmpty(lstrTipoDat) Then
                    Select Case .StrTipoDatos.ToUpper
                        Case "BOOLEAN"
                            lstrTipoDat = "BOOLEAN"
                        Case "STRING"
                            If lushTamano > 0 Then
                                lstrTipoDat = "CHAR" & "(" & lushTamano.ToString & ")"
                            Else
                                lstrMensErr = "El tamaño de la Columna esta declarado erroneamente: " & .StrTipoDatos
                            End If
                        Case "VSTRING"
                            If lushTamano > 0 Then
                                lstrTipoDat = "VARCHAR" & "(" & lushTamano.ToString & ")"
                            Else
                                lstrMensErr = "El tamaño de la Columna esta declarado erroneamente: " & .StrTipoDatos
                            End If
                        Case "OBJECT"
                            lstrTipoDat = "LONGBLOB"
                        Case Else
                            lstrMensErr = "El tipo de la Columna no está definido."
                    End Select
                End If
            End If
        End With
        If Not String.IsNullOrEmpty(lstrMensErr) Then
            Throw New ArchivoXmlPanException(lstrMensErr)
        End If
        Return lstrTipoDat
    End Function

    Private Function FstrTipoDatoMySql(aobjColumna As ClsColumna,
                aushTamano As UShort, abytDecimales As Byte) As String
        Dim lstrTipoDato = String.Empty
        Select Case aobjColumna.StrTipoDatos.ToUpper
            Case "BYTE"
                lstrTipoDato = "TINYINT"
            Case "UBYTE"
                lstrTipoDato = "TINYINT UNSIGNED"
            Case "SHORT"
                lstrTipoDato = "SMALLINT"
            Case "USHORT"
                lstrTipoDato = "SMALLINT UNSIGNED"
            Case "INTEGER"
                lstrTipoDato = "INTEGER"
            Case "UINTEGER"
                lstrTipoDato = "INTEGER UNSIGNED"
            Case "LONG"
                lstrTipoDato = "BIGINT"
            Case "ULONG"
                lstrTipoDato = "BIGINT UNSIGNED"
            Case "CURRENCY"
                If aushTamano > 0 Then
                    If aushTamano > 0 Then
                        lstrTipoDato = "DECIMAL(" & aushTamano & "," & abytDecimales & ")"
                    Else
                        lstrTipoDato = "DECIMAL(" & aushTamano & ",2)"
                    End If
                Else
                    lstrTipoDato = "DECIMAL(12,2)"
                End If
        End Select
        Return lstrTipoDato
    End Function

    Private Function FstrTipoDatoMySql_1(aobjColumna As ClsColumna) As String
        Dim lstrTipoDat = String.Empty
        Select Case aobjColumna.StrTipoDatos.ToUpper
            Case "SINGLE"
                lstrTipoDat = "FLOAT"
            Case "DOUBLE"
                lstrTipoDat = "DOUBLE"
            Case "DATE"
                lstrTipoDat = "DATE"
            Case "DATETIME"
                lstrTipoDat = "DATETIME"
        End Select
        Return lstrTipoDat
    End Function

    Friend Function FstrDefColumnaSql(aobjColumna As ClsColumna,
                astrTipoDatoSql As String) As String
        Dim lstrSql As String = aobjColumna.StrNombre & " "
        Select Case GenuProveedorBD
            Case EnuProveedorBD.enuMySql
                With aobjColumna
                    lstrSql &= astrTipoDatoSql
                    If Not String.IsNullOrEmpty(.StrCharacterSet) Then
                        lstrSql &= " CHARACTER SET " & .StrCharacterSet
                    End If
                    If Not String.IsNullOrEmpty(.StrCollationName) Then
                        lstrSql &= " COLLATE " & .StrCollationName
                    End If
                    If .BlnRequerido Then
                        lstrSql &= " NOT NULL"
                        If Not .BlnAutoNumerico Then
                            lstrSql &= FstrValorDefecto(aobjColumna)
                        End If
                    Else
                        lstrSql &= " NULL"
                    End If
                    If .BlnAutoNumerico Then
                        lstrSql &= " AUTO_INCREMENT"
                    End If
                    If Not String.IsNullOrEmpty(.StrComentario) Then
                        lstrSql &= " COMMENT '" & .StrComentario & ("', ")
                    Else
                        lstrSql &= (", ")
                    End If
                End With
            Case EnuProveedorBD.enuOracle
            Case EnuProveedorBD.enuSQLServer
        End Select
        Return lstrSql
    End Function

    Friend Function FstrValorDefecto(aobjCalumna As ClsColumna) As String
        Dim lstrValorPorDefecto = " DEFAULT ", lstrValorDefecto As String
        If aobjCalumna IsNot Nothing Then
            lstrValorDefecto = FstrValorDefecto_1(aobjCalumna)
            If String.IsNullOrEmpty(lstrValorDefecto) Then
                Select Case aobjCalumna.StrTipoDatos.ToUpper
                    Case "CURRENCY"
                        If IsNothing(aobjCalumna.ObjValorDefault) Then
                            lstrValorDefecto &= "0.00"
                        Else
                            lstrValorDefecto &= aobjCalumna.ObjValorDefault
                        End If
                    Case "DATE"
                        If IsNothing(aobjCalumna.ObjValorDefault) Then
                            lstrValorDefecto &= "'" & "1900-01-01" & "'"
                        Else
                            lstrValorDefecto &= "'" & aobjCalumna.ObjValorDefault & "'"
                        End If
                    Case "DATETIME"
                        If IsNothing(aobjCalumna.ObjValorDefault) Then
                            lstrValorDefecto &= "'" & "1900-01-01 00:00:00" & "'"
                        Else
                            lstrValorDefecto &= "'" & aobjCalumna.ObjValorDefault & "'"
                        End If
                    Case "BOOLEAN"
                        If IsNothing(aobjCalumna.ObjValorDefault) Then
                            lstrValorDefecto &= "False"
                        Else
                            lstrValorDefecto &= aobjCalumna.ObjValorDefault
                        End If
                    Case "STRING", "VSTRING"
                        If IsNothing(aobjCalumna.ObjValorDefault) Then
                            lstrValorDefecto &= "'*'"
                        Else
                            lstrValorDefecto &= "'" & aobjCalumna.ObjValorDefault & "'"
                        End If
                    Case "OBJECT"
                        lstrValorDefecto = String.Empty
                End Select
            End If
            lstrValorPorDefecto &= lstrValorDefecto
        Else
            Throw New ArgumentNullException(NameOf(aobjCalumna))
        End If
        Return lstrValorPorDefecto
    End Function

    Private Function FstrValorDefecto_1(aobjCalumna As ClsColumna) As String
        Dim lstrTipoDat = aobjCalumna.StrTipoDatos.ToUpper
        Dim lstrValorDefecto As String = String.Empty
        If lstrTipoDat = "BYTE" OrElse lstrTipoDat = "SHORT" OrElse lstrTipoDat = "INTEGER" OrElse
                    lstrTipoDat = "LONG" OrElse lstrTipoDat = "UBYTE" OrElse lstrTipoDat = "USHORT" OrElse
                    lstrTipoDat = "UINTEGER" OrElse lstrTipoDat = "ULONG" OrElse lstrTipoDat = "SINGLE" OrElse
                    lstrTipoDat = "DOUBLE" Then
            If IsNothing(aobjCalumna.ObjValorDefault) Then
                lstrValorDefecto = "0"
            Else
                lstrValorDefecto = aobjCalumna.ObjValorDefault
            End If
        End If
        Return lstrValorDefecto
    End Function

    Friend Function FstrDefIndice(aobjIndice As ClsIndice) As String
        If aobjIndice IsNot Nothing Then
            Dim lstrSql As String = String.Empty
            Select Case GenuProveedorBD
                Case EnuProveedorBD.enuMySql
                    If aobjIndice.BlnPrincipal AndAlso aobjIndice.BlnUnico Then
                        lstrSql = " CONSTRAINT PRIMARY KEY " & aobjIndice.StrNombre & " ("
                    ElseIf aobjIndice.BlnUnico Then
                        lstrSql = " CONSTRAINT UNIQUE INDEX " & aobjIndice.StrNombre & " ("
                    Else
                        lstrSql = " KEY " & aobjIndice.StrNombre & " ("
                    End If
                    For Each lcolIndice As ClsColumnaIndice In aobjIndice.ColColumnasIndice
                        lstrSql &= lcolIndice.StrNombre
                        If lcolIndice.BlnAscendente Then
                            lstrSql &= " ASC, "
                        Else
                            lstrSql &= " DESC, "
                        End If
                    Next
                    lstrSql = lstrSql.Substring(0, Len(lstrSql) - 2) & ")"
                Case EnuProveedorBD.enuOracle
                Case EnuProveedorBD.enuSQLServer
            End Select
            Return lstrSql
        Else
            Throw New ArgumentNullException(NameOf(aobjIndice))
        End If
    End Function
#End Region
#Region "Funciones a partir de la informacion devuelta por la funcion GetSchema"
    ''' <summary>
    ''' Devuelve el string que representa un tipo de datos usado en el script, 
    ''' a partir del tipo de datos en la información devuelta por GetSchema
    ''' </summary>
    ''' <param name="astrTipoDatosSchema">String del tipo de datos devuelto por GetSchema</param>
    ''' <returns>String utilizado en el script</returns>
    ''' <remarks></remarks>
    Friend Function FstrTipoDatosScriptSchema(astrTipoDatosSchema As String) As String
        Select Case GenuProveedorBD
            Case EnuProveedorBD.enuMySql
                Return FstrTipoDatosScriptSchemaMySql(astrTipoDatosSchema)
            Case EnuProveedorBD.enuOracle
            Case EnuProveedorBD.enuSQLServer

        End Select
        Return Nothing
    End Function

    ''' <summary>
    ''' Devuelve la longitud que representa la longitud de una cadena usada en el script, 
    ''' a partir del tipo de datos en la información devuelta por GetSchema
    ''' </summary>
    ''' <param name="astrTipoDatosSchema">String del tipo de datos devuelto por GetSchema</param>
    ''' <returns>String utilizado en el script</returns>
    ''' <remarks></remarks>
    Friend Function FstrLongitudStringScriptSchema(astrTipoDatosSchema As String) As String
        Select Case GenuProveedorBD
            Case EnuProveedorBD.enuMySql
                Return FstrLongitudStringScriptSchemaMySql(astrTipoDatosSchema)
            Case EnuProveedorBD.enuOracle
            Case EnuProveedorBD.enuSQLServer

        End Select
        Return Nothing
    End Function

    Friend Function FstrTipoDatosScriptSchemaMySql(astrTipoDatosSchema As String) As String
        If String.IsNullOrEmpty(astrTipoDatosSchema) Then
            Throw New ArgumentNullException(NameOf(astrTipoDatosSchema))
        End If
        Dim lstrTipoDato As String
        Dim lstrLongitud As String = String.Empty
        Dim lblnSinSigno As Boolean = astrTipoDatosSchema.Contains("unsigned")
        If astrTipoDatosSchema.Contains("(") Then
            Dim lstrParametros As String =
                        astrTipoDatosSchema.Substring(astrTipoDatosSchema.IndexOf("(") + 1)
            lstrParametros = lstrParametros.Substring(0, lstrParametros.IndexOf(")"))
            If lstrParametros.Contains(",") Then
                Dim lstrpar() As String = lstrParametros.Split(",")
                lstrLongitud = lstrpar(0)
            Else
                lstrLongitud = lstrParametros
            End If
            lstrTipoDato = astrTipoDatosSchema.Substring(0, astrTipoDatosSchema.IndexOf("("))
        Else
            lstrTipoDato = astrTipoDatosSchema
        End If
        Dim lstrTipoDatoScript = FstrTipoDato(lstrTipoDato.ToUpper, lstrLongitud, lblnSinSigno)
        Return lstrTipoDatoScript
    End Function

    Private Function FstrTipoDato(astrTipoDat As String, astrLong As String,
                ablnSinSigno As Boolean) As String
        Dim lstrTipoDat = FstrTipoDat(astrTipoDat, astrLong, ablnSinSigno)
        If String.IsNullOrEmpty(lstrTipoDat) Then
            Select Case astrTipoDat
                Case "DECIMAL"
                    lstrTipoDat = "CURRENCY"
                Case "FLOAT"
                    lstrTipoDat = "SINGLE"
                Case "DOUBLE"
                    lstrTipoDat = "DOUBLE"
                Case "DATE"
                    lstrTipoDat = "DATE"
                Case "DATETIME"
                    lstrTipoDat = "DATETIME"
                Case "CHAR"
                    lstrTipoDat = "STRING"
                Case "VARCHAR"
                    lstrTipoDat = "VSTRING"
                Case "LONGBLOB"
                    lstrTipoDat = "OBJOLE"
                Case Else
                    Throw New ArchivoXmlPanException("El tipo de dato " & astrTipoDat.ToUpper &
                            " no es soportado por Panorama.net.")
            End Select
        End If
        Return lstrTipoDat
    End Function

    Private Function FstrTipoDat(astrTipoDat As String, astrLong As String,
                ablnSinSigno As Boolean) As String
        Dim lstrDat = String.Empty
        If astrTipoDat = "TINYINT" Then
            If astrLong = "1" Then
                lstrDat = "BOOLEAN"
            ElseIf ablnSinSigno Then
                lstrDat = "UBYTE"
            Else
                lstrDat = "BYTE"
            End If
        ElseIf astrTipoDat = "TINYINT UNSIGNED" Then
            lstrDat = "UBYTE"
        ElseIf astrTipoDat = "SMALLINT" Then
            If ablnSinSigno Then
                lstrDat = "USHORT"
            Else
                lstrDat = "SHORT"
            End If
        ElseIf astrTipoDat = "INT" Then
            If ablnSinSigno Then
                lstrDat = "UINTEGER"
            Else
                lstrDat = "INTEGER"
            End If
        ElseIf astrTipoDat = "BIGINT" Then
            If ablnSinSigno Then
                lstrDat = "ULONG"
            Else
                lstrDat = "LONG"
            End If
        End If
        Return lstrDat
    End Function

    Friend Function FstrLongitudStringScriptSchemaMySql(astrTipoDatosSchema As String) As String
        If String.IsNullOrEmpty(astrTipoDatosSchema) Then
            Throw New ArgumentNullException(NameOf(astrTipoDatosSchema))
        End If
        Dim lstrTipoDato As String
        Dim lstrLongitud As String = String.Empty
        If astrTipoDatosSchema.Contains("(") Then
            Dim lstrParametros As String =
                        astrTipoDatosSchema.Substring(astrTipoDatosSchema.IndexOf("(") + 1)
            lstrParametros = lstrParametros.Substring(0, lstrParametros.IndexOf(")"))
            If lstrParametros.Contains(",") Then
                Dim lstrpar() As String = lstrParametros.Split(",")
                lstrLongitud = lstrpar(0)
            Else
                lstrLongitud = lstrParametros
            End If
            lstrTipoDato = astrTipoDatosSchema.Substring(0, astrTipoDatosSchema.IndexOf("("))
        Else
            lstrTipoDato = astrTipoDatosSchema
        End If
        Select Case lstrTipoDato.ToUpper
            Case "CHAR", "VARCHAR"
                Return lstrLongitud
            Case Else
                Return ""
        End Select
        Return Nothing
    End Function
#End Region
#Region "Funciones que devuelven expresiones sql"
#Region "Construye expresiones Sql de tabla"
    ''' <summary>
    ''' Devueleve la expresion Sql que crea una tabla
    ''' </summary>
    ''' <param name="aobjTablaXml">El objeto tabla que contiene toda la definición de la tabla</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function FstrConstruyaExpSqlCreeTabla(aobjTablaXml As ClsTabla) As String
        Select Case GenuProveedorBD
            Case EnuProveedorBD.enuMySql
                Return FstrConstruyaExpSqlCreeTablaMySql(aobjTablaXml)
            Case EnuProveedorBD.enuOracle
            Case EnuProveedorBD.enuSQLServer
        End Select
        Return ""
    End Function

    Friend Function FstrConstruyaExpSqlCreeTablaMySql(aobjTablaXml As ClsTabla) As String
        Dim lstrTipoDatoCol As String
        Dim lstrSql As String = "CREATE TABLE " & aobjTablaXml.StrNombre & " ("
        For Each lobjCol As ClsColumna In aobjTablaXml.ColColumnas
            lstrTipoDatoCol = FstrTipoDatoColumna(lobjCol)
            lstrSql &= FstrDefColumnaSql(lobjCol, lstrTipoDatoCol)
        Next
        lstrSql = lstrSql.Substring(0, Len(lstrSql) - 2)
        For Each lobjIndice As ClsIndice In aobjTablaXml.ColIndices
            lstrSql &= ","
            lstrSql &= FstrDefIndice(lobjIndice)
        Next
        lstrSql &= ")"
        If Not String.IsNullOrEmpty(aobjTablaXml.StrCollationName) Then
            lstrSql &= " DEFAULT COLLATE " & aobjTablaXml.StrCollationName
        End If
        Return lstrSql
    End Function

    ''' <summary>
    ''' Devuelve la expresión Sql que elimina una tabla de la base de datos 
    ''' tiene en cuenta el motor de base de datos actus
    ''' </summary>
    ''' <param name="astrNombreTabla">Nombre de la tabla a suprimir</param>
    ''' <returns>La expresion Sql como un string</returns>
    ''' <remarks>Se tiene en cuenta el motor de base de datos actual</remarks>
    Friend Function FstrConstruyaExpSqlElimineTabla(astrNombreTabla As String) As String
        Select Case GenuProveedorBD
            Case EnuProveedorBD.enuMySql
                Return "DROP TABLE IF EXISTS " & astrNombreTabla
            Case EnuProveedorBD.enuOracle
            Case EnuProveedorBD.enuSQLServer

        End Select
        Return ""
    End Function
#End Region
#Region "Construye expresiones Sql de Columna"
    ''' <summary>
    ''' Devueleve la expresion Sql que crea un Columna(columna) en una tabla
    ''' </summary>
    ''' <param name="aobjColumna">Objeto columna que contiene la definición completa de la columna</param>
    ''' <returns>La expresion Sql como un string</returns>
    ''' <remarks>Se teniene en cuenta el proveedor de datos actual</remarks>
    Friend Function FstrConstruyaExpSqlCreeColumna(aobjColumna As ClsColumna)
        Select Case GenuProveedorBD
            Case EnuProveedorBD.enuMySql
                Return FstrConstruyaExpSqlCreeColumnaMySql(aobjColumna)
            Case EnuProveedorBD.enuOracle
            Case EnuProveedorBD.enuSQLServer
        End Select
        Return ""
    End Function

    Friend Function FstrConstruyaExpSqlCreeColumnaMySql(aobjColumna As ClsColumna)
        Dim lstrNombreTabla As String = aobjColumna.ObjPadre.StrNombre
        Dim lstrTipoDatoCol As String = FstrTipoDatoColumna(aobjColumna)
        Dim lstrSql As String = "ALTER TABLE " & lstrNombreTabla & " ADD " & FstrDefColumnaSql(aobjColumna,
                    lstrTipoDatoCol)
        lstrSql = lstrSql.Substring(0, Len(lstrSql) - 2)
        Return lstrSql
    End Function

    ''' <summary>
    ''' Devuelve la expresion sql para copiar los datos de un Columna en una tabla a otro Columna de la 
    ''' misma tabla
    ''' </summary>
    ''' <param name="astrNombreTabla">Tabla a la cual pertenecen los Columnas de origen y de destino</param>
    ''' <param name="astrNombreColumnaOrigen">Columna que contiene los datos que van a ser copiados</param>
    ''' <param name="astrNombreColumnaDestino">Columna que va ha recibir los datos de la Columna de origen</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function FstrConstruyaExpSqlCopiarColumna(astrNombreTabla As String,
                astrNombreColumnaOrigen As String, astrNombreColumnaDestino As String) As String
        Dim lstrSql As String = "UPDATE " & astrNombreTabla & " SET " & astrNombreColumnaDestino &
                    " = " & astrNombreColumnaOrigen
        Return lstrSql
    End Function

    ''' <summary>
    ''' Devuelve la expresión sql para cambiar alguna o varias de las propiedades de una columna.
    ''' </summary>
    ''' <param name="aobjColumnaXml">Definición xml de la columna como debe quedar.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function FstrConstruyaExpSqlCambiarColumna(aobjColumnaXml As ClsColumna) As String
        Dim lstrSql As String = String.Empty
        Select Case GenuProveedorBD
            Case EnuProveedorBD.enuMySql
                lstrSql = FstrConstruyaExpSqlCambiarColumnaMySql(aobjColumnaXml)
        End Select
        Return lstrSql
    End Function

    Friend Function FstrConstruyaExpSqlCambiarColumnaMySql(aobjColumnaXml As ClsColumna) As String
        Dim lstrSql As String
        Dim lstrNombreTabla As String = aobjColumnaXml.ObjPadre.StrNombre
        Dim lstrNombreColumna As String = aobjColumnaXml.StrNombre
        Dim lstrTipoDatoCol = FstrTipoDatoColumna(aobjColumnaXml)
        lstrSql = "ALTER TABLE " & lstrNombreTabla & " CHANGE " & lstrNombreColumna & " "
        lstrSql &= FstrDefColumnaSql(aobjColumnaXml, lstrTipoDatoCol)
        lstrSql = lstrSql.Substring(0, Len(lstrSql) - 2)
        Return lstrSql
    End Function

    Friend Function FstrConstruyaExpSqlCambiarNull(aobjColumnaBd As ClsColumna) As String
        Dim lstrSql As String
        Dim lstrNombreTabla As String = aobjColumnaBd.ObjPadre.StrNombre
        Dim lstrNombreColumna As String = aobjColumnaBd.StrNombre
        Dim lstrTipoDatoCol = FstrTipoDatoColumna(aobjColumnaBd)
        lstrSql = "ALTER TABLE " & lstrNombreTabla & " CHANGE " & lstrNombreColumna & " "
        lstrSql &= FstrDefColumnaSql(aobjColumnaBd, lstrTipoDatoCol)
        lstrSql = lstrSql.Substring(0, Len(lstrSql) - 2)
        lstrSql &= " DEFAULT '*'"
        Return lstrSql
    End Function

    ''' <summary>
    ''' Devuelve la expresion sql para cambiar un Columna de una tabla (Nombre, tipo de dato y comentarios) 
    ''' </summary>
    ''' <param name="aobjColumna">Objeto columna al cual se le cambiara de nombre</param>
    ''' <param name="astrNombreColumnaNuevo">Nombre que tendrá la Columna</param>
    ''' <returns>Expresion SQL</returns>
    ''' <remarks></remarks>
    Friend Function FstrConstruyaExpSqlRenombrarColumna(aobjColumna As ClsColumna,
                astrNombreColumnaNuevo As String) As String
        Dim lstrNombreTabla As String = aobjColumna.ObjPadre.StrNombre
        Dim lstrNombreColumnaOriginal As String = aobjColumna.StrNombre
        Dim lstrTipoDatoCol As String = FstrTipoDatoColumna(aobjColumna)
        Dim lstrDefColumna = FstrDefColumnaSql(aobjColumna, lstrTipoDatoCol)
        lstrDefColumna = lstrDefColumna.Replace(lstrNombreColumnaOriginal, astrNombreColumnaNuevo)
        Dim lstrSql As String = String.Empty
        Select Case GenuProveedorBD
            Case EnuProveedorBD.enuMySql
                lstrSql = "ALTER TABLE " & lstrNombreTabla & " CHANGE " &
                            lstrNombreColumnaOriginal & " " & lstrDefColumna
                lstrSql = lstrSql.Substring(0, lstrSql.Length - 2)
        End Select
        Return lstrSql
    End Function
#End Region
#Region "Construye expresiones Sql de Registro"
    ''' <summary>
    ''' Devuelve la expresion sql para insertar un registro en una tabla
    ''' </summary>
    ''' <param name="astrNombreTabla">Nombre de la tabla donde se insertará el registro</param>
    ''' <param name="acolNombresColumnas">Colección que contiene los nombres de los Columnas de la tabla</param>
    ''' <param name="acolDatos">Colección que contiene los datos de los Columnas del registro a ser inseratado</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function FstrConstruyaExpSqlInsertarReg(astrNombreTabla As String,
                    acolNombresColumnas As Collection, acolDatos As Collection,
                    acolTipoDatos As Collection) As String
        Dim lobjDato As Object
        Dim lstrTipoDato As String
        Dim lstrSql As String = "INSERT INTO " & astrNombreTabla & " ("
        If acolNombresColumnas.Count <> acolDatos.Count Then
            Throw New ArgumentoInvalidoPanException("Los Parametros pasados son incoherentes.")
        End If
        For Each lobjObjeto In acolNombresColumnas
            lstrSql &= lobjObjeto.ToString & ", "
        Next
        lstrSql = lstrSql.Remove(lstrSql.Length - 2, 2) & ") VALUES ("
        For i As Short = 1 To acolDatos.Count
            lobjDato = acolDatos(i)
            lstrTipoDato = acolTipoDatos(i)
            lstrSql &= FstrValorCampo(lobjDato, lstrTipoDato)
        Next
        lstrSql = lstrSql.Remove(lstrSql.Length - 2, 2) & ")"
        Return lstrSql
    End Function
#End Region
#End Region
#End Region
End Module