Imports System.IO
Namespace ActualizaBd
    Friend Class ClsEstructuraBD
#Region "Definiciones"
        ' Constantes
        Private Const CSTRNRaiz As String = "ActualiceBD"
        Private Const CSTRNBD As String = "BD"
        Friend Const CSTRNComandos As String = "Comandos"
        Friend Const CSTRNComando As String = "Comando"
        Private Const CSTRNTablas As String = "Tablas"
        Private Const CSTRNTabla As String = "Tabla"
        Private Const CSTRNRelaciones As String = "Relaciones"
        Private Const CSTRNRelacion As String = "Relacion"
        Private Const CSTRNColumnas As String = "Columnas"
        Private Const CSTRNColumna As String = "Columna"
        Private Const CSTRNIndices As String = "Indices"
        Private Const CSTRNIndice As String = "Indice"
        Private Const CSTRNColumnasIndice As String = "ColumnasIndice"
        Private Const CSTRNColumnaIndice As String = "ColumnaIndice"
        Private Const CSTRNRegistros As String = "Registros"
        Private Const CSTRNRegistro As String = "Registro"
        ' Variables de  propiedad
        Private ReadOnly Property MobjBaseDatos As ClsBaseDatos = Nothing
        Private MstrNombreNodo As String = String.Empty
        ' Variables de  control
        Private MenuTipoRegistro As EnuTipoLinea
        Private MstrArchivoXml As String = String.Empty
        Private MsrArchivoXml As StreamReader = Nothing
#End Region
#Region "Constructores"
        Public Sub New()
            If gobjPanDat.blnRegistrado Then
                MobjBaseDatos = New ClsBaseDatos(Me)
            Else
                Throw New System.InvalidOperationException("El módulo no esta registrado")
            End If
        End Sub
#End Region
#Region "Estructura BD"
        ''' <summary>
        ''' Devuelve el objeto objBaseDatos de la estructura. Este objeto contiene toda la estructura 
        ''' de la BD.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend ReadOnly Property ObjBaseDatos() As ClsBaseDatos
            Get
                Return MobjBaseDatos
            End Get
        End Property
        Friend ReadOnly Property StrNombreNodo() As String
            Get
                Return MstrNombreNodo
            End Get
        End Property
        ''' <summary>
        ''' Crea la estructura de la base de datos a partir del archivo XML que representa la base de datos.
        ''' </summary>
        ''' <param name="astrArchivoXml">Ruta completa del archivo XML que contiene la estructura de la 
        ''' base de datos</param>
        ''' <remarks></remarks>
        ''' 
        Friend Sub SCreaEstructuraXml(astrArchivoXml As String)
            Dim lstrMensError As String = String.Empty
            Dim lstrLinea As String
            Dim lstrPropiedades() As String
            MstrArchivoXml = astrArchivoXml
            MsrArchivoXml = ClsPanoramaDat.FsrStreamReader(MstrArchivoXml)
            SVerificaScript(MsrArchivoXml, astrArchivoXml)
            lstrLinea = FstrLineaScript()
            If MstrNombreNodo = CSTRNRaiz Then
                lstrLinea = FstrLineaScript()
                Do While lstrLinea <> CSTRNRaiz.Insert(0, "</") & ">"
                    If MstrNombreNodo = CSTRNBD Then
                        lstrPropiedades = FstrPropiedades(lstrLinea)
                        MobjBaseDatos.SAsignaPropiedadesBD(lstrPropiedades)
                        lstrLinea = FstrLineaScript()
                        Do While lstrLinea <> CSTRNBD.Insert(0, "</") & ">"
                            Select Case MstrNombreNodo
                                Case CSTRNComandos
                                    MobjBaseDatos.SAdicioneComandos()
                                Case CSTRNTablas
                                    MobjBaseDatos.SAdicioneTablas()
                                Case Else
                                    lstrMensError = "La linea '" & lstrLinea & vbTab &
                                            "' no cumple con el esquema de la Aplicación!"
                                    Exit Do
                            End Select
                            lstrLinea = FstrLineaScript()
                        Loop
                    Else
                        lstrMensError = "La linea '" & lstrLinea & vbTab &
                                "' no cumple con el esquema de la Aplicación!"
                    End If
                    If Not String.IsNullOrEmpty(lstrMensError) Then
                        Exit Do
                    End If
                    lstrLinea = FstrLineaScript()
                Loop
            Else
                lstrMensError = "La linea '" & lstrLinea & vbTab & "' no cumple con el esquema de la Aplicación!"
            End If
            If Not String.IsNullOrEmpty(lstrMensError) Then
                MsrArchivoXml.Close()
                Throw New ArchivoXmlPanException(lstrMensError)
            End If
        End Sub
        ''' <summary>
        ''' Lee la siguiente linea del archivo XML y verifica que cumpla con las directivas del esquema, y 
        ''' que este bien formada de lo contrario levanta uns excepcion "ArchivoXmlPanException"
        ''' </summary>
        ''' <returns>Indica un boolean que indica si la linea leida esta bien formada o no.</returns>
        ''' <remarks>Excepxiones: ArchivoXmlPanException</remarks>
        Friend Function FstrLineaScript() As String
            Dim lstrLinea As String = String.Empty
            Dim lstrLinCom As String
            Dim lblnLeyoLinea As Boolean = True
            Dim lstrNodosCon() As String = {CSTRNComando, CSTRNColumna, CSTRNColumnaIndice, CSTRNRegistro, CSTRNRelacion}
            Dim lstrNodosInicioCon() As String = {CSTRNBD, CSTRNTabla, CSTRNIndice} 'Nodos cuya linea de inicio contienen propiedades
            Dim lstrNodosInicio() As String = {CSTRNRaiz, CSTRNTablas, CSTRNComandos, CSTRNRelaciones,
                    CSTRNColumnas, CSTRNIndices, CSTRNColumnasIndice, CSTRNRegistros}
            Dim lblnCumpleEsquema As Boolean = True
            Dim lstrMensErr As String = String.Empty
            MstrNombreNodo = String.Empty
            MenuTipoRegistro = EnuTipoLinea.None
            Do While True
                Try
                    SLimpieLineasBlanco(lstrLinea)
                    If lstrLinea = Nothing Then
                        Exit Do
                    End If
                    Do While Not lstrLinea.EndsWith(">")
                        lstrLinCom = msrArchivoXml.ReadLine()
                        If lstrLinCom Is Nothing Then
                            lblnLeyoLinea = False
                            Exit Do
                        Else
                            lstrLinea = lstrLinea.Trim & " " & lstrLinCom.Trim
                        End If
                    Loop
                    If Not lblnLeyoLinea Then
                        Exit Do
                    End If
                    If lstrLinea.StartsWith("<!--") Then
                        If Not lstrLinea.EndsWith("-->") Then
                            lblnLeyoLinea = False
                            Exit Do
                        End If
                    Else
                        SRemplaceSignos(lstrLinea)
                        Exit Do
                    End If
                Catch ex As IOException
                    Throw New ArchivoXmlPanException(ex.Message)
                Catch ex As OutOfMemoryException
                    Throw New ArchivoXmlPanException(ex.Message)
                Catch ex As Exception
                    Throw New ArchivoXmlPanException(ex.Message)
                End Try
            Loop
            If lblnLeyoLinea AndAlso lstrLinea Is Nothing Then
                Return lstrLinea
            End If
            If lblnLeyoLinea Then
                If lstrLinea.StartsWith("<") AndAlso (Not lstrLinea.Contains(" ")) AndAlso lstrLinea.EndsWith(">") Then
                    SNombreNodo1(lstrLinea, lstrNodosInicio, lblnCumpleEsquema)
                ElseIf lstrLinea.StartsWith("<") AndAlso lstrLinea.EndsWith(Chr(34) & ">") Then
                    SNombreNodo2(lstrLinea, lstrNodosInicioCon, lblnCumpleEsquema)
                ElseIf lstrLinea.StartsWith("<") AndAlso lstrLinea.EndsWith("/>") Then
                    SNombreNodo3(lstrLinea, lstrNodosCon, lblnCumpleEsquema)
                Else
                    menuTipoRegistro = EnuTipoLinea.None
                End If
                If lblnCumpleEsquema AndAlso menuTipoRegistro = EnuTipoLinea.None Then
                    lblnLeyoLinea = False
                End If
            End If
            If Not lblnLeyoLinea Then
                lstrMensErr = "El archivo " & mstrArchivoXml & " no es un archivo XML Bien Formado." _
                         & vbCrLf & "Hay errores de sintaxis en la linea: " & lstrLinea & "!"
            End If
            If Not lblnCumpleEsquema Then
                lstrMensErr = "La linea '" & lstrLinea &
                        "' no cumple con el esquema de la Aplicación!"
            End If
            If Not String.IsNullOrEmpty(lstrMensErr) Then
                Throw New ArchivoXmlPanException(lstrMensErr)
            End If
            Return lstrLinea
        End Function
        Private Sub SLimpieLineasBlanco(ByRef astrLinea As String)
            Dim lstrLinea = String.Empty
            Do While String.IsNullOrEmpty(lstrLinea)
                lstrLinea = MsrArchivoXml.ReadLine
                If Not IsNothing(lstrLinea) Then
                    lstrLinea = lstrLinea.Trim
                Else
                    Exit Do
                End If
            Loop
            astrLinea = lstrLinea
        End Sub
        Private Shared Sub SRemplaceSignos(ByRef astrLinea As String)
            If astrLinea.Contains("&lt;") Then
                astrLinea = astrLinea.Replace("&lt;", "<")
            End If
            If astrLinea.Contains("&gt;") Then
                astrLinea = astrLinea.Replace("&gt;", ">")
            End If
        End Sub
        Private Sub SNombreNodo1(astrLinea As String, astrNodosInicio As String(),
                                 ByRef ablnCumpleEsquema As Boolean)
            If astrLinea.StartsWith("</") Then
                MstrNombreNodo = astrLinea.Substring(2, astrLinea.Length - 3)
                MenuTipoRegistro = EnuTipoLinea.enuFinNodo
            ElseIf astrLinea.IndexOf(">") <> astrLinea.Length - 1 AndAlso astrLinea.IndexOf("<", 2) > 0 _
                            AndAlso astrLinea.IndexOf("/") = astrLinea.IndexOf("<", 2) + 1 Then
                MstrNombreNodo = astrLinea.Substring(1, astrLinea.IndexOf(">") - 1)
                If MstrNombreNodo = astrLinea.Substring(astrLinea.IndexOf("/") + 1, astrLinea.IndexOf(">") - 1) Then
                    MenuTipoRegistro = EnuTipoLinea.enuPropiedad
                Else
                    MenuTipoRegistro = EnuTipoLinea.None
                End If
            ElseIf Not astrLinea.Contains("/") Then
                MstrNombreNodo = astrLinea.Substring(1, astrLinea.Length - 2)
                MenuTipoRegistro = EnuTipoLinea.enuInicioNodo
                If Not astrNodosInicio.Contains(MstrNombreNodo) Then
                    ablnCumpleEsquema = False
                End If
            Else
                MenuTipoRegistro = EnuTipoLinea.None
            End If
        End Sub
        Private Sub SNombreNodo2(astrLinea As String, astrNodosInicioCon As String(),
                                 ByRef ablnCumpleEsquema As Boolean)
            ' Linea de inicio de nodos Genral, BD y Tabla
            If astrLinea.Contains(" ") AndAlso astrLinea.Contains("=") Then
                ' Linea de nodo que contiene propiedades
                MstrNombreNodo = astrLinea.Substring(1, astrLinea.IndexOf(" ") - 1)
                MenuTipoRegistro = EnuTipoLinea.enuInicioNodoConProp
                If Not astrNodosInicioCon.Contains(MstrNombreNodo) Then
                    ablnCumpleEsquema = False
                End If
            Else
                MenuTipoRegistro = EnuTipoLinea.None
            End If
        End Sub
        Private Sub SNombreNodo3(astrLinea As String, astrNodosInicio As String(),
                                 ByRef ablnCumpleEsquema As Boolean)
            If astrLinea.Contains(" ") AndAlso astrLinea.Contains(Chr(34)) _
                            AndAlso astrLinea.Contains("=") Then
                MstrNombreNodo = astrLinea.Substring(1, astrLinea.IndexOf(" ") - 1)
                MenuTipoRegistro = EnuTipoLinea.enuNodoConProp
                If Not astrNodosInicio.Contains(MstrNombreNodo) Then
                    ablnCumpleEsquema = False
                End If
            Else
                MenuTipoRegistro = EnuTipoLinea.None
            End If
        End Sub
        Friend Shared Function FstrPropiedades(astrLineaProp As String) As Object
            If astrLineaProp Is Nothing Then
                Throw New ArgumentNullException(NameOf(astrLineaProp))
            End If
            Dim lstrPartes() As String
            Dim lstrProp() As String
            Dim lstrNomProp As String
            Dim lstrValPro As String
            Dim i, k As Integer
            k = -1
            If astrLineaProp.EndsWith("/>") OrElse astrLineaProp.EndsWith(Chr(34) & ">") Then
                astrLineaProp = astrLineaProp.Remove(0, astrLineaProp.IndexOf(" ") + 1)
                lstrPartes = astrLineaProp.Split(Chr(34))
                ReDim lstrProp((lstrPartes.GetUpperBound(0) - 2) / 2)
                For i = 0 To lstrPartes.GetUpperBound(0) - 1 Step 2
                    k += 1
                    lstrNomProp = lstrPartes(i).Trim.Replace("=", "&")
                    lstrValPro = lstrPartes(i + 1).Trim
                    lstrProp(k) = lstrNomProp & lstrValPro
                Next
            Else
                Throw New ArchivoXmlPanException("La linea '" & astrLineaProp & vbTab &
                        "' no cumple con el esquema de la Aplicación!")
            End If
            Return lstrProp
        End Function
        Private Shared Sub SVerificaScript(asrScript As StreamReader, astrArchivoScript As String)
            Dim lstrLinea As String = String.Empty
            Dim lstrMens As String = String.Empty
            If asrScript Is Nothing Then
                Throw New ArgumentNullException(NameOf(asrScript))
            End If
            Try
                Do While String.IsNullOrEmpty(lstrLinea)
                    lstrLinea = asrScript.ReadLine
                    If IsNothing(lstrLinea) Then
                        lstrMens = "El archivo " & astrArchivoScript & " es un archivo vacio."
                        Exit Do
                    End If
                Loop
                If Not IsNothing(lstrLinea) Then
                    If Not (lstrLinea.StartsWith("<?xml") AndAlso lstrLinea.EndsWith("?>")) Then
                        lstrMens = "El archivo " & astrArchivoScript & " no es un archivo XML bien formado." & vbCrLf &
                                "La linea de encabezado no es válida!"
                    End If
                End If
            Catch ex As IOException
                Throw New ArchivoXmlPanException(ex.Message)
            Catch ex As OutOfMemoryException
                Throw New ArchivoXmlPanException(ex.Message)
            Catch ex As Exception
                Throw New ArchivoXmlPanException(ex.Message)
            End Try
            If Not String.IsNullOrEmpty(lstrMens) Then
                Throw New ArchivoXmlPanException(lstrMens)
            End If
        End Sub
#End Region
    End Class
#Region "Clases de la estructura"
    Friend Class ClsBaseDatos
        ' Constantes
        Private Const CSTRNComandos As String = "Comandos"
        Private Const CSTRNComando As String = "Comando"
        Private Const CSTRNTablas As String = "Tablas"
        Private Const CSTRNTabla As String = "Tabla"
        Private Const CSTRNColumnas As String = "Columnas"
        Private Const CSTRNIndices As String = "Indices"
        Private Const CSTRNRegistros As String = "Registros"
        ' Variables    
        Private MintVersion As Integer = 0
        Private MstrNombreBD As String = String.Empty
        Private MstrCharacterSet As String = String.Empty
        Private MstrCollationName As String = String.Empty
        Friend ReadOnly Property ObjPadre As ClsEstructuraBD = Nothing
        Private MobjTabla As ClsTabla = Nothing
        Public Sub New(aobjPadre As ClsEstructuraBD)
            ObjPadre = aobjPadre
            ColTablas = New Collection
            ColComandos = New Collection
        End Sub
        ''' <summary>
        ''' Devuelve el nombre de la base de datos a la cual se refiere la estructura
        ''' </summary>
        ''' <returns>Nombre de la base de datos (String)</returns>
        ''' <remarks></remarks>
        Friend ReadOnly Property StrNombreBD() As String
            Get
                Return MstrNombreBD
            End Get
        End Property
        ''' <summary>
        ''' Devuelve la Versión de la base de datos en la estructura
        ''' </summary>
        ''' <returns>Versión de la base de datos (Integer)</returns>
        ''' <remarks></remarks>
        Friend ReadOnly Property EntVersion() As Integer
            Get
                Return MintVersion
            End Get
        End Property
        ''' <summary>
        ''' Devuelve la colección de los objetos tablas de la estructura
        ''' </summary>
        ''' <returns>Colección de objetos tabla (Collection)</returns>
        ''' <remarks></remarks>
        Friend ReadOnly Property ColTablas() As Collection = Nothing
        ''' <summary>
        ''' Devuelve la colección de los objetos comandos en la estructura
        ''' </summary>
        ''' <returns>Colección de objetos comando (Collection)</returns>
        ''' <remarks></remarks>
        Friend ReadOnly Property ColComandos() As Collection = Nothing
        Friend ReadOnly Property StrCharacterSet As String
            Get
                Return MstrCharacterSet
            End Get
        End Property
        Friend ReadOnly Property StrCollationName As String
            Get
                Return MstrCollationName
            End Get
        End Property
        Friend Sub SAsignaPropiedadesBD(astrNombreBD As String, astrCharacterSet As String,
                astrCollationName As String, aushVersion As UShort)
            Dim lstrPropiedades(3) As String
            lstrPropiedades(0) = "Nombre&" & astrNombreBD
            lstrPropiedades(1) = "CharacterSet&" & astrCharacterSet
            lstrPropiedades(2) = "CollationName&" & astrCollationName
            lstrPropiedades(3) = "Version&" & aushVersion
            SAsignaPropiedadesBD(lstrPropiedades)
        End Sub
        Friend Sub SAsignaPropiedadesBD(astrPropiedades() As String)
            Dim i As Integer
            Dim lstrNomProp As String
            Dim lstrValProp As String
            Dim lstrMensError As String = String.Empty
            ColTablas.Clear()
            ColComandos.Clear()
            For i = 0 To astrPropiedades.GetUpperBound(0)
                lstrNomProp = astrPropiedades(i).Split("&")(0)
                lstrValProp = astrPropiedades(i).Split("&")(1)
                Select Case lstrNomProp
                    Case Is = "Nombre"
                        MstrNombreBD = lstrValProp
                    Case Is = "Version"
                        MintVersion = CType(lstrValProp, Integer)
                    Case "CharacterSet"
                        MstrCharacterSet = lstrValProp
                    Case "CollationName"
                        MstrCollationName = lstrValProp
                    Case Else
                        lstrMensError = "El Nombre de Propiedad '" & lstrNomProp & "' no es valido"
                        Exit For
                End Select
            Next
            If Not String.IsNullOrEmpty(lstrMensError) Then
                Throw New ArchivoXmlPanException(lstrMensError)
            End If
        End Sub
        Friend Sub SAsigneNombreBD(astrNombreBD)
            MstrNombreBD = astrNombreBD
        End Sub
        Friend Sub SAdicioneComandos()
            Dim lstrLinea As String
            Dim lstrPropiedades() As String
            Dim lstrMensError As String = String.Empty
            Dim lobjComando As ClsComando
            lstrLinea = ObjPadre.FstrLineaScript()
            Do While lstrLinea <> CSTRNComandos.Insert(0, "</") & ">"
                If ObjPadre.StrNombreNodo = CSTRNComando Then
                    lstrPropiedades = ClsEstructuraBD.FstrPropiedades(lstrLinea)
                    lobjComando = New ClsComando(Me)
                    lobjComando.SAsignePropComado(lstrPropiedades)
                    ColComandos.Add(lobjComando)
                Else
                    lstrMensError = "La linea '" & lstrLinea & vbTab & "' no cumple con el esquema de la Aplicación!"
                    Exit Do
                End If
                lstrLinea = ObjPadre.FstrLineaScript()
            Loop
            If Not String.IsNullOrEmpty(lstrMensError) Then
                Throw New ArchivoXmlPanException(lstrMensError)
            End If
        End Sub
        Friend Sub SAdicioneTablas()
            Dim lstrPropiedades() As String
            Dim lstrLinea As String
            Dim lstrMensError As String = String.Empty
            lstrLinea = ObjPadre.FstrLineaScript
            Do While lstrLinea <> CSTRNTablas.Insert(0, "</") & ">"
                If ObjPadre.StrNombreNodo = CSTRNTabla Then
                    lstrPropiedades = ClsEstructuraBD.FstrPropiedades(lstrLinea)
                    MobjTabla = New ClsTabla(Me)
                    MobjTabla.SAsignePropTabla(lstrPropiedades)
                    lstrLinea = ObjPadre.FstrLineaScript
                    Do While lstrLinea <> CSTRNTabla.Insert(0, "</") & ">"
                        Select Case ObjPadre.StrNombreNodo
                            Case CSTRNComandos
                                MobjTabla.SAdicioneComandos()
                            Case CSTRNColumnas
                                MobjTabla.SAdicioneColumnas()
                            Case CSTRNIndices
                                MobjTabla.SAdicioneIndices()
                                MobjTabla.SVerifiqueColumnasRequeridas()
                            Case CSTRNRegistros
                                MobjTabla.SAdicioneRegistros()
                            Case Else
                                lstrMensError = "La linea '" & lstrLinea & "' tiene errores de sintaxis"
                                Exit Do
                        End Select
                        lstrLinea = ObjPadre.FstrLineaScript()
                    Loop
                    If Not String.IsNullOrEmpty(lstrMensError) Then
                        Exit Do
                    End If
                    ColTablas.Add(MobjTabla, MobjTabla.StrNombre)
                Else
                    lstrMensError = "La linea '" & lstrLinea & "' tiene errores de sintaxis"
                    Exit Do
                End If
                lstrLinea = ObjPadre.FstrLineaScript()
            Loop
            If Not String.IsNullOrEmpty(lstrMensError) Then
                Throw New ArchivoXmlPanException(lstrMensError)
            End If
        End Sub
        Friend Function FobjNuevaTabla(astrNombreTabla As String,
                astrCollationName As String) As ClsTabla
            Dim lobjTabla As New ClsTabla(Me)
            Dim lstrPropiedades(0) As String
            lstrPropiedades(0) = "Nombre&" & astrNombreTabla
            If Not String.IsNullOrEmpty(astrCollationName) Then
                ReDim Preserve lstrPropiedades(1)
                lstrPropiedades(1) = "CollationName&" & astrCollationName
            End If
            With lobjTabla
                .SAsignePropTabla(lstrPropiedades)
            End With
            ColTablas.Add(lobjTabla, lobjTabla.StrNombre)
            Return lobjTabla
        End Function
        Friend Sub SAdicioneTabla(aobjTabla As ClsTabla)
            Dim lobjTabla As ClsTabla
            If Not ColTablas.Contains(aobjTabla.StrNombre) Then
                lobjTabla = aobjTabla
                ColTablas.Add(lobjTabla, aobjTabla.StrNombre)
            Else
                Throw New ArchivoXmlPanException("La tabla '" & aobjTabla.StrNombre & "' ya existe en la estructura")
            End If
        End Sub
        Friend Sub SRenombreTabla(astrNombreTablaOri As String,
                astrNombreTablaNuevo As String)
            Dim lobjTabla As ClsTabla
            If ColTablas.Contains(astrNombreTablaOri) Then
                lobjTabla = ColTablas(astrNombreTablaOri)
                ColTablas.Remove(astrNombreTablaOri)
                lobjTabla.StrNombre = astrNombreTablaNuevo
                ColTablas.Add(lobjTabla, lobjTabla.StrNombre)
            End If
        End Sub
        Friend Sub SRemuevaTabla(astrNombreTabla As String)
            If ColTablas.Contains(astrNombreTabla) Then
                ColTablas.Remove(astrNombreTabla)
            End If
        End Sub
        Friend Sub SConviertaCharacterSet(astrCharacterSet As String, astrCollationName As String)
            MstrCharacterSet = astrCharacterSet
            MstrCollationName = astrCollationName
        End Sub
    End Class
    Friend Class ClsTabla
        ' Constantes
        Private Const CSTRNComandos As String = "Comandos"
        Private Const CSTRNComando As String = "Comando"
        Private Const CSTRNColumnas As String = "Columnas"
        Private Const CSTRNColumna As String = "Columna"
        Private Const CSTRNIndices As String = "Indices"
        Private Const CSTRNIndice As String = "Indice"
        Private Const CSTRNColumnasIndice As String = "ColumnasIndice"
        Private Const CSTRNRegistros As String = "Registros"
        Private Const CSTRNRegistro As String = "Registro"
        ' Variables de propiedad
        Private MstrNombre As String = String.Empty
        Private MstrBDVinculada As String = "N"
        Private MstrCollationName As String = String.Empty
        Private MblnVinculada As Boolean = False
        Public Sub New(aobjPadre As ClsBaseDatos)
            ObjPadre = aobjPadre
            ColComandos = New Collection
            ColColumnas = New Collection
            ColIndices = New Collection
            ColRegistros = New Collection
        End Sub
        ' PROPIEDADES
        Friend ReadOnly Property ObjPadre As ClsBaseDatos = Nothing
        Friend ReadOnly Property ColComandos() As Collection = Nothing
        Friend ReadOnly Property ColColumnas() As Collection = Nothing
        Friend ReadOnly Property ColIndices() As Collection = Nothing
        Friend ReadOnly Property ColRegistros() As Collection = Nothing
        Friend Property StrNombre() As String
            Get
                Return MstrNombre
            End Get
            Set(astrValor As String)
                If astrValor <> MstrNombre Then
                    MstrNombre = astrValor
                End If
            End Set
        End Property
        Friend ReadOnly Property BlnVinculada() As Boolean
            Get
                Return MblnVinculada
            End Get
        End Property
        Friend ReadOnly Property StrBDVinculada() As String
            Get
                Return MstrBDVinculada
            End Get
        End Property
        Friend ReadOnly Property StrCollationName As String
            Get
                Return MstrCollationName
            End Get
        End Property
        ' PROCEDIMIENTOS Y FUNCIONES
        Friend Sub SAsignePropTabla(astrPropiedades() As String)
            Dim lstrNomProp As String
            Dim lstrValProp As String
            Dim lstrMensError As String = String.Empty
            Dim i As Integer
            For i = 0 To astrPropiedades.GetUpperBound(0)
                lstrNomProp = astrPropiedades(i).Split("&")(0)
                lstrValProp = astrPropiedades(i).Split("&")(1)
                Select Case lstrNomProp
                    Case Is = "Nombre"
                        If lstrValProp.Length >= 4 Then
                            MstrNombre = lstrValProp
                        Else
                            lstrMensError = "El nombre de la tabla no es valido;" & vbCrLf &
                                    "este debe tener más de 4 caracteres."
                            Exit For
                        End If
                    Case Is = "Vinculada"
                        If lstrValProp = "N" Then
                            MblnVinculada = False
                        ElseIf lstrValProp = "S" Then
                            MblnVinculada = True
                        Else
                            lstrMensError = "El valor de la Propiedad 'Vinculada' de la tabla no es soportado por la Aplicación."
                            Exit For
                        End If
                    Case "BaseDatosVinculada"
                        MstrBDVinculada = lstrValProp
                    Case "CollationName"
                        MstrCollationName = lstrValProp
                    Case Else
                        lstrMensError = "El nombre de Propiedad '" & lstrNomProp & "' no es valido."
                        Exit For
                End Select
                If Not String.IsNullOrEmpty(lstrMensError) Then
                    Exit For
                End If
            Next
            If Not String.IsNullOrEmpty(lstrMensError) Then
                Throw New ArchivoXmlPanException(lstrMensError)
            End If
        End Sub
        Friend Sub SAdicioneComandos()
            Dim lstrLinea As String
            Dim lstrPropiedades() As String
            Dim lstrMensError As String = String.Empty
            Dim lobjComando As ClsComando
            lstrLinea = ObjPadre.ObjPadre.FstrLineaScript
            Do While lstrLinea <> CSTRNComandos.Insert(0, "</") & ">"
                If ObjPadre.ObjPadre.StrNombreNodo = CSTRNComando Then
                    lstrPropiedades = ClsEstructuraBD.FstrPropiedades(lstrLinea)
                    lobjComando = New ClsComando(Me)
                    lobjComando.SAsignePropComado(lstrPropiedades)
                    ColComandos.Add(lobjComando)
                Else
                    lstrMensError = "La linea '" & lstrLinea & vbTab & "' no cumple con el esquema de la Aplicación!"
                    Exit Do
                End If
                lstrLinea = ObjPadre.ObjPadre.FstrLineaScript()
            Loop
            If Not String.IsNullOrEmpty(lstrMensError) Then
                Throw New ArchivoXmlPanException(lstrMensError)
            End If
        End Sub
        Friend Sub SAdicioneColumnas()
            Dim lstrLinea As String
            Dim lstrPropiedades() As String
            Dim lstrMensError As String = String.Empty
            Dim lobjColumna As ClsColumna
            lstrLinea = ObjPadre.ObjPadre.FstrLineaScript
            Do While lstrLinea <> CSTRNColumnas.Insert(0, "</") & ">"
                If ObjPadre.ObjPadre.StrNombreNodo = CSTRNColumna Then
                    lstrPropiedades = ClsEstructuraBD.FstrPropiedades(lstrLinea)
                    lobjColumna = New ClsColumna(Me)
                    lobjColumna.SAsignePropColumna(lstrPropiedades)
                    If ColColumnas.Contains(lobjColumna.StrNombre) Then
                        lstrMensError = "La Columna '" & lobjColumna.StrNombre & "' esta duplicada" _
                                & " en la definición de la tabla '" & MstrNombre & "'"
                        Exit Do
                    Else
                        ColColumnas.Add(lobjColumna, lobjColumna.StrNombre)
                    End If
                Else
                    lstrMensError = "La linea '" & lstrLinea & vbTab & "' no cumple con el esquema de la Aplicación!"
                    Exit Do
                End If
                lstrLinea = ObjPadre.ObjPadre.FstrLineaScript
            Loop
            If Not String.IsNullOrEmpty(lstrMensError) Then
                Throw New ArchivoXmlPanException(lstrMensError)
            End If
        End Sub
        Friend Sub SAdicioneColumna(astrNombreColumna As String,
                ablnRequerido As Boolean, astrTipoDato As String,
                ablnAutoNumerico As Boolean, astrLongitud As String,
                astrCharacterSetCol As String, astrCollationNameCol As String,
                aobjValorDefault As Object, astrComentario As String)
            Dim lstrPropiedades(2) As String
            Dim i As Byte = 2
            Dim lobjColumna As New ClsColumna(Me)
            lstrPropiedades(0) = "Nombre&" & astrNombreColumna
            lstrPropiedades(1) = "Requerido&"
            If ablnRequerido Then
                lstrPropiedades(1) &= "S"
            Else
                lstrPropiedades(1) &= "N"
            End If
            lstrPropiedades(2) = "TipoDato&" & astrTipoDato
            If ablnAutoNumerico Then
                i += 1
                ReDim Preserve lstrPropiedades(i)
                lstrPropiedades(i) = "AutoNumerico&S"
            End If
            If Not String.IsNullOrEmpty(astrLongitud) Then
                i += 1
                ReDim Preserve lstrPropiedades(i)
                lstrPropiedades(i) = "Longitud&" & astrLongitud
            End If
            If Not String.IsNullOrEmpty(astrCharacterSetCol) Then
                i += 1
                ReDim Preserve lstrPropiedades(i)
                lstrPropiedades(i) = "CharacterSet&" & astrCharacterSetCol
            End If
            If Not String.IsNullOrEmpty(astrCollationNameCol) Then
                i += 1
                ReDim Preserve lstrPropiedades(i)
                lstrPropiedades(i) = "CollationName&" & astrCollationNameCol
            End If
            If Not (IsNothing(aobjValorDefault) OrElse String.IsNullOrEmpty(aobjValorDefault)) Then
                i += 1
                ReDim Preserve lstrPropiedades(i)
                lstrPropiedades(i) = "ValorDefecto&" & aobjValorDefault.ToString
            End If
            If Not String.IsNullOrEmpty(astrComentario) Then
                i += 1
                ReDim Preserve lstrPropiedades(i)
                lstrPropiedades(i) = "Comentario&" & astrComentario
            End If
            lobjColumna.SAsignePropColumna(lstrPropiedades)
            ColColumnas.Add(lobjColumna, astrNombreColumna)
        End Sub
        Friend Sub SAdicioneColumna(aobjColumna As ClsColumna)
            If Not ColColumnas.Contains(aobjColumna.StrNombre) Then
                ColColumnas.Add(aobjColumna, aobjColumna.StrNombre)
            End If
        End Sub
        Friend Sub SCambieColumna(ByRef aobjColumnaOriginal As ClsColumna,
                aobjColumnaNueva As ClsColumna)
            If ColColumnas.Contains(aobjColumnaOriginal.StrNombre) Then
                ColColumnas.Remove(aobjColumnaOriginal.StrNombre)
            End If
            If Not ColColumnas.Contains(aobjColumnaNueva.StrNombre) Then
                ColColumnas.Add(aobjColumnaNueva, aobjColumnaNueva.StrNombre)
            End If
        End Sub
        Friend Sub SRenombreColumna(astrNombreColumna As String,
                astrNuevoNombreColumna As String)
            Dim lobjColumna As ClsColumna
            If ColColumnas.Contains(astrNombreColumna) Then
                lobjColumna = ColColumnas(astrNombreColumna)
                lobjColumna.StrNombre = astrNuevoNombreColumna
                ColColumnas.Remove(astrNombreColumna)
                ColColumnas.Add(lobjColumna, lobjColumna.StrNombre)
            End If
        End Sub
        Friend Sub SRemuevaColumna(astrNombreColumna As String)
            If ColColumnas.Contains(astrNombreColumna) Then
                ColColumnas.Remove(astrNombreColumna)
            End If
        End Sub
        Friend Sub SAdicioneIndices()
            Dim lstrLinea As String
            Dim lstrMensError As String = String.Empty
            Dim lstrPropiedades() As String
            Dim lobjIndice As ClsIndice = Nothing
            lstrLinea = ObjPadre.ObjPadre.FstrLineaScript
            Do While lstrLinea <> CSTRNIndices.Insert(0, "</") & ">"
                Select Case ObjPadre.ObjPadre.StrNombreNodo
                    Case CSTRNIndice
                        If lstrLinea <> CSTRNIndice.Insert(0, "</") & ">" Then
                            lstrPropiedades = ClsEstructuraBD.FstrPropiedades(lstrLinea)
                            lobjIndice = New ClsIndice(Me)
                            lobjIndice.SAsignePropIndice(lstrPropiedades)
                            If ColIndices.Contains(lobjIndice.StrNombre) Then
                                lstrMensError = "El indice '" & lobjIndice.StrNombre & "' esta duplicado" _
                                        & " en la definición de la tabla '" & MstrNombre & "'"
                                Exit Do
                            Else
                                ColIndices.Add(lobjIndice, lobjIndice.StrNombre)
                            End If
                        End If
                    Case CSTRNColumnasIndice
                        lobjIndice.SAdicioneColumnasIndice()
                    Case Else
                        lstrMensError = "La linea '" & lstrLinea & vbTab & "' no cumple con el esquema de la Aplicación!"
                        Exit Do
                End Select
                lstrLinea = ObjPadre.ObjPadre.FstrLineaScript
            Loop
            If Not String.IsNullOrEmpty(lstrMensError) Then
                Throw New ArchivoXmlPanException(lstrMensError)
            End If
        End Sub
        Friend Function FobjNuevoIndice(astrNombreIndice As String,
                ablnPrincipal As Boolean, ablnUnico As Boolean) As ClsIndice
            Dim lobjIndice As ClsIndice
            Dim lstrPropiedades(2) As String
            lstrPropiedades(0) = "Nombre&" & astrNombreIndice
            If ablnPrincipal Then
                lstrPropiedades(1) = "Principal&" & "S"
            Else
                lstrPropiedades(1) = "Principal&" & "N"
            End If
            If ablnUnico Then
                lstrPropiedades(2) = "Unico&" & "S"
            Else
                lstrPropiedades(2) = "Unico&" & "N"
            End If
            lobjIndice = New ClsIndice(Me)
            lobjIndice.SAsignePropIndice(lstrPropiedades)
            ColIndices.Add(lobjIndice, lobjIndice.StrNombre)
            Return lobjIndice
        End Function
        Friend Sub SAdicioneIndice(aobjIndice As ClsIndice)
            If Not ColIndices.Contains(aobjIndice.StrNombre) Then
                ColIndices.Add(aobjIndice, aobjIndice.StrNombre)
            End If
        End Sub
        Friend Sub SRemuevaIndice(astrNombreIndice As String)
            If ColIndices.Contains(astrNombreIndice) Then
                ColIndices.Remove(astrNombreIndice)
            End If
        End Sub
        Friend Sub SVerifiqueColumnasRequeridas()
            Dim lobjColumnaIndice As ClsColumnaIndice
            Dim i As UShort
            For Each lobjIndice As ClsIndice In ColIndices
                If lobjIndice.BlnPrincipal OrElse lobjIndice.BlnUnico Then
                    For i = 1 To lobjIndice.ColColumnasIndice.Count
                        lobjColumnaIndice = lobjIndice.ColColumnasIndice(i)
                        If Not FblnColumnaEsRequerida(lobjColumnaIndice.StrNombre) Then
                            Throw New ArchivoXmlPanException("La Columna " & lobjColumnaIndice.StrNombre & " de la tabla " &
                                    StrNombre & " debe declararse como requerido en el archivo xml" & vbCrLf &
                                    " por cuanto forma parte del indice principal")
                            Exit For
                        End If
                    Next i
                End If
            Next
        End Sub
        Private Function FblnColumnaEsRequerida(astrNombreColumna As String) As Boolean
            Dim lobjColumna As ClsColumna = ColColumnas(astrNombreColumna)
            Return lobjColumna.BlnRequerido
        End Function
        Friend Sub SAdicioneRegistros()
            Dim lstrMensError As String = String.Empty
            Dim lstrLinea As String
            Dim lobjRegistro As ClsRegistro
            Dim lstrPropiedades() As String
            Dim lstrKeyComponente As String
            Dim lstrKey As String
            lstrLinea = ObjPadre.ObjPadre.FstrLineaScript
            Do While lstrLinea <> CSTRNRegistros.Insert(0, "</") & ">"
                If ObjPadre.ObjPadre.StrNombreNodo = CSTRNRegistro Then
                    lstrPropiedades = ClsEstructuraBD.FstrPropiedades(lstrLinea)
                    lstrKey = String.Empty
                    For Each lstrPropiedad As String In lstrPropiedades
                        lstrKeyComponente = lstrPropiedad.Substring(lstrPropiedad.IndexOf("&") + 1)
                        lstrKey &= lstrKeyComponente
                    Next
                    lobjRegistro = New ClsRegistro(Me)
                    lobjRegistro.SAsignePropRegistro(lstrPropiedades)
                    If ColRegistros.Contains(lstrKey) Then
                        lstrMensError = "El Registro " & lstrKey & " esta reptido en el archivo de esquema Xml"
                        Exit Do
                    Else
                        ColRegistros.Add(lobjRegistro, lstrKey)
                    End If
                Else
                    lstrMensError = "La linea '" & lstrLinea & vbTab & "' no cumple con el esquema de la Aplicación!"
                    Exit Do
                End If
                lstrLinea = ObjPadre.ObjPadre.FstrLineaScript
            Loop
            If Not String.IsNullOrEmpty(lstrMensError) Then
                Throw New ArchivoXmlPanException(lstrMensError)
            End If
        End Sub
        Friend Function FcolColumnasRef() As Collection
            If Me.ColIndices.Count > 0 Then
                For Each lobjIndice As ClsIndice In Me.ColIndices
                    If lobjIndice.BlnPrincipal Then
                        Return lobjIndice.ColColumnasIndice
                    End If
                Next
            End If
            Return Nothing
        End Function
        Friend Sub SConviertaCollationName(astrCollationName As String)
            MstrCollationName = astrCollationName
        End Sub
    End Class
    Friend Class ClsColumna
        Private MblnRequerido As Boolean = False
        Private MblnAutonumerico As Boolean = False
        Private MstrCharacterSet As String = String.Empty
        Private MstrCollationName As String = String.Empty
        Private MstrNombre As String = String.Empty
        Private MstrTipoDato As String = String.Empty
        Private MobjValorDefault As Object = Nothing
        Private MstrLongitud As String = String.Empty
        ' ===========
        ' PROPIEDADES
        ' ===========
        Public Sub New(aobjPadre As ClsTabla)
            ObjPadre = aobjPadre
        End Sub
        Friend ReadOnly Property ObjPadre() As ClsTabla = Nothing
        Friend Property StrComentario() As String = String.Empty
        Friend ReadOnly Property BlnAutoNumerico() As Boolean
            Get
                Return MblnAutonumerico
            End Get
        End Property
        Friend ReadOnly Property BlnRequerido() As Boolean
            Get
                Return MblnRequerido
            End Get
        End Property
        Friend ReadOnly Property StrCharacterSet As String
            Get
                Return MstrCharacterSet
            End Get
        End Property
        Friend ReadOnly Property StrCollationName As String
            Get
                Return MstrCollationName
            End Get
        End Property
        Friend Property StrNombre() As String
            Get
                Return MstrNombre
            End Get
            Set(astrValor As String)
                If Not astrValor = MstrNombre Then
                    MstrNombre = astrValor
                End If
            End Set
        End Property
        Friend ReadOnly Property StrTipoDatos() As String
            Get
                Return MstrTipoDato
            End Get
        End Property
        Friend ReadOnly Property StrLongitud As String
            Get
                Return MstrLongitud
            End Get
        End Property
        Friend ReadOnly Property ObjValorDefault() As Object
            Get
                Return MobjValorDefault
            End Get
        End Property
        ' ==========================
        ' PROCEDIMIENTOS Y FUNCIONES
        ' ==========================
        Friend Sub SAsignePropColumna(astrPropiedades() As String)
            Dim lstrNomProp As String
            Dim lstrValProp As String
            Dim lstrMensError As String = String.Empty
            Dim i As Integer
            For i = 0 To astrPropiedades.GetUpperBound(0)
                lstrNomProp = astrPropiedades(i).Split("&")(0)
                lstrValProp = astrPropiedades(i).Split("&")(1)
                If Not FblnAsignoPropCol(lstrNomProp, lstrValProp, lstrMensError) Then
                    If String.IsNullOrEmpty(lstrMensError) Then
                        Select Case lstrNomProp
                            Case "AutoNumerico"
                                If lstrValProp = "N" Then
                                    MblnAutonumerico = False
                                ElseIf lstrValProp = "S" Then
                                    MblnAutonumerico = True
                                Else
                                    lstrMensError = "El valor de la Propiedad 'Autonumerico' de la columna no es valido."
                                    Exit For
                                End If
                            Case "Longitud"
                                If IsNumeric(lstrValProp) Then
                                    MstrLongitud = lstrValProp
                                Else
                                    lstrMensError = "El Valor de la Propiedad '" & lstrNomProp & "' no es valido."
                                    Exit For
                                End If
                            Case Else
                                lstrMensError = "El Nombre de Propiedad '" & lstrNomProp & "' no es valido."
                                Exit For
                        End Select
                    Else
                        Exit For
                    End If
                End If
            Next
            If Not String.IsNullOrEmpty(lstrMensError) Then
                Throw New ArchivoXmlPanException(lstrMensError)
            End If
        End Sub
        Private Function FblnAsignoPropCol(astrNomPro As String,
                astrValPro As String, ByRef astrMens As String) As Boolean
            Dim lblnAsigno = True
            Select Case astrNomPro
                Case "Nombre"
                    MstrNombre = astrValPro
                Case "Requerido"
                    If astrValPro = "N" Then
                        MblnRequerido = False
                    ElseIf astrValPro = "S" Then
                        MblnRequerido = True
                    Else
                        lblnAsigno = False
                        astrMens = "El valor de la Propiedad 'Requerido' de la Columna no es soportado por la Aplicación"
                    End If
                Case "TipoDato"
                    MstrTipoDato = astrValPro
                Case "CharacterSet"
                    MstrCharacterSet = astrValPro
                Case "CollationName"
                    MstrCollationName = astrValPro
                Case "ValorDefecto"
                    MobjValorDefault = astrValPro
                Case "Comentario"
                    StrComentario = astrValPro
                Case Else
                    lblnAsigno = False
            End Select
            Return lblnAsigno
        End Function
        Friend Function FblnEsNumerico() As Boolean
            Dim lblnEsNumerico As Boolean
            lblnEsNumerico = (MstrTipoDato = "BYTE" OrElse MstrTipoDato = "UBYTE" OrElse MstrTipoDato = "SHORT" OrElse
                    MstrTipoDato = "USHORT" OrElse MstrTipoDato = "INTEGER" OrElse MstrTipoDato = "UINTEGER" OrElse
                    MstrTipoDato = "LONG" OrElse MstrTipoDato = "ULONG" OrElse MstrTipoDato = "CURRENCY" OrElse
                    MstrTipoDato = "SINGLE" OrElse MstrTipoDato = "DOUBLE")
            Return lblnEsNumerico
        End Function
        Friend Function FblnEsCadena() As Boolean
            Dim lblnEsCadena As Boolean
            lblnEsCadena = (MstrTipoDato = "STRING" OrElse MstrTipoDato = "VSTRING")
            Return lblnEsCadena
        End Function
    End Class
    Friend Class ClsIndice
        ' Constantes
        Private Const CSTRNColumnasIndice As String = "ColumnasIndice"
        Private Const CSTRNColumnaIndice As String = "ColumnaIndice"
        ' Variables
        Private MstrNombre As String = String.Empty
        Private MblnPrincipal As Boolean = False
        Private MblnUnico As Boolean = False
        Public Sub New(aobjPadre As ClsTabla)
            ObjPadre = aobjPadre
            ColColumnasIndice = New Collection
        End Sub
        ' ===========
        ' PROPIEDADES
        ' ===========
        Friend ReadOnly Property ObjPadre() As ClsTabla = Nothing
        Friend ReadOnly Property StrNombre() As String
            Get
                Return MstrNombre
            End Get
        End Property
        Friend ReadOnly Property BlnPrincipal() As Boolean
            Get
                Return MblnPrincipal
            End Get
        End Property
        Friend ReadOnly Property BlnUnico() As Boolean
            Get
                Return MblnUnico
            End Get
        End Property
        Friend ReadOnly Property ColColumnasIndice() As Collection = Nothing
        ' ==========================
        ' PROCEDIMIENTOS Y FUNCIONES
        ' ==========================
        Friend Sub SAsignePropIndice(astrPropiedades() As String)
            Dim lstrNomProp As String
            Dim lstrValProp As String
            Dim lstrMensError As String = String.Empty
            Dim i As Integer
            For i = 0 To astrPropiedades.GetUpperBound(0)
                lstrNomProp = astrPropiedades(i).Split("&")(0)
                lstrValProp = astrPropiedades(i).Split("&")(1)
                Select Case lstrNomProp
                    Case "Nombre"
                        MstrNombre = lstrValProp
                    Case "Principal"
                        If lstrValProp = "N" Then
                            MblnPrincipal = False
                        ElseIf lstrValProp = "S" Then
                            MblnPrincipal = True
                        Else
                            lstrMensError = "El valor de la Propiedad 'Principal' de la tabla no es soportado por la Aplicación'"
                            Exit For
                        End If
                    Case "Unico"
                        If lstrValProp = "N" Then
                            MblnUnico = False
                        ElseIf lstrValProp = "S" Then
                            MblnUnico = True
                        Else
                            lstrMensError = "El valor de la Propiedad 'Unico' de la tabla no es soportado por la Aplicación'"
                            Exit For
                        End If
                    Case "NulosIndice"
                        ' Solo se usa para Access
                    Case Else
                        lstrMensError = "EL nombre de Propiedad '" & lstrNomProp & "' no es valido."
                        Exit For
                End Select
            Next
            If Not String.IsNullOrEmpty(lstrMensError) Then
                Throw New ArchivoXmlPanException(lstrMensError)
            End If
        End Sub
        Friend Sub SAdicioneColumnasIndice()
            Dim lstrLinea As String
            Dim lstrPropiedades() As String
            Dim lstrMensError As String = String.Empty
            Dim lobjColumnaIndice As ClsColumnaIndice
            lstrLinea = ObjPadre.ObjPadre.ObjPadre.FstrLineaScript
            Do While lstrLinea <> CSTRNColumnasIndice.Insert(0, "</") & ">"
                If ObjPadre.ObjPadre.ObjPadre.StrNombreNodo = CSTRNColumnaIndice Then
                    lstrPropiedades = ClsEstructuraBD.FstrPropiedades(lstrLinea)
                    lobjColumnaIndice = New ClsColumnaIndice(Me)
                    lobjColumnaIndice.SAsignePropColumnaIndice(lstrPropiedades)
                    If ColColumnasIndice.Contains(lobjColumnaIndice.StrNombre) Then
                        lstrMensError = "La Columna de Indice '" & lobjColumnaIndice.StrNombre &
                                "' esta duplicado en la definición del Indice" & vbCrLf &
                                "'" & StrNombre & "' de la Tabla '" & ObjPadre.StrNombre & "'."
                        Exit Do
                    End If
                    ColColumnasIndice.Add(lobjColumnaIndice, lobjColumnaIndice.StrNombre)
                Else
                    lstrMensError = "La linea '" & lstrLinea & vbTab & "' no cumple con el esquema de la Aplicación!"
                    Exit Do
                End If
                lstrLinea = ObjPadre.ObjPadre.ObjPadre.FstrLineaScript
            Loop
            If Not String.IsNullOrEmpty(lstrMensError) Then
                Throw New ArchivoXmlPanException(lstrMensError)
            End If
        End Sub
        Friend Sub SAdicioneColumnaIndice(astrNombreColumnaIndice As String,
                astrOrdenIndice As String)
            Dim lobjColumnaIndice As ClsColumnaIndice
            Dim lstrPropiedades(1) As String
            lstrPropiedades(0) = "Nombre&" & astrNombreColumnaIndice
            lstrPropiedades(1) = "Orden&" & astrOrdenIndice
            lobjColumnaIndice = New ClsColumnaIndice(Me)
            lobjColumnaIndice.SAsignePropColumnaIndice(lstrPropiedades)
            ColColumnasIndice.Add(lobjColumnaIndice, lobjColumnaIndice.StrNombre)
        End Sub
    End Class
    Friend Class ClsRegistro
        ' Variables de propiedad
        Public Sub New(aobjPadre As ClsTabla)
            ObjPadre = aobjPadre
            ColNombresColumna = New Collection
            ColDatos = New Collection
            ColTipoDatos = New Collection
        End Sub
        ' PROPIEDADES
        Friend ReadOnly Property ObjPadre() As ClsTabla = Nothing
        Friend ReadOnly Property ColNombresColumna() As Collection = Nothing
        Friend ReadOnly Property ColDatos() As Collection = Nothing
        Friend ReadOnly Property ColTipoDatos() As Collection = Nothing
        ' PROCEDIMIENTOS Y FUNCIONES
        Friend Sub SAsignePropRegistro(astrPropiedades() As String)
            Dim lstrDato As String
            Dim lstrNombreColumna As String
            Dim lobjColumna As ClsColumna
            For i As Integer = 0 To astrPropiedades.GetUpperBound(0)
                lstrNombreColumna = astrPropiedades(i).Split("&")(0)
                ColNombresColumna.Add(lstrNombreColumna, lstrNombreColumna)
                lstrDato = astrPropiedades(i).Split("&")(1)
                ColDatos.Add(lstrDato, lstrNombreColumna)
                If Me.ObjPadre.ColColumnas.Contains(lstrNombreColumna) Then
                    lobjColumna = Me.ObjPadre.ColColumnas(lstrNombreColumna)
                    ColTipoDatos.Add(lobjColumna.StrTipoDatos, lstrNombreColumna)
                Else
                    Throw New ArchivoXmlPanException("Se esta tratando de asignar un valor a un campo que no existe")
                End If
            Next
        End Sub
        Friend Sub SAsigneValorCampo(astrNombreColumna As String, aobjDato As Object)
            ColDatos.Remove(astrNombreColumna)
            ColDatos.Add(aobjDato, astrNombreColumna)
        End Sub
    End Class
    Friend Class ClsComando
        ' Variables de propiedad
        Private MenuTipoComando As EnuTipoComando = EnuTipoComando.None
        Private MstrCondicion As String = String.Empty
        Private MstrAccion As String = String.Empty
        Private MstrParametros As String = String.Empty
        Private MenuSecuencia As EnuSecuenciaAccion = EnuSecuenciaAccion.enuInicio
        Public Sub New(aobjPadre As Object)
            ObjPadre = aobjPadre
        End Sub
        ' ===========
        ' PROPIEDADES
        ' ===========
        Friend ReadOnly Property ObjPadre As Object = Nothing
        Friend ReadOnly Property EnuTipoComando() As EnuTipoComando
            Get
                Return MenuTipoComando
            End Get
        End Property
        Friend ReadOnly Property StrCondicion() As String
            Get
                Return MstrCondicion
            End Get
        End Property
        Friend ReadOnly Property StrAccion() As String
            Get
                Return MstrAccion
            End Get
        End Property
        Friend ReadOnly Property StrParametros() As String
            Get
                Return MstrParametros
            End Get
        End Property
        Friend ReadOnly Property EnuSecuencia() As EnuSecuenciaAccion
            Get
                Return MenuSecuencia
            End Get
        End Property
        ' ==========================
        ' PROCEDIMIENTOS Y FUNCIONES
        ' ==========================
        Friend Sub SAsignePropComado(astrPropiedades() As String)
            Dim lstrNomProp As String
            Dim lstrValProp As String
            Dim lstrMensError As String = String.Empty
            Dim i As Integer
            For i = 0 To astrPropiedades.GetUpperBound(0)
                lstrNomProp = astrPropiedades(i).Split("&")(0)
                lstrValProp = astrPropiedades(i).Split("&")(1)
                Select Case lstrNomProp
                    Case "Condicion"
                        MstrCondicion = lstrValProp
                    Case "Tipo"
                        Select Case lstrValProp
                            Case "Propio"
                                MenuTipoComando = EnuTipoComando.enuPropio
                            Case "SQL"
                                MenuTipoComando = EnuTipoComando.enuInstruccionSQL
                            Case "Proc"
                                MenuTipoComando = EnuTipoComando.enuLlamadaProc
                            Case Else
                                lstrMensError = "El valor de la Propiedad '" & lstrValProp & "' no es soportado por la Aplicación."
                                Exit For
                        End Select
                    Case "Accion"
                        MstrAccion = lstrValProp
                    Case "Parametros"
                        MstrParametros = lstrValProp
                    Case "Secuencia"
                        Select Case lstrValProp
                            Case "Inicio"
                                MenuSecuencia = EnuSecuenciaAccion.enuInicio
                            Case "Fin"
                                MenuSecuencia = EnuSecuenciaAccion.enuFin
                            Case Else
                                lstrMensError = "El valor de la Propiedad '" & lstrValProp & "' no es soportado por la Aplicación."
                                Exit For
                        End Select
                    Case Else
                        lstrMensError = "El nombre de propiedad '" & lstrNomProp & "' no es valido."
                        Exit For
                End Select
            Next i
            If Not String.IsNullOrEmpty(lstrMensError) Then
                Throw New ArchivoXmlPanException(lstrMensError)
            End If
        End Sub
    End Class
    Friend Class ClsColumnaIndice
        Private MstrNombre As String = String.Empty
        Private MblnAscendente As Boolean = True
        ' ===========
        ' PROPIEDADES
        ' ===========
        Public Sub New(aobjPadre As ClsIndice)
            ObjPadre = aobjPadre
        End Sub
        Friend ReadOnly Property ObjPadre() As ClsIndice = Nothing
        Friend ReadOnly Property StrNombre() As String
            Get
                Return MstrNombre
            End Get
        End Property
        Friend ReadOnly Property BlnAscendente() As Boolean
            Get
                Return MblnAscendente
            End Get
        End Property
        ' ==========================
        ' PROCEDIMIENTOS Y FUNCIONES
        ' ==========================
        Friend Sub SAsignePropColumnaIndice(astrPropiedades() As String)
            Dim lstrNomProp As String
            Dim lstrValProp As String
            Dim lstrMensError As String = String.Empty
            For i As Integer = 0 To astrPropiedades.GetUpperBound(0)
                lstrNomProp = astrPropiedades(i).Split("&")(0)
                lstrValProp = astrPropiedades(i).Split("&")(1)
                Select Case lstrNomProp
                    Case "Nombre"
                        MstrNombre = lstrValProp
                    Case "Orden"
                        If lstrValProp = "ASC" Then
                            MblnAscendente = True
                        ElseIf lstrValProp = "DEC" Then
                            MblnAscendente = False
                        Else
                            lstrMensError = "El valor de la propiedad Orden(" & lstrValProp & ") no es valido"
                            Exit For
                        End If
                    Case Else
                        lstrMensError = "El nombre de propiedad '" & lstrNomProp & "' no es valido."
                        Exit For
                End Select
            Next i
            If Not String.IsNullOrEmpty(lstrMensError) Then
                Throw New ArchivoXmlPanException(lstrMensError)
            End If
        End Sub
    End Class
#End Region
End Namespace