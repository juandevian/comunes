Imports System.Net
Imports System.Net.Mail
Friend Class ClsPanorama
#Region "Definiciones"
#Region "Constantes"
    Private Const HKEY_LOCAL_MACHINE As Integer = &H80000002
    Private Const SYNCHRONIZE = &H100000

    Private Const READ_CONTROL = &H20000
    Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
    Private Const KEY_QUERY_VALUE = &H1
    Private Const KEY_ENUMERATE_SUB_KEYS = &H8
    Private Const KEY_NOTIFY = &H10

    Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or
            KEY_NOTIFY) And (Not SYNCHRONIZE))
    Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
    Private Const KEY_SET_VALUE = &H2
    Private Const KEY_CREATE_SUB_KEY = &H4
    Private Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY

    Private Const CINTDiasPrueba As Integer = 60
#End Region
    ' Variables
    Private ReadOnly MblnRegistrado As Boolean = False
    ' LogApp
    Private MdtbLogApp As DataTable
    ' Ubicacion Actual
    Private MobjCarpetaActual As ClsCarpeta = Nothing
#End Region

#Region "Constructores"
    Friend Sub New(aobjRegistro As Object)
        If aobjRegistro Is Nothing OrElse Not (aobjRegistro.GetType.Name = "String" AndAlso
                aobjRegistro = GCOBJREGISTRO) Then
            Throw New ModuloNoRegistradoPanException()
        Else
            MblnRegistrado = True
        End If
    End Sub
#End Region

#Region "Propiedades"
    Friend Property ObjUsuarioActual As ClsUsuario = Nothing
    Friend Property ObjAppActual As ClsAplicacion = Nothing
    Friend ReadOnly Property BlnRegistrado() As Boolean
        Get
            Return MblnRegistrado
        End Get
    End Property
    Friend ReadOnly Property ObjCarpetaActual As ClsCarpeta
        Get
            Return MobjCarpetaActual
        End Get
    End Property
#End Region

#Region "Procedimientos" 'Ok
    '''<summary>
    ''' Actualiza los objetos de una colección cuyo estado de actualización (enuEstadoActualizacion) sea diferente
    ''' a ninguno (enuEstadoObjetoDef.enuConsultando)
    ''' </summary>
    ''' <param name="acolObjetos">Colección que contiene los objetos a actualizar</param>
    Friend Shared Sub SActualiceCol(acolObjetos As Collection)
        If acolObjetos Is Nothing Then
            Throw New ArgumentNullException(NameOf(acolObjetos))
        End If
        If acolObjetos.Count > 0 Then
            GobjPanDat.SControleProcesoObj(True)
            Try
                For Each lobjObjeto As ClsCBObjetoPan In acolObjetos
                    If lobjObjeto.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                        lobjObjeto.SActualice(True)
                    End If
                Next
            Catch ex As PanLException
                Throw
            Catch ex As PanDatException
                Throw
            Catch ex As Exception
                Throw
            End Try
            GobjPanDat.SControleProcesoObj(False)
        End If
    End Sub
    ''' <summary>
    ''' Actualiza los objetos de una colección cuyo estado de actualización (enuEstadoActualizacion) sea diferente
    ''' a ninguno (enuEstadoObjetoDef.enuConsultando)
    ''' </summary>
    ''' <param name="acolObjetos">Colección que contiene los objetos a actualizar</param>
    ''' <param name="ablnLimpiarCol">Parametro que indica si se eliminan los objetos de la colleción después de
    ''' actualizados</param>
    ''' <remarks></remarks>
    Shared Sub SActualiceCol(acolObjetos As Collection, ablnLimpiarCol As Boolean)
        If acolObjetos Is Nothing Then
            Throw New ArgumentNullException(NameOf(acolObjetos))
        End If
        If acolObjetos.Count > 0 Then
            For Each lobjObjeto As ClsCBObjetoPan In acolObjetos
                If lobjObjeto.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                    lobjObjeto.SActualice(True)
                End If
            Next
            If ablnLimpiarCol Then
                acolObjetos.Clear()
            End If
        End If
    End Sub
    Friend Shared Function FblnSuprimioCol(acolObjetos As Collection) As Boolean
        Dim lblnSuprimioCol = True
        If acolObjetos.Count > 0 Then
            For i = acolObjetos.Count To 1 Step -1
                Dim lobjObjeto As ClsCBObjetoPan = acolObjetos(i)
                lblnSuprimioCol = lobjObjeto.FblnSuprimio()
                If Not lblnSuprimioCol Then
                    Exit For
                End If
            Next
            If lblnSuprimioCol Then
                acolObjetos.Clear()
            End If
        End If
        Return lblnSuprimioCol
    End Function
    Friend Shared Sub SEncriptePassword(ByRef astrPassword As String)
        Dim i As Integer = 1, lblnPrimera = True
        Dim lstrPassword As String = String.Empty, lstrParte As String, lentChar As Integer
        Dim lentEx As Integer() = {9, 10, 13, 32, 127, 129, 141, 143, 144, 157, 160, 173}
        Dim lstrClave As String = "adhadhadhadhadhadhadhadhadhadhadhadhadha"
        For Each lchr As Char In astrPassword
            If lblnPrimera Then
                lentChar = Asc(lchr) + Asc(lstrClave.Chars(i - 1)) + Int((1.525 ^ (i + 1)))
                lblnPrimera = lentChar <= 254
                If Not lblnPrimera Then
                    i = 1
                    lentChar = Asc(lchr) + Asc(lstrClave.Chars(i - 1)) + Math.Round((1.75 ^ (i + 0.5)), 0)
                End If
            Else
                lentChar = Asc(lchr) + Asc(lstrClave.Chars(i - 1)) + Math.Round((1.75 ^ (i + 0.5)), 0)
            End If
            If lentChar <= 254 AndAlso Not lentEx.Contains(lentChar) Then
                lstrParte = Chr(lentChar)
                lstrPassword &= lstrParte
            End If
            i += 1
        Next
        astrPassword = lstrPassword
    End Sub
    Friend Shared Function FobjNuevoTercero(adblIdTercero As Double) As ClsTercero
        Dim lobjNuevoTercero As New ClsTercero(EnuModoInstanciaObjDef.enuNavegable)
        lobjNuevoTercero.SCreeObj(Nothing)
        lobjNuevoTercero.ObjIdTerceroDbl.ObjValorPro = adblIdTercero
        Return lobjNuevoTercero
    End Function
    Shared Function FobjValorProNormal(aobjValorPro As Object, aenuTipoValor As EnuTipoValor) As Object
        If aobjValorPro Is Nothing Then
            Throw New ArgumentNullException(NameOf(aobjValorPro))
        End If
        Dim lobjValorPro As Object = aobjValorPro
        If IsNumeric(lobjValorPro) AndAlso aenuTipoValor <> EnuTipoValor.enuString Then
            lobjValorPro = Val(aobjValorPro)
        End If
        Select Case aenuTipoValor
            Case EnuTipoValor.enuBoolean
                lobjValorPro = FblnValorProBol(lobjValorPro)
            Case EnuTipoValor.enuByte
                lobjValorPro = FbytValorProByte(lobjValorPro)
            Case EnuTipoValor.enuDate
                lobjValorPro = FdtmValorProFecha(lobjValorPro)
            Case EnuTipoValor.enuDecimal
                lobjValorPro = FdecValorProDecimal(lobjValorPro)
            Case EnuTipoValor.enuDouble
                lobjValorPro = FdblValorProDouble(lobjValorPro)
            Case EnuTipoValor.enuImagen
                '
            Case EnuTipoValor.enuInteger, EnuTipoValor.enuUInteger
                lobjValorPro = FentValorProEntero(lobjValorPro)
            Case EnuTipoValor.enuLong, EnuTipoValor.enuULong
                lobjValorPro = FlngValorProLng(lobjValorPro)
            Case EnuTipoValor.enuShort, EnuTipoValor.enuUShort
                lobjValorPro = FshrValorProShr(lobjValorPro)
            Case EnuTipoValor.enuSingle
                If IsNumeric(lobjValorPro) Then
                    If Val(lobjValorPro) < Single.MinValue OrElse Val(lobjValorPro) > Single.MaxValue Then
                        lobjValorPro = 0
                    Else
                        lobjValorPro = CType(lobjValorPro, Single)
                    End If
                Else
                    lobjValorPro = 0
                End If
            Case EnuTipoValor.enuString
                If IsNumeric(lobjValorPro) OrElse lobjValorPro.GetType.Name = "String" Then
                    lobjValorPro = Trim(CType(lobjValorPro, String))
                Else
                    lobjValorPro = String.Empty
                End If
        End Select
        Return lobjValorPro
    End Function
    Shared Function FblnValorProBol(aobjValorPro As Object) As Boolean
        If aobjValorPro Is Nothing Then
            Throw New ArgumentNullException(NameOf(aobjValorPro))
        End If
        Dim lobjValorPro = False
        If IsNumeric(aobjValorPro) Then
            If Not aobjValorPro.GetType.Name = "Boolean" Then
                lobjValorPro = (aobjValorPro <> 0)
            End If
        ElseIf aobjValorPro.GetType.Name = "String" Then
            lobjValorPro = FblnValor(aobjValorPro.ToString)
        End If
        Return lobjValorPro
    End Function
    Private Shared Function FblnValor(astrValor As String) As Boolean
        Dim lblnValor = False
        Select Case astrValor.ToUpper
            Case "S", "SI", "YES", "TRUE", "VERDADERO"
                lblnValor = True
            Case "N", "NO", "FALSE", "FALSO"
                lblnValor = False
        End Select
        Return lblnValor
    End Function
    Shared Function FbytValorProByte(aobjValorPro As Object) As Byte
        Dim lbytValorPro As Byte
        If IsNumeric(aobjValorPro) Then
            If aobjValorPro > Byte.MaxValue OrElse aobjValorPro < Byte.MinValue Then
                lbytValorPro = 0
            Else
                lbytValorPro = CType(aobjValorPro, Byte)
            End If
        Else
            lbytValorPro = 0
        End If
        Return lbytValorPro
    End Function
    Shared Function FdtmValorProFecha(aobjValorPro As Object) As Date
        Dim ldtmFecha As Date
        If IsDate(aobjValorPro) Then
            If aobjValorPro < Date.MinValue OrElse aobjValorPro > Date.MaxValue Then
                ldtmFecha = GCDTMFECHANULA
            Else
                ldtmFecha = CType(aobjValorPro, Date)
            End If
        Else
            ldtmFecha = GCDTMFECHANULA
        End If
        Return ldtmFecha
    End Function
    Shared Function FdecValorProDecimal(aobjValorPro As Object) As Decimal
        Dim ldecValorPro As Decimal
        If IsNumeric(aobjValorPro) Then
            If Val(aobjValorPro) < Decimal.MinValue OrElse Val(aobjValorPro) > Decimal.MaxValue Then
                ldecValorPro = 0
            Else
                ldecValorPro = CType(aobjValorPro, Decimal)
            End If
        Else
            ldecValorPro = 0
        End If
        Return ldecValorPro
    End Function
    Shared Function FdblValorProDouble(aobjValorPro As Object) As Double
        Dim ldblValorPro As Double
        If IsNumeric(aobjValorPro) Then
            If Val(aobjValorPro) < Double.MinValue OrElse Val(aobjValorPro) > Double.MaxValue Then
                ldblValorPro = 0
            Else
                ldblValorPro = CType(aobjValorPro, Double)
            End If
        Else
            ldblValorPro = 0
        End If
        Return ldblValorPro
    End Function
    Shared Function FentValorProEntero(aobjValorPro As Object) As Integer
        Dim lentValorPro As Integer
        If IsNumeric(aobjValorPro) Then
            If Val(aobjValorPro) < Decimal.MinValue OrElse Val(aobjValorPro) > Decimal.MaxValue Then
                lentValorPro = 0
            Else
                lentValorPro = CType(aobjValorPro, Integer)
            End If
        Else
            lentValorPro = 0
        End If
        Return lentValorPro
    End Function
    Shared Function FlngValorProLng(aobjValorPro As Object) As Long
        Dim llngValorPro As Long
        If IsNumeric(aobjValorPro) Then
            If Val(aobjValorPro) < Long.MinValue OrElse Val(aobjValorPro) > Long.MaxValue Then
                llngValorPro = 0
            Else
                llngValorPro = CType(Val(aobjValorPro), Long)
            End If
        Else
            llngValorPro = 0
        End If
        Return llngValorPro
    End Function
    Shared Function FshrValorProShr(aobjValorPro As Object) As Short
        Dim lshrValorPro As Short
        If IsNumeric(aobjValorPro) Then
            If Val(aobjValorPro) < Short.MinValue OrElse Val(aobjValorPro) > Short.MaxValue Then
                lshrValorPro = 0
            Else
                lshrValorPro = CType(Val(aobjValorPro), Short)
            End If
        Else
            lshrValorPro = 0
        End If
        Return lshrValorPro
    End Function
    Friend Shared Sub SEscribaArchivoError(astrMensajeError As String)
        Dim lstrMensEx = Date.Now.ToString & ": " & astrMensajeError
        Dim lstrArchivoErr = GstrTrayDatPrg & "errsori.txt"
        Try
            If Not String.IsNullOrEmpty(lstrMensEx) Then
                Using lswArchivoErr = File.AppendText(lstrArchivoErr)
                    lswArchivoErr.WriteLine(lstrMensEx)
                    lswArchivoErr.WriteLine("")
                    lswArchivoErr.Flush()
                End Using
#If DES = 1 Then
                Process.Start("notepad.exe", lstrArchivoErr)
#End If
            End If
        Catch ex As ArgumentException
            Throw
        Catch ex As PathTooLongException
            Throw
        Catch ex As DirectoryNotFoundException
            Throw
        Catch ex As NotSupportedException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Sub
    Friend Shared Function FsrStreamReader(astrArchivo As String) As StreamReader
        Dim lsrArchivo As StreamReader = Nothing
        If String.IsNullOrEmpty(astrArchivo) Then
            Throw New ArgumentNullException(NameOf(astrArchivo))
        End If
        Dim lblnNoHayError = False
        Try
            If My.Computer.FileSystem.FileExists(astrArchivo) Then
                lsrArchivo = File.OpenText(astrArchivo)
            End If
            lblnNoHayError = True
        Catch ex As FileNotFoundException
            Throw
        Catch ex As IOException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not lblnNoHayError Then
                If Not IsNothing(lsrArchivo) Then
                    lsrArchivo.Dispose()
                End If
            End If
        End Try
        Return lsrArchivo
    End Function
    Friend Shared Function FswStreamWriter(astrArchivo As String) As StreamWriter
        Dim lswArchivo As StreamWriter = Nothing
        If String.IsNullOrEmpty(astrArchivo) Then
            Throw New ArgumentNullException(NameOf(astrArchivo))
        End If
        Dim lblnNoHayError = False
        Try
            lswArchivo = File.CreateText(astrArchivo)
            lblnNoHayError = True
        Catch ex As ArgumentException
            Throw
        Catch ex As PathTooLongException
            Throw
        Catch ex As DirectoryNotFoundException
            Throw
        Catch ex As NotSupportedException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not lblnNoHayError AndAlso lswArchivo IsNot Nothing Then
                lswArchivo.Dispose()
            End If
        End Try
        Return lswArchivo
    End Function
    Friend Shared Function FswStreamWriterAppend(astrArchivo As String) As StreamWriter
        Dim lswArchivo As StreamWriter = Nothing
        If String.IsNullOrEmpty(astrArchivo) Then
            Throw New ArgumentNullException(NameOf(astrArchivo))
        End If
        Dim lblnNoHayError = False
        Try
            lswArchivo = File.AppendText(astrArchivo)
            lblnNoHayError = True
        Catch ex As ArgumentException
            Throw
        Catch ex As PathTooLongException
            Throw
        Catch ex As DirectoryNotFoundException
            Throw
        Catch ex As NotSupportedException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not lblnNoHayError Then
                lswArchivo.Dispose()
            End If
        End Try
        Return lswArchivo
    End Function
    Friend Shared Sub SReasigneIdTer(adblIdTerceroActua As Double, adblIdTerceroNuevo As Double)
        Dim lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            Dim lobjBaseDatosPan = ClsPanoramaDat.FobjEstructuraBD(EnuListaAplicaciones.EnuAdministrador)
            Dim lobjBaseDatos = ClsPanoramaDat.FobjEstructuraBD(EnuListaAplicaciones.EnuOrionCop)
            Dim lobjCol As ClsColumna = Nothing
            For Each lobjTabla As ClsTabla In lobjBaseDatosPan.ColTablas
                For i = 1 To lobjTabla.ColColumnas.Count
                    lobjCol = lobjTabla.ColColumnas(i)
                    If lobjCol.StrNombre.Contains("IdTercero") Then
                        SModifiqueIdTer(lobjTabla.StrNombre, lobjCol.StrNombre, adblIdTerceroActua,
                                adblIdTerceroNuevo)
                    End If
                Next
            Next
            For Each lobjTabla As ClsTabla In lobjBaseDatos.ColTablas
                For i = 1 To lobjTabla.ColColumnas.Count
                    lobjCol = lobjTabla.ColColumnas(i)
                    If lobjCol.StrNombre.Contains("IdTercero") Then
                        SModifiqueIdTer(lobjTabla.StrNombre, lobjCol.StrNombre, adblIdTerceroActua,
                                adblIdTerceroNuevo)
                    End If
                    If lobjCol.StrNombre.Contains("AliasContable") Then
                        SModifiqueIdTer(lobjTabla.StrNombre, lobjCol.StrNombre, adblIdTerceroActua,
                                adblIdTerceroNuevo)
                    End If
                Next
            Next
            GobjPanorama.SRegistreAccionLogApp("Tercero", "Reasignó Id. " & adblIdTerceroActua.ToString &
                    " al " & adblIdTerceroNuevo.ToString)
            lblnNoHayError = True
        Catch ex As ProveedorBdPanException
            Throw
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Private Shared Sub SModifiqueIdTer(astrNombreTabla As String, astrNombreCampo As String,
            adblIdTerceroActua As Double, adblIdTerceroNuevo As Double)
        Dim lcolCamposCambio As New Collection, lcolCamposRef As New Collection
        Dim lcolDatosNue As New Collection, lcolDatosRef As New Collection
        lcolCamposCambio.Add(astrNombreCampo)
        lcolDatosNue.Add(adblIdTerceroNuevo, astrNombreCampo)
        lcolCamposRef.Add(astrNombreCampo)
        lcolDatosRef.Add(adblIdTerceroActua, astrNombreCampo)
        GobjPanDat.SActualiceRegistro(astrNombreTabla, lcolCamposCambio, lcolDatosNue,
                lcolCamposRef, lcolDatosRef)
    End Sub
    Friend Shared Function FstrVersionActual(aenuIdAplicacion As EnuListaAplicaciones) As String
        Dim lstrfiltro = Aplicacion.ClsIdAppShr.SstrNombreCampoBd & " = " & aenuIdAplicacion
        Dim ldtbAplicaciones As DataTable = FdtbDataTable(Aplicacion.ClsAplicacion.SstrNombreTabla,
                {Aplicacion.ClsVersionStr.SstrNombreCampoBd}, {{"", ""}}, lstrfiltro)
        Dim lstrVersion As String = ClsPanorama.
                FobjValorCampo(ldtbAplicaciones.Rows(0)(Aplicacion.ClsVersionStr.SstrNombreCampoBd),
                EnuTipoValor.enuString)
        Dim lstrPartes As String() = lstrVersion.Split(".")
        Dim lstrPrte1 = Format(CType(lstrPartes(0), Integer), "00")
        Dim lstrPrte2 = Format(CType(lstrPartes(1), Integer), "00")
        Dim lstrPrte3 = Format(CType(lstrPartes(2), Integer), "000")
        lstrVersion = lstrPrte1 & "." & lstrPrte2 & "." & lstrPrte3
        Return lstrVersion
    End Function
#End Region

#Region "Procedimientos de validación generales"
    ''' <summary>
    ''' Valida que un objeto sea numérico string no buleano; si es numerico, no buleano ni string, lo convierte 
    ''' a string; además valida que su longitud este entre los argumantos pasados.
    ''' debe cumplir los valores de los parametros pasados 
    ''' </summary>
    ''' <param name="aobjValor">Objeto string a validar</param>
    ''' <param name="abytLongitudMin">Mínima longitud permitida del objeto a validar.</param>
    ''' <param name="abytLongitudMax">Máxima longitud permitida del objeto a validar.</param>
    ''' <param name="ablnRequerido">Indica si el objeto a validar puede, o no, ser Nothing o Empty.</param>
    ''' <returns>Valor buleano que indica si el objeto es valido o no.</returns>
    ''' <remarks></remarks>
    Friend Shared Function FblnEsValidoStringNumerico(ByRef aobjValor As Object, abytLongitudMin As Byte,
                abytLongitudMax As Byte, ablnRequerido As Boolean) As Boolean
        Dim lblnEsValido As Boolean = True
        If Not IsNothing(aobjValor) Then
            If Not String.IsNullOrEmpty(aobjValor) Then
                If IsNumeric(aobjValor) AndAlso Not aobjValor.GetType.Name = "Boolean" Then
                    If Not aobjValor.GetType.Name = "String" Then
                        aobjValor = CType(aobjValor, String)
                    End If
                Else
                    lblnEsValido = False
                End If
                If lblnEsValido Then
                    If aobjValor.Length > abytLongitudMax OrElse aobjValor.Length < abytLongitudMin Then
                        If aobjValor = "0" Then
                            lblnEsValido = Not ablnRequerido
                        Else
                            lblnEsValido = False
                        End If
                    End If
                Else
                    lblnEsValido = False
                End If
            Else
                lblnEsValido = Not ablnRequerido
            End If
        Else
            aobjValor = "0"
            lblnEsValido = Not ablnRequerido
        End If
        Return lblnEsValido
    End Function
    ''' <summary>
    ''' Valida si un objeto contiene un valor de tipo numérico y si su valor esta entre los rangos de
    ''' valor pasados en los argumentos, teniendo en cuenta si el valor es requerido.
    ''' </summary>
    ''' <param name="aobjValor">Objeto a ser validado</param>
    ''' <param name="adblValorMin">Mínimo valor permitido</param>
    ''' <param name="adblValorMax">Máximo valor permitido</param>
    ''' <param name="ablnRequerido">Indica si el objeto a validar puede, o no, ser Nothing o Empty.</param>
    ''' <param name="aenuTipoValor">Es el tipo del dato que contiene "aobjValor"</param>
    ''' <returns>Valor buleano que indica si el objeto es valido o no.</returns>
    ''' <remarks>El objeto a validar es pasado por referencia por lo que su contenido puede ser modificado</remarks>
    Friend Shared Function FblnEsValidoNumero(ByRef aobjValor As Object, adblValorMin As Double,
            adblValorMax As Double, ablnRequerido As Boolean,
            Optional aenuTipoValor As EnuTipoValor = EnuTipoValor.enuDouble) As Boolean
        Dim lblnEsValido As Boolean = True
        If Not (aobjValor Is Nothing OrElse String.IsNullOrEmpty(aobjValor)) Then
            If IsNumeric(aobjValor) AndAlso Not aobjValor.GetType.Name = "Boolean" Then
                If aobjValor.GetType.Name = "String" Then
                    If String.IsNullOrEmpty(aobjValor) Then aobjValor = 0.0
                    aobjValor = CType(aobjValor, Double)
                End If
                If aobjValor < adblValorMin OrElse aobjValor > adblValorMax Then
                    lblnEsValido = False
                    If aobjValor < adblValorMin AndAlso aobjValor = 0 Then
                        lblnEsValido = Not ablnRequerido
                    End If
                End If
            Else
                lblnEsValido = False
                aobjValor = 0.0
            End If
        Else
            aobjValor = 0.0
            If ablnRequerido Then
                lblnEsValido = aobjValor >= adblValorMin
            Else
                lblnEsValido = True
            End If
        End If
        If lblnEsValido AndAlso aenuTipoValor <> EnuTipoValor.enuDouble AndAlso Not IsNothing(aobjValor) Then
            aobjValor = FobjValor(aobjValor, aenuTipoValor)
        End If
        Return lblnEsValido
    End Function
    Private Shared Function FobjValor(aobjValor As Object, aenuTipoValor As EnuTipoValor) As Object
        Dim lobjValor As Object
        Select Case aenuTipoValor
            Case EnuTipoValor.enuByte
                lobjValor = CType(aobjValor, Byte)
            Case EnuTipoValor.enuSByte
                lobjValor = CType(aobjValor, SByte)
            Case EnuTipoValor.enuDecimal
                lobjValor = CType(aobjValor, Decimal)
            Case EnuTipoValor.enuUInteger
                lobjValor = CType(aobjValor, UInteger)
            Case EnuTipoValor.enuInteger
                lobjValor = CType(aobjValor, Integer)
            Case EnuTipoValor.enuULong
                lobjValor = CType(aobjValor, ULong)
            Case EnuTipoValor.enuLong
                lobjValor = CType(aobjValor, Long)
            Case EnuTipoValor.enuUShort
                lobjValor = CType(aobjValor, UShort)
            Case EnuTipoValor.enuShort
                lobjValor = CType(aobjValor, Short)
            Case EnuTipoValor.enuSingle
                lobjValor = CType(aobjValor, Single)
            Case Else
                Throw New ErrorInesperadoPanLException("Tipo de valor no esperado!")
        End Select
        Return lobjValor
    End Function
    ''' <summary>
    ''' Valida si un objeto contiene un valor de tipo byte y si su valor esta entre los rangos de
    ''' valor pasados en los argumentos.
    ''' </summary>
    ''' <param name="aobjValor">Objeto a ser validado</param>
    ''' <param name="abytValorMin">Mínimo valor permitido</param>
    ''' <param name="abytValorMax">Máximo valor permitido</param>
    ''' <returns>Valor buleano que indica si el objeto es valido o no.</returns>
    ''' <remarks>El objeto a validar es pasado por referencia por lo que su contenido puede ser modificado</remarks>
    Friend Shared Function FblnEsValidoEnumByte(ByRef aobjValor As Object, abytValorMin As Byte,
            abytValorMax As Byte, ablnRequerido As Boolean) As Boolean
        Dim lblnEsValido As Boolean
        Dim lbytValor As Byte
        If Not IsNothing(aobjValor) Then
            If IsNumeric(aobjValor) Then
                If Not aobjValor.GetType.Name = "Boolean" Then
                    If aobjValor.GetType.Name = "String" Then
                        aobjValor = CType(aobjValor, Byte)
                    End If
                    lblnEsValido = (aobjValor >= Byte.MinValue AndAlso aobjValor <= Byte.MaxValue)
                    If lblnEsValido Then
                        lblnEsValido = (aobjValor >= abytValorMin AndAlso aobjValor <= abytValorMax)
                        If Not lblnEsValido Then
                            lbytValor = aobjValor
                            If lbytValor = 0 Then
                                lblnEsValido = Not ablnRequerido
                            End If
                        End If
                    End If
                Else
                    lblnEsValido = False
                End If
            Else
                lblnEsValido = False
            End If
        Else
            lblnEsValido = Not ablnRequerido
            aobjValor = 0
        End If
        Return lblnEsValido
    End Function
    Friend Shared Function FblnEsValidoEMail(astrEmail As String) As Boolean
        Dim lstrExpValidaMail As String = "^([\w-\.]+)@((\[[0-9]{1,3}\." &
                    "[0-9]{1,3}\.)|(([\w-]+\.)+))" &
                    "([a-zA-z]{2,4}|[0-9]{1,3})(\]?)$"
        Return RegularExpressions.Regex.IsMatch(astrEmail, lstrExpValidaMail)
    End Function
    ''' <summary>
    ''' Valida si un objeto es tipo string y si su longitud esta esta entre los limites pasados en los argumentos
    ''' </summary>
    ''' <param name="aobjValor">Objeto a ser validado</param>
    ''' <param name="ashrLongitudMin">Longitud mínima permitida para la cadena.</param>
    ''' <param name="ashrLongitudMax">Longitud máxima permitida para la cadena.</param>
    ''' <param name="ablnRequerido">Indica si el objeto a validar puede, o no, ser Nothing o Empty.</param>
    ''' <returns>Valor buleano que indica si el objeto es valido o no.</returns>
    ''' <remarks>Si el valor no es requerido puede ser Nothing o una cadena de tipo Empty</remarks>
    Friend Shared Function FblnEsValidoString(ByRef aobjValor As Object,
            ashrLongitudMin As Short, ashrLongitudMax As Short, ablnRequerido As Boolean)
        Dim lblnEsValido As Boolean = True
        If Not IsNothing(aobjValor) Then
            If aobjValor.GetType.Name = "String" Then
                aobjValor = aobjValor.ToString().Trim
                If aobjValor.Length < ashrLongitudMin OrElse aobjValor.Length > ashrLongitudMax Then
                    If Not ablnRequerido Then
                        lblnEsValido = String.IsNullOrEmpty(aobjValor)
                    Else
                        lblnEsValido = False
                    End If
                End If
            Else
                aobjValor = String.Empty
                lblnEsValido = False
            End If
        Else
            aobjValor = String.Empty
            lblnEsValido = Not ablnRequerido
        End If
        Return lblnEsValido
    End Function
    ''' <summary>
    ''' Valida si el objeto "aobjValor" es un buleano."
    ''' </summary>
    ''' <param name="aobjValor">Objeto a validar</param>
    ''' <returns>Valor buleano que indica si el objeto es valido o no.</returns>
    ''' <remarks></remarks>
    Friend Shared Function FblnEsValidoBuleano(aobjValor As Object) As Boolean
        Dim lblnEsValido As Boolean = Not IsNothing(aobjValor)
        If lblnEsValido Then
            If IsNumeric(aobjValor) Then
                lblnEsValido = (aobjValor.GetType.Name = "Boolean")
            Else
                lblnEsValido = False
            End If
        End If
        Return lblnEsValido
    End Function
    ''' <summary>
    ''' Valida si el objeto "aobjValor" es una fecha valida y esta comprendida entre "adtmFechaMinima" y
    ''' "adtmFechaMaxima" teniendo en cuenta si el valor es requerido o no.
    ''' </summary>
    ''' <param name="aobjValor">Objeto de dipo "Date" que contiene la fecha a ser validada</param>
    ''' <param name="adtmFechaMinima">Fecha mínima permitida.</param>
    ''' <param name="adtmFechaMaxima">Fecha máxima permitida.</param>
    ''' <param name="ablnEsRequerido">Indica si el objeto a validar puede, o no, ser Nothing o "01/01/1900".</param>
    ''' <returns>Valor buleano que indica si el objeto es valido o no.</returns>
    ''' <remarks></remarks>
    Friend Shared Function FblnEsValidoFecha(aobjValor As Object, adtmFechaMinima As Date,
            adtmFechaMaxima As Date, ablnEsRequerido As Boolean) As Boolean
        Dim lblnEsValido As Boolean = True
        If Not IsNothing(aobjValor) Then
            If IsDate(aobjValor) Then
                If ablnEsRequerido Then
                    If aobjValor < adtmFechaMinima OrElse aobjValor > adtmFechaMaxima Then
                        lblnEsValido = False
                    End If
                Else
                    If aobjValor <> GCDTMFECHANULA Then
                        If aobjValor < adtmFechaMinima OrElse aobjValor > adtmFechaMaxima Then
                            lblnEsValido = False
                        End If
                    End If
                End If
            Else
                lblnEsValido = False
            End If
        Else
            If ablnEsRequerido Then
                lblnEsValido = False
            End If
        End If
        Return lblnEsValido
    End Function
#End Region

#Region "Funciones y procedimientos de fechas"
    Friend Shared Function FblnEsFechaNula(aobjObjeto As Object) As Object
        Dim lblnEsDate = IsDate(aobjObjeto)
        Dim lblnEsNula = False
        If lblnEsDate Then
            lblnEsNula = (aobjObjeto = GCDTMFECHANULA)
        End If
        Return lblnEsNula
    End Function
    ''' <summary>
    ''' Devuelve un string compuesto por los cuatro digitos del año de la fecha pasada en el argumento, seguido de 
    ''' los dos digitos del mes de la misma fecha
    ''' </summary>
    ''' <param name="adtmFecha">Fecha de la cual se devolvera el Periodo</param>
    ''' <returns>String: "AAAAMM" donde "AAAA" es el año de ñla fecha y "MM" es el mes de la fecha</returns>
    ''' <remarks></remarks>
    Friend Shared Function FstrPeriodo(adtmFecha As Date) As String
        Dim lstrPeriodo As String = String.Empty
        If IsDate(adtmFecha) Then
            lstrPeriodo = adtmFecha.Year.ToString & Format(adtmFecha.Month, "0#")
        End If
        Return lstrPeriodo
    End Function
    Friend Shared Function FstrFechayyyymmdd(adtmFecha As Date) As String
        Dim lstrFecha = adtmFecha.Year.ToString & Format(adtmFecha.Month, "0#") & Format(adtmFecha.Day, "0#")
        Return lstrFecha
    End Function
    Friend Shared Function FstrFechaddmmaaaaSepSlash(adtmFecha As Date) As String
        Dim lstrFecha = Format(adtmFecha.Day, "0#") & "/" & Format(adtmFecha.Month, "0#") & "/" &
                adtmFecha.Year.ToString
        Return lstrFecha
    End Function
    ''' <summary>
    ''' Devuelve la cantidad de dias que hay entre la fecha "adtmFecha1" y la fecha "adtmFecha2"
    ''' </summary>
    ''' <param name="adtmFecha1">La primera fecha (anterior a la fecha adtmFecha2)</param>
    ''' <param name="adtmFecha2">La segunda fecha (posterior a la fecha adtmFehca1)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Shared Function FentDiasEntreFechas(adtmFecha1 As Date, adtmFecha2 As Date) As Integer
        Dim lentCanDias As Integer
        lentCanDias = DateDiff(DateInterval.Day, adtmFecha1, adtmFecha2)
        Return lentCanDias
    End Function
    ''' <summary>
    ''' Devuelve la fecha del último día del mes correspondiente a la fecha pasada en el argumento
    ''' </summary>
    Friend Shared Function FdtmFecUltimoDiaMes(adtmfecha As Date) As Date
        Dim lstrPer As String = ClsPanorama.FstrPeriodo(adtmfecha)
        Dim lentAno As Integer = CInt(lstrPer.Substring(0, 4))
        Dim lentMes As Integer = CInt(lstrPer.Substring(4))
        Dim lentDia As Integer
        If lentMes = 2 Then
            If lentAno Mod 4 = 0 Then
                lentDia = 29
            Else
                lentDia = 28
            End If
        ElseIf lentMes = 1 OrElse lentMes = 3 OrElse lentMes = 5 OrElse lentMes = 7 OrElse
                    lentMes = 8 OrElse lentMes = 10 Or lentMes = 12 Then
            lentDia = 31
        Else
            lentDia = 30
        End If
        Dim ldtmFechaFinMes = DateSerial(lentAno, lentMes, lentDia)
        Return ldtmFechaFinMes
    End Function
#End Region

#Region "Funciones" 'Ok
    ''' <summary>
    ''' Devuelve el valor que contiene el campo correspondiente a la propiedad "aobjPropiedad" convertido al
    ''' tipo de dato respectivo
    ''' </summary>
    ''' <param name="adrwRegistro">DataRow que contiene el registro correspondiente a la propiedad "aobjPropiedad"</param>
    ''' <param name="aobjPropiedad">Objeto que define la propiedad de la cual se esta obteniendo el valor</param>
    ''' <returns>Objeto que tiene el valor de la propiedad</returns>
    ''' <remarks></remarks>
    Friend Shared Function FobjValorCampo(adrwRegistro As DataRow, aobjPropiedad As ClsCBPropiedad) As Object
        If adrwRegistro Is Nothing OrElse aobjPropiedad Is Nothing Then
            Dim lstrNombreArgNulo As String
            If adrwRegistro Is Nothing Then
                lstrNombreArgNulo = "adrwRegistro"
            Else
                lstrNombreArgNulo = "aobjPropiedad"
            End If
            Throw New ArgumentNullException(lstrNombreArgNulo)
        End If
        Dim lobjValor As Object
        Dim lstrNombreCampo = aobjPropiedad.StrNombreCampoBD
        If String.IsNullOrEmpty(lstrNombreCampo) OrElse adrwRegistro.IsNull(lstrNombreCampo) Then
            Return Nothing
        End If
        lobjValor = adrwRegistro.Item(lstrNombreCampo)
        Try
            Select Case aobjPropiedad.EnuTipoValor
                Case EnuTipoValor.enuBoolean
                    lobjValor = CType(lobjValor, Boolean)
                Case EnuTipoValor.enuByte
                    lobjValor = CType(lobjValor, Byte)
                Case EnuTipoValor.enuDate
                    lobjValor = CType(lobjValor, Date)
                Case EnuTipoValor.enuDecimal
                    lobjValor = CType(lobjValor, Decimal)
                Case EnuTipoValor.enuDouble
                    lobjValor = CType(lobjValor, Double)
                Case EnuTipoValor.enuInteger
                    lobjValor = CType(lobjValor, Integer)
                Case EnuTipoValor.enuLong
                    lobjValor = CType(lobjValor, Long)
                Case EnuTipoValor.enuShort
                    lobjValor = CType(lobjValor, Short)
                Case EnuTipoValor.enuSingle
                    lobjValor = CType(lobjValor, Single)
                Case EnuTipoValor.enuString
                    lobjValor = CType(lobjValor, String)
                Case EnuTipoValor.enuImagen
                    lobjValor = CType(lobjValor, Byte())
                Case EnuTipoValor.enuUInteger
                    lobjValor = CType(lobjValor, UInteger)
                Case EnuTipoValor.enuULong
                    lobjValor = CType(lobjValor, ULong)
                Case EnuTipoValor.enuUShort
                    lobjValor = CType(lobjValor, UShort)
            End Select
        Catch ex As OverflowException
            Throw New ValorPropiedadInvalidoException()
        Catch ex As Exception
            Throw New ValorPropiedadInvalidoException()
        End Try
        Return lobjValor
    End Function
    ''' <summary>
    ''' Devuelve el valor contenido en un campo de un DataRow convertido al tipo correspondiente  
    ''' </summary>
    ''' <param name="aobjValorCampo">Valor del campo a convertir</param>
    ''' <param name="aenuTipoValor">Tipo de valor que contiene el campo</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Shared Function FobjValorCampo(aobjValorCampo As Object, aenuTipoValor As EnuTipoValor) _
                As Object
        Dim lobjValor As Object = Nothing
        If Not IsDBNull(aobjValorCampo) Then
            lobjValor = aobjValorCampo
            Select Case aenuTipoValor
                Case EnuTipoValor.enuBoolean
                    lobjValor = CType(lobjValor, Boolean)
                Case EnuTipoValor.enuByte
                    lobjValor = CType(lobjValor, Byte)
                Case EnuTipoValor.enuDate
                    If TypeOf lobjValor Is Date Then
                        lobjValor = CType(lobjValor, Date)
                    Else
                        lobjValor = GCDTMFECHANULA
                    End If
                Case EnuTipoValor.enuDecimal
                    lobjValor = CType(lobjValor, Decimal)
                Case EnuTipoValor.enuDouble
                    lobjValor = CType(lobjValor, Double)
                Case EnuTipoValor.enuInteger
                    lobjValor = CType(lobjValor, Integer)
                Case EnuTipoValor.enuUInteger
                    lobjValor = CType(lobjValor, UInteger)
                Case EnuTipoValor.enuLong
                    lobjValor = CType(lobjValor, Long)
                Case EnuTipoValor.enuULong
                    lobjValor = CType(lobjValor, ULong)
                Case EnuTipoValor.enuShort
                    lobjValor = CType(lobjValor, Short)
                Case EnuTipoValor.enuUShort
                    lobjValor = CType(lobjValor, UShort)
                Case EnuTipoValor.enuSingle
                    lobjValor = CType(lobjValor, Single)
                Case EnuTipoValor.enuString
                    lobjValor = CType(lobjValor, String).Trim
                Case EnuTipoValor.enuImagen
                    lobjValor = CType(lobjValor, Byte())
                Case Else
                    lobjValor = lobjValor
            End Select
        End If
        Return lobjValor
    End Function
    Friend Shared Function FobjValorNuloPropiedad(aobjProp As ClsCBPropiedad)
        Dim lobjValor As Object
        Dim lenuTipoValor As EnuTipoValor = aobjProp.EnuTipoValor
        Select Case lenuTipoValor
            Case EnuTipoValor.enuByte, EnuTipoValor.enuDecimal, EnuTipoValor.enuDouble,
                    EnuTipoValor.enuInteger, EnuTipoValor.enuLong, EnuTipoValor.enuShort,
                    EnuTipoValor.enuSingle, EnuTipoValor.enuUInteger, EnuTipoValor.enuULong,
                    EnuTipoValor.enuUShort
                lobjValor = Nothing
            Case EnuTipoValor.enuBoolean
                lobjValor = False
            Case EnuTipoValor.enuDate
                lobjValor = GCDTMFECHANULA
            Case EnuTipoValor.EnuString
                lobjValor = String.Empty
            Case EnuTipoValor.EnuDateTime
                lobjValor = GCDTMTIMENULA
            Case Else
                lobjValor = Nothing
        End Select
        Return lobjValor
    End Function
    ''' <summary>
    ''' Devuelve la identificación numérica del último registro de una tabla. 
    ''' </summary>
    ''' <param name="astrNombreTabla">Nombre de la tabla</param>
    ''' <param name="astrNombreColumna">Nombre de la columna donde se buscara el valor más alto.</param>
    ''' <param name="aenuTipoValor">Tipo de valor que debe devolver esta función</param>
    ''' <param name="astrFiltro">Expresión que selecciona los registros a ser tenidos en cuenta. Para no filtrar
    ''' se debe enviar una cadena vacia:("")</param>
    ''' <returns>Un objeto con el valor requerido</returns>
    ''' <remarks></remarks>
    Shared Function FobjUltimaIdNumericaObjeto(astrNombreTabla As String, astrNombreColumna As String,
            aenuTipoValor As EnuTipoValor, astrFiltro As String) As Object
        Dim lobjUltimoValor As Object
        lobjUltimoValor = GobjPanDat.FobjMaxValorCampo(astrNombreTabla,
                astrNombreColumna, astrFiltro)
        If Not IsNothing(lobjUltimoValor) Then
            If IsDBNull(lobjUltimoValor) Then
                lobjUltimoValor = 0
            Else
                lobjUltimoValor = FobjValorCampo(lobjUltimoValor, aenuTipoValor)
            End If
        End If
        Return lobjUltimoValor
    End Function
    ''' <summary>
    ''' Examina todas las tablas de la base de datos buscando las que contengan la columna pasada en el argumento
    ''' "astrNombreColumna"; luego examina todas las tablas encontradas y determina si alguno de los valores 
    ''' contenidos en la columna cumple la condición pasada en el argumento "astrCondicion" caso en el cual
    ''' devuelve falso, de lo contrario devuelve verdadero.
    ''' </summary>
    ''' <param name="astrNombreTablasExcluir"> Array que contiene los Nombres de las Tablas que contienen una
    ''' referencia al registro que se va a eliminar y aun asi no impiden suprimir el registro.</param>
    ''' <param name="astrNombreColumna">Nombre de la columna que sera buscado en todas las tablas 
    ''' de la base de datos</param>
    ''' <param name="astrCondicion">La condicion que determina el valor devuelto por la función.Si la condición
    ''' se cumple al menos una vez la función devuelve "False"</param>
    ''' <param name="ablnComo">Indica si el nombre de la columna debe coincidir exactamente con el
    ''' string pasado en el parametro "astrNombreColumna" o lo contiene</param>
    ''' <param name="ablnPan">Indica si se buscará en las tablas de Panorama o en las tablas de la Aplicación.</param>
    ''' <returns>Un booleano que indica si el registro puede ser eliminado o no</returns>
    ''' <remarks></remarks>
    Shared Function FblnEsEliminableReg(astrNombreTablasExcluir As String(), astrNombreColumna As String,
            astrCondicion As String, ablnComo As Boolean, ablnPan As Boolean) As Boolean
        Dim lblnEsEliminable As Boolean = True
        Dim lstrCondicion As String
        Dim lshrCantReg As Short
        Dim lblnExcluir As Boolean
        Dim lstrNombreTablas As String()
        lstrNombreTablas = GobjPanDat.FstrNombreTablasContienenColumna(astrNombreColumna, ablnComo, ablnPan)
        If Not IsNothing(lstrNombreTablas) AndAlso lstrNombreTablas.Count > 0 Then
            For Each lstrElemento As String In lstrNombreTablas
                Dim lstrNomTabla As String = lstrElemento.Split(",")(0)
                Dim lstrNomColumna As String = lstrElemento.Split(",")(1)
                lblnExcluir = False
                For j As Byte = 0 To astrNombreTablasExcluir.GetUpperBound(0)
                    If astrNombreTablasExcluir(j).ToUpper = lstrNomTabla.ToUpper Then
                        lblnExcluir = True
                        Exit For
                    End If
                Next
                If Not lblnExcluir Then
                    lstrCondicion = lstrNomColumna & astrCondicion
                    lshrCantReg = GobjPanDat.FshrCantidadRegistros(lstrNomTabla,
                                lstrNomColumna, lstrCondicion)
                    lblnEsEliminable = (lshrCantReg = 0)
                End If
                If Not lblnEsEliminable Then
                    Exit For
                End If
            Next
        End If
        Return lblnEsEliminable
    End Function
    ''' <summary>
    ''' Examina todas las tablas de la base de datos buscando las que contengan la ultima columna del array
    ''' pasado en el argumento "astrNombreColumnas"; luego examina todas las tablas encontradas y 
    ''' determina si existe algun registro que cumpla la condición pasada en el argumento "astrCondiciones" 
    ''' caso en el cual devuelve falso, de lo contrario devuelve verdadero.
    ''' </summary>
    ''' <param name="astrNombreTablasExcluir"> Array que contiene los Nombres de las Tablas que contienen una
    ''' referencia al registro que se va a eliminar y aun asi no impiden suprimir el registro.</param>
    ''' <param name="astrNombreColumnas">Array que contiene el nombre de las columnas que identifican 
    ''' el registro a ser buscado en todas las tablas de la base de datos</param>
    ''' <param name="astrCondiciones">Array que contiene el valor de las columnas que identifican 
    ''' inequivocamente el registro a ser buscado. El valor debe estar antecedido por el operador de
    ''' comparación.</param>
    ''' <param name="ablnComo">Indica si el nombre de la columna debe coincidir exactamente con el
    ''' string pasado en el parametro "astrNombreColumna" o lo contiene</param>
    ''' <param name="ablnPan">Indica si se buscará en las tablas de Panorama o en las tablas de la Aplicación.</param>
    ''' <returns>Un booleano que indica si el registro puede ser eliminado o no</returns>
    ''' <remarks></remarks>
    Shared Function FblnEsEliminableReg(astrNombreTablasExcluir As String(),
            astrNombreColumnas As String(), astrCondiciones As String(),
            ablnComo As Boolean, ablnPan As Boolean) As Boolean
        Dim lblnEsEliminable As Boolean = True
        Dim lstrCondicion As String = String.Empty
        Dim lblnExcluir As Boolean
        Dim lshrCantReg As Short
        ' Indice de ubicacion de la columna que identifica el objeto
        Dim lbytIndiceColumna As Byte = astrNombreColumnas.GetUpperBound(0)
        Dim lstrNombreTablas As String()
        lstrNombreTablas = GobjPanDat.FstrNombreTablasContienenColumna(
                    astrNombreColumnas(lbytIndiceColumna), ablnComo, ablnPan)
        If Not IsNothing(lstrNombreTablas) AndAlso lstrNombreTablas.Count > 0 Then
            For Each lstrElemento As String In lstrNombreTablas
                Dim lstrNomTab As String = lstrElemento.Split(",")(0)
                Dim lstrNomCol As String = lstrElemento.Split(",")(1)
                lblnExcluir = False
                For j As Byte = 0 To astrNombreTablasExcluir.GetUpperBound(0)
                    If astrNombreTablasExcluir(j).ToUpper = lstrNomTab.ToUpper Then
                        lblnExcluir = True
                        Exit For
                    End If
                Next
                If Not lblnExcluir Then
                    If lbytIndiceColumna > 0 Then
                        lstrCondicion = String.Empty
                        For i As Byte = 0 To lbytIndiceColumna - 1
                            lstrCondicion &= astrNombreColumnas(i) & astrCondiciones(i) & " AND "
                        Next i
                        lstrCondicion &= lstrNomCol & astrCondiciones(lbytIndiceColumna)
                    Else
                        lstrCondicion &= lstrNomCol & astrCondiciones(0)
                    End If
                    lshrCantReg = GobjPanDat.FshrCantidadRegistros(lstrNomTab,
                                    lstrNomCol, lstrCondicion)
                    lblnEsEliminable = (lshrCantReg = 0)
                End If
                If Not lblnEsEliminable Then
                    Exit For
                End If
            Next
        End If
        Return lblnEsEliminable
    End Function
    Shared Function FblnEsEliminableReg(astrNombreTablasExcluir As String(),
            astrNombreColumnas As String(), acolValoresRef As Collection,
            ablnComo As Boolean, ablnPan As Boolean) As Boolean
        If astrNombreTablasExcluir Is Nothing OrElse astrNombreColumnas Is Nothing OrElse
                    acolValoresRef Is Nothing OrElse acolValoresRef.Count = 0 Then
            Dim lstrNombreArgNulo As String = String.Empty
            Select Case True
                Case astrNombreTablasExcluir Is Nothing
                    lstrNombreArgNulo = "astrNombreTablasExcluir"
                Case astrNombreColumnas Is Nothing
                    lstrNombreArgNulo = "astrNombreColumnas"
                Case acolValoresRef Is Nothing OrElse acolValoresRef.Count = 0
                    lstrNombreArgNulo = "astrCondiciones"
            End Select
            Throw New ArgumentNullException(lstrNombreArgNulo)
        End If
        Dim lblnEsEliminable As Boolean = True
        Dim lstrCondicion As String
        Dim lblnExcluir As Boolean
        Dim lshrCantReg As Short
        ' Indice de ubicacion de la columna que identifica el objeto
        Dim lstrNombreTablas As String()()
        lstrNombreTablas = GobjPanDat.FstrNombreTablasContienenColumnas(astrNombreColumnas, ablnComo, ablnPan)
        If Not IsNothing(lstrNombreTablas) AndAlso lstrNombreTablas.Count > 0 Then
            For Each lstrColsTabla As String() In lstrNombreTablas
                Dim lstrNomTab As String = lstrColsTabla(0).Split(",")(0)
                lblnExcluir = False
                For j As Byte = 0 To astrNombreTablasExcluir.GetUpperBound(0)
                    If astrNombreTablasExcluir(j).ToUpper = lstrNomTab.ToUpper Then
                        lblnExcluir = True
                        Exit For
                    End If
                Next
                If Not lblnExcluir Then
                    lstrCondicion = String.Empty
                    Dim lstrCol As String = String.Empty
                    Dim i = 0
                    For Each lstrElemento As String In lstrColsTabla
                        i += 1
                        lstrCol = lstrElemento.Split(",")(1)
                        lstrCondicion &= lstrCol & " = " & acolValoresRef(i).ToString & " AND "
                    Next
                    lstrCondicion = lstrCondicion.Substring(0, lstrCondicion.Length - 5)
                    lshrCantReg = GobjPanDat.FshrCantidadRegistros(lstrNomTab,
                                    lstrCol, lstrCondicion)
                    lblnEsEliminable = (lshrCantReg = 0)
                    If Not lblnEsEliminable Then Exit For
                End If
            Next
        End If
        Return lblnEsEliminable
    End Function
    ''' <summary>
    ''' Devuelve los permisos del cliente en la clase correpondiente al parametro aenuIdClase
    ''' </summary>
    Friend Function FenuTipoPermisos(aenuIdClase As EnuIdClasesPanDef) As EnuPermisosDef
        Dim lenuTipopermisoObj As EnuPermisosDef = EnuPermisosDef.None
        If Not IsNothing(ObjUsuarioActual) Then
            If ObjUsuarioActual.ObjIdUsuarioStr.ToString().ToUpper = GCSTRADMIN.ToUpper() OrElse
                    ObjUsuarioActual.ObjIdUsuarioStr.ObjValorPro = GCSTRUSUARIOU Then
                lenuTipopermisoObj = EnuPermisosDef.enuTodos
            Else
                lenuTipopermisoObj = ObjUsuarioActual.FenuTipoPermisosObj(aenuIdClase)
                If lenuTipopermisoObj = EnuPermisosDef.None Then
                    If aenuIdClase = EnuIdClasesPanDef.enuAno Then
                        lenuTipopermisoObj = EnuPermisosDef.enuConsultar
                    Else
                        lenuTipopermisoObj = EnuPermisosDef.enuTodos
                    End If
                End If
            End If
        End If
        Return lenuTipopermisoObj
    End Function
    Friend Shared Function FstrBuleanoToString(ablnBuleano As Boolean) As String
        Dim lstrValor = "No"
        If ablnBuleano Then
            lstrValor = "Si"
        End If
        Return lstrValor
    End Function
    Friend Shared Function FstrPrefijoDcto(astrNumeroDcto As String) As String
        If astrNumeroDcto Is Nothing Then
            Throw New ArgumentNullException(NameOf(astrNumeroDcto))
        End If
        If astrNumeroDcto.Contains("-") Then
            Return astrNumeroDcto.Split("-")(0)
        Else
            Return ""
        End If
    End Function
    Friend Shared Function FentIdDcto(astrNumeroDcto As String) As Integer
        If astrNumeroDcto Is Nothing Then
            Throw New ArgumentNullException(NameOf(astrNumeroDcto))
        End If
        Dim lentIdDocumento = 0, lstrMens = String.Empty
        If astrNumeroDcto.Contains("-") Then
            If Not IsNumeric(astrNumeroDcto.Split("-")(1)) OrElse (Val(astrNumeroDcto.Split("-")(1)) >
                    Integer.MaxValue OrElse Val(astrNumeroDcto.Split("-")(1)) <= 0) Then
                lstrMens = "El Número del Documento no es valido!"
            Else
                lentIdDocumento = CType(astrNumeroDcto.Split("-")(1), Integer)
            End If
        Else
            If Not IsNumeric(astrNumeroDcto) OrElse (Val(astrNumeroDcto) > Integer.MaxValue OrElse
                        Val(astrNumeroDcto) <= 0) Then
                lstrMens = "El Número del Documento '" & astrNumeroDcto & "' no es valido!"
            Else
                lentIdDocumento = CType(astrNumeroDcto, Integer)
            End If
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            Throw New ErrorInesperadoPanLException(lstrMens & " en FentIdDoc!")
        End If
        Return lentIdDocumento
    End Function
    Shared Function FstrNumeroDcto(astrPrefijoDcto As String, aentIdDcto As Integer) As String
        Dim lstrNroDcto = astrPrefijoDcto
        If Not String.IsNullOrEmpty(lstrNroDcto) Then
            lstrNroDcto &= "-"
        End If
        lstrNroDcto &= aentIdDcto.ToString
        Return lstrNroDcto
    End Function
    Shared Function FblnExisteTercero(adblIdTercero As Double) As Boolean
        Dim lstrTabla = ClsTercero.SstrNombreTabla
        Dim lstrCamposSelect = {ClsIdTerceroDbl.SstrNombreCampoBd}
        Dim lstrFiltro = ClsIdTerceroDbl.SstrNombreCampoBd & " = " & adblIdTercero.ToString
        Dim ldtbTercero = FdtbDataTable(lstrTabla, lstrCamposSelect, {{"", ""}}, lstrFiltro)
        Dim lblnExiste = (ldtbTercero.Rows.Count > 0)
        Return lblnExiste
    End Function
    ' Manejo Contraseñas correo y proveedor eFactura
    Friend Shared Function FstrContrasena(ablnIn As Boolean, astrIn As String) As String
        If Not IsNothing(astrIn) Then
            Dim lstrCar(astrIn.Length) As String, lstrOut = String.Empty, lstrIn = String.Empty
            Dim i = 0
            If ablnIn Then
                For Each lchrIn As Char In astrIn
                    lstrCar(i) = Chr((Asc(lchrIn) + (astrIn.Length * 3) + astrIn.Length - i))
                    i += 1
                Next
                For j = (astrIn.Length - 1) To 0 Step -1
                    lstrOut += lstrCar(j)
                Next j
                Return lstrOut
            Else
                For Each lchrIn As Char In astrIn
                    lstrCar(i) = lchrIn
                    i += 1
                Next
                For j = (astrIn.Length - 1) To 0 Step -1
                    lstrOut += lstrCar(j)
                Next j
                i = 0
                For Each lchrOut As Char In lstrOut
                    lstrCar(i) = Chr(Asc(lchrOut) - (astrIn.Length * 3) - astrIn.Length + i)
                    lstrIn += lstrCar(i)
                    i += 1
                Next
                Return lstrIn
            End If
        Else
            Return ""
        End If
    End Function
    Friend Shared Function FstrContrasena(aobjContrasena As ClsCBPropiedad)
        Dim lstrContr = String.Empty
        If Not IsNothing(aobjContrasena.ObjValorPro) AndAlso aobjContrasena.BlnEsValido AndAlso
                aobjContrasena.ObjValorPro.Length > 0 Then
            lstrContr = FstrContrasena(False, aobjContrasena.ObjValorPro)
        End If
        Return lstrContr
    End Function
#Region "Funciones Objetos BD" 'Ok
    ''' <summary>
    ''' Devuelve la DataTable definida en el parametro "astrSql" que debe ser una sentencia SELECT
    ''' </summary>
    ''' <param name="astrSql">Expresion SQL del tipo SELECT</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Shared Function FdtbDataTable(astrSql As String) As DataTable
        If String.IsNullOrEmpty(astrSql) Then
            Throw New ArgumentNullException(NameOf(astrSql))
        End If
        Dim ldtbDataTable As DataTable = Nothing
        Using ldstDataSet As New DataSet
            GobjPanDat.SdsDataSet(ldstDataSet, astrSql)
            If Not IsNothing(ldstDataSet) Then
                ldtbDataTable = ldstDataSet.Tables(0)
            End If
        End Using
        Return ldtbDataTable
    End Function
    ''' <summary>
    ''' Devuelve un DataTable conteniendo los datarows que corresponden a los parametros pasados
    ''' </summary>
    ''' <param name="astrNombreTabla">Nombre de la tabla a la cual pertenecen los DataRow</param>
    ''' <param name="astrCamposSelect">Array con los nombres de los campos que contenra cada DataRow</param>
    ''' <param name="astrIndice">Array de columnas que forman el indice. Cada elemento de este array es un array
    ''' que contiene dos elementos: nombre de la columna y Orden (Valores aceptados:"ASC" y "DESC") </param>
    ''' <param name="astrFiltro">Expresión de filtro para indicar los registros a ser tenidos en cuenta. 
    ''' Para no filtrar los registros este argumento debe ser una cadena vacia:("") </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Shared Function FdtbDataTable(astrNombreTabla As String, astrCamposSelect() As String,
            astrIndice(,) As String, astrFiltro As String) As DataTable
        Dim ldtbDataTable As DataTable = Nothing
        Using ldstDataSet As New DataSet
            GobjPanDat.SdsDataSet(ldstDataSet, astrNombreTabla, astrCamposSelect, astrIndice, astrFiltro,
                True, Array.Empty(Of String))
            If Not IsNothing(ldstDataSet) Then
                ldtbDataTable = ldstDataSet.Tables(0)
            End If
        End Using
        Return ldtbDataTable
    End Function
    ''' <summary>
    ''' Devuelve un DataTable con datos de dos tablas de acuerdo a a los argumentos pasados.
    ''' </summary>
    ''' <param name="astrNombreTablaPri">Nombre de la tabla principal que contiene los campos que se convinarán con los campos
    ''' de la tabla secundaria mediante INNER JOIN.</param>
    ''' <param name="astrCamposTablaPri">Array que contiene los nombres de los campos de la tabla primaria que se tendrán 
    ''' en cuenta.</param>
    ''' <param name="astrNombreTablaSec">Nombre de la tabla secundaria que contiene los campos que se convinarán con los campos
    ''' de la tabla primaria mediante INNER JOIN.</param>
    ''' <param name="astrCamposTablaSec">Array que contiene los nombres de los campos de la tabla secundaria que se tendrán 
    ''' en cuenta.</param>
    ''' <param name="astrCamposPriRel">Array que contiene los nombres de los campos de la tabla primaria que definen
    ''' la relación con la tabla secundaria.</param>
    ''' <param name="astrCamposSecRel">Array que contiene los nombres de los campos de la tabla secundaria que definen
    ''' la relación con la tabla primaria.</param>
    ''' <param name="astrIndice">Un array de dos dimensiones que definen el orden de los registros. La primer dimensión
    ''' contiene el nombre de las columnas del indice y la segunda dimensión define el orden; esta segunda dimensión
    '''  debe contener una de las siguientes expresiones: "ASC" o "DESC"</param>
    ''' <param name="astrFiltro">Un string que define los criterios de selección de los registros. Para no filtrar
    ''' los registros se pasa una cadena vacia:("")</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Shared Function FdtbDataTable(astrNombreTablaPri As String, astrCamposTablaPri() As String,
            astrNombreTablaSec As String, astrCamposTablaSec() As String,
            astrCamposPriRel() As String, astrCamposSecRel() As String, astrIndice(,) As String,
            astrFiltro As String, astrCamposGrupo() As String, ablnPkIndice As Boolean) As DataTable
        Dim ldtbDataTable As DataTable = Nothing
        Using ldstDataSet As New DataSet
            GobjPanDat.SdsDataSet(ldstDataSet, astrNombreTablaPri, astrCamposTablaPri,
                        astrNombreTablaSec, astrCamposTablaSec, astrCamposPriRel, astrCamposSecRel,
                        astrIndice, astrFiltro, ablnPkIndice, astrCamposGrupo)
            If Not IsNothing(ldstDataSet) Then
                ldtbDataTable = ldstDataSet.Tables(0)
            End If
        End Using
        Return ldtbDataTable
    End Function
    ''' <summary>
    ''' Devuelve un DataTable con datos de dos tablas de acuerdo a a los argumentos pasados.
    ''' </summary>
    ''' <param name="astrNombreTablaPri">Nombre de la tabla principal que contiene los campos que se convinarán con los campos
    ''' de la tabla secundaria mediante INNER JOIN.</param>
    ''' <param name="astrCamposTablaPri">Array que contiene los nombres de los campos de la tabla primaria que se tendrán 
    ''' en cuenta.</param>
    ''' <param name="astrNombreTablaSec">Nombre de la tabla secundaria que contiene los campos que se convinarán con los campos
    ''' de la tabla primaria mediante INNER JOIN.</param>
    ''' <param name="astrCamposTablaSec">Array que contiene los nombres de los campos de la tabla secundaria que se tendrán 
    ''' en cuenta.</param>
    ''' <param name="astrCamposPriRel">Array que contiene los nombres de los campos de la tabla primaria que definen
    ''' la relación con la tabla secundaria.</param>
    ''' <param name="astrCamposSecRel">Array que contiene los nombres de los campos de la tabla secundaria que definen
    ''' la relación con la tabla primaria.</param>
    ''' <param name="astrIndice">Un array de dos dimensiones que definen el orden de los registros. La primer dimensión
    ''' contiene el nombre de las columnas del indice y la segunda dimensión define el orden; esta segunda dimensión
    '''  debe contener una de las siguientes expresiones: "ASC" o "DESC"</param>
    ''' <param name="ablnIndiceUnico">Boolean que indica si el indice esperado es unico o no.</param>
    ''' <param name="astrFiltro">Un string que define los criterios de selección de los registros. Para no filtrar
    ''' los registros se pasa una cadena vacia:("")</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Shared Function FdtbDataTable(astrNombreTablaPri As String, astrCamposTablaPri As String(),
            astrNombreTablaSec As String, astrCamposTablaSec As String(),
            astrCamposPriRel As String(), astrCamposSecRel As String(), astrIndice As String(,),
            ablnIndiceUnico As Boolean, astrFiltro As String,
            astrCamposGrupo As String()) As DataTable
        Dim ldtbDataTable As DataTable = Nothing
        Using ldstDataSet As New DataSet
            GobjPanDat.SdsDataSet(ldstDataSet, astrNombreTablaPri, astrCamposTablaPri,
                        astrNombreTablaSec, astrCamposTablaSec, astrCamposPriRel, astrCamposSecRel,
                        astrIndice, astrFiltro, ablnIndiceUnico, astrCamposGrupo)
            If Not IsNothing(ldstDataSet) Then
                ldtbDataTable = ldstDataSet.Tables(0)
            End If
        End Using
        Return ldtbDataTable
    End Function
    Friend Shared Function FdtbDataTable(astrNombreTabla As String, astrCampos() As String,
            astrIndice(,) As String, astrFiltro As String, ablnPKIndice As Boolean,
            astrCamposGrupo() As String) As DataTable
        Dim ldtbDataTable As DataTable = Nothing
        Using ldstDataSet As New DataSet
            GobjPanDat.SdsDataSet(ldstDataSet, astrNombreTabla, astrCampos, astrIndice, astrFiltro,
                        ablnPKIndice, astrCamposGrupo)
            If Not IsNothing(ldstDataSet) Then
                ldtbDataTable = ldstDataSet.Tables(0)
            End If
        End Using
        Return ldtbDataTable
    End Function
    ''' <summary>
    ''' Devuelve el DataTable correspondiente a la tabla "astrNombreTabla" de la base de datos Access contenida
    ''' en el archivo "astrArchivo".
    ''' </summary>
    ''' <param name="astrArchivo">Trayectoria completa del archivo que contiene los datos que poblarán 
    ''' el DataTable.</param>
    ''' <param name="astrContrasena">Contraseña de la base de datos Access.</param>
    ''' <param name="astrNombreTabla">Nombre de la tabla que contiene los datos que se quieren importar.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Shared Function FdtbTablaAccess(astrArchivo As String, astrContrasena As String,
            astrNombreTabla As String) As DataTable
        Dim ldtbDataTable As DataTable = Nothing
        Using ldstDataSet As New DataSet
            GobjPanDat.SdstDataSetAccess(ldstDataSet, astrArchivo, astrContrasena, astrNombreTabla)
            If Not IsNothing(ldstDataSet) Then
                ldtbDataTable = ldstDataSet.Tables(0)
            End If
        End Using
        Return ldtbDataTable
    End Function
    ''' <summary>
    ''' Devuelve un Array de DatRows según los argumentos pasados 
    ''' </summary>
    ''' <param name="astrNombreTabla">Nombre de la tabla a la cual pertenecen los registros</param>
    ''' <param name="astrCamposSelect">Un array con los nombres de las columnas que incluira cada DataRow</param>
    ''' <param name="astrIndice">Un array de dos dimensiones que definen el orden de los registros. La primer dimensión
    ''' contiene el nombre de las columnas del indice y la segunda dimensión define el orden; esta segunda dimensión
    '''  debe contener una de las siguientes expresiones: "ASC" o "DESC"</param>
    ''' <param name="astrFiltro">Un string que define los criterios de selección de los registros. Si no se quiere
    ''' filtrar los registros se debe pasar una cadena vacia:("")</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Shared Function FdrwDataRow(astrNombreTabla As String, astrCamposSelect() As String,
        astrIndice(,) As String, astrFiltro As String) As DataRow()
        Dim ldtbDataTable As DataTable
        Dim ldrwDataRow As DataRow() = Nothing
        Using ldstDataSet As New DataSet
            GobjPanDat.SdsDataSet(ldstDataSet, astrNombreTabla, astrCamposSelect, astrIndice, astrFiltro,
                        True, Array.Empty(Of String))
            If Not IsNothing(ldstDataSet) Then
                ldtbDataTable = ldstDataSet.Tables(0)
                ldrwDataRow = ldtbDataTable.Select
            End If
        End Using
        Return ldrwDataRow
    End Function
    Friend Shared Function FdtbUltimosReg(astrNombreTabla As String, astrFiltro As String,
                                  astrIndice As String, aentCantRegistros As Integer) As DataTable
        Dim ldtbOrigen As DataTable
        Select Case GenuProveedorBD
            Case EnuProveedorBD.enuMySql
                Dim lstrSql = "SELECT * FROM " & astrNombreTabla
                If Not String.IsNullOrEmpty(astrFiltro) Then
                    lstrSql += " WHERE " & astrFiltro
                End If
                lstrSql += " ORDER BY " & astrIndice & " DESC LIMIT " & aentCantRegistros.ToString
                ldtbOrigen = FdtbDataTable(lstrSql)
            Case Else
                Throw New PanDatException("Función no disponible para este Proveedor de BD")
        End Select
        Return ldtbOrigen
    End Function
    ''' <summary>
    ''' Devuelve el valor contenido en el campo "astrNombreCampoBuscado" del DataRow que en el array de
    ''' DataRows "adrwDataRow", su campo "astrNombreCampoRef" coincide con el valor "aobjValorCampoRef". 
    ''' </summary>
    ''' <param name="adrwDataRow">Array de DataRows donde se hara la busqueda</param>
    ''' <param name="astrNombreCampoRef">Campo que contiene el valor de referencia</param>
    ''' <param name="astrNombreCampoBuscado">Campo que contiene el valor buscado</param>
    ''' <param name="aobjValorCampoRef">Valor que contiene el campo de referencia</param>
    ''' <returns>El objeto conteniendo el valor solicitado</returns>
    ''' <remarks></remarks>
    Friend Shared Function FobjValorCampoDataRow(adrwDataRow() As DataRow, astrNombreCampoRef As String,
            astrNombreCampoBuscado As String, aobjValorCampoRef As Object) As Object
        Dim lobjValor As Object = False
        If Not IsNothing(adrwDataRow) AndAlso adrwDataRow.Count > 0 Then
            For Each ldrwDataRow As DataRow In adrwDataRow
                If ldrwDataRow(astrNombreCampoRef) = aobjValorCampoRef Then
                    lobjValor = ldrwDataRow(astrNombreCampoBuscado)
                End If
            Next
        End If
        Return lobjValor
    End Function
    ''' <summary>
    ''' Devuelve el indice (base cero) del elemento que contiene el valor pasado en el argumento "aobjValorRef" en
    ''' el campo pasado en el argumento "astrNombreCampoRef"
    ''' </summary>
    ''' <param name="adrwDataRow">El DataRow en el cual se buscará el indice.</param>
    ''' <param name="astrNombreCampoRef">Nombre del campo que contiene el valor pasado en el argumento "aobjValorRef"</param>
    ''' <param name="aobjValorRef">Valor a buscar en el campo pasado en el argumento "aobjValorRef"</param>
    ''' <returns>Un valor Short que indica el indice donde esta ubicado el elemento en el array</returns>
    ''' <remarks></remarks>
    Friend Shared Function FstrNombreTercero(adblIdTercero As Double) As String
        Dim lobjTercero As New ClsTercero(EnuModoInstanciaObjDef.enuUnico)
        Dim lobjValorLlave() As Object = {adblIdTercero}, lstrNombreTer = String.Empty
        lobjTercero.SAbra(lobjValorLlave)
        If lobjTercero.BlnExiste Then
            lstrNombreTer = lobjTercero.FstrNombreCompleto()
        End If
        Return lstrNombreTer
    End Function
#End Region
#End Region

#Region "Auriga"
    ''' <summary>
    ''' Devuelve un datatable con la información del contrato actual 
    ''' </summary>
    ''' <returns></returns>
    Friend Shared Function FdtbAuriga(astrIdCopropiedad As String) As DataTable
        Dim lstrTabla = "Contratos"
        Dim lstrCampSel As String() = {"*"}
        Dim lstrFiltro As String = "IdTerceroCopropiedad = '" & astrIdCopropiedad & "'"
        Dim lstrOrden = {{"", ""}}
        Dim lstrExpSql As String = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla,
                lstrCampSel, lstrOrden, lstrFiltro, {})
        Dim ldtbAuriga As DataTable = Nothing
        Using ldstDataSet As New DataSet
            GobjPanDat.SdstAuriga(ldstDataSet, lstrExpSql)
            If Not IsNothing(ldstDataSet) Then
                ldtbAuriga = ldstDataSet.Tables(0)
            End If
        End Using
        Return ldtbAuriga
    End Function
    Friend Shared Sub SRegistreNit(alngNit As Long)
        Dim lstrTabla = "Contratos"
        Dim lcolColumnas As New Collection
        Dim lcolDatos As New Collection
        lcolColumnas.Add("IdEstado")
        lcolColumnas.Add("FechaVence")
        lcolColumnas.Add("IdContrato")
        lcolColumnas.Add("IdTerceroCopropiedad")
        lcolColumnas.Add("EsLicencia")
        lcolDatos.Add(EnuEstadoContrato.EnuRegistradoNit, "IdEstado")
        lcolDatos.Add("1900-01-01", "FechaVence")
        lcolDatos.Add(10000, "IdContrato")
        lcolDatos.Add(alngNit, "IdTerceroCopropiedad")
        lcolDatos.Add(False, "EsLicencia")
        Dim lentRegInsertados = GobjPanDat.SInserteRegistro_Au(lstrTabla, lcolColumnas, lcolDatos)
    End Sub
    Friend Shared Sub SRegistreIngreso(astrNit As String)
        Dim lstrNombreTabla = "Contratos"
        Dim lcolCamposCambio As New Collection From {
            "UltimoIngreso"
        }
        Dim lcolDatosNuevos As New Collection From {
            {Today, "UltimoIngreso"}
        }
        Dim lcolCampRef As New Collection From {
            "IdTerceroCopropiedad"
        }
        Dim lcolDatosRef As New Collection From {
            {astrNit, "IdTerceroCopropiedad"}
        }
        Dim lentRegActu = GobjPanDat.SActualiceReg_Au(lstrNombreTabla, lcolCamposCambio,
                lcolDatosNuevos, lcolCampRef, lcolDatosRef)
    End Sub
#End Region

#Region "Envio de Emails"
    Friend Shared Function FblnEmailsHabilitado() As Boolean
        Dim lstrTra = GstrTraySegApp & "plamarten.dll"
        Return Computer.FileSystem.FileExists(lstrTra)
    End Function
    Friend Shared Function FblnEnvioMensaje(astrEmailDestinatario As ArrayList,
            astrAsunto As String, astrMensaje As String, astrArchivoAdjunto As String,
            astrServidor As String, ablnHabilitarSSL As Boolean, aentPuerto As Integer,
            astrMailOrigen As String, ablnRequiereAutenticacion As Boolean,
            astrContraseña As String, ByRef ablnTimeOut As Boolean) As Boolean
        Dim lblnNoHayError = False
        GobjPanDat.SControleProcesoObj(True)
        Dim lmlmMensaje As New MailMessage
        Dim smcCliente As New SmtpClient
        Dim lattAdjunto As Attachment = Nothing
        Dim lstrUsuario = String.Empty, lstrCon = String.Empty
        Try
            If astrArchivoAdjunto.Length > 0 Then
                If Computer.FileSystem.FileExists(astrArchivoAdjunto) Then
                    lattAdjunto = New Attachment(astrArchivoAdjunto)
                    lmlmMensaje.Attachments.Add(lattAdjunto)
                Else
                    Throw New ErrorInesperadoPanLException("El archivo para adjuntar no existe en el equipo!")
                End If
            End If
            lmlmMensaje.Body = astrMensaje
            lmlmMensaje.Subject = astrAsunto
            lmlmMensaje.IsBodyHtml = False
            For Each lstrDir As String In astrEmailDestinatario
                lmlmMensaje.To.Clear()
                lmlmMensaje.To.Add(lstrDir)
                lmlmMensaje.From = New MailAddress(astrMailOrigen)
                ' 
                If ablnRequiereAutenticacion Then
                    lstrUsuario = astrMailOrigen
                    lstrCon = astrContraseña
                End If
                smcCliente.Host = astrServidor
                smcCliente.Port = aentPuerto
                smcCliente.EnableSsl = ablnHabilitarSSL
                smcCliente.Timeout = 30000
                smcCliente.Credentials = New NetworkCredential(lstrUsuario, lstrCon)
                smcCliente.Send(lmlmMensaje)
            Next
            lblnNoHayError = True
        Catch ex As TimeoutException
            ablnTimeOut = True
        Catch ex As System.Net.Mail.SmtpException
            ablnTimeOut = True
        Catch ex As Exception
            lblnNoHayError = False
        Finally
            GobjPanDat.SControleProcesoObj(False)
            If Not IsNothing(lattAdjunto) Then
                lattAdjunto.Dispose()
            End If
            lmlmMensaje.Dispose()
        End Try
        Return lblnNoHayError
    End Function
    Friend Shared Function FblnEnvioMensaje(astrEmailDestinatario As ArrayList, astrAsunto As String,
            astrMensaje As String, astrArchivoAdjunto As String,
            aobjCentroUtil As ClsCentroUtilidad, astrCon As String,
            ByRef astrMens As String) As Boolean
        Dim lblnNoHayError = False, lblnTimeOut = False
        GobjPanDat.SControleProcesoObj(True)
        Dim lmlmMensaje As New MailMessage
        Dim smcCliente As New SmtpClient
        Dim lattAdjunto As Attachment = Nothing
        Dim lstrUsuario = String.Empty, lstrCon = String.Empty
        Dim lstrServidor As String = aobjCentroUtil.ObjServidorSmtpStr.ObjValorPro
        Dim lblnHabilitarSSL As Boolean = aobjCentroUtil.ObjHabilitarSslBln.ObjValorPro
        Dim lentPuerto As Integer = aobjCentroUtil.ObjPuertoHostShr.ObjValorPro
        Try
            If astrArchivoAdjunto.Length > 0 Then
                If Computer.FileSystem.FileExists(astrArchivoAdjunto) Then
                    lattAdjunto = New Attachment(astrArchivoAdjunto)
                    lmlmMensaje.Attachments.Add(lattAdjunto)
                Else
                    Throw New ErrorInesperadoPanLException("El archivo para adjuntar no existe en el equipo!")
                End If
            End If
            lmlmMensaje.Body = astrMensaje
            lmlmMensaje.Subject = astrAsunto
            lmlmMensaje.IsBodyHtml = False
            For Each lstrDir As String In astrEmailDestinatario
                lmlmMensaje.To.Clear()
                lmlmMensaje.To.Add(lstrDir)
                lmlmMensaje.From = New MailAddress(aobjCentroUtil.ObjEmailOrigenStr.ObjValorPro)
                ' 
                If aobjCentroUtil.ObjRequiereAutenticacionBln.ObjValorPro Then
                    lstrUsuario = aobjCentroUtil.ObjEmailOrigenStr.ObjValorPro
                    lstrCon = astrCon
                End If
                smcCliente.Host = lstrServidor
                smcCliente.Port = lentPuerto
                smcCliente.EnableSsl = lblnHabilitarSSL
                smcCliente.Timeout = 30000
                smcCliente.Credentials = New NetworkCredential(lstrUsuario, lstrCon)
                smcCliente.Send(lmlmMensaje)
            Next
            lblnNoHayError = True
        Catch ex As TimeoutException
            lblnNoHayError = True
            lblnTimeOut = True
        Catch ex As System.Net.Mail.SmtpException
            lblnNoHayError = True
            lblnTimeOut = True
        Catch ex As Exception
            lblnNoHayError = False
        Finally
            If lblnTimeOut Then
                astrMens = "No es posible enviar mensajes en este momento. " &
                    "Intentelo de nuevo más tarde!"
            ElseIf Not lblnNoHayError Then
                astrMens = "Se presento un error al enviar el mensaje!"
            End If
            GobjPanDat.SControleProcesoObj(False)
            If Not IsNothing(lattAdjunto) Then
                lattAdjunto.Dispose()
            End If
            lmlmMensaje.Dispose()
        End Try
        Return lblnNoHayError
    End Function
#End Region

#Region "Procedimientos Instancia actual" 'Ok
    ''' <summary>
    ''' Establece la propiedad "objCarpetaActual" que representa la Carpeta de Trabajo.
    ''' </summary>
    ''' <param name="ashrIdCarpeta">Id de la Carpeta que va a ser instanciada.</param>
    ''' <remarks></remarks>
    Friend Sub SEstablezcaCarpetaActual(ashrIdCarpeta As Short)
        MobjCarpetaActual = ClsAdministrador.FobjCarpeta(ashrIdCarpeta)
        If MobjCarpetaActual.BlnExiste Then
            GshrIdCarpeta = MobjCarpetaActual.ObjIdCarpetaShr.ObjValorPro
            ClsPanorama.SEstablezcaTrayaectoriaDatos(MobjCarpetaActual.ObjNombreStr.ObjValorPro)
        End If
    End Sub
    Private Sub SCargueDtbLogApp()
        Static lblnDtbLogAppCargada As Boolean
        If Not lblnDtbLogAppCargada Then
            Dim lstrSql As String = "SELECT * FROM PanLogApp LIMIT 1"
            MdtbLogApp = FdtbDataTable(lstrSql)
            lblnDtbLogAppCargada = True
        End If
    End Sub
    ''' <summary>
    ''' Crea un nuevo objeto LogApp de tipo "Cambio" y lo registra en la base de datos
    ''' </summary>
    ''' <param name="astrIdObjeto">Valor que identifica el objeto en el cual se produjo un cambio.</param>
    ''' <param name="astrNombreClase">Es el nombre de la clase a la cual pertenece el objeto que generó el Log.</param>
    ''' <param name="astrNombrePropiedad">Nombre de la propiedad que tuvo el cambio que se está registrando.</param>
    ''' <param name="aobjvalorAnterior">Valor de la propiedad antes del cambio.</param>
    ''' <param name="aobjValorNuevo">Valor de la propiedad después del cambio.</param>
    ''' <remarks></remarks>
    Friend Sub SRegistreCambioLogApp(astrIdObjeto As String, astrNombreClase As String,
                astrNombrePropiedad As String, aobjValorAnterior As Object, aobjValorNuevo As Object)
        If aobjValorAnterior Is Nothing OrElse aobjValorNuevo Is Nothing Then
            Dim lstrNombreArgNulo As String = String.Empty
            If aobjValorAnterior Is Nothing Then
                lstrNombreArgNulo = "aobjValorAnterior"
            Else
                lstrNombreArgNulo = "aobjValorNuevo"
            End If
            Throw New ArgumentNullException(lstrNombreArgNulo)
        End If
        SCargueDtbLogApp()
        Dim ldrwLogApp As DataRow = MdtbLogApp.NewRow
        Dim lobjLogApp As New ClsLogApp(Me, ldrwLogApp)
        If Not CType(lobjLogApp.EnuPermisosObj And EnuPermisosDef.enuCrear, Boolean) Then
            lobjLogApp.EnuPermisosObj += EnuPermisosDef.enuCrear
        End If
        lobjLogApp.SCreeObj(Nothing)
        With lobjLogApp
            .ObjFechaCreacionDtm.ObjValorPro = Now
            .ObjIdAppLogShr.ObjValorPro = CType(GenuIdAplicacion, UShort)
            .ObjIdLogAppInt.ObjValorPro = 0
            If GenuIdAplicacion = EnuListaAplicaciones.EnuAdministrador Then
                .ObjIdCarpetaLogShr.ObjValorPro = 0
                .ObjIdCentroUtilLogShr.ObjValorPro = 0
            Else
                .ObjIdCarpetaLogShr.ObjValorPro = GshrIdCarpeta
                .ObjIdCentroUtilLogShr.ObjValorPro = GshrIdCentroUtil
            End If
            If Not IsNothing(astrIdObjeto) Then
                .ObjIdObjetoStr.ObjValorPro = astrIdObjeto
            Else
                .ObjIdObjetoStr.ObjValorPro = String.Empty
            End If
            If String.IsNullOrEmpty(GstrIdUsuario) Then
                .ObjIdUsuarioLogStr.ObjValorPro = "Auto"
            Else
                .ObjIdUsuarioLogStr.ObjValorPro = GstrIdUsuario
            End If
            .ObjNombreClaseStr.ObjValorPro = astrNombreClase
            .ObjNombreCampoLogStr.ObjValorPro = astrNombrePropiedad
            .ObjTipoLogByt.ObjValorPro = EnuTipoLogDef.enuCambio
            If Not IsNothing(aobjValorAnterior) Then
                .ObjValorAnteriorStr.ObjValorPro = aobjValorAnterior.ToString
            End If
            If Not IsNothing(aobjValorNuevo) Then
                .ObjValorNuevoStr.ObjValorPro = aobjValorNuevo.ToString
            End If
            .SActualice(True)
        End With
    End Sub
    ''' <summary>
    ''' Crea un nuevo objeto LogApp de tipo "Acción" y lo registra en la base de datos 
    ''' </summary>
    ''' <param name="astrNombreClase">Es el nombre de la clase a la cual pertenece el objeto generó el Log.</param>
    ''' <param name="astrNombreAccion">Nombre de la acción que se está registrando.</param>
    ''' <remarks></remarks>
    Friend Sub SRegistreAccionLogApp(astrNombreClase As String, astrNombreAccion As String)
        If astrNombreAccion Is Nothing Then
            Throw New ArgumentNullException(NameOf(astrNombreAccion))
        End If
        SCargueDtbLogApp()
        Dim lstrAccion = astrNombreAccion
        Dim ldrwLogApp As DataRow = MdtbLogApp.NewRow
        Dim lobjLogApp As New ClsLogApp(Me, ldrwLogApp)
        If Not CType(lobjLogApp.EnuPermisosObj And EnuPermisosDef.enuCrear, Boolean) Then
            lobjLogApp.EnuPermisosObj += EnuPermisosDef.enuCrear
        End If
        If lstrAccion.Length > 250 Then
            lstrAccion = lstrAccion.Substring(0, 250)
        End If
        lobjLogApp.SCreeObj(Nothing)
        With lobjLogApp
            .ObjFechaCreacionDtm.ObjValorPro = Now
            .ObjIdAppLogShr.ObjValorPro = CType(GenuIdAplicacion, Short)
            If GenuIdAplicacion = EnuListaAplicaciones.EnuAdministrador Then
                .ObjIdCarpetaLogShr.ObjValorPro = 0
                .ObjIdCentroUtilLogShr.ObjValorPro = 0
            Else
                .ObjIdCarpetaLogShr.ObjValorPro = GshrIdCarpeta
                .ObjIdCentroUtilLogShr.ObjValorPro = GshrIdCentroUtil
            End If
            .ObjIdObjetoStr.ObjValorPro = String.Empty
            .ObjIdUsuarioLogStr.ObjValorPro = GstrIdUsuario
            .ObjNombreClaseStr.ObjValorPro = astrNombreClase
            .ObjNombreCampoLogStr.ObjValorPro = lstrAccion
            .ObjIdLogAppInt.ObjValorPro = 0
            .ObjValorAnteriorStr.ObjValorPro = String.Empty
            .ObjValorNuevoStr.ObjValorPro = String.Empty
            .ObjTipoLogByt.ObjValorPro = CType(EnuTipoLogDef.enuAccion, Byte)
            .SActualice(True)
        End With
    End Sub
    Shared Function FdtbLogApp() As DataTable
        Dim lstrCampos As String = "FechaCreacion, IdLogApp, Dato, NombreClase, Nombre, IdObjeto, IdUsuario," &
                "ValorAnterior, ValorNuevo"
        Dim lstrFiltro As String = FstrFiltroUbicacion()
        Dim lstrSql As String = "SELECT " & lstrCampos & " FROM PanLogApp INNER JOIN PanTblConstantes" &
                " ON PanTblConstantes.IdGrupo = " & EnuGrupoConstantesPanDef.enuTipoLog &
                " AND PanLogApp.IdAplicacion = " & GenuIdAplicacion &
                " AND PanLogApp.IdTblTipoLog = PanTblConstantes.IdConstante" &
                lstrFiltro & " ORDER BY IdLogApp"
        Dim ldtbLogApp As DataTable = FdtbDataTable(lstrSql)
        Return ldtbLogApp
    End Function
    Shared Function FdrwTiposLog() As DataRow()
        Dim ldrwTiposLog As DataRow()
        ldrwTiposLog = ClsAdministrador.FdrwConstantesPan(EnuGrupoConstantesPanDef.enuTipoLog)
        Return ldrwTiposLog
    End Function
    Shared Function FdrwNombresClases() As DataRow()
        Dim ldrwNombresClases As DataRow()
        Dim lstrSql As String = "SELECT DISTINCT NombreClase FROM PanLogApp" & FstrFiltroUbicacion() &
                " AND IdAplicacion = " & GenuIdAplicacion & " ORDER BY NombreClase"
        ldrwNombresClases = FdtbDataTable(lstrSql).Select
        Return ldrwNombresClases
    End Function
    Shared Function FdrwNombresDatos() As DataRow()
        Dim ldrwNombresDatos As DataRow()
        Dim lstrSql As String = "SELECT DISTINCT Nombre FROM PanLogApp" & FstrFiltroUbicacion() &
                " AND IdAplicacion = " & GenuIdAplicacion & " ORDER BY Nombre"
        ldrwNombresDatos = FdtbDataTable(lstrSql).Select
        Return ldrwNombresDatos
    End Function
    Shared Function FdrwUsuarios() As DataRow()
        Dim ldrwUsuarios As DataRow()
        Dim lstrSql As String = "SELECT DISTINCT IdUsuario FROM PanLogApp" & FstrFiltroUbicacion() &
                " AND IdAplicacion = " & GenuIdAplicacion & " ORDER BY IdUsuario"
        ldrwUsuarios = FdtbDataTable(lstrSql).Select
        Return ldrwUsuarios
    End Function
    Private Shared Function FstrFiltroUbicacion() As String
        Return " WHERE IdCarpeta = " & GshrIdCarpeta & " AND IdCentroUtil = " & GshrIdCentroUtil
    End Function
    ''' <summary>
    ''' Establece la variable global que contiene la trayectoria de la carpeta de datos actual y crea,
    ''' si no existe, la Carpetas de Reportes e Imagenes.
    ''' </summary>
    ''' <param name="astrNombreCarpeta"></param>
    ''' <remarks></remarks>
    Friend Shared Sub SEstablezcaTrayaectoriaDatos(astrNombreCarpeta As String)
        GstrTrayDatos = GstrTrayDat & astrNombreCarpeta
        GstrTrayReportes = GstrTrayDatos & "\" & "Reportes"
        If Not My.Computer.FileSystem.DirectoryExists(GstrTrayReportes) Then
            My.Computer.FileSystem.CreateDirectory(GstrTrayReportes)
        End If
        GstrTrayEmails = GstrTrayReportes & "\" & "Emails"
        If Not My.Computer.FileSystem.DirectoryExists(GstrTrayEmails) Then
            My.Computer.FileSystem.CreateDirectory(GstrTrayEmails)
        End If
        GstrTrayEFac = GstrTrayReportes & "\" & "EFactura"
        If Not My.Computer.FileSystem.DirectoryExists(GstrTrayEFac) Then
            My.Computer.FileSystem.CreateDirectory(GstrTrayEFac)
        End If
        GstrTrayInterfContable = GstrTrayReportes & "\Interfaz Contable"
        If Not My.Computer.FileSystem.DirectoryExists(GstrTrayInterfContable) Then
            My.Computer.FileSystem.CreateDirectory(GstrTrayInterfContable)
        End If
    End Sub
#End Region
End Class
