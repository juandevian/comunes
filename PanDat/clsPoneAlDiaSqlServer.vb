Imports System.Data.SqlClient
Imports TyS.Panorama.GesDat.clsEstructuraBD
Friend Class clsPoneAlDiaSqlServer
    'Herencia
    Inherits clsCBPoneAlDiaBD
    ' Objeto error
    'variables de propiedad
    Private mstrTrayectoriaMDF As String = ""
    Private mstrTrayectoriaLDF As String = ""
    Private mintTamañoMDF As Integer = 50
    Private mintTamañoLDF As Integer = 25
    Private mintCrecimientoMDF As Integer = 20
    Private mintCrecimientoLDF As Integer = 10
    Private mstrTabla As String = ""
    Private mstrDefColumna As String = ""
    Private mstrBaseDatos As String = ""
    'Otras Vriables
    Private menuTipoAutenticacion As clsActualizaBD.enuTATipoAutenticacionDef
    Private mblnHayError As Boolean = False
    Private mstrCommandText As String = ""
    Private mobjPadre As clsActualizaBD = Nothing

    Public Sub New(ByVal aobjPadre As clsActualizaBD,ByVal astrArchivoXml As String)
        If gobjPanDat.blnRegistrado Then
            mobjPadre = aobjPadre
            hobjEstructuraBD = New clsEstructuraBD
            sActualiceBaseDatos(astrArchivoXml)
        Else
            Throw New System.InvalidOperationException("El módulo no esta registrado")
        End If
    End Sub

    '=======================
    'Propiedades de la clase
    '=======================
    Public Property strNombreBaseDatos() As String
        Get
            Return mstrBaseDatos
        End Get
        Set(ByVal value As String)
            mstrBaseDatos = value
        End Set
    End Property

    Public Property strTrayectoriaArchivoMDF() As String
        Get
            Return mstrTrayectoriaMDF
        End Get
        Set(ByVal value As String)
            mstrTrayectoriaMDF = value
            clsActualizaBD.sVerifiqueExistenciaDir(mstrTrayectoriaMDF)
            If clsActualizaBD.fblnHayError() Then
                mstrTrayectoriaMDF = ""
                MsgBox("La trayectoria para el archivo MDF es invalida.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Error")
            End If
            mblnHayError = False
        End Set
    End Property

    Public Property strTrayectoriaArchivoLDF() As String
        Get
            Return mstrTrayectoriaLDF
        End Get
        Set(ByVal value As String)
            mstrTrayectoriaLDF = value
            clsActualizaBD.sVerifiqueExistenciaDir(mstrTrayectoriaLDF)
            If clsActualizaBD.fblnHayError() Then
                mstrTrayectoriaLDF = ""
                MsgBox("La trayectoria para el archivo LDF es invalida.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Error")
            End If
            mblnHayError = False
        End Set
    End Property

    Public Property intTamañoArchivoMDF() As Integer
        Get
            Return mintTamañoMDF
        End Get
        Set(ByVal value As Integer)
            mintTamañoMDF = value
        End Set
    End Property

    Public Property intTamañoArchivoLDF() As Integer
        Get
            Return mintTamañoLDF
        End Get
        Set(ByVal value As Integer)
            mintTamañoLDF = value
        End Set
    End Property

    Public Property intPorcentajeCrecimientoMDF() As Integer
        Get
            Return mintCrecimientoMDF
        End Get
        Set(ByVal value As Integer)
            mintCrecimientoMDF = value
        End Set
    End Property

    Public Property intPorcentajeCrecimientoLDF() As Integer
        Get
            Return mintCrecimientoLDF
        End Get
        Set(ByVal value As Integer)
            mintCrecimientoLDF = value
        End Set
    End Property

    Public Property strNombreTabla() As String
        Get
            Return mstrTabla
        End Get
        Set(ByVal value As String)
            mstrTabla = value
        End Set
    End Property

    Public Property strDefinicionColumna() As String
        Get
            Return mstrDefColumna
        End Get
        Set(ByVal value As String)
            mstrDefColumna = value
        End Set
    End Property
    '====================================
    ' Funciones y procedimientos Publicos
    '====================================
    Friend Sub sCreeBaseDatosSqlSer(Optional ByVal astrNombreBaseDatos As String = "", _
            Optional ByVal astrTrayectoriaMDF As String = "", _
            Optional ByVal astrTrayectoriaLDF As String = "", Optional ByVal aintTamañoMDF As Integer = 50, _
            Optional ByVal aintTamañoLDF As Integer = 25, Optional ByVal aintPorcentajeCrecimientoMDF As Integer = 20, _
            Optional ByVal aintPorcentajeCrecimientoLDF As Integer = 10, _
            Optional ByVal aenuTipoAutenticacion As clsActualizaBD.enuTATipoAutenticacionDef = _
                   clsActualizaBD.enuTATipoAutenticacionDef.enuTANinguna)
        Dim lstrMensaje As String = ""
        If astrNombreBaseDatos <> "" Then
            mstrBaseDatos = astrNombreBaseDatos
        End If
        If astrTrayectoriaMDF <> "" Then
            mstrTrayectoriaMDF = astrTrayectoriaMDF
        End If
        If astrTrayectoriaLDF <> "" Then
            mstrTrayectoriaLDF = astrTrayectoriaLDF
        End If
        If aintTamañoMDF <> 50 Then
            mintCrecimientoLDF = aintTamañoMDF
        End If
        If aintTamañoLDF <> 25 Then
            mintCrecimientoLDF = aintTamañoMDF
        End If
        If aintPorcentajeCrecimientoMDF <> 20 Then
            mintCrecimientoMDF = aintPorcentajeCrecimientoMDF
        End If
        If aintPorcentajeCrecimientoLDF <> 10 Then
            mintCrecimientoMDF = aintPorcentajeCrecimientoMDF
        End If
        If aenuTipoAutenticacion <> clsActualizaBD.enuTATipoAutenticacionDef.enuTANinguna Then
            menuTipoAutenticacion = aenuTipoAutenticacion
        End If

        If mstrBaseDatos = "" Then
            lstrMensaje &= "Falta el nombre de la base de datos"
        End If
        If mstrTrayectoriaMDF = "" Then
            If lstrMensaje = "" Then
                lstrMensaje &= "Falta la trayectoria para el archivo MDF de la base de datos"
            Else
                lstrMensaje &= ", la trayectoria para el archivo MDF de la base de datos"
            End If
        End If
        If mstrTrayectoriaLDF = "" Then
            If lstrMensaje = "" Then
                lstrMensaje &= "Falta la trayectoria para el archivo LDF de la base de datos"
            Else
                lstrMensaje &= ", la trayectoria para el archivo LDF de la base de datos"
            End If
        End If
        If menuTipoAutenticacion = clsActualizaBD.enuTATipoAutenticacionDef.enuTANinguna Then
            If lstrMensaje = "" Then
                lstrMensaje &= "Falta establecer el tipo de autenticación"
            Else
                lstrMensaje &= " y el tipo de autenticación"
            End If
        End If
        If lstrMensaje <> "" Then
            lstrMensaje &= "."
            mblnHayError = True
            ''sActualiceError(lstrMensaje, "clsActualiceBD")
        Else
            sCreeBD()
        End If
    End Sub

    'Friend Sub sCreeTabla(Optional ByVal astrNombreBaseDatos As String = "", _
    '        Optional ByVal astrNombreTabla As String = "", _
    '        Optional ByVal astrDefColumnas() As String = Nothing, _
    '        Optional ByVal aenuTipoAutenticacion As clsActualizaBD.enuTATipoAutenticacionDef = _
    '                clsActualizaBD.enuTATipoAutenticacionDef.enuTANinguna)
    '    Dim lstrMensaje As String = ""
    '    If astrNombreBaseDatos <> "" Then
    '        'mstrBaseDatos = astrNombreBaseDatos
    '    End If
    '    If astrNombreTabla <> "" Then
    '        mstrTabla = astrNombreTabla
    '    End If
    '    If aenuTipoAutenticacion <> clsActualizaBD.enuTATipoAutenticacionDef.enuTANinguna Then
    '        menuTipoAutenticacion = aenuTipoAutenticacion
    '    End If
    '    'If mstrBaseDatos = "" Then
    '    '    lstrMensaje &= "Falta el nombre de la base de datos"
    '    'End If
    '    If mstrTabla = "" Then
    '        If lstrMensaje = "" Then
    '            lstrMensaje &= "Falta el nombre de la tabla"
    '        Else
    '            lstrMensaje &= ", el nombre de la tabla"
    '        End If
    '    End If
    '    If menuTipoAutenticacion = clsActualizaBD.enuTATipoAutenticacionDef.enuTANinguna Then
    '        If lstrMensaje = "" Then
    '            lstrMensaje &= "Falta establecer el tipo de autenticación"
    '        Else
    '            lstrMensaje &= " y el tipo de autenticación"
    '        End If
    '    End If
    '    If lstrMensaje <> "" Then
    '        lstrMensaje &= "."
    '        mblnHayError = True
    '        'sActualiceError(lstrMensaje, "clsActualiceBD")
    '        MsgBox(lstrMensaje, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Error")
    '    Else
    '        sCreeTabla(astrDefColumnas)
    '    End If
    '    mblnHayError = False
    'End Sub

    Friend Sub sCreeColumnas(ByVal astrDefColumnas() As String, Optional ByVal astrNombreBaseDatos As String = "", _
            Optional ByVal astrNombreTabla As String = "")
        Dim lstrMensaje As String = ""
        If astrNombreBaseDatos <> "" Then
            'mstrBaseDatos = astrNombreBaseDatos
        End If
        If astrNombreTabla <> "" Then
            mstrTabla = astrNombreTabla
        End If
        'If mstrBaseDatos = "" Then
        '    lstrMensaje &= "Falta el nombre de la base de datos"
        'End If
        If mstrTabla = "" Then
            If lstrMensaje = "" Then
                lstrMensaje &= "Falta el nombre de la tabla"
            Else
                lstrMensaje &= ", el nombre de la tabla"
            End If
        End If
        If menuTipoAutenticacion = clsActualizaBD.enuTATipoAutenticacionDef.enuTANinguna Then
            If lstrMensaje = "" Then
                lstrMensaje &= "Falta establecer el tipo de autenticación"
            Else
                lstrMensaje &= " y el tipo de autenticación"
            End If
        End If
        If astrDefColumnas.Count = 0 Or astrDefColumnas Is Nothing Then
            If lstrMensaje = "" Then
                lstrMensaje &= "Falta establecer las definiciones de Columna"
            Else
                lstrMensaje &= " y las definiciones de Columna"
            End If
        End If
        If lstrMensaje <> "" Then
            lstrMensaje &= "."
            mblnHayError = True
            'sActualiceError(lstrMensaje, "clsActualiceBD")
            MsgBox(lstrMensaje, "clsActualiceBD")
        Else
            sModifiqueColumnas()
        End If
    End Sub

    Friend Sub sModifiqueColumna(Optional ByVal astrNombreBaseDatos As String = "",
            Optional ByVal astrNombreTabla As String = "",
            Optional ByVal astrDefColumna As String = "",
            Optional ByVal aenuTipoAutenticacion As clsActualizaBD.enuTATipoAutenticacionDef = _
                    clsActualizaBD.enuTATipoAutenticacionDef.enuTANinguna)
        Dim lstrMensaje As String = ""
        Dim lstrDefColumna(0) As String
        If astrNombreBaseDatos <> "" Then
            'mstrBaseDatos = astrNombreBaseDatos
        End If
        If astrNombreTabla <> "" Then
            mstrTabla = astrNombreTabla
        End If
        If aenuTipoAutenticacion <> clsActualizaBD.enuTATipoAutenticacionDef.enuTANinguna Then
            menuTipoAutenticacion = aenuTipoAutenticacion
        End If
        'If mstrBaseDatos = "" Then
        '    lstrMensaje &= "Falta el nombre de la base de datos"
        'End If
        If mstrTabla = "" Then
            If lstrMensaje = "" Then
                lstrMensaje &= "Falta el nombre de la tabla"
            Else
                lstrMensaje &= ", el nombre de la tabla"
            End If
        End If

        If astrDefColumna <> mstrDefColumna Then
            If lstrMensaje = "" Then
                lstrMensaje &= "Falta la definicion de Columna"
            Else
                lstrMensaje &= ", la definicion de Columna"
            End If
        End If
        If menuTipoAutenticacion = clsActualizaBD.enuTATipoAutenticacionDef.enuTANinguna Then
            If lstrMensaje = "" Then
                lstrMensaje &= "Falta establecer el tipo de autenticación"
            Else
                lstrMensaje &= " y el tipo de autenticación"
            End If
        End If
        If lstrMensaje <> "" Then
            lstrMensaje &= "."
            mblnHayError = True
            'sActualiceError(lstrMensaje, "clsActualiceBD")
        Else
            lstrDefColumna(0) = mstrDefColumna
            sModifiqueColumnas()
        End If

    End Sub
    '====================================
    ' Funciones y procedimientos Publicos
    '====================================
    Private Sub sCreeBD()
        If Not fblnExisteEnBaseDatosSQL(enuTOBDTipoObjetoBDDef.enuTOBDBaseDatos, mstrBaseDatos) And _
                (Not mblnHayError) Then
            mstrCommandText &= "Create database " & mstrBaseDatos & vbCrLf
            mstrCommandText &= "on" & vbCrLf
            mstrCommandText &= "( name = N'" & mstrBaseDatos & "DB'," & vbCrLf
            mstrCommandText &= "    filename = '" & mstrTrayectoriaMDF & mstrBaseDatos & ".mdf'," & vbCrLf
            mstrCommandText &= "    size = " & mintTamañoMDF & ","
            mstrCommandText &= "    filegrowth = " & mintCrecimientoMDF & " %)"
            mstrCommandText &= "log on" & vbCrLf
            mstrCommandText &= "( name = N'" & mstrBaseDatos & "BD'," & vbCrLf
            mstrCommandText &= "    filename = '" & mstrTrayectoriaLDF & mstrBaseDatos & ".ldf'," & vbCrLf
            mstrCommandText &= "    size = " & mintTamañoLDF & ","
            mstrCommandText &= "    filegrowth = " & mintCrecimientoLDF & " %)"
            sEjecuteComandoTexto()
        End If
        mblnHayError = False
    End Sub

    'Private Sub sCreeTabla(ByVal astrDefColumnas() As String)
    '    Dim lstrColumnas As String
    '    Dim lstrColumna As String
    '    If astrDefColumnas Is Nothing Then
    '        lstrColumnas = "(cInicio int)"
    '    Else
    '        lstrColumnas = "(" & vbCrLf
    '        For Each lstrColumna In astrDefColumnas
    '            If Not lstrColumna.EndsWith(",") Then
    '                lstrColumna &= ","
    '            End If
    '            lstrColumnas &= lstrColumna & vbCrLf
    '        Next
    '        lstrColumnas &= ")"
    '    End If
    '    If Not fblnExisteEnBaseDatosSQL(enuTOBDTipoObjetoBDDef.enuTOBDTabla, mstrBaseDatos, mstrTabla) And _
    '            Not (mblnHayError) Then
    '        mstrCommandText = "create table " & mstrTabla & vbCrLf
    '        mstrCommandText &= lstrColumnas & vbCrLf
    '        sEjecuteComandoTexto()
    '    End If
    'End Sub

    Private Sub sModifiqueColumnas()
        Dim lstrNomColumna As String = "cInicio"
        Dim lstrDefColumna As String = ""
        ' Aurelio: Manejo de Columnas via transact
        If fblnExisteEnBaseDatosSQL(enuTOBDTipoObjetoBDDef.enuTOBDColumna, mstrBaseDatos, mstrTabla, lstrNomColumna) Then

        End If
        mstrCommandText &= "alter table " & mstrTabla & vbCrLf
        mstrCommandText &= "    with nocheck add" & vbCrLf
        mstrCommandText &= "    " & lstrDefColumna & vbCrLf
        mstrCommandText &= "    if exists(select [name] from sys.columns where object_id = OBJECT_ID('" & mstrTabla & "') and [name] = 'cInicio')" & vbCrLf
        mstrCommandText &= "    begin" & vbCrLf
        mstrCommandText &= "        alter table " & mstrTabla & vbCrLf
        mstrCommandText &= "            drop column cInicio" & vbCrLf
        mstrCommandText &= "    end" & vbCrLf
        sEjecuteComandoTexto()

    End Sub

    Private Function fblnExisteEnBaseDatosSQL(ByVal aenuTipoObjetoBD As enuTOBDTipoObjetoBDDef,
            ByVal astrBaseDatos As String,
            Optional ByVal astrTabla As String = "", Optional ByVal astrColumna As String = "", _
            Optional ByVal astrIndice As String = "", Optional ByVal ablnIndicePrimaryKey As Boolean = False, _
            Optional ByVal astrProcedimientoAlmacenado As String = "")
        Dim lblnExiste As Boolean = True
        ' Base de datos
        If aenuTipoObjetoBD >= enuTOBDTipoObjetoBDDef.enuTOBDBaseDatos Then
            mstrCommandText = "use master;" & vbCrLf
            mstrCommandText &= "select count(*) from sys.sysdatabases where [name] = '" & astrBaseDatos & "'" & vbCrLf
            lblnExiste = fblnExiste()
        End If
        ' Tabla
        If (Not mblnHayError) And lblnExiste Then
            If aenuTipoObjetoBD >= enuTOBDTipoObjetoBDDef.enuTOBDTabla Then
                If astrTabla = "" Then
                    lblnExiste = False
                Else
                    mstrCommandText = "use " & mstrBaseDatos & vbCrLf
                    mstrCommandText &= "select count(*) from sys.tables where [name] = '" & astrTabla & "'" & vbCrLf
                    lblnExiste = fblnExiste()
                End If
            End If
        End If
        ' Columna
        If (Not mblnHayError) And lblnExiste Then
            If aenuTipoObjetoBD >= enuTOBDTipoObjetoBDDef.enuTOBDColumna Then
                If astrColumna = "" Then
                    lblnExiste = False
                Else
                    mstrCommandText = "select count(*) from sys.columns where object_id = OBJECT_ID('" & astrTabla & "') and [name] = '" & astrColumna & "'" & vbCrLf
                    lblnExiste = fblnExiste()
                End If
            End If
        End If
        ' Indice
        If (Not mblnHayError) And lblnExiste Then
            If aenuTipoObjetoBD >= enuTOBDTipoObjetoBDDef.enuTOBDIndice Then
                If astrIndice = "" Then
                    lblnExiste = False
                Else
                    If ablnIndicePrimaryKey Then
                        mstrCommandText = "select count(*) from sys.indexes where object_id = OBJECT_ID('" & astrTabla & "') and Left([name], 2) = 'PK'" & vbCrLf
                    Else
                        mstrCommandText = "select count(*) from sys.indexes where object_id = OBJECT_ID('" & astrTabla & "') and [name] = '" & astrIndice & "'" & vbCrLf
                    End If
                    lblnExiste = fblnExiste()
                End If
            End If
        End If
        ' Procedimiento almacenado
        If (Not mblnHayError) And lblnExiste Then
            If aenuTipoObjetoBD >= enuTOBDTipoObjetoBDDef.enuTOBProcedimientoAlmacenado Then
                If astrProcedimientoAlmacenado = "" Then
                    lblnExiste = False
                Else
                    mstrCommandText = "select count(*) from sys.procedures where [name] = '" & astrProcedimientoAlmacenado & "'" & vbCrLf
                    lblnExiste = fblnExiste()
                End If
            End If
        End If
        Return lblnExiste
    End Function

    Private Function fblnExiste() As Boolean
        Dim lintRes As Integer
        Dim lblnExiste As Boolean = True
        Dim lcmdComando = New SqlClient.SqlCommand
        lcmdComando.Connection = gobjPanDat.cnnConexion
        lcmdComando.CommandType = CommandType.Text
        lcmdComando.CommandText = mstrCommandText
        Try
            lintRes = lcmdComando.ExecuteScalar()
        Catch exep As Exception
            mblnHayError = True
            'sActualiceError(exep.Message, "clsActualiceBD")
            MsgBox(exep.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Error")
            Return False
        End Try
        If Not mblnHayError Then
            If lintRes <= 0 Then
                lblnExiste = False
            End If
        End If
        Return lblnExiste
    End Function

    Private Sub sEjecuteComandoTexto()
        Dim lcmdComando = New SqlClient.SqlCommand
        lcmdComando.Connection = gobjPanDat.cnnConexion
        lcmdComando.CommandType = CommandType.Text
        lcmdComando.CommandText = mstrCommandText
        Try
            lcmdComando.ExecuteNonQuery()
        Catch exep As Exception
            'sActualiceError(exep.Message, "clsActualiceBD")
            mblnHayError = True
            MsgBox(exep.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Error")
            Exit Sub
        End Try

    End Sub

#Region "Procedimientos sobreescritos"
    Protected Overrides Sub sGenereEstructuraActualBD()
        Stop
    End Sub

    Protected Overrides Sub sCreeBaseDatos()
        Stop
    End Sub

    Protected Overrides Sub sActualiceBD()

    End Sub

    Protected Overrides Sub sVerifiqueCollationTabla(ByVal astrCoolationNameOriginalBD As String)
        '
    End Sub

    Protected Overrides Sub sVerifiqueCollationColumna(ByVal aobjColumnaXml As clsEstructuraBD.clsColumna,
            ByRef ablnCambiarColumna As Boolean)
        '
    End Sub

    Protected Overrides Sub sCreeTabla()

    End Sub

    Protected Overrides Sub sRenombreTabla(ByVal astrNombreTablaOri As String,
            ByVal astrNombreTablaDes As String)
        '
    End Sub

    Protected Overrides Sub sVinculeTablas(ByVal astrNombreTabla As String,
        ByVal astrBaseDatosExterna As String)
        '
    End Sub

    Protected Overrides Sub sVersioneBD(ByVal ablnActualiceTamano As Boolean)
        Stop
    End Sub

    Protected Overrides Sub sCompareIndices(ByVal aobjIndiceXml As clsEstructuraBD.clsIndice,
            ByRef ablnDifieren As Boolean)
        '
    End Sub

    Protected Overrides Sub sElimineIndice(ByVal astrNombreTabla As String, ByVal astrNombreIndice As String)
        Stop
    End Sub

    Protected Overrides Sub sCreeIndice(ByVal aobjIndiceXml As clsIndice)
        Stop
    End Sub
#End Region
End Class