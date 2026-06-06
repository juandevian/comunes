Friend Class ClsMensajePan
#Region "Definiciones"
    ' Variables
    Private MstrMensaje As String = String.Empty
    Private MstrMensajeErr As String = String.Empty
    Private MenuSeveridad As EnuSeveridadMen = EnuSeveridadMen.None
    Private MintNumeroError As Integer = 0
#End Region
#Region "Constructores"
    Public Sub New(aobjRegistro As Object)
        If aobjRegistro Is Nothing OrElse Not (aobjRegistro.GetType.Name = "String" AndAlso
                aobjRegistro = GCOBJREGISTRO) Then
            Throw New ModuloNoRegistradoPanException()
        End If
    End Sub
#End Region
#Region "Propiedades"
    Public Property BlnErrorIndeterminado As Boolean = False
    Public Property BlnNotificacionSonora As Boolean = True
    Friend ReadOnly Property BlnHayMensaje() As Boolean
        Get
            Return Not String.IsNullOrEmpty(MstrMensaje)
        End Get
    End Property
    Friend ReadOnly Property BlnHayMensajeErr() As Boolean
        Get
            Return Not String.IsNullOrEmpty(MstrMensajeErr)
        End Get
    End Property
    Friend ReadOnly Property StrMensaje As String
        Get
            Return MstrMensaje
        End Get
    End Property
    Friend ReadOnly Property EntNumeroError As Integer
        Get
            Return MintNumeroError
        End Get
    End Property
    Friend ReadOnly Property StrMensajeErr As String
        Get
            Return MstrMensajeErr
        End Get
    End Property
    Friend ReadOnly Property EnuSeveridad As EnuSeveridadMen
        Get
            Return MenuSeveridad
        End Get
    End Property
#End Region
#Region "Procedimientos y funciones"
    Friend Sub SInicialiceMensaje()
        MstrMensaje = String.Empty
        MstrMensajeErr = String.Empty
        MintNumeroError = 0
        BlnErrorIndeterminado = False
    End Sub
    Friend Sub SInicialiceMensajeErr()
        MstrMensajeErr = String.Empty
        MintNumeroError = 0
        BlnErrorIndeterminado = False
    End Sub
    Friend Sub SRegistreMensajeEx(astrMsgEx As String)
        MstrMensajeErr = astrMsgEx
        MenuSeveridad = EnuSeveridadMen.enuError
    End Sub
    Friend Sub SRegistreMensaje(astrMsg As String, aenuSeveridad As EnuSeveridadMen)
        If aenuSeveridad = EnuSeveridadMen.enuError Then
            MstrMensaje = astrMsg
            MenuSeveridad = EnuSeveridadMen.enuError
        Else
            If Not BlnHayMensaje Then
                SIncorporeMensaje(astrMsg, aenuSeveridad)
            End If
        End If
    End Sub
    Friend Sub SRegistreMensaje(astrMsg As String, aintNroError As Integer)
        MintNumeroError = aintNroError
        MstrMensaje = astrMsg
    End Sub
    Private Sub SIncorporeMensaje(astrMsg As String, aenuSeveridad As EnuSeveridadMen)
        MstrMensaje = astrMsg
        If Not String.IsNullOrEmpty(MstrMensaje) Then
            MenuSeveridad = aenuSeveridad
        Else
            MenuSeveridad = EnuSeveridadMen.None
        End If
    End Sub
#End Region
End Class
