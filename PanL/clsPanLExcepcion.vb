Imports System.Security
<Serializable()>
Friend Class PanLException
    Inherits Exception
    Protected Property EntNumeroError As Integer = 0
    Protected Property EnuTipoError As EnuSeveridadMen = EnuSeveridadMen.None
    Public Sub New()
        MyBase.New()
    End Sub
    Public Sub New(astrMensaje As String)
        MyBase.New(astrMensaje)
    End Sub
    Public Sub New(astrMensaje As String, aenuTipoError As EnuSeveridadMen)
        MyBase.New(astrMensaje)
        enuTipoError = aenuTipoError
    End Sub
    Public Sub New(astrMensaje As String, aexcInner As Exception)
        MyBase.New(astrMensaje, aexcInner)
    End Sub
    Protected Sub New(info As Runtime.Serialization.SerializationInfo,
        context As _
        Runtime.Serialization.StreamingContext)
        MyBase.New(info, context)
    End Sub
    Friend ReadOnly Property NumeroErr As Integer
        Get
            Return entNumeroError
        End Get
    End Property
    Friend ReadOnly Property Severidad As EnuSeveridadMen
        Get
            Return enuTipoError
        End Get
    End Property
    <SecurityCriticalAttribute()>
    Public Overrides Sub GetObjectData(info As System.Runtime.Serialization.SerializationInfo, context As System.Runtime.Serialization.StreamingContext)
        MyBase.GetObjectData(info, context)
        info.AddValue("NroError", EntNumeroError)
    End Sub
End Class
<Serializable()>
Friend Class ErrorInesperadoPanLException
    Inherits PanLException
    Public Sub New()
        MyBase.New()
        EntNumeroError = 101
        EnuTipoError = EnuSeveridadMen.enuError
    End Sub
    Public Sub New(astrMensaje As String)
        MyBase.New(astrMensaje)
        EntNumeroError = 101
        EnuTipoError = EnuSeveridadMen.enuError
    End Sub
    Public Sub New(astrMensaje As String, aexcInner As Exception)
        MyBase.New(astrMensaje, aexcInner)
        EntNumeroError = 101
        EnuTipoError = EnuSeveridadMen.enuError
    End Sub
    Protected Sub New(info As Runtime.Serialization.SerializationInfo,
    context As _
    Runtime.Serialization.StreamingContext)
        MyBase.New(info, context)
    End Sub
End Class
<Serializable()>
Friend Class ValorPropiedadInvalidoException
    Inherits PanLException
    Public Sub New()
        MyBase.New()
        EntNumeroError = 102
        EnuTipoError = EnuSeveridadMen.enuError
    End Sub
    Public Sub New(astrMensaje As String)
        MyBase.New(astrMensaje)
        EntNumeroError = 102
        EnuTipoError = EnuSeveridadMen.enuError
    End Sub
    Public Sub New(astrMensaje As String, aexcInner As Exception)
        MyBase.New(astrMensaje, aexcInner)
        EntNumeroError = 102
        EnuTipoError = EnuSeveridadMen.enuError
    End Sub
    Protected Sub New(info As Runtime.Serialization.SerializationInfo,
    context As _
    Runtime.Serialization.StreamingContext)
        MyBase.New(info, context)
    End Sub
End Class
<Serializable()>
Friend Class ValorArgumentoInvalidoException
    Inherits PanLException
    Public Sub New()
        MyBase.New()
        entNumeroError = 103
        enuTipoError = EnuSeveridadMen.enuError
    End Sub
    Public Sub New(astrMensaje As String)
        MyBase.New(astrMensaje)
        entNumeroError = 103
        enuTipoError = EnuSeveridadMen.enuError
    End Sub
    Public Sub New(astrMensaje As String, aexcInner As Exception)
        MyBase.New(astrMensaje, aexcInner)
        entNumeroError = 103
        enuTipoError = EnuSeveridadMen.enuError
    End Sub
    Protected Sub New(info As Runtime.Serialization.SerializationInfo,
    context As _
    Runtime.Serialization.StreamingContext)
        MyBase.New(info, context)
    End Sub
End Class
