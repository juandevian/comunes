Imports System.Security
<Serializable()>
Friend Class PanDatException
    Inherits Exception
    Public Sub New()
        MyBase.New()
    End Sub
    Public Sub New(astrMensaje As String)
        MyBase.New(astrMensaje)
    End Sub
    Public Sub New(astrMensaje As String, aexcInner As Exception)
        MyBase.New(astrMensaje, aexcInner)
    End Sub
    Protected Sub New(info As Runtime.Serialization.SerializationInfo,
        context As _
        Runtime.Serialization.StreamingContext)
        MyBase.New(info, context)
    End Sub
    Protected Property EntNumeroError As Integer = 0
    Friend ReadOnly Property NumeroErr As Integer
        Get
            Return entNumeroError
        End Get
    End Property
    <SecurityCriticalAttribute()> _
    Public Overrides Sub GetObjectData(info As System.Runtime.Serialization.SerializationInfo, _
                                              context As System.Runtime.Serialization.StreamingContext)
        MyBase.GetObjectData(info, context)
        info.AddValue("NroError", entNumeroError)
    End Sub
End Class
<Serializable()>
Friend Class ModuloNoRegistradoPanException
    Inherits PanDatException
    Public Sub New()
        MyBase.New()
        entNumeroError = 1
    End Sub
    Public Sub New(astrMensaje As String)
        MyBase.New(astrMensaje)
        entNumeroError = 1
    End Sub
    Public Sub New(astrMensaje As String, aexcInner As Exception)
        MyBase.New(astrMensaje, aexcInner)
        entNumeroError = 1
    End Sub
    Protected Sub New(info As Runtime.Serialization.SerializationInfo,
        context As _
        Runtime.Serialization.StreamingContext)
        MyBase.New(info, context)
    End Sub
    Public Overrides ReadOnly Property Message As String
        Get
            Dim lstrMens = "El módulo no ha sido registrado debidamente."
            Return lstrMens
        End Get
    End Property
End Class
<Serializable()>
Friend Class ArgumentoInvalidoPanException
    Inherits PanDatException
    Public Sub New()
        MyBase.New()
        EntNumeroError = 2
    End Sub
    Public Sub New(astrMensaje As String)
        MyBase.New(astrMensaje)
        EntNumeroError = 2
    End Sub
    Public Sub New(astrMensaje As String, aexcInner As Exception)
        MyBase.New(astrMensaje, aexcInner)
        EntNumeroError = 2
    End Sub
    Protected Sub New(info As Runtime.Serialization.SerializationInfo,
    context As _
    Runtime.Serialization.StreamingContext)
        MyBase.New(info, context)
    End Sub
End Class
<Serializable()>
Friend Class SintaxisInvalidaPanException
    Inherits PanDatException
    Public Sub New()
        MyBase.New()
        entNumeroError = 3
    End Sub
    Public Sub New(astrMensaje As String)
        MyBase.New(astrMensaje)
        entNumeroError = 3
    End Sub
    Public Sub New(astrMensaje As String, aexcInner As Exception)
        MyBase.New(astrMensaje, aexcInner)
        entNumeroError = 3
    End Sub
    Protected Sub New(info As Runtime.Serialization.SerializationInfo,
    context As _
    Runtime.Serialization.StreamingContext)
        MyBase.New(info, context)
    End Sub
End Class
<Serializable()>
Friend Class ProveedorBdPanException
    Inherits PanDatException
    Public Sub New()
        MyBase.New()
        entNumeroError = 4
    End Sub
    Public Sub New(astrMensaje As String)
        MyBase.New(astrMensaje)
        entNumeroError = 4
    End Sub
    Public Sub New(astrMensaje As String, aexcInner As Exception)
        MyBase.New(astrMensaje, aexcInner)
        entNumeroError = 4
    End Sub
    Protected Sub New(info As Runtime.Serialization.SerializationInfo,
    context As _
    Runtime.Serialization.StreamingContext)
        MyBase.New(info, context)
    End Sub
End Class
<Serializable()>
Friend Class ParametrosConexionBDPanException
    Inherits PanDatException
    Public Sub New()
        MyBase.New()
        entNumeroError = 5
    End Sub
    Public Sub New(astrMensaje As String)
        MyBase.New(astrMensaje)
        entNumeroError = 5
    End Sub
    Public Sub New(astrMensaje As String, aexcInner As Exception)
        MyBase.New(astrMensaje, aexcInner)
        entNumeroError = 5
    End Sub
    Protected Sub New(info As Runtime.Serialization.SerializationInfo,
    context As _
    Runtime.Serialization.StreamingContext)
        MyBase.New(info, context)
    End Sub
End Class
<Serializable()>
Friend Class ConexionBdPanException
    Inherits PanDatException
    Public Sub New()
        MyBase.New()
        entNumeroError = 6
    End Sub
    Public Sub New(astrMensaje As String)
        MyBase.New(astrMensaje)
        entNumeroError = 6
    End Sub
    Public Sub New(astrMensaje As String, aexcInner As Exception)
        MyBase.New(astrMensaje, aexcInner)
        entNumeroError = 6
    End Sub
    Protected Sub New(info As Runtime.Serialization.SerializationInfo,
    context As _
    Runtime.Serialization.StreamingContext)
        MyBase.New(info, context)
    End Sub
End Class
<Serializable()>
Friend Class ArchivoXmlPanException
    Inherits PanDatException
    Public Sub New()
        MyBase.New()
        entNumeroError = 7
    End Sub
    Public Sub New(astrMensaje As String)
        MyBase.New(astrMensaje)
        entNumeroError = 7
    End Sub
    Public Sub New(astrMensaje As String, aexcInner As Exception)
        MyBase.New(astrMensaje, aexcInner)
        entNumeroError = 7
    End Sub
    Protected Sub New(info As Runtime.Serialization.SerializationInfo,
    context As _
    Runtime.Serialization.StreamingContext)
        MyBase.New(info, context)
    End Sub
End Class
<Serializable()>
Friend Class ErrorInesperadoPanDatException
    Inherits PanDatException
    Public Sub New()
        MyBase.New()
        entNumeroError = 8
    End Sub
    Public Sub New(astrMensaje As String)
        MyBase.New(astrMensaje)
        entNumeroError = 8
    End Sub
    Public Sub New(astrMensaje As String, aexcInner As Exception)
        MyBase.New(astrMensaje, aexcInner)
        entNumeroError = 8
    End Sub
    Protected Sub New(info As Runtime.Serialization.SerializationInfo,
    context As _
    Runtime.Serialization.StreamingContext)
        MyBase.New(info, context)
    End Sub
End Class
<Serializable()>
Friend Class AplicacionSinInstalarException
    Inherits PanDatException
    Public Sub New()
        MyBase.New()
        entNumeroError = 9
    End Sub
    Public Sub New(astrMensaje As String)
        MyBase.New(astrMensaje)
        entNumeroError = 9
    End Sub
    Public Sub New(astrMensaje As String, aexcInner As Exception)
        MyBase.New(astrMensaje, aexcInner)
        entNumeroError = 9
    End Sub
    Protected Sub New(info As Runtime.Serialization.SerializationInfo,
    context As _
    Runtime.Serialization.StreamingContext)
        MyBase.New(info, context)
    End Sub
End Class
<Serializable()>
Friend Class AplicacionInstaladaException
    Inherits PanDatException
    Public Sub New()
        MyBase.New()
        entNumeroError = 9
    End Sub
    Public Sub New(astrMensaje As String)
        MyBase.New(astrMensaje)
        entNumeroError = 9
    End Sub
    Public Sub New(astrMensaje As String, aexcInner As Exception)
        MyBase.New(astrMensaje, aexcInner)
        entNumeroError = 9
    End Sub
    Protected Sub New(info As Runtime.Serialization.SerializationInfo,
    context As _
    Runtime.Serialization.StreamingContext)
        MyBase.New(info, context)
    End Sub
End Class