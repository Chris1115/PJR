Public Class ClsClient


    Private _socketId As String
    Public Property SocketID() As String
        Get
            Return _socketId
        End Get
        Set(ByVal value As String)
            _socketId = value
        End Set
    End Property


    Private _ipAddresss As String
    Public Property IpAddress() As String
        Get
            Return _ipAddresss
        End Get
        Set(ByVal value As String)
            _ipAddresss = value
        End Set
    End Property

    Private _login As ClsLogin
    Public Property Login() As ClsLogin
        Get
            Return _login
        End Get
        Set(ByVal value As ClsLogin)
            _login = value
        End Set
    End Property

    Private _so As ClsSo
    Public Property SO() As ClsSo
        Get
            Return _so
        End Get
        Set(ByVal value As ClsSo)
            _so = value
        End Set
    End Property

    Private _soED As ClsSOED
    Public Property SOED() As ClsSOED
        Get
            Return _soED
        End Get
        Set(ByVal value As ClsSOED)
            _soED = value
        End Set
    End Property

End Class
