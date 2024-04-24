Public Class ClsHandheld
    Private _ipAddress As String
    Public Property ipAddress() As String
        Get
            Return _ipAddress
        End Get
        Set(ByVal value As String)
            _ipAddress = value
        End Set
    End Property

    Private _socketID As String
    Public Property socketID() As String
        Get
            Return _socketID
        End Get
        Set(ByVal value As String)
            _socketID = value
        End Set
    End Property

    Private _jenis_so As String
    Public Property jenis_so() As String
        Get
            Return _jenis_so
        End Get
        Set(ByVal value As String)
            _jenis_so = value
        End Set
    End Property

    Private _lokasi_so As String
    Public Property lokasi_so() As String
        Get
            Return _lokasi_so
        End Get
        Set(ByVal value As String)
            _lokasi_so = value
        End Set
    End Property

End Class
