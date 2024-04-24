Public Class ClsLogin


    Private _status As Integer
    Public Property Status() As Integer
        Get
            Return _status
        End Get
        Set(ByVal value As Integer)
            _status = value
        End Set
    End Property


    Private _message As String
    Public Property Message() As String
        Get
            Return _message
        End Get
        Set(ByVal value As String)
            _message = value
        End Set
    End Property


    Private _user As ClsUser
    Public Property User() As ClsUser
        Get
            Return _user
        End Get
        Set(ByVal value As ClsUser)
            _user = value
        End Set
    End Property



End Class
