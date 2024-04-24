Public Class ClsBPBBKL
    Private _statusQTY As String
    Public Property StatusQTY() As String
        Get
            Return _statusQTY
        End Get
        Set(ByVal value As String)
            _statusQTY = value
        End Set
    End Property

    Private _statusExp As String
    Public Property StatusExp() As String
        Get
            Return _statusExp
        End Get
        Set(ByVal value As String)
            _statusExp = value
        End Set
    End Property

    Private _statusDesc As String
    Public Property StatusDesc() As String
        Get
            Return _statusDesc
        End Get
        Set(ByVal value As String)
            _statusDesc = value
        End Set
    End Property

    Private _feedback As String
    Public Property Feedback() As String
        Get
            Return _feedback
        End Get
        Set(ByVal value As String)
            _feedback = value
        End Set
    End Property

    Private _bkl As ClsBKL
    Public Property BKL() As ClsBKL
        Get
            Return _bkl
        End Get
        Set(ByVal value As ClsBKL)
            _bkl = value
        End Set
    End Property


End Class
