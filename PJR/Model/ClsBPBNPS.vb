Public Class ClsBPBNPS
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

    Private _nps As CLSNPS
    Public Property NPS() As CLSNPS
        Get
            Return _nps
        End Get
        Set(ByVal value As CLSNPS)
            _nps = value
        End Set
    End Property


End Class
