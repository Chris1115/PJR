Public Class ClsAktiva

    Private _nSeri As String
    Public Property NSeri() As String
        Get
            Return _nSeri
        End Get
        Set(ByVal value As String)
            _nSeri = value
        End Set
    End Property

    Private _deskripsi As String
    Public Property Deskripsi() As String
        Get
            Return _deskripsi
        End Get
        Set(ByVal value As String)
            _deskripsi = value
        End Set
    End Property

    Private _deskripsi2 As String
    Public Property Deskripsi2() As String
        Get
            Return _deskripsi2
        End Get
        Set(ByVal value As String)
            _deskripsi2 = value
        End Set
    End Property

    Private _qtySO As Integer
    Public Property QtySO() As Integer
        Get
            Return _qtySO
        End Get
        Set(ByVal value As Integer)
            _qtySO = value
        End Set
    End Property

    Private _qtyMax As String
    Public Property QtyMax() As String
        Get
            Return _qtyMax
        End Get
        Set(ByVal value As String)
            _qtyMax = value
        End Set
    End Property

    Private _statusQty As Boolean
    Public Property statusQty() As Boolean
        Get
            Return _statusQty
        End Get
        Set(ByVal value As Boolean)
            _statusQty = value
        End Set
    End Property

    Private _recid As String
    Public Property Recid() As String
        Get
            Return _recid
        End Get
        Set(ByVal value As String)
            _recid = value
        End Set
    End Property
End Class
