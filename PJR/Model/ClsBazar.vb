Public Class ClsBazar
    Private _recid As String
    Public Property Recid() As String
        Get
            Return _recid
        End Get
        Set(ByVal value As String)
            _recid = value
        End Set
    End Property

    Private _barcodePlu As String
    Public Property BarcodePlu() As String
        Get
            Return _barcodePlu
        End Get
        Set(ByVal value As String)
            _barcodePlu = value
        End Set
    End Property

    Private _prdcd As String
    Public Property PRDCD() As String
        Get
            Return _prdcd
        End Get
        Set(ByVal value As String)
            _prdcd = value
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
    Private _qtyInput As String
    Public Property QTYInput() As String
        Get
            Return _qtyInput
        End Get
        Set(ByVal value As String)
            _qtyInput = value
        End Set
    End Property

    Private _qtyTotal As String
    Public Property QTYTotal() As String
        Get
            Return _qtyTotal
        End Get
        Set(ByVal value As String)
            _qtyTotal = value
        End Set
    End Property

    Private _qtyCom As String
    Public Property QTYCom() As String
        Get
            Return _qtyCom
        End Get
        Set(ByVal value As String)
            _qtyCom = value
        End Set
    End Property

    Private _rak As String
    Public Property Rak() As String
        Get
            Return _rak
        End Get
        Set(ByVal value As String)
            _rak = value
        End Set
    End Property

    Private _tgl_Exp As String
    Public Property Tgl_exp() As String
        Get
            Return _tgl_Exp
        End Get
        Set(ByVal value As String)
            _tgl_Exp = value
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

    Private _feedback As String
    Public Property Feedback() As String
        Get
            Return _feedback
        End Get
        Set(ByVal value As String)
            _feedback = value
        End Set
    End Property

    Private _bzr As ClsBazar
    Public Property BZR() As ClsBazar
        Get
            Return _bzr
        End Get
        Set(ByVal value As ClsBazar)
            _bzr = value
        End Set
    End Property
End Class
