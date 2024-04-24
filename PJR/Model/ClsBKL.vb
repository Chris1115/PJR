Public Class ClsBKL

    Private _prdcd As String
    Public Property Prdcd() As String
        Get
            Return _prdcd
        End Get
        Set(ByVal value As String)
            _prdcd = value
        End Set
    End Property

    Private _desc As String
    Public Property Desc() As String
        Get
            Return _desc
        End Get
        Set(ByVal value As String)
            _desc = value
        End Set
    End Property

    Private _qty As String
    Public Property Qty() As String
        Get
            Return _qty
        End Get
        Set(ByVal value As String)
            _qty = value
        End Set
    End Property

    Private _sjqty As String
    Public Property sjQty() As String
        Get
            Return _sjqty
        End Get
        Set(ByVal value As String)
            _sjqty = value
        End Set
    End Property

    Private _supco As String
    Public Property Supco() As String
        Get
            Return _supco
        End Get
        Set(ByVal value As String)
            _supco = value
        End Set
    End Property

    Private _toko As String
    Public Property Toko() As String
        Get
            Return _toko
        End Get
        Set(ByVal value As String)
            _toko = value
        End Set
    End Property

    Private _docno As String
    Public Property Docno() As String
        Get
            Return _docno
        End Get
        Set(ByVal value As String)
            _docno = value
        End Set
    End Property

    Private _tglexp As String
    Public Property TglEXP() As String
        Get
            Return _tglexp
        End Get
        Set(ByVal value As String)
            _tglexp = value
        End Set
    End Property
    Private _totalqty As String

    Public Property totalqty() As String
        Get
            Return _totalqty
        End Get
        Set(ByVal value As String)
            _totalqty = value
        End Set
    End Property
    Private _batasLayak As String
    Public Property batasLayak() As String
        Get
            Return _batasLayak
        End Get
        Set(ByVal value As String)
            _batasLayak = value
        End Set
    End Property
    Private _fraction_pcs As String
    Public Property fraction_pcs() As String
        Get
            Return _fraction_pcs
        End Get
        Set(ByVal value As String)
            _fraction_pcs = value
        End Set
    End Property
End Class
