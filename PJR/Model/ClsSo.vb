Public Class ClsSo

    'Tambah atrribute NoContainer untuk menyimpan nomor barcode bronjong pada proses BPB
    Private _noContainer As String
    Public Property NoContainer() As String
        Get
            Return _noContainer
        End Get
        Set(ByVal value As String)
            _noContainer = value
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


    Private _rak As String
    Public Property Rak() As String
        Get
            Return _rak
        End Get
        Set(ByVal value As String)
            _rak = value
        End Set
    End Property

    Private _unit As String
    Public Property Unit() As String
        Get
            Return _unit
        End Get
        Set(ByVal value As String)
            _unit = value
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

    Private _qtyToko As String
    Public Property QTYToko() As String
        Get
            Return _qtyToko
        End Get
        Set(ByVal value As String)
            _qtyToko = value
        End Set
    End Property

    Private _qtyGudang As String
    Public Property QTYGudang() As String
        Get
            Return _qtyGudang
        End Get
        Set(ByVal value As String)
            _qtyGudang = value
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

    Private _totalRak As Integer
    Public Property TotalRak() As Integer
        Get
            Return _totalRak
        End Get
        Set(ByVal value As Integer)
            _totalRak = value
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

    Private _qtyInputBaik As String
    Public Property QTYInputBaik() As String
        Get
            Return _qtyInputBaik
        End Get
        Set(ByVal value As String)
            _qtyInputBaik = value
        End Set
    End Property

    Private _qtyInputRusak As String
    Public Property QTYInputRusak() As String
        Get
            Return _qtyInputRusak
        End Get
        Set(ByVal value As String)
            _qtyInputRusak = value
        End Set
    End Property

    Private _indeksList As String
    Public Property indeksList() As String
        Get
            Return _indeksList
        End Get
        Set(ByVal value As String)
            _indeksList = value
        End Set
    End Property

    Private _statusBarcode As String
    Public Property statusBarcode() As String
        Get
            Return _statusBarcode
        End Get
        Set(ByVal value As String)
            _statusBarcode = value
        End Set
    End Property
    Private _qtyTTL1_OLD As String
    Public Property qtyTTL1_OLD() As String
        Get
            Return _qtyTTL1_OLD
        End Get
        Set(ByVal value As String)
            _qtyTTL1_OLD = value
        End Set
    End Property

    Private _qtyReturExpired As String
    Public Property qtyReturExpired() As String
        Get
            Return _qtyReturExpired
        End Get
        Set(value As String)
            _qtyReturExpired = value
        End Set
    End Property

    Private _qtyReturKemasan As String
    Public Property qtyReturKemasan() As String
        Get
            Return _qtyReturKemasan
        End Get
        Set(value As String)
            _qtyReturKemasan = value
        End Set
    End Property

    Private _qtyReturDigigit As String
    Public Property qtyReturDigigit() As String
        Get
            Return _qtyReturDigigit
        End Get
        Set(value As String)
            _qtyReturDigigit = value
        End Set
    End Property

    Private _isBADraft As Boolean
    Public Property isBADraft() As Boolean
        Get
            Return _isBADraft
        End Get
        Set(value As Boolean)
            _isBADraft = value
        End Set
    End Property

    Private _isWtran As Boolean
    Public Property isWtran() As Boolean
        Get
            Return _isWtran
        End Get
        Set(value As Boolean)
            _isWtran = value
        End Set
    End Property

End Class
