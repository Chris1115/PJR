Public Class ClsSOED
    Private _prdcd As String
    Private _lokasi As String
    Private _counter As String
    Private _expDate As String
    Private _exptDateInput As String
    Private _qtyExpDate As String
    Private _updtimeExp As Date
    Private _deskripsi As String
    Private _feedback As String
    Private _statusBarcode As String
    Private _noPropED As String
    Public Property PRDCD() As String
        Get
            Return _prdcd
        End Get
        Set(value As String)
            _prdcd = value
        End Set
    End Property

    Public Property Counter() As String
        Get
            Return _counter
        End Get
        Set(value As String)
            _counter = value
        End Set
    End Property

    Public Property ExpDate() As String
        Get
            Return _expDate
        End Get
        Set(value As String)
            _expDate = value
        End Set
    End Property

    Public Property ExpDateInput() As String
        Get
            Return _exptDateInput
        End Get
        Set(value As String)
            _exptDateInput = value
        End Set
    End Property

    Public Property QTYExpDate() As String
        Get
            Return _qtyExpDate
        End Get
        Set(value As String)
            _qtyExpDate = value
        End Set
    End Property

    Public Property UpdtimeEXP() As Date
        Get
            Return _updtimeExp
        End Get
        Set(value As Date)
            _updtimeExp = value
        End Set
    End Property

    Public Property Deskripsi() As String
        Get
            Return _deskripsi
        End Get
        Set(value As String)
            _deskripsi = value
        End Set
    End Property

    Public Property Feedback() As String
        Get
            Return _feedback
        End Get
        Set(ByVal value As String)
            _feedback = value
        End Set
    End Property

    Public Property Lokasi() As String
        Get
            Return _lokasi
        End Get
        Set(value As String)
            _lokasi = value
        End Set
    End Property

    Public Property StatusBarcode() As String
        Get
            Return _statusBarcode
        End Get
        Set(value As String)
            _statusBarcode = value
        End Set
    End Property

    Public Property noPropED() As String
        Get
            Return _noPropED
        End Get
        Set(value As String)
            _noPropED = value
        End Set
    End Property
End Class
