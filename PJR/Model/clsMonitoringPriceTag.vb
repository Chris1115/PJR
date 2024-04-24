Public Class ClsMonitoringPriceTag

    Private _barcode As String
    Private _keterangan As String
    Private _setMenu As String

    Public Property barcode() As String
        Get
            Return _barcode
        End Get
        Set(value As String)
            _barcode = value
        End Set
    End Property

    Public Property keterangan() As String
        Get
            Return _keterangan
        End Get
        Set(value As String)
            _keterangan = value
        End Set
    End Property

    Property setMenu() As String
        Get
            Return _setMenu
        End Get
        Set(value As String)
            _setMenu = value
        End Set
    End Property

End Class
