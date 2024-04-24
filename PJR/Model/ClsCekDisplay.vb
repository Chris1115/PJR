Public Class ClsCekDisplay

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

    Private _Norak As String
    Public Property Norak() As String
        Get
            Return _Norak
        End Get
        Set(ByVal value As String)
            _Norak = value
        End Set
    End Property

    Private _shelf As String
    Public Property Shelf() As String
        Get
            Return _shelf
        End Get
        Set(ByVal value As String)
            _shelf = value
        End Set
    End Property

    Private _kirikanan As String
    Public Property Kirikanan() As String
        Get
            Return _kirikanan
        End Get
        Set(ByVal value As String)
            _kirikanan = value
        End Set
    End Property

    Private _kodemodis As String
    Public Property KodeModis() As String
        Get
            Return _kodemodis
        End Get
        Set(ByVal value As String)
            _kodemodis = value
        End Set
    End Property

    Private _kapdisp As String
    Public Property Kap_disp() As String
        Get
            Return _kapdisp
        End Get
        Set(ByVal value As String)
            _kapdisp = value
        End Set
    End Property

    Private _qty As String
    Public Property QTY() As String
        Get
            Return _qty
        End Get
        Set(ByVal value As String)
            _qty = value
        End Set
    End Property

    Private _kebutuhanDisp As String
    Public Property KebutuhanDisp() As String
        Get
            Return _kebutuhanDisp
        End Get
        Set(ByVal value As String)
            _kebutuhanDisp = value
        End Set
    End Property

End Class
