﻿Public Class ClsPriceTag

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
    Private _keterangan As String

    Public Property Keterangan() As String
        Get
            Return _keterangan
        End Get
        Set(ByVal value As String)
            _keterangan = value
        End Set
    End Property
End Class
