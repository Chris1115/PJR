Imports MySql.Data.MySqlClient

Public Class ClsTokoController
    Private utility As New Utility

    ''' <summary>
    ''' untuk mendapatkan info toko
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetInfoToko() As ClsToko
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim tmpToko As New ClsToko
        Dim mcom As New MySqlCommand("", conn)

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            mcom.CommandText = "Select `DESC` from const where RKEY='TOK';"
            tmpToko.Kode = mcom.ExecuteScalar & ""

            mcom.CommandText = "Select `DESC` from const where RKEY='BPB';"
            tmpToko.Cabang = mcom.ExecuteScalar & ""

            mcom.CommandText = "Select `DESC` from const where RKEY='" & tmpToko.Cabang.Substring(1) & "';"
            tmpToko.CabangName = mcom.ExecuteScalar & ""

            mcom.CommandText = "Select `DESC` from const where RKEY='CON';"
            tmpToko.Lokasi = mcom.ExecuteScalar & ""

            mcom.CommandText = "Select `DESC` from const where RKEY='CPT';"
            tmpToko.PtName = mcom.ExecuteScalar & ""

            Dim Odar As Object
            mcom.CommandText = "Select Nama,ALMT,KOTA,NPWP,SKP,TELP_1 from TOKO WHERE KDTK='" & tmpToko.Kode & "';"
            Odar = mcom.ExecuteReader()
            While Odar.Read()
                tmpToko.Nama = Odar("Nama") & ""
                tmpToko.Alamat = Odar("ALMT") & ""
                tmpToko.Kota = Odar("KOTA") & ""
                tmpToko.NPWP = Odar("SKP") & ""
                tmpToko.Telephone = Odar("TELP_1") & ""
            End While

            If tmpToko.Kode = "" Or tmpToko.Lokasi = "" Or tmpToko.PtName = "" Or tmpToko.Nama = "" _
               Or tmpToko.Alamat = "" Or tmpToko.Kota = "" Or tmpToko.NPWP = "" Then
                utility.Tracelog("Error", "Kesalahan Terjadi Saat Pembacaan Informasi Toko.", "GetInfoToko", conn)
                Err.Raise(-1)
            End If

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetInfoToko", conn)
        Finally
            conn.Close()
        End Try

        Return tmpToko
    End Function

End Class
