Imports MySql.Data.MySqlClient

Public Class frmTampilLBTD
    Public Shared TglAwal As DateTime
    Public Shared TglAkhir As DateTime
    Private Mcon As MySqlConnection
    Private Madp As New MySqlDataAdapter
    Private Mcmd As New MySqlCommand
    Private Mrdr As MySqlDataReader
    Dim Rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
    Public Shared jenis As String = ""
    Public Shared nik_as As String = ""
    Private Sub frmTampilLBTD_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If jenis = "" Or jenis.ToUpper = "CEKPJR" Then

            Dim cpjr As New ClsPJRController

            Mcon = ClsConnection.GetConnection.Clone
            Madp = New MySqlDataAdapter("", Mcon)
            Mcmd = New MySqlCommand("", Mcon)
            Rpt = New rptBarangTidakTerpajang
            Dim NikToko As String = ""
            NikToko = cpjr.getConstNIKPJR

            Dim rakpjr As String = ""
            rakpjr = cpjr.getConstRakPJR
            Try
                If Mcon.State <> ConnectionState.Open Then
                    Mcon.Open()
                End If
                Mcmd.CommandText = " Select PLU,Nama as `Desc`, NamaRak,NoShelf,TglScan,KiriKanan, Divisi, NoRak, Stock, JenisBarang from cekpjr"
                Mcmd.CommandText &= " where date(tglscan) = CURDATE() and jenisBarang = 'TT'"
                Mcmd.CommandText &= " and NIK = '" & NikToko & "' "
                Mcmd.CommandText &= " AND NAMARAK = '" & rakpjr.Split("-")(0).Trim & "'"
                Mcmd.CommandText &= " AND norak = '" & rakpjr.Split("-")(2).Trim & "'"
                Mcmd.CommandText &= " order by noshelf asc "
                Console.WriteLine(Mcmd.CommandText)
                Dim DtCp As New DataTable
                DtCp.Clear()
                Madp = New MySqlDataAdapter(Mcmd)
                Madp.Fill(DtCp)

                Rpt.SetDataSource(DtCp)

                Rpt.SetParameterValue("toko", FormMain.Toko.Kode & "/" & FormMain.Toko.Lokasi)
                Rpt.SetParameterValue("Tgl", Now())
                Rpt.SetParameterValue("nik", NikToko)


                CrystalReportViewer1.ReportSource = Rpt
                CrystalReportViewer1.Zoom(1)


            Catch ex As Exception
                MsgBox(ex.Message & ex.StackTrace)
            End Try

        Else
            Dim cpjr As New ClsPJRController
            Mcon = ClsConnection.GetConnection.Clone
            Madp = New MySqlDataAdapter("", Mcon)
            Mcmd = New MySqlCommand("", Mcon)
            Rpt = New rptListItemBAPJR_AS

            Try
                If Mcon.State <> ConnectionState.Open Then
                    Mcon.Open()
                End If
                Mcmd.CommandText = " SELECT PRDCD AS PLU,`DESC`,HARI,NORAK, NOSHELF AS SHELFING,KODE_MODIS AS KODEMODIS, NIK FROM ITEMSO_PJR_BA_AS
                                     WHERE RECID = ''"
                Dim DtCp As New DataTable
                DtCp.Clear()
                Madp = New MySqlDataAdapter(Mcmd)
                Madp.Fill(DtCp)

                Rpt.SetDataSource(DtCp)

                Rpt.SetParameterValue("toko", FormMain.Toko.Kode & "/" & FormMain.Toko.Lokasi)

                CrystalReportViewer1.ReportSource = Rpt
                CrystalReportViewer1.Zoom(1)


            Catch ex As Exception
                MsgBox(ex.Message & ex.StackTrace)
            End Try

        End If

    End Sub
End Class