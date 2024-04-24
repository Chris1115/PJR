Imports CrystalDecisions.Shared
Imports MySql.Data.MySqlClient
Imports PJR.clsFinger
Imports PJR.ClsPJRController

Public Class frmApprovePJR
    Private Sub btnCari_Click(sender As Object, e As EventArgs) Handles btnCari.Click
        Try
            MsgBox("HARAP SIAPKAN PRINTER BESAR", MsgBoxStyle.Exclamation)
            Dim tanggal As String = ""
            Dim tanggalawal As String = ""
            Dim tanggalakhir As String = ""
            Dim NikToko As String = ""
            Dim hari As String = ""
            Dim hx As String = ""
            'tanggal = tglAwalDateTimePicker.Value.ToString("yyyy-MM-dd")
            hari = cbboxHari.Text
            If hari.ToLower.Trim = "senin" Then
                hx = "h1"
            ElseIf hari.ToLower.Trim = "selasa" Then
                hx = "h2"

            ElseIf hari.ToLower.Trim = "rabu" Then
                hx = "h3"

            ElseIf hari.ToLower.Trim = "kamis" Then
                hx = "h4"

            ElseIf hari.ToLower.Trim = "jumat" Then
                hx = "h5"

            ElseIf hari.ToLower.Trim = "sabtu" Then
                hx = "h6"

            End If

            Dim cPJR As New ClsPJRController
            Dim dt As New DataTable
            Dim Rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Rpt = New rptJadwalPJRWaktuAprove
            Rpt.PrintOptions.PaperOrientation = PaperOrientation.Landscape
            Rpt.PrintOptions.PaperSize = PaperSize.PaperA4
            'dt = cPJR.cariJadwal(tanggalawal, tanggalakhir)

            dt = cPJR.cariJadwalHari(hari)
            'Console.WriteLine(dt.Rows.Count)
            If dt.Rows.Count <> 0 Then
                Try


                    NikToko = cPJR.getConstNIKPJR
                    Rpt.SetDataSource(dt)

                    Rpt.SetParameterValue("toko", FormMain.Toko.Kode & "/" & FormMain.Toko.Lokasi)
                    Rpt.SetParameterValue("user", "")
                    Rpt.SetParameterValue("HX", hx)
                    Rpt.SetParameterValue("HARI", hari)

                    CrystalReportViewer1.ReportSource = Rpt
                Catch ex As Exception

                End Try
                btnApprove.Enabled = True
                btnTolak.Enabled = True
            Else
                MessageBox.Show("Maaf, Jadwal PJR Hari " & hari & " sudah selesai di Approve !", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)

            End If

        Catch ex As Exception
            MsgBox("MAAF, HARAP GUNAKAN PRINTER BESAR", MsgBoxStyle.Critical)
            IDM.Fungsi.TraceLog("Error, BTN_Cari_Aprrove : " & ex.Message & ex.StackTrace)
        End Try
    End Sub

    Private Sub frmApprovePJR_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim cPJR As New ClsPJRController
        'create table untuk menghitung timesecond rak
        cPJR.getJadwal_menit()

    End Sub

    Private Sub btnApprove_Click(sender As Object, e As EventArgs) Handles btnApprove.Click
        Dim Result As DialogResult
        Dim cPJR As New ClsPJRController
        Dim tanggal As String


        Result = MessageBox.Show("Apakah Anda yakin akan Approve Jadwal PJR?", "Approval PJR Selesai..", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If Result = Windows.Forms.DialogResult.Yes Then
            'Revisi PRPK 848/06-23/E/PMO 'Kukuh
            Dim tgl As String = tglAwalDateTimePicker.Value.ToString("dd")
            Dim tes As String() = {"a", "b", "c"}

            MsgBox("Proses Approval PJR menggunakan Scan FINGER", MsgBoxStyle.Information)

            'Revisi PRPK 848/06-23/E/PMO 'Kukuh
            If (tgl >= 1 And tgl <= 5) Or (tgl >= 16 And tgl <= 20) Then
                tes = Panggil_CekFingerprintV3("frmApprovePJR", "WDCP")
            Else
                tes = Panggil_CekFingerprintV3("frmApprovePJR", "WDCP_PJR 2")
            End If

            'If cPJR.scanFinger("WDCP") = True Then
            If tes(0) = "" And tes(1) = "" And tes(2) = "" Then
                Debug.WriteLine("Password Salah")
            Else
                ConstNIKPJR(tes(2).Split("|")(0).Trim)

                tanggal = tglAwalDateTimePicker.Value.ToString("yyyy-MM-dd")
                'cPJR.approvePJR("Y", tanggal)
                cPJR.approvePJR("Y", cbboxHari.Text)

                MsgBox("Berhasil Approve")
                CrystalReportViewer1.ReportSource = Nothing
                CrystalReportViewer1.Refresh()
            End If


        Else
            MsgBox("Maaf, Validasi scanfinger tidak berhasil !", MsgBoxStyle.Exclamation)

        End If
    End Sub

    Private Sub btnTolak_Click(sender As Object, e As EventArgs) Handles btnTolak.Click
        Dim Result As DialogResult
        Dim cPJR As New ClsPJRController
        Dim tanggal As String

        Result = MessageBox.Show("Apakah Anda yakin akan Menolak Jadwal PJR?", "Approval PJR Selesai..", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If Result = Windows.Forms.DialogResult.Yes Then
            Dim tes As String() = {"a", "b", "c"}

            MsgBox("Proses Approval PJR menggunakan Scan FINGER", MsgBoxStyle.Information)
            tes = Panggil_CekFingerprintV3("frmApprovePJR", "WDCP")

            'If cPJR.scanFinger("WDCP") = True Then
            If tes(0) = "" And tes(1) = "" And tes(2) = "" Then
                Debug.WriteLine("Password Salah")
            Else
                ConstNIKPJR(tes(2).Split("|")(0).Trim)

                tanggal = tglAwalDateTimePicker.Value.ToString("yyyy-MM-dd")

                'cPJR.approvePJR("N", tanggal)
                cPJR.approvePJR("N", cbboxHari.Text)

                MsgBox("Berhasil Tolak")
                CrystalReportViewer1.ReportSource = Nothing
                CrystalReportViewer1.Refresh()
            End If
        End If
    End Sub
End Class