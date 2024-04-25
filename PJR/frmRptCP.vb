Imports MySql.Data.MySqlClient
Imports IDM.Fungsi

Public Class frmRptCP
    Public Shared TglAwal As DateTime
    Public Shared TglAkhir As DateTime
    Private Mcon As MySqlConnection
    Private Madp As New MySqlDataAdapter
    Private Mcmd As New MySqlCommand
    Private Mrdr As MySqlDataReader
    Dim Rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument

    Private Sub frmRptCP_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Mcon = ClsConnection.GetConnection.Clone
        Madp = New MySqlDataAdapter("", Mcon)
        Mcmd = New MySqlCommand("", Mcon)
        settingawal()
    End Sub

    Private Sub btnProses_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProses.Click
        If cmbJenisLap.Text.Length = 0 Or cmbJenisLap.SelectedIndex = -1 Then
            MsgBox("Pilih Jenis Laporan Terlebih Dahulu")
            Exit Sub
        End If

        If cmbJenisLap.SelectedIndex = 2 And (txtNik.Text = "" Or txtNik.Text.Length = 0) Then
            MsgBox("NIK harus diisi")
            Exit Sub
        ElseIf cmbJenisLap.SelectedIndex = 3 And (txtNik.Text = "" Or txtNik.Text.Length = 0) Then
            MsgBox("NIK harus diisi")
            Exit Sub
        End If

        Dim dtCP As New DataTable

        If cmbJenisLap.SelectedIndex = 0 Then
            dtCP = New dsBarang.dtPlanogramDataTable
            'If IO.File.Exists(Application.StartupPath & "\rptBarangTidakTerdisplay.rpt") Then
            '    Rpt.Load(Application.StartupPath & "\rptBarangTidakTerdisplay.rpt")
            'Else
            '    MsgBox("file rptBarangTidakTerdisplay.rpt tidak ditemukan", MsgBoxStyle.Critical)
            '    Exit Sub
            'End If
            Rpt = New rptBarangTidakTerdisplay
        ElseIf cmbJenisLap.SelectedIndex = 1 Then
            dtCP = New dsBarang.dtPlanogramDataTable
            'If IO.File.Exists(Application.StartupPath & "\rptTrendBarangTidakTerdisplay.rpt") Then
            '    Rpt.Load(Application.StartupPath & "\rptTrendBarangTidakTerdisplay.rpt")
            'Else
            '    MsgBox("file rptTrendBarangTidakTerdisplay.rpt tidak ditemukan", MsgBoxStyle.Critical)
            '    Exit Sub
            'End If
            Rpt = New rptTrendBarangTidakTerdisplay
        ElseIf cmbJenisLap.SelectedIndex = 2 Then
            dtCP = New dsBarang.dtPeriodeReturDataTable
            Rpt = New rptListingPeriodeRetur
        ElseIf cmbJenisLap.SelectedIndex = 3 Then
            dtCP = New dsBarang.dtRLBTDDataTable
            Rpt = New RptRekapRLBTD
        End If
        'Rpt = New rptPlanogram

        Try
            If Mcon.State <> ConnectionState.Open Then
                Mcon.Open()
            End If

            If cmbJenisLap.SelectedIndex = 0 Then
                Mcmd.CommandText = " Select PLU,Nama as `Desc`, NamaRak,NoShelf,TglScan,KiriKanan, Divisi, NoRak, Stock, JenisBarang from CekPlanogram"
                Mcmd.CommandText &= " where date(tglscan) = '" & Format(dtpTglAwal.Value, "yyyy-MM-dd") & "' and jenisBarang <> '' and `status` <> 'S' "
                If txtNik.Text <> "" Or txtNik.Text.Length <> 0 Then
                    Mcmd.CommandText &= " and NIK = '" & txtNik.Text & "'"
                End If
                Mcmd.CommandText &= " order by noshelf asc "
                Madp = New MySqlDataAdapter(Mcmd)
                Madp.Fill(dtCP)

                For i As Integer = 0 To dtCP.Rows.Count - 1
                    If (dtCP.Rows(i)(8).ToString = "TT" And dtCP.Rows(i)(7).ToString = "0") Then
                        dtCP.Rows(i).Item("JenisBarang") = "SO"
                    End If
                Next
            ElseIf cmbJenisLap.SelectedIndex = 1 Then
                Mcmd.CommandText = " Select PLU,Nama as `Desc`, NamaRak,NoShelf,KiriKanan, NoRak,Divisi,COUNT(PLU) AS Intensitas from CekPlanogram"
                Mcmd.CommandText &= " where date(tglscan) between '" & Format(dtpTglAwal.Value, "yyyy-MM-dd") & "' and '" & Format(dtpTglAkhir.Value, "yyyy-MM-dd") & "' "
                Mcmd.CommandText &= " and stock > 0  and jenisbarang = 'TT' and `status` <> 'S'"
                If txtNik.Text <> "" Or txtNik.Text.Length <> 0 Then
                    Mcmd.CommandText &= " and NIK = '" & txtNik.Text & "'"
                End If
                Mcmd.CommandText &= " group by PLU order by noshelf asc "
                Madp = New MySqlDataAdapter(Mcmd)
                Madp.Fill(dtCP)
            ElseIf cmbJenisLap.SelectedIndex = 2 Then
                Mcmd.CommandText = " select plu, nama as `desc`, cast(concat(MAXBATASRETUR_S,' - ',MAXBATASRETUR) as char) as BatasRetur from cekplanogram"
                Mcmd.CommandText &= " where date(tglscan) = '" & Format(dtpTglAwal.Value, "yyyy-MM-dd") & "' "
                Mcmd.CommandText &= " and NIK = '" & txtNik.Text & "' "
                Mcmd.CommandText &= " and `Status` = 'B' and nama <> '' "
                Madp = New MySqlDataAdapter(Mcmd)
                Madp.Fill(dtCP)
            ElseIf cmbJenisLap.SelectedIndex = 3 Then
                Mcmd.CommandText = "  select distinct Date(tglscan) as tglscan, namarakinput as namarak "
                Mcmd.CommandText &= " from cekplanogram a "
                Mcmd.CommandText &= " where namarakinput <> '' "
                Mcmd.CommandText &= " and date(tglscan) between '" & Format(dtpTglAwal.Value, "yyyy-MM-dd") & "' and '" & Format(dtpTglAkhir.Value, "yyyy-MM-dd") & "' "
                Mcmd.CommandText &= " and nik = '" & txtNik.Text & "'"
                Mcmd.CommandText &= " group by a.tglscan, namarak"
                Mcmd.CommandText &= " order by a.tglscan asc;"
                Madp = New MySqlDataAdapter(Mcmd)
                Madp.Fill(dtCP)

                For a As Integer = 0 To dtCP.Rows.Count - 1
                    Mcmd.CommandText = " select"
                    Mcmd.CommandText &= " ("
                    Mcmd.CommandText &= " select cast(concat(min(noshelf), ' s/d ', max(noshelf)) as char) from cekplanogram where date(tglscan) = '" & Format(dtCP.Rows(a)("tglscan"), "yyyy-MM-dd") & "' and namarakinput = '" & dtCP.Rows(a)("namarak") & "' and nik = '" & txtNik.Text & "'"
                    Mcmd.CommandText &= " )as shelf,"
                    Mcmd.CommandText &= " ("
                    Mcmd.CommandText &= " select count(*)  from cekplanogram where jenisbarang = 'TT' and date(tglscan) = '" & Format(dtCP.Rows(a)("tglscan"), "yyyy-MM-dd") & "' and namarakinput = '" & dtCP.Rows(a)("namarak") & "' and nik = '" & txtNik.Text & "'"
                    Mcmd.CommandText &= " )as tidakterdisplay,"
                    Mcmd.CommandText &= " ("
                    Mcmd.CommandText &= " select count(*) from cekplanogram where jenisbarang = 'SO' and date(tglscan) = '" & Format(dtCP.Rows(a)("tglscan"), "yyyy-MM-dd") & "' and namarakinput = '" & dtCP.Rows(a)("namarak") & "' and nik = '" & txtNik.Text & "'"
                    Mcmd.CommandText &= " )as stockout,"
                    Mcmd.CommandText &= " ("
                    Mcmd.CommandText &= " select count(*) from cekplanogram where jenisbarang = 'SD' and date(tglscan) = '" & Format(dtCP.Rows(a)("tglscan"), "yyyy-MM-dd") & "' and namarakinput = '" & dtCP.Rows(a)("namarak") & "' and nik = '" & txtNik.Text & "'"
                    Mcmd.CommandText &= " )as salahdisplay"
                    Mrdr = Mcmd.ExecuteReader
                    While Mrdr.Read
                        dtCP.Rows(a)("shelf") = Mrdr("shelf")
                        dtCP.Rows(a)("tidakterdisplay") = Mrdr("tidakterdisplay")
                        dtCP.Rows(a)("stockout") = Mrdr("stockout")
                        dtCP.Rows(a)("salahdisplay") = Mrdr("salahdisplay")
                    End While
                    Mrdr.Close()
                Next
            End If

            TraceLog(Mcmd.CommandText)

            If cmbJenisLap.SelectedIndex = 0 Then
                Mcmd.CommandText = "select count(*) from cekplanogram where date(tglscan) = '" & Format(dtpTglAwal.Value, "yyyy-MM-dd") & "'"
            ElseIf cmbJenisLap.SelectedIndex = 1 Then
                Mcmd.CommandText = "select count(*) from cekplanogram where date(tglscan) between '" & Format(dtpTglAwal.Value, "yyyy-MM-dd") & "' and '" & Format(dtpTglAkhir.Value, "yyyy-MM-dd") & "' "
            End If

            If cmbJenisLap.SelectedIndex = 0 Or cmbJenisLap.SelectedIndex = 1 Then
                If Mcmd.ExecuteScalar > 0 And dtCP.Rows.Count = 0 Then
                    MsgBox("Tidak ada barang yang tidak terdisplay!")
                    CRVCP.ReportSource = Nothing
                    CRVCP.Refresh()
                    btnCetak.Enabled = False
                    Exit Sub
                ElseIf Mcmd.ExecuteScalar = 0 Then
                    MsgBox("Tidak Ada Data!")
                    CRVCP.ReportSource = Nothing
                    CRVCP.Refresh()
                    btnCetak.Enabled = False
                    Exit Sub
                End If
            ElseIf cmbJenisLap.SelectedIndex = 2 Then
                If dtCP.Rows.Count = 0 Then
                    MsgBox("Tidak Ada Data!")
                    CRVCP.ReportSource = Nothing
                    CRVCP.Refresh()
                    btnCetak.Enabled = False
                    Exit Sub
                End If
            ElseIf cmbJenisLap.SelectedIndex = 3 Then
                If dtCP.Rows.Count = 0 Then
                    MsgBox("Tidak Ada Data!")
                    CRVCP.ReportSource = Nothing
                    CRVCP.Refresh()
                    btnCetak.Enabled = False
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
            Exit Sub
        Finally
            Mcon.Close()
        End Try

        Rpt.SetDataSource(dtCP)

        If cmbJenisLap.SelectedIndex = 0 Then
            Rpt.SetParameterValue("toko", FormMain.Toko.Kode & "/" & FormMain.Toko.Lokasi)
            Rpt.SetParameterValue("Tgl", DateTime.Parse(dtpTglAwal.Value))
            If txtNik.Text <> "" Or txtNik.Text.Length <> 0 Then
                Rpt.SetParameterValue("nik", txtNik.Text)
            Else
                Rpt.SetParameterValue("nik", "ALL")
            End If
        ElseIf cmbJenisLap.SelectedIndex = 1 Then
            Rpt.SetParameterValue("toko", FormMain.Toko.Kode & "/" & FormMain.Toko.Lokasi)
            Rpt.SetParameterValue("TglAwal", DateTime.Parse(dtpTglAwal.Value))
            Rpt.SetParameterValue("TglAkhir", DateTime.Parse(dtpTglAkhir.Value))
            If txtNik.Text <> "" Or txtNik.Text.Length <> 0 Then
                Rpt.SetParameterValue("nik", txtNik.Text)
            Else
                Rpt.SetParameterValue("nik", "ALL")
            End If
        ElseIf cmbJenisLap.SelectedIndex = 2 Then
            Rpt.SetParameterValue("toko", FormMain.Toko.Kode & "/" & FormMain.Toko.Lokasi)
            Rpt.SetParameterValue("user", txtNik.Text)
        ElseIf cmbJenisLap.SelectedIndex = 3 Then
            Rpt.SetParameterValue("toko", FormMain.Toko.Kode & "/" & FormMain.Toko.Lokasi)
            Rpt.SetParameterValue("user", txtNik.Text)
            Rpt.SetParameterValue("tanggalawal", DateTime.Parse(dtpTglAwal.Value))
            Rpt.SetParameterValue("tanggalakhir", DateTime.Parse(dtpTglAkhir.Value))
        End If

        CRVCP.ReportSource = Rpt
        CRVCP.Zoom(1)

        btnCetak.Enabled = True
    End Sub

    Sub settingawal()
        If cmbJenisLap.SelectedIndex = -1 Then
            lblTglAwal.Visible = False
            lblTglAkhir.Visible = False
            dtpTglAwal.Visible = False
            dtpTglAkhir.Visible = False
        ElseIf cmbJenisLap.SelectedIndex = 0 Then
            dtpTglAwal.Text = Date.Now
        ElseIf cmbJenisLap.SelectedIndex = 1 Then
            dtpTglAwal.Text = Convert.ToDateTime(Month(Now) & "/01" & "/" & Year(Now))
            dtpTglAkhir.Text = Date.Now
        End If

        dtpTglAwal.Format = DateTimePickerFormat.Custom
        dtpTglAwal.CustomFormat = "dd/MM/yyyy"
        dtpTglAkhir.Format = DateTimePickerFormat.Custom
        dtpTglAkhir.CustomFormat = "dd/MM/yyyy"
    End Sub

    Private Sub btnKeluar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKeluar.Click
        Me.Dispose()
    End Sub

    Private Sub cmbJenisLap_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbJenisLap.SelectedIndexChanged
        If cmbJenisLap.SelectedIndex = 0 Then
            lblTglAwal.Text = "Tanggal"
            lblTglAwal.Visible = True
            dtpTglAwal.Visible = True
            lblTglAkhir.Visible = False
            dtpTglAkhir.Visible = False
        ElseIf cmbJenisLap.SelectedIndex = 1 Then
            lblTglAwal.Text = "Tanggal Awal"
            lblTglAkhir.Visible = True
            dtpTglAkhir.Visible = True
            lblTglAwal.Visible = True
            dtpTglAwal.Visible = True
        ElseIf cmbJenisLap.SelectedIndex = 2 Then
            lblTglAwal.Text = "Tanggal"
            lblTglAwal.Visible = True
            dtpTglAwal.Visible = True
            lblTglAkhir.Visible = False
            dtpTglAkhir.Visible = False
        ElseIf cmbJenisLap.SelectedIndex = 3 Then
            lblTglAwal.Text = "Tanggal Awal"
            lblTglAkhir.Visible = True
            dtpTglAkhir.Visible = True
            lblTglAwal.Visible = True
            dtpTglAwal.Visible = True
        End If
        lblUser.Visible = True
        txtNik.Visible = True
    End Sub

    Private Sub btnCetak_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCetak.Click
        MsgBox("Pastikan menggunakan printer besar!")
        Rpt.PrintToPrinter(1, False, 0, 0)
        MsgBox("Cetak Selesai!")
    End Sub

End Class