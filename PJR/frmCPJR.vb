Imports MySql.Data.MySqlClient
Imports IDM.Fungsi
Public Class frmCPJR
    Private Mcon As MySqlConnection
    Private Madp As New MySqlDataAdapter
    Private Mcom As New MySqlCommand
    Private Mrdr As MySqlDataReader
    Dim Rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
    Private Sub frmCPJR_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Mcon = ClsConnection.GetConnection.Clone
        Madp = New MySqlDataAdapter("", Mcon)
        Mcom = New MySqlCommand("", Mcon)

    End Sub

    Private Sub btnProses_Click(sender As Object, e As EventArgs) Handles btnProses.Click
        If cmbJenisLap.Text.Length = 0 Or cmbJenisLap.SelectedIndex = -1 Then
            MsgBox("Pilih Jenis Laporan Terlebih Dahulu")
            Exit Sub
        End If
        Dim dtCP As New DataTable
        Dim tanggal As String = ""
        Dim tanggalawal As String = ""
        Dim tanggalakhir As String = ""
        Dim cPJR As New ClsPJRController
        Dim sqltampung As String = ""
        Dim NikToko As String = ""
        NikToko = cPJR.getConstNIKPJR
        If cmbJenisLap.SelectedIndex = 0 Then
            'dtCP = New dsBarang.dtPlanogramDataTable

            Rpt = New rptLaporanJadwalPJR
        ElseIf cmbJenisLap.SelectedIndex = 1 Then
            'dtCP = New dsBarang.dtPlanogramDataTable

            Rpt = New rptJadwalPJR
        ElseIf cmbJenisLap.SelectedIndex = 2 Then
            'dtCP = New dsBarang.dtPlanogramDataTable

            Rpt = New rptJadwalPJRWaktu
        ElseIf cmbJenisLap.SelectedIndex = 3 Then
            'dtCP = New dsBarang.dtPlanogramDataTable

            Rpt = New rptLaporanJadwalPJRperRak
        ElseIf cmbJenisLap.SelectedIndex = 4 Then
            'dtCP = New dsBarang.dtPlanogramDataTable

            Rpt = New rptLaporanJadwalPJRperKary
        ElseIf cmbJenisLap.SelectedIndex = 5 Then
            'dtCP = New dsBarang.dtPlanogramDataTable

            Rpt = New rptLaporanJadwalPJRperTanggal

        ElseIf cmbJenisLap.SelectedIndex = 6 Then
            Rpt = New rptFinalBarangTidakTerpajang

        End If

        Try
            If Mcon.State <> ConnectionState.Open Then
                Mcon.Open()
            End If

            If cmbJenisLap.SelectedIndex <> 3 Then
                tanggal = dtpTglAwal.Value.ToString("yyyy-MM-dd")
            End If
            'Laporan Jadwal Penanggung Jawab Rak
            If cmbJenisLap.SelectedIndex = 0 Then
                'Mcom.CommandText = "SELECT DAYNAME('" & tanggal & "')"
                'If Mcom.ExecuteScalar <> "Sunday" Then
                Mcom.CommandText = "SELECT a.NIK, NAMA,JABATAN, 
                                    IF(a.hari='Senin',SUBSTRING(a.kode_modis,1,3),'') AS Modis1,
                                    IF(a.hari='Senin',a.shelfing,'') AS Shelf1,
                                    IF(a.hari='Selasa',SUBSTRING(a.kode_modis,1,3),'') AS Modis2,
                                    IF(a.hari='Selasa',a.shelfing,'') AS Shelf2,
                                    IF(a.hari='Rabu',SUBSTRING(a.kode_modis,1,3),'') AS Modis3,
                                    IF(a.hari='Rabu',a.shelfing,'') AS Shelf3,
                                    IF(a.hari='Kamis',SUBSTRING(a.kode_modis,1,3),'') AS Modis4,
                                    IF(a.hari='Kamis',a.shelfing,'') AS Shelf4,
                                    IF(a.hari='Jumat',SUBSTRING(a.kode_modis,1,3),'') AS Modis5,
                                    IF(a.hari='Jumat',a.shelfing,'') AS Shelf5,
                                    IF(a.hari='Sabtu',SUBSTRING(a.kode_modis,1,3),'') AS Modis6,
                                    IF(a.hari='Sabtu',a.shelfing,'') AS Shelf6,
                                    IF(a.hari='Minggu',SUBSTRING(a.kode_modis,1,3),'') AS Modis7,
                                    IF(a.hari='Minggu',a.shelfing,'') AS Shelf7
                                    FROM temp_jadwal_pjr a 
                                    WHERE StatusApproval = 'Y' AND tanggal
                                     >= (SELECT DATE_ADD(CAST('" & tanggal & "' AS DATE), INTERVAL(1-DAYOFWEEK(CAST('" & tanggal & "' AS DATE))) DAY))

                                    AND tanggal
                                     <= (SELECT DATE_ADD('" & tanggal & "', INTERVAL(7-DAYOFWEEK('" & tanggal & "')) DAY))"
                'Else
                '    Mcom.CommandText = "SELECT a.NIK, NAMA,JABATAN, 
                '                    IF(a.hari='Senin',SUBSTRING(a.kode_modis,1,3),'') AS Modis1,
                '                    IF(a.hari='Senin',a.shelfing,'') AS Shelf1,
                '                    IF(a.hari='Selasa',SUBSTRING(a.kode_modis,1,3),'') AS Modis2,
                '                    IF(a.hari='Selasa',a.shelfing,'') AS Shelf2,
                '                    IF(a.hari='Rabu',SUBSTRING(a.kode_modis,1,3),'') AS Modis3,
                '                    IF(a.hari='Rabu',a.shelfing,'') AS Shelf3,
                '                    IF(a.hari='Kamis',SUBSTRING(a.kode_modis,1,3),'') AS Modis4,
                '                    IF(a.hari='Kamis',a.shelfing,'') AS Shelf4,
                '                    IF(a.hari='Jumat',SUBSTRING(a.kode_modis,1,3),'') AS Modis5,
                '                    IF(a.hari='Jumat',a.shelfing,'') AS Shelf5,
                '                    IF(a.hari='Sabtu',SUBSTRING(a.kode_modis,1,3),'') AS Modis6,
                '                    IF(a.hari='Sabtu',a.shelfing,'') AS Shelf6,
                '                    IF(a.hari='Minggu',SUBSTRING(a.kode_modis,1,3),'') AS Modis7,
                '                    IF(a.hari='Minggu',a.shelfing,'') AS Shelf7
                '                    FROM temp_jadwal_pjr a 
                '                    WHERE CAST(CONCAT(SUBSTRING(TANGGAL,7,4),'-',SUBSTRING(TANGGAL,4,2),'-',SUBSTRING(TANGGAL,1,2)) AS DATE)
                '                     >= (SELECT DATE_ADD(CAST('" & tanggal & "' AS DATE), INTERVAL(2-DAYOFWEEK(CAST('" & tanggal & "' AS DATE))-7) DAY))

                '                    AND CAST(CONCAT(SUBSTRING(TANGGAL,7,4),'-',SUBSTRING(TANGGAL,4,2),'-',SUBSTRING(TANGGAL,1,2)) AS DATE)
                '                     <= (SELECT DATE_ADD('" & tanggal & "', INTERVAL(8-DAYOFWEEK('" & tanggal & "')-7) DAY))"
                'End If


                TraceLog("WDCP_CetakLaporan_0 : " & Mcom.CommandText & "")
                Madp.SelectCommand.CommandText = Mcom.CommandText
                Console.WriteLine(Madp.SelectCommand.CommandText)
                dtCP.Clear()
                Madp.Fill(dtCP)

                tanggalawal = cPJR.getTanggal(tanggal, "awal")
                tanggalakhir = cPJR.getTanggal(tanggal, "akhir")
                Console.WriteLine(tanggalawal)
                Console.WriteLine(tanggalakhir)

                Rpt.SetDataSource(dtCP)

                Rpt.SetParameterValue("toko", FormMain.Toko.Kode & "/" & FormMain.Toko.Lokasi)
                Rpt.SetParameterValue("tanggalawal", tanggalawal)
                Rpt.SetParameterValue("tanggalakhir", tanggalakhir)
                Rpt.SetParameterValue("user", NikToko)

                CrystalReportViewer1.ReportSource = Rpt
                CrystalReportViewer1.Zoom(1)

                btnCetak.Enabled = True


                'Jadwal Penangngung Jawab Rak
            ElseIf cmbJenisLap.SelectedIndex = 1 Then
                'Mcom.CommandText = "SELECT DAYNAME('" & tanggal & "')"
                'If Mcom.ExecuteScalar <> "Sunday" Then
                Mcom.CommandText = "SELECT SUBSTRING(kode_modis,1,3) AS NamaRak,shelfing AS Shelf, totalestimasi AS Estimasi,nik AS NIK,nama AS Nama,jabatan AS Jabatan,hari AS Hari FROM temp_jadwal_pjr
                                    WHERE  StatusApproval = 'Y' AND tanggal
                                     >= (SELECT DATE_ADD(CAST('" & tanggal & "' AS DATE), INTERVAL(1-DAYOFWEEK(CAST('" & tanggal & "' AS DATE))) DAY))

                                    AND tanggal
                                     <= (SELECT DATE_ADD('" & tanggal & "', INTERVAL(7-DAYOFWEEK('" & tanggal & "')) DAY))
                ORDER BY nik,FIELD(hari, 'Senin','Selasa','Rabu','Kamis','Jumat','Sabtu','Minggu')"
                'Else
                '    Mcom.CommandText = "SELECT SUBSTRING(kode_modis,1,3) AS NamaRak,shelfing AS Shelf, totalestimasi AS Estimasi,nik AS NIK,nama AS Nama,jabatan AS Jabatan,hari AS Hari FROM temp_jadwal_pjr
                '                    WHERE CAST(CONCAT(SUBSTRING(TANGGAL,7,4),'-',SUBSTRING(TANGGAL,4,2),'-',SUBSTRING(TANGGAL,1,2)) AS DATE)
                '                     >= (SELECT DATE_ADD(CAST('" & tanggal & "' AS DATE), INTERVAL(2-DAYOFWEEK(CAST('" & tanggal & "' AS DATE))-7) DAY))

                '                    AND CAST(CONCAT(SUBSTRING(TANGGAL,7,4),'-',SUBSTRING(TANGGAL,4,2),'-',SUBSTRING(TANGGAL,1,2)) AS DATE)
                '                     <= (SELECT DATE_ADD('" & tanggal & "', INTERVAL(8-DAYOFWEEK('" & tanggal & "')-7) DAY))
                'ORDER BY nik,FIELD(hari, 'Senin','Selasa','Rabu','Kamis','Jumat','Sabtu','Minggu')"
                'End If

                TraceLog("WDCP_CetakLaporan_1 : " & Mcom.CommandText & "")

                Madp.SelectCommand.CommandText = Mcom.CommandText
                Console.WriteLine(Madp.SelectCommand.CommandText)
                dtCP.Clear()
                Madp.Fill(dtCP)
                tanggalawal = cPJR.getTanggal(tanggal, "awal")
                tanggalakhir = cPJR.getTanggal(tanggal, "akhir")
                Console.WriteLine(tanggalawal)
                Console.WriteLine(tanggalakhir)

                Rpt.SetDataSource(dtCP)

                Rpt.SetParameterValue("toko", FormMain.Toko.Kode & "/" & FormMain.Toko.Lokasi)
                Rpt.SetParameterValue("tanggalawal", tanggalawal)
                Rpt.SetParameterValue("tanggalakhir", tanggalakhir)
                Rpt.SetParameterValue("user", NikToko)

                CrystalReportViewer1.ReportSource = Rpt
                CrystalReportViewer1.Zoom(1)

                btnCetak.Enabled = True
                'Jadwal Penanggung Jawab Rak dengan ESTIMASI WAKTU
            ElseIf cmbJenisLap.SelectedIndex = 2 Then
                Madp.SelectCommand.CommandText = "SELECT DISTINCT NIK FROM TEMP_JADWAL_PJR where StatusApproval = 'Y' AND  tanggal
                                     >= (SELECT DATE_ADD('" & tanggal & "', INTERVAL(1-DAYOFWEEK('" & tanggal & "')) DAY)) 
                                      AND
                                    tanggal
                                     <= (SELECT DATE_ADD('" & tanggal & "', INTERVAL(7-DAYOFWEEK('" & tanggal & "')) DAY)) "
                dtCP.Clear()
                Madp.Fill(dtCP)
                'buat tabel temp_jadwal_pjr_estimasi_ kosong
                For k As Integer = 1 To 7
                    Mcom.CommandText = "DROP TABLE IF EXISTS temp_jadwal_pjr_estimasi_" & k & ""
                    Mcom.ExecuteNonQuery()
                    Mcom.CommandText = "CREATE TABLE temp_jadwal_pjr_estimasi_" & k & "  
                                    SELECT nik,nama,hari,kode_modis,SUM(totalestimasi) AS h" & k & " FROM TEMP_JADWAL_PJR LIMIT 0"
                    Mcom.ExecuteNonQuery()
                Next
                'ambil estimasi pernik, per hari (h1 - h7)
                For j As Integer = 0 To dtCP.Rows.Count - 1

                    For i As Integer = 1 To 7
                        'Mcom.CommandText = "DROP TABLE IF EXISTS temp_jadwal_pjr_estimasi_" & j & "_" & i & ""
                        'Mcom.ExecuteNonQuery()

                        Mcom.CommandText = "INSERT INTO temp_jadwal_pjr_estimasi_" & i & " 
                                    SELECT nik,nama,hari,kode_modis,SUM(totalestimasi) as h" & i & " FROM TEMP_JADWAL_PJR "
                        If i = 1 Then
                            Mcom.CommandText &= "WHERE HARI = 'Senin' "
                        ElseIf i = 2 Then
                            Mcom.CommandText &= "WHERE HARI = 'Selasa' "
                        ElseIf i = 3 Then
                            Mcom.CommandText &= "WHERE HARI = 'Rabu' "
                        ElseIf i = 4 Then
                            Mcom.CommandText &= "WHERE HARI = 'Kamis' "
                        ElseIf i = 5 Then
                            Mcom.CommandText &= "WHERE HARI = 'Jumat' "
                        ElseIf i = 6 Then
                            Mcom.CommandText &= "WHERE HARI = 'Sabtu' "
                        ElseIf i = 7 Then
                            Mcom.CommandText &= "WHERE HARI = 'Minggu' "
                        End If
                        'If i <> 7 Then
                        Mcom.CommandText &= " AND StatusApproval = 'Y' AND NIK = '" & dtCP.Rows(j)("nik") & "' AND 
                                     tanggal
                                     >= (SELECT DATE_ADD('" & tanggal & "', INTERVAL(1-DAYOFWEEK('" & tanggal & "')) DAY)) 
                                      AND
                                    tanggal
                                     <= (SELECT DATE_ADD('" & tanggal & "', INTERVAL(7-DAYOFWEEK('" & tanggal & "')) DAY)) 
                                      "
                        'Else
                        '    Mcom.CommandText &= " AND NIK = " & dtCP.Rows(j)("nik") & " AND 
                        '             (CAST(CONCAT(SUBSTRING(TANGGAL,7,4),'-',SUBSTRING(TANGGAL,4,2),'-',SUBSTRING(TANGGAL,1,2)) AS DATE) 
                        '             >= (SELECT DATE_ADD('" & tanggal & "', INTERVAL(2-DAYOFWEEK('" & tanggal & "')-7) DAY)) 
                        '              AND
                        '            CAST(CONCAT(SUBSTRING(TANGGAL,7,4),'-',SUBSTRING(TANGGAL,4,2),'-',SUBSTRING(TANGGAL,1,2)) AS DATE) 
                        '             <= (SELECT DATE_ADD('" & tanggal & "', INTERVAL(8-DAYOFWEEK('" & tanggal & "')-7) DAY))) 
                        '              "
                        'End If
                        TraceLog("WDCP_CetakLaporan_2_" & j & "_" & i & " : " & Mcom.CommandText & "")

                        Console.WriteLine(Mcom.CommandText)

                        Mcom.ExecuteNonQuery()
#Region "bck"
                        'Mcom.CommandText = "ALTER TABLE `pos`.temp_jadwal_pjr_estimasi_" & j & "_" & i & " CHANGE `h1` `h1` VARCHAR(10) NULL;"
                        'Mcom.ExecuteNonQuery()
                        'Mcom.CommandText = "SELECT COUNT(*) FROM temp_jadwal_pjr_estimasi_" & j & " "
                        'Console.WriteLine(Mcom.CommandText)

                        'If Mcom.ExecuteScalar = 0 Then
                        '    Mcom.CommandText = "INSERT INTO temp_jadwal_pjr_estimasi_" & j & "_" & i & " 
                        '                SELECT nik,nama,"
                        '    If i = 1 Then
                        '        Mcom.CommandText &= "'Senin' "
                        '    ElseIf i = 2 Then
                        '        Mcom.CommandText &= "'Selasa' "
                        '    ElseIf i = 3 Then
                        '        Mcom.CommandText &= "'Rabu' "
                        '    ElseIf i = 4 Then
                        '        Mcom.CommandText &= "'Kamis' "
                        '    ElseIf i = 5 Then
                        '        Mcom.CommandText &= "'Jumat' "
                        '    ElseIf i = 6 Then
                        '        Mcom.CommandText &= "'Sabtu' "
                        '    ElseIf i = 7 Then
                        '        Mcom.CommandText &= "'Minggu' "
                        '    End If

                        '    Mcom.CommandText &= ",'',0 as h" & i & " FROM TEMP_JADWAL_PJR WHERE  NIK = '" & dtCP.Rows(j)("nik") & "'  AND RECID = '1' GROUP BY NIK"
                        '    Console.WriteLine(Mcom.CommandText)

                        '    Mcom.ExecuteNonQuery()

                        '    Mcom.CommandText = "INSERT INTO temp_jadwal_pjr_estimasi_" & i & " 
                        '                SELECT nik,nama,"
                        '    If i = 1 Then
                        '        Mcom.CommandText &= "'Senin' "
                        '    ElseIf i = 2 Then
                        '        Mcom.CommandText &= "'Selasa' "
                        '    ElseIf i = 3 Then
                        '        Mcom.CommandText &= "'Rabu' "
                        '    ElseIf i = 4 Then
                        '        Mcom.CommandText &= "'Kamis' "
                        '    ElseIf i = 5 Then
                        '        Mcom.CommandText &= "'Jumat' "
                        '    ElseIf i = 6 Then
                        '        Mcom.CommandText &= "'Sabtu' "
                        '    ElseIf i = 7 Then
                        '        Mcom.CommandText &= "'Minggu' "
                        '    End If

                        '    Mcom.CommandText &= ",'',0 as h" & i & " FROM TEMP_JADWAL_PJR WHERE  NIK = '" & dtCP.Rows(j)("nik") & "'  AND RECID = '1' GROUP BY NIK"
                        '    Mcom.ExecuteNonQuery()
                        'Else
                        '    Mcom.CommandText = "INSERT INTO temp_jadwal_pjr_estimasi_" & i & " 
                        '                SELECT nik,nama,"
                        '    If i = 1 Then
                        '        Mcom.CommandText &= "'Senin' "
                        '    ElseIf i = 2 Then
                        '        Mcom.CommandText &= "'Selasa' "
                        '    ElseIf i = 3 Then
                        '        Mcom.CommandText &= "'Rabu' "
                        '    ElseIf i = 4 Then
                        '        Mcom.CommandText &= "'Kamis' "
                        '    ElseIf i = 5 Then
                        '        Mcom.CommandText &= "'Jumat' "
                        '    ElseIf i = 6 Then
                        '        Mcom.CommandText &= "'Sabtu' "
                        '    ElseIf i = 7 Then
                        '        Mcom.CommandText &= "'Minggu' "
                        '    End If

                        '    Mcom.CommandText &= ",'',0 as h" & i & " FROM TEMP_JADWAL_PJR WHERE  NIK = '" & dtCP.Rows(j)("nik") & "'  AND RECID = '1' GROUP BY NIK"
                        '    Mcom.ExecuteNonQuery()

                        'End If

                        'sqltampung &= "LEFT JOIN temp_jadwal_pjr_estimasi_" & j & "_" & i & " a" & j & "_" & i & " ON a.nik = a" & j & "_" & i & ".nik "
#End Region

                    Next


                Next
                'buat tampungan variabel temp_jadwal_pjr_estimasi_1 - 7
                For k As Integer = 1 To 7
                    sqltampung &= " LEFT JOIN temp_jadwal_pjr_estimasi_" & k & " a_" & k & " ON a.nik = a_" & k & ".nik "
                Next

                Console.WriteLine(sqltampung)
                'gabung
                Madp.SelectCommand.CommandText = "SELECT DISTINCT a.NIK,a.nama,a.jabatan,h1,h2,h3,h4,h5,h6,h7 FROM TEMP_JADWAL_PJR a "
                Madp.SelectCommand.CommandText &= sqltampung & "
                                     where StatusApproval = 'Y' AND  tanggal
                                     >= (SELECT DATE_ADD('" & tanggal & "', INTERVAL(1-DAYOFWEEK('" & tanggal & "')) DAY)) 
                                      AND
                                    tanggal
                                     <= (SELECT DATE_ADD('" & tanggal & "', INTERVAL(7-DAYOFWEEK('" & tanggal & "')) DAY)) 
                                              GROUP BY NIK"
                TraceLog("WDCP_CetakLaporan_2 : " & Madp.SelectCommand.CommandText & "")

                Console.WriteLine(Madp.SelectCommand.CommandText)
                dtCP.Clear()
                Madp.Fill(dtCP)
                Rpt.SetDataSource(dtCP)
                tanggalawal = cPJR.getTanggal(tanggal, "awal")
                tanggalakhir = cPJR.getTanggal(tanggal, "akhir")
                Rpt.SetParameterValue("toko", FormMain.Toko.Kode & "/" & FormMain.Toko.Lokasi)
                Rpt.SetParameterValue("tanggalawal", tanggalawal)
                Rpt.SetParameterValue("tanggalakhir", tanggalakhir)
                Rpt.SetParameterValue("user", NikToko)

                CrystalReportViewer1.ReportSource = Rpt
                CrystalReportViewer1.Zoom(1)

                btnCetak.Enabled = True

            ElseIf cmbJenisLap.SelectedIndex = 3 Then
                'Mcom.CommandText = "  SELECT kode_modis,shelfing,tanggal,nik,nama,
                '                      (SELECT COUNT(*) FROM cekpjr
                '                    WHERE jenisBarang <> '' AND `status` <> 'S' AND TGLSCAN = '') AS itt,
                '                    (SELECT COUNT(*) FROM TINDAKLBTD
                '                    WHERE jenisBarang = '' AND `status` = 'B' AND TGLSCAN = '') AS ada,
                '                    (SELECT COUNT(*) FROM TINDAKLBTD
                '                    WHERE jenisBarang <> '' AND `status` <> 'S' AND TGLSCAN = '') AS tidak
                '                     FROM temp_jadwal_pjr WHERE tanggal = ''"
                tanggal = dtpTglAwal.Value.ToString("yyyy-MM-dd")
                Console.WriteLine(tanggal)
                Madp.SelectCommand.CommandText = "SELECT * FROM temp_tindaklbtd_detail WHERE  TANGGAL LIKE '" & tanggal.Substring(0, 7) & "%'"
                Console.WriteLine(Madp.SelectCommand.CommandText)
                TraceLog("WDCP_CetakLaporan_3 : " & Madp.SelectCommand.CommandText & "")

                dtCP.Clear()
                Madp.Fill(dtCP)
                Rpt.SetDataSource(dtCP)
                tanggalawal = "01" & dtpTglAwal.Value.ToString("dd-MMMM-yyyy").Substring(2)
                Mcom.CommandText = "SELECT LAST_DAY('" & tanggal & "')"
                tanggalakhir = Mcom.ExecuteScalar
                Rpt.SetParameterValue("toko", FormMain.Toko.Kode & "/" & FormMain.Toko.Lokasi)
                Rpt.SetParameterValue("tanggalawal", tanggalawal)
                Rpt.SetParameterValue("tanggalakhir", tanggalakhir)
                Rpt.SetParameterValue("user", NikToko)

                CrystalReportViewer1.ReportSource = Rpt
                CrystalReportViewer1.Zoom(1)

                btnCetak.Enabled = True


            ElseIf cmbJenisLap.SelectedIndex = 4 Then
                Madp.SelectCommand.CommandText = "SELECT a.tanggal as Tanggal,a.modis as Modis,a.shelf as Shelfing,
                                                  a.PJR as `Status`,a.nama as NamaPersonil,a.nik as NIKPersonil,
                                                    b.jabatan AS Jabatan FROM temp_tindaklbtd_detail a LEFT JOIN temp_jadwal_pjr b 
                                                    ON a.nik = b.nik AND a.tanggal = b.tanggal AND a.modis=b.kode_modis WHERE StatusApproval = 'Y' AND
                                                 b.tanggal
                                                 >= (SELECT DATE_ADD('" & tanggal & "', INTERVAL(1-DAYOFWEEK('" & tanggal & "')) DAY)) 
                                                  AND
                                                b.tanggal
                                                 <= (SELECT DATE_ADD('" & tanggal & "', INTERVAL(7-DAYOFWEEK('" & tanggal & "')) DAY))"
                Console.WriteLine(Madp.SelectCommand.CommandText)
                TraceLog("WDCP_CetakLaporan_4 : " & Madp.SelectCommand.CommandText & "")

                dtCP.Clear()
                Madp.Fill(dtCP)
                Rpt.SetDataSource(dtCP)
                tanggalawal = cPJR.getTanggal(tanggal, "awal")
                tanggalakhir = cPJR.getTanggal(tanggal, "akhir")
                Rpt.SetParameterValue("toko", FormMain.Toko.Kode & "/" & FormMain.Toko.Lokasi)
                Rpt.SetParameterValue("tanggalawal", tanggalawal)
                Rpt.SetParameterValue("tanggalakhir", tanggalakhir)
                Rpt.SetParameterValue("user", NikToko)

                CrystalReportViewer1.ReportSource = Rpt
                CrystalReportViewer1.Zoom(1)

                btnCetak.Enabled = True

            ElseIf cmbJenisLap.SelectedIndex = 5 Then
                Madp.SelectCommand.CommandText = "SELECT DATE_FORMAT(a.tanggal,'%d %b %Y') as Tanggal,a.modis as Modis,a.shelf as Shelfing,
                                                a.PJR as `Status`,a.nama as NamaPersonil,a.nik as NIKPersonil,b.jabatan AS Jabatan 
                                                FROM temp_tindaklbtd_detail a LEFT JOIN temp_jadwal_pjr b ON a.nik = b.nik 
                                                AND a.tanggal = b.tanggal AND a.modis=b.kode_modis WHERE StatusApproval = 'Y' AND
                                                b.tanggal
                                                 >= (SELECT DATE_ADD('" & tanggal & "', INTERVAL(1-DAYOFWEEK('" & tanggal & "')) DAY)) 
                                                  AND
                                                b.tanggal
                                                 <= (SELECT DATE_ADD('" & tanggal & "', INTERVAL(7-DAYOFWEEK('" & tanggal & "')) DAY)) "
                Console.WriteLine(Madp.SelectCommand.CommandText)
                TraceLog("WDCP_CetakLaporan_5 : " & Madp.SelectCommand.CommandText & "")

                dtCP.Clear()
                Madp.Fill(dtCP)
                Rpt.SetDataSource(dtCP)
                tanggalawal = cPJR.getTanggal(tanggal, "awal")
                tanggalakhir = cPJR.getTanggal(tanggal, "akhir")
                Rpt.SetParameterValue("toko", FormMain.Toko.Kode & "/" & FormMain.Toko.Lokasi)
                Rpt.SetParameterValue("tanggalawal", tanggalawal)
                Rpt.SetParameterValue("tanggalakhir", tanggalakhir)
                Rpt.SetParameterValue("user", NikToko)

                CrystalReportViewer1.ReportSource = Rpt
                CrystalReportViewer1.Zoom(1)

                btnCetak.Enabled = True

            ElseIf cmbJenisLap.SelectedIndex = 6 Then

                Mcom.CommandText = "SHOW TABLES LIKE 'tindaklbtd_bapjr'"
                If Mcom.ExecuteScalar = "" Then
                    MessageBox.Show("Maaf, Anda belum melakukan Tindak LBTD BA AS !", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)

                    Exit Sub
                End If


                Mcom.CommandText = "DROP TABLE IF EXISTS `temp_itt_ketemu`"
                Mcom.ExecuteNonQuery()
                Mcom.CommandText = "CREATE TABLE temp_itt_ketemu SELECT * FROM(SELECT prdcd, stock AS QTY, stock*PRICE AS QTY_RP FROM tindaklbtd_bapjr a INNER JOIN prodmast b ON a.plu = b.prdcd 
                                    WHERE a.plu IN (SELECT PLU FROM tindaklbtd_bapjr WHERE    DATE(tglscan) = '" & tanggal & "'  AND `STATUS` = 'B')) b;"
                TraceLog("WDCP_CetakLaporan_6_a : " & Mcom.CommandText & "")

                Mcom.ExecuteNonQuery()

                Mcom.CommandText = "DROP TABLE IF EXISTS `temp_itt`"
                Mcom.ExecuteNonQuery()
                Mcom.CommandText = "CREATE TABLE temp_itt SELECT * FROM(SELECT prdcd, stock AS QTY, stock*PRICE AS QTY_RP FROM tindaklbtd_bapjr a INNER JOIN prodmast b ON a.plu = b.prdcd 
                                   WHERE a.plu IN (SELECT PLU FROM tindaklbtd_bapjr WHERE   DATE(tglscan) = '" & tanggal & "'  AND `STATUS` = 'I' AND JENISBARANG = 'TT')) c;"
                TraceLog("WDCP_CetakLaporan_6_b : " & Mcom.CommandText & "")

                Mcom.ExecuteNonQuery()

                'Madp.SelectCommand.CommandText = " Select a.NamaRak,a.NoRak,a.NoShelf,a.KiriKanan,a.Divisi,a.PLU,a.Nama as `Desc`, "
                'Madp.SelectCommand.CommandText &= " IF(b.qty IS NULL,0,b.qty) AS ITT_QTY, IF(b.qty_rp IS NULL,0,b.qty_rp) AS ITT_RP, IF(c.qty IS NULL,0,c.qty) AS ITT_KETEMU_QTY,"
                'Madp.SelectCommand.CommandText &= "  IF(c.qty_rp IS NULL,0,c.qty_rp) AS ITT_KETEMU_RP,IF(b.qty IS NULL,0,b.qty) AS BA_QTY,IF(b.qty_rp IS NULL,0,b.qty_rp) AS BA_RP FROM TINDAKLBTD a"
                'Madp.SelectCommand.CommandText &= " LEFT JOIN temp_itt b ON a.PLU = b.prdcd LEFT JOIN temp_itt_ketemu c ON a.PLU = c.prdcd "
                'Madp.SelectCommand.CommandText &= " WHERE nik = '" & cbNik.Text & "' and  DATE(tglscan) = '" & tanggal & "' "
                'Madp.SelectCommand.CommandText &= " AND (JENISBARANG = 'TT' OR STATUS = 'B') "

                Madp.SelectCommand.CommandText = " Select a.NamaRak,a.NoRak,a.NoShelf,a.KiriKanan,a.Divisi,a.PLU,a.Nama as `Desc`, "
                Madp.SelectCommand.CommandText &= " IF(b.qty IS NULL,0,b.qty) AS ITT_QTY, IF(b.qty_rp IS NULL,0,b.qty_rp) AS ITT_RP, IF(c.qty IS NULL,0,c.qty) AS ITT_KETEMU_QTY,"
                Madp.SelectCommand.CommandText &= "  IF(c.qty_rp IS NULL,0,c.qty_rp) AS ITT_KETEMU_RP,IF(d.qty IS NULL,0,d.qty) AS BA_QTY,IF(d.price_jual IS NULL,0,d.qty * d.price_jual) AS BA_RP FROM tindaklbtd_bapjr a"
                Madp.SelectCommand.CommandText &= " LEFT JOIN temp_itt b ON a.PLU = b.prdcd LEFT JOIN temp_itt_ketemu c ON a.PLU = c.prdcd "
                Madp.SelectCommand.CommandText &= " LEFT JOIN (SELECT prdcd,qty,price_jual FROM mstran WHERE istype = 'BS' AND KETER = 'PJR' AND DATE(BUKTI_TGL) = '" & tanggal & "') d ON a.plu = d.prdcd "
                Madp.SelectCommand.CommandText &= " WHERE   DATE(tglscan) = '" & tanggal & "' "
                Madp.SelectCommand.CommandText &= " AND (JENISBARANG = 'TT' OR STATUS = 'B') GROUP BY PLU"

                Console.WriteLine(Madp.SelectCommand.CommandText)
                TraceLog("WDCP_CetakLaporan_6 : " & Madp.SelectCommand.CommandText & "")

                dtCP.Clear()
                Madp.Fill(dtCP)
                Rpt.SetDataSource(dtCP)

                Rpt.SetParameterValue("toko", FormMain.Toko.Kode & "/" & FormMain.Toko.Lokasi)
                Rpt.SetParameterValue("Tgl", Now())
                Rpt.SetParameterValue("nik", cbNik.Text)

                CrystalReportViewer1.ReportSource = Rpt
                CrystalReportViewer1.Zoom(1)

                btnCetak.Enabled = True
            End If



        Catch ex As Exception
            MsgBox(ex.Message & ex.StackTrace)
        End Try


    End Sub

    Private Sub btnCetak_Click(sender As Object, e As EventArgs) Handles btnCetak.Click
        MsgBox("Pastikan menggunakan printer besar!")
        Rpt.PrintToPrinter(1, False, 0, 0)
        MsgBox("Cetak Selesai!")
    End Sub

    Private Sub btnKeluar_Click(sender As Object, e As EventArgs) Handles btnKeluar.Click
        Me.Close()

    End Sub

    Private Sub cmbJenisLap_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbJenisLap.SelectedIndexChanged
        Dim CPJR As New ClsPJRController
        Dim DtUser As New DataTable
        FormMain.isPengganti = False
        'If cmbJenisLap.SelectedIndex = 6 Then
        '    DtUser = CPJR.GetPersonilCetak
        '    cbNik.Items.Clear()

        '    If DtUser.Rows.Count > 0 Then
        '        For Each Dr As DataRow In DtUser.Rows
        '            cbNik.Items.Add(Dr(0))
        '        Next

        '    Else

        '    End If
        '    cbNik.Visible = True
        '    lblNIK.Visible = True

        '    btnProses.Enabled = False
        'Else
        cbNik.Visible = False
            lblNIK.Visible = False
            btnProses.Enabled = True
            btnCetak.Enabled = True
        'End If
    End Sub



    Private Sub cbNik_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbNik.SelectedIndexChanged
        btnProses.Enabled = True
        btnCetak.Enabled = True

    End Sub
End Class