Imports MySql.Data.MySqlClient
Imports IDM.Fungsi
Public Class FrmRegistPJR
    Private Sub FrmRegistPJR_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim cUser As New ClsUserController
        Dim cPJR As New ClsPJRController
        Dim DtRak As New DataTable
        Dim DtUser As New DataTable
        Dim dt As New DataTable
        Dim dtHari As New DataTable

        Try
            cbHariBuka.Text = FormMain.cbHariBukaToko
            cPJR.ReloadJadwal()
            cPJR.createTabelJadwal()
            DtUser = cPJR.GetPersonil
            If DtUser.Rows.Count > 0 Then
                For Each Dr As DataRow In DtUser.Rows
                    cmbNIK.Items.Add(Dr(0))
                Next

            Else

            End If

            cmbHari.Enabled = False
            cmbModis.Enabled = False

            btnTambahPJR.Enabled = False


            dt = cPJR.GETJADWAL

            dgvJadwalPJR.DataSource = dt

            dgvJadwalPJR.Columns(0).Width = 117
            dgvJadwalPJR.Columns(1).Width = 160
            dgvJadwalPJR.Columns(2).Width = 50
            dgvJadwalPJR.Columns(3).Width = 70
            dgvJadwalPJR.Columns(4).Width = 73
            dgvJadwalPJR.Columns(5).Width = 160
            dgvJadwalPJR.Columns(6).Width = 45
            dgvJadwalPJR.Columns(7).Width = 45
            dgvJadwalPJR.ReadOnly = True
            dgvJadwalPJR.Refresh()


        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    Private Sub cmbNIK_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbNIK.SelectedIndexChanged
        Dim ClsPJR As New ClsPJRController
        Dim DtModis As New DataTable
        Dim cls As New ClsPJR
        Dim Count As Integer = 0
        Dim NamaPersonil As String = ""
        Dim CountModis As Integer = 0
        Dim NamaModis As String = ""
        Dim dtHari As New DataTable
        Dim temp_tanggal As Date
        Dim tanggal As String
        If cmbNIK.SelectedIndex <> -1 Then

            Try
                txtNamaPersonil.Text = ClsPJR.CekPersonil(cmbNIK.Text, NamaPersonil)

                'cmbNIK.SelectedIndex = -1
                cmbHari.SelectedIndex = -1
                cmbHari.Enabled = False
                cmbModis.Enabled = False

                'txtNamaPersonil.Text = ""
                txtNamaModis.Text = ""
                cmbModis.Text = ""
                cmbNorak.Text = ""

                TextBoxFrom.Text = ""
                TextBoxTo.Text = ""

                dtHari = ClsPJR.GetHari(cmbNIK.Text)
                cmbHari.Items.Clear()
                cmbModis.Items.Clear()


                If dtHari.Rows.Count > 0 Then
                    For Each Dr As DataRow In dtHari.Rows
                        temp_tanggal = Dr(0)
                        tanggal = temp_tanggal.ToString("dddd", New Globalization.CultureInfo("id-ID"))
                        Console.WriteLine(tanggal)
                        cmbHari.Items.Add(tanggal)
                    Next
                    cmbHari.Enabled = True
                    cmbModis.Enabled = True

                    btnTambahPJR.Enabled = False
                Else
                    cmbHari.Enabled = False
                    cmbModis.Enabled = False

                    btnTambahPJR.Enabled = False
                    MsgBox("Tidak ada jadwal minggu ini ! untuk nik - " & cmbNIK.Text & "")
                End If



            Catch ex As Exception
                IDM.Fungsi.TraceLog("Error cmbNIK_SelectedIndexChanged " & ex.Message & ex.StackTrace)
                Exit Sub
            End Try
        End If
    End Sub

    Private Sub btnCari_byNIK_Click(sender As Object, e As EventArgs)
        Dim cpjr As New ClsPJRController
        Dim dt As New DataTable
        Dim dtHari As New DataTable
        Dim temp_tanggal As Date
        Dim tanggal As String
        dt = cpjr.GETJADWAL_BYNIK(cmbNIK.Text)

        dgvJadwalPJR.DataSource = dt

        dgvJadwalPJR.Columns(0).Width = 70
        dgvJadwalPJR.Columns(1).Width = 100
        dgvJadwalPJR.Columns(2).Width = 360
        dgvJadwalPJR.Columns(3).Width = 70
        dgvJadwalPJR.ReadOnly = True
        dgvJadwalPJR.Refresh()


        dtHari = cpjr.GetHari(cmbNIK.Text)
        cmbHari.Items.Clear()
        cmbModis.Items.Clear()


        If dtHari.Rows.Count > 0 Then
            For Each Dr As DataRow In dtHari.Rows
                temp_tanggal = Dr(0)
                tanggal = temp_tanggal.ToString("dddd, yyyy-MM-dd", New Globalization.CultureInfo("id-ID"))

                cmbHari.Items.Add(tanggal)
            Next
            cmbHari.Enabled = True
            cmbModis.Enabled = True

            btnTambahPJR.Enabled = True
        Else
            cmbHari.Enabled = False
            cmbModis.Enabled = False

            btnTambahPJR.Enabled = False
            MsgBox("tidak ada jadwal minggu ini ! untuk nik - " & cmbNIK.Text & "")
        End If
    End Sub

    Private Sub cmbHari_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbHari.SelectedIndexChanged
        Dim ClsPJR As New ClsPJRController
        Dim DtModis As New DataTable
        Dim cls As New ClsPJR
        Dim Count As Integer = 0
        Dim NamaPersonil As String = ""
        Dim CountModis As Integer = 0
        Dim NamaModis As String = ""
        If cmbHari.SelectedIndex <> -1 Then
            Try
                cmbModis.Items.Clear()
                DtModis = ClsPJR.GetNamaRak_2
                If DtModis.Rows.Count > 0 Then
                    For Each Dr As DataRow In DtModis.Rows
                        cmbModis.Items.Add(Dr(0))
                    Next
                    cmbHari.Enabled = True
                Else
                    'MsgBox("Tidak ada Modis yang tersedia atau semua modis telah terdaftar!")

                End If
            Catch ex As Exception
                Exit Sub
            End Try
        End If
        'If cmbHari.SelectedIndex <> -1 Then
        '    Try
        '        cmbModis.Items.Clear()
        '        DtModis = ClsPJR.GetNamaRak(cmbHari.Text.Split(",")(1))
        '        If DtModis.Rows.Count > 0 Then
        '            For Each Dr As DataRow In DtModis.Rows
        '                cmbModis.Items.Add(Dr(0))
        '            Next
        '        Else

        '        End If
        '    Catch ex As Exception
        '        Exit Sub
        '    End Try
        'End If
    End Sub

    Private Sub btnTambahPJR_Click(sender As Object, e As EventArgs) Handles btnTambahPJR.Click
        Dim cpjr As New ClsPJRController
        Dim result As Boolean
        Dim dt As New DataTable
        Dim DtUser As New DataTable
        'Console.WriteLine(cmbHari.Text.Split(",")(1).Trim)
        result = cpjr.tambahPersonilPJR_Temp(cmbNIK.Text, txtNamaPersonil.Text, cmbHari.Text,
                               cmbModis.Text, txtNamaModis.Text, cmbNorak.Text,
                               TextBoxFrom.Text & "-" & TextBoxTo.Text)
        If result = True Then
            'uat
            DtUser = cpjr.GetPersonil
            cmbNIK.Items.Clear()

            If DtUser.Rows.Count > 0 Then
                For Each Dr As DataRow In DtUser.Rows
                    cmbNIK.Items.Add(Dr(0))
                Next

            Else

            End If

            cmbHari.Enabled = False
            cmbModis.Enabled = False

            btnTambahPJR.Enabled = False



            dt = cpjr.GETJADWAL

            dgvJadwalPJR.DataSource = dt

            dgvJadwalPJR.Columns(0).Width = 117
            dgvJadwalPJR.Columns(1).Width = 160
            dgvJadwalPJR.Columns(2).Width = 50
            dgvJadwalPJR.Columns(3).Width = 70
            dgvJadwalPJR.Columns(4).Width = 73
            dgvJadwalPJR.Columns(5).Width = 160
            dgvJadwalPJR.Columns(6).Width = 45
            dgvJadwalPJR.Columns(7).Width = 45
            dgvJadwalPJR.ReadOnly = True
            dgvJadwalPJR.Refresh()

            cmbNIK.Text = ""
            cmbHari.SelectedIndex = -1
            cmbModis.SelectedIndex = -1

            txtNamaPersonil.Text = ""
            txtNamaModis.Text = ""
            cmbNorak.Text = ""
            TextBoxFrom.Text = ""
            TextBoxTo.Text = ""

        Else
            MsgBox("Gagal")
        End If

    End Sub

    Private Sub cmbModis_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbModis.SelectedIndexChanged
        Dim ClsPJR As New ClsPJRController
        Dim DtModis As New DataTable
        Dim CountModis As Integer = 0
        Dim NamaModis As String = ""
        Dim NomorRak As String = ""

        If cmbModis.SelectedIndex <> -1 Then
            'cmbShelfFrom.Enabled = True
            'cmbShelfTo.Enabled = True
            TextBoxFrom.Text = ""
            TextBoxTo.Text = ""
            txtNamaModis.Text = ""
            cmbNorak.Items.Clear()

            Try
                DtModis = ClsPJR.CekModis(cmbModis.Text, NamaModis, NomorRak)

                'If DtModis.Rows.Count > 1 Then

                txtNamaModis.Text = NamaModis

                For Each Dr As DataRow In DtModis.Rows
                    cmbNorak.Items.Add(Dr(0))
                Next
                Console.WriteLine(DtModis.Rows.Count)
                If DtModis.Rows.Count = 0 Then
                    cmbNorak.Items.Add("1")

                End If
                cmbNorak.Enabled = True
                'TextBoxFrom.Text = DtModis.Rows(0)("noshelf").ToString
                'TextBoxTo.Text = DtModis.Rows(DtModis.Rows.Count - 1)("noshelf").ToString


                'txtNomorRaK.Text = NomorRak

                'End If


                'DtModis = ClsPJR.CekModis(cmbModis.Text, cmbHari.Text.Split(",")(1).Trim, CountModis, NamaModis)
                'If CountModis = 0 Then
                '    MsgBox("Modis ini sudah selesai diproses!")
                '    cmbModis.SelectedIndex = -1
                'Else
                '    txtNamaModis.Text = NamaModis
                '    If DtModis.Rows.Count > 0 Then
                '        Console.WriteLine(DtModis.Rows(0)("noshelf").ToString)

                '        Console.WriteLine(DtModis.Rows(DtModis.Rows.Count - 1)("noshelf").ToString)
                '        TextBoxFrom.Text = DtModis.Rows(0)("noshelf").ToString
                '        TextBoxTo.Text = DtModis.Rows(DtModis.Rows.Count - 1)("noshelf").ToString

                '    End If

                'End If

            Catch ex As Exception
                MsgBox(ex.Message & ex.StackTrace)
                Exit Sub
            End Try
        End If
    End Sub

    'Private Sub dgvJadwalPJR_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvJadwalPJR.CellMouseDoubleClick
    '    Dim clsPJR As New ClsPJR
    '    Dim PJR As New ClsPJRController

    '    Dim cell As Integer = dgvJadwalPJR.CurrentRow.Index
    '    Dim hari As String = dgvJadwalPJR(0, cell).Value.ToString
    '    Dim NIK As String = dgvJadwalPJR(1, cell).Value.ToString
    '    Dim MODIS As String = dgvJadwalPJR(2, cell).Value.ToString
    '    Dim norak As String = dgvJadwalPJR(3, cell).Value.ToString
    '    Dim noshelf As String = dgvJadwalPJR(4, cell).Value.ToString

    '    Console.WriteLine(hari)
    '    Console.WriteLine(NIK)
    '    Console.WriteLine(modis)
    '    Console.WriteLine(norak)
    '    Console.WriteLine(noshelf)
    '    'dgvJadwalPJR.Columns(0).Width = 70  'HARI
    '    'dgvJadwalPJR.Columns(1).Width = 100 'NIK
    '    'dgvJadwalPJR.Columns(2).Width = 200 'MODIS
    '    'dgvJadwalPJR.Columns(3).Width = 70  'NORAK
    '    'dgvJadwalPJR.Columns(4).Width = 70  'SHELFING
    '    clsPJR = PJR.ambilPersonilPJR(hari, NIK, MODIS, norak, noshelf)

    '    cmbNIK.SelectedIndex = -1
    '    txtNamaPersonil.Text = ""

    '    cmbHari.SelectedIndex = -1
    '    cmbModis.SelectedIndex = -1

    '    'txtNamaPersonil.Text = ""
    '    txtNamaModis.Text = ""
    '    txtNomorRaK.Text = ""
    '    TextBoxFrom.Text = ""
    '    TextBoxTo.Text = ""


    '    cmbNIK.SelectedText = clsPJR.NIK
    '    txtNamaPersonil.Text = clsPJR.NAMA

    '    cmbHari.SelectedText = clsPJR.HARI
    '    cmbModis.SelectedText = clsPJR.MODIS

    '    txtNamaModis.Text = clsPJR.NAMAMODIS
    '    txtNomorRaK.Text = clsPJR.NORAK
    '    TextBoxFrom.Text = clsPJR.SHELFFROM
    '    TextBoxTo.Text = clsPJR.SHELFTO

    'End Sub

    Private Sub btnSimpanPJR_Click(sender As Object, e As EventArgs) Handles btnSimpanPJR.Click
        Dim confirm As DialogResult
        Dim cPJR As New ClsPJRController
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As Boolean = False
        Dim Mcom As New MySqlCommand("", Conn)
        Dim dt1 As New DataTable
        Dim dtrak As New DataTable
        Dim timesecond As Double = 0.0
        Dim jumlahitem As Integer = 0
        Dim totalitem As Integer = 0
        Dim totalestimasi As Double = 0.0
        Dim progres As Integer = 1
        Dim jabatan As String = ""
        Dim hariBuka As Integer = 0
        Dim shelfing As String = ""
        Dim shelfing_min As String = ""
        Dim shelfing_max As String = ""
        Dim jamProses As Date

        confirm = MessageBox.Show("Apakah Anda yakin akan Proses Simpan Jadwal PJR?", "Jadwal PJR", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If confirm = Windows.Forms.DialogResult.Yes Then


            btnSimpanPJR.Enabled = False
            Label8.Visible = True
            'cPJR.SimpanJadwalPJR()

            Try
                If Conn.State = ConnectionState.Closed Then
                    Conn.Open()
                End If


                jamProses = Date.Now
                jabatan = cPJR.getJabatanVirbacaprod
                Mcom.CommandText = "SELECT COUNT(1) FROM SOPPAGENT.ABSPEGAWAIMST WHERE JABATAN IN (" & jabatan & ")
                                    AND MENOIN NOT IN (SELECT NIK FROM TEMP_JADWAL_PENANGGUNGJAWABRAK ) AND pinjaman = 0 "
                TraceLog(Mcom.CommandText)
                If Mcom.ExecuteScalar > 0 Then
                    MessageBox.Show("Maaf, Proses Simpan belum dapat dilakukan jika seluruh Personil belum didaftarkan !", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Rtn = False
                    Exit Try
                End If
                If cbHariBuka.Text.Contains("5") Then
                    hariBuka = 5
                ElseIf cbHariBuka.Text.Contains("6") Then
                    hariBuka = 6
                End If

                Dim dtMenoin As New DataTable
                Madp.SelectCommand.CommandText = "SELECT menoin FROM SOPPAGENT.ABSPEGAWAIMST WHERE JABATAN IN (" & jabatan & ") AND pinjaman = 0 group by menoin"
                Console.WriteLine(Madp.SelectCommand.CommandText)

                dtMenoin.Clear()
                Madp.Fill(dtMenoin)

                For i As Integer = 1 To hariBuka
                    For j As Integer = 0 To dtMenoin.Rows.Count - 1
                        Mcom.CommandText = "SELECT count(1) FROM TEMP_JADWAL_PENANGGUNGJAWABRAK WHERE NIK  = '" & dtMenoin.Rows(j)("MENOIN").ToString & "' "
                        If i = 1 Then
                            Mcom.CommandText &= "  AND HARI = 'Senin'"

                        ElseIf i = 2 Then
                            Mcom.CommandText &= "  AND HARI = 'Selasa'"

                        ElseIf i = 3 Then
                            Mcom.CommandText &= "  AND HARI = 'Rabu'"

                        ElseIf i = 4 Then
                            Mcom.CommandText &= "  AND HARI = 'Kamis'"

                        ElseIf i = 5 Then
                            Mcom.CommandText &= "  AND HARI = 'Jumat'"

                        ElseIf i = 6 Then
                            Mcom.CommandText &= "  AND HARI = 'Sabtu'"


                        End If
                        Console.WriteLine(Mcom.CommandText)
                        If Mcom.ExecuteScalar = 0 Then
                            MessageBox.Show("Maaf, Proses Simpan belum dapat dilakukan, Setiap Nik wajib didaftarkan setiap hari !", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Rtn = False
                            Exit Try
                        End If

                    Next
                Next
                ProgressBar1.Value = 50

                Mcom.CommandText = "DROP TABLE IF EXISTS `pos`.`temp_jadwal_pjr_estimasi_Detail` "
                Mcom.ExecuteNonQuery()

                Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`temp_jadwal_pjr_estimasi_Detail` ( 
                                `NIK` VARCHAR(12), 
                                `NAMA` VARCHAR(99), 
                                `JABATAN` VARCHAR(50),
                                `HARI` VARCHAR(10),
                                `Kode_Modis` VARCHAR(20),
                                `MODIS` VARCHAR(99),
                                `Shelfing` VARCHAR(10),
                                `NORAK` VARCHAR(10),
                                `Addtime` DATE,
                                `Cat_cod` Varchar(8),
                                `Kemasan` Varchar(5),
                                `Totalitem` Varchar(5),
                                `Timesecond` Varchar(5)
                                ) ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci;"
                Mcom.ExecuteNonQuery()

                'Memo 447/cps/23
                'PJR hnya FJP=Y

                Mcom.CommandText = "SELECT COUNT(DISTINCT modisp) FROM bracket a INNER JOIN rak b ON a.modisp = b.kodemodis
                                    WHERE (modisp,no_rak) NOT  IN (SELECT KODE_MODIS,norak FROM temp_jadwal_penanggungjawabrak WHERE NIK <> '') AND b.flagprod LIKE '%FJP=Y%'"

                'Mcom.CommandText = "SELECT COUNT(DISTINCT modisp) FROM bracket a INNER JOIN rak b ON a.modisp = b.kodemodis
                '                    WHERE (modisp,no_rak) NOT  IN (SELECT KODE_MODIS,norak FROM temp_jadwal_penanggungjawabrak WHERE NIK <> '') AND b.flagprod NOT LIKE '%FJP=N%'"
                If Mcom.ExecuteScalar <> 0 Then
                    'Memo 447/cps/23
                    'PJR hnya FJP=Y

                    Mcom.CommandText = "SELECT COUNT(DISTINCT KODEMODIS) FROM RAK WHERE KODEMODIS NOT IN (SELECT KODE_MODIS FROM TEMP_JADWAL_PENANGGUNGJAWABRAK WHERE NIK <> '') and flagprod LIKE '%FJP=Y%'"
                    'Mcom.CommandText = "SELECT COUNT(DISTINCT KODEMODIS) FROM RAK WHERE KODEMODIS NOT IN (SELECT KODE_MODIS FROM TEMP_JADWAL_PENANGGUNGJAWABRAK WHERE NIK <> '') and flagprod NOT LIKE '%FJP=N%'"

                    If Mcom.ExecuteScalar <> 0 Then
                        MessageBox.Show("Maaf, Proses Simpan belum dapat dilakukan jika seluruh MODIS belum didaftarkan !", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        ProgressBar1.Value = 0

                        Rtn = False
                        Exit Try
                    End If

                Else
                    Label10.Visible = False
                    'Madp.SelectCommand.CommandText = "SELECT nik,nama,jabatan,hari,kode_modis,modis,shelfing,norak FROM temp_jadwal_penanggungjawabrak GROUP BY kode_modis,nik,norak"
                    Madp.SelectCommand.CommandText = "SELECT a.nik,a.nama,a.jabatan,a.hari,a.kode_modis,a.modis,a.shelfing,a.norak, b.statusapproval FROM temp_jadwal_penanggungjawabrak a LEFT JOIN
                                       jadwal_penanggungjawabrak b ON a.nik=b.nik AND a.hari=b.hari AND a.kode_modis=b.kode_modis AND a.norak=b.norak 
                                       WHERE b.statusapproval <> 'Y'
                                       GROUP BY kode_modis,nik,norak"

                    dtrak.Clear()
                    Madp.Fill(dtrak)
                    'ProgressBar1.Value = ProgressBar1.Value + 10
                    ProgressBar1.Value = 75

                    'progres = (80) / (dtrak.Rows.Count + progres)
                    For j As Integer = 0 To dtrak.Rows.Count - 1
                        totalitem = 0

                        Madp.SelectCommand.CommandText = "SELECT a.cat_cod, kemasan From prodmast a 
                                LEFT Join rak b ON a.prdcd = b.prdcd inner join TEMP_JADWAL_PENANGGUNGJAWABRAK c on b.kodemodis=c.kode_modis  
                                Where kodemodis = '" & dtrak.Rows(j)("Kode_Modis").ToString & "' AND c.NORAK = '" & dtrak.Rows(j)("NORAK").ToString & "'  
                                GROUP BY CAT_COD,KEMASAN ORDER BY CAT_COD "

                        'Console.Writeline(Madp.SelectCommand.CommandText)
                        dt1.Clear()
                        Madp.Fill(dt1)
                        'Console.Writeline(dt1.Rows.Count)

                        For i As Integer = 0 To dt1.Rows.Count - 1
                            'ambil timesecond per cat_cod dan kemasan
                            If dt1.Rows(i)("cat_cod").ToString.StartsWith("0") Then
                                dt1.Rows(i)("cat_cod") = dt1.Rows(i)("cat_cod").ToString.Substring(1)
                                ''Console.Writeline(dt1.Rows(i)("cat_cod"))
                            End If
                            Mcom.CommandText = "Select `TIMESECOND` FROM penanggungjawab_rak 
                                    where ctgr like '%" & dt1.Rows(i)("cat_cod") & "%'
                                    AND KEMASAN = '" & dt1.Rows(i)("kemasan") & "'"
                            ''Console.Writeline(Mcom.CommandText)
                            TraceLog("Kueri  : " & Mcom.CommandText)

                            timesecond = Mcom.ExecuteScalar

                            Mcom.CommandText = "SELECT shelfing FROM TEMP_JADWAL_PENANGGUNGJAWABRAK WHERE kode_modis = '" & dtrak.Rows(j)("Kode_Modis").ToString & "'
                                                AND NORAK = '" & dtrak.Rows(j)("NORAK").ToString & "'"
                            shelfing = Mcom.ExecuteScalar

                            shelfing_min = shelfing.Split("-")(0)
                            shelfing_max = shelfing.Split("-")(1)


                            Mcom.CommandText = "SELECT COUNT(DISTINCT a.prdcd,c.norak) FROM prodmast a
                                    LEFT JOIN rak b ON a.prdcd = b.prdcd inner join TEMP_JADWAL_PENANGGUNGJAWABRAK c on b.kodemodis = c.kode_modis 
                                    WHERE kodemodis = '" & dtrak.Rows(j)("Kode_Modis").ToString & "' AND CAT_COD LIKE '%" & dt1.Rows(i)("cat_cod") & "%' 
                                    AND KEMASAN LIKE '%" & dt1.Rows(i)("kemasan") & "%' AND c.NORAK = '" & dtrak.Rows(j)("NORAK").ToString & "' 
                                    AND b.noshelf >= " & shelfing_min & " AND b. noshelf <= " & shelfing_max & "
                                    ORDER BY CAT_COD "

                            ''Console.Writeline(Mcom.CommandText)
                            TraceLog("Kueri  : " & Mcom.CommandText)

                            jumlahitem = Mcom.ExecuteScalar
                            ''Console.Writeline(jumlahitem)
                            totalitem += jumlahitem

                            Mcom.CommandText = "INSERT IGNORE INTO TEMP_JADWAL_PJR_estimasi_detail VALUES(
                                '" & dtrak.Rows(j)("nik").ToString & "', '" & dtrak.Rows(j)("NAMA").ToString & "','" & dtrak.Rows(j)("JABATAN").ToString & "'
                                ,'" & dtrak.Rows(j)("HARI").ToString & "','" & dtrak.Rows(j)("Kode_Modis").ToString & "','" & dtrak.Rows(j)("Modis").ToString & "'
                                ,'" & dtrak.Rows(j)("Shelfing").ToString & "' ,'" & dtrak.Rows(j)("norak").ToString & "'
                                ,NOW(),'" & dt1.Rows(i)("cat_cod") & "','" & dt1.Rows(i)("kemasan") & "'
                                ," & jumlahitem & ", " & timesecond & ")"
                            TraceLog("Kueri  : " & Mcom.CommandText)

                            ''Console.Writeline(Mcom.CommandText)
                            Mcom.ExecuteNonQuery()

                        Next

                        If dt1.Rows.Count <> 0 Then
                            Mcom.CommandText = "SELECT CEILING(SUM(TOTALITEM*timesecond)/60)  FROM temp_jadwal_pjr_estimasi_DETAIL 
                                WHERE NIK = '" & dtrak.Rows(j)("nik").ToString & "'  AND KODE_MODIS = '" & dtrak.Rows(j)("Kode_Modis").ToString & "' and norak = '" & dtrak.Rows(j)("NORAK").ToString & "' GROUP BY KODE_MODIS,NORAK ;"
                            TraceLog("Kueri  : " & Mcom.CommandText)
                            totalestimasi = Mcom.ExecuteScalar
                            Mcom.CommandText = "UPDATE temp_jadwal_penanggungjawabrak SET `Totalitem` = '" & totalitem & "' , TotalEstimasi = '" & totalestimasi & "'
                                WHERE NIK = '" & dtrak.Rows(j)("nik").ToString & "' AND KODE_MODIS = '" & dtrak.Rows(j)("Kode_Modis").ToString & "' and norak = '" & dtrak.Rows(j)("NORAK").ToString & "'; "
                            TraceLog("Kueri  : " & Mcom.CommandText)

                            Mcom.ExecuteNonQuery()
                        End If

                        'If j = dtrak.Rows.Count / 2 Then
                        '    'ProgressBar1.Value = ProgressBar1.Value + 40

                        'End If

                    Next

                    ProgressBar1.Value = 100

                    Mcom.CommandText = "DELETE FROM jadwal_penanggungjawabrak WHERE KODE_MODIS NOT IN (SELECT KODE_MODIS FROM TEMP_JADWAL_PENANGGUNGJAWABRAK);"
                    Mcom.ExecuteNonQuery()
                    Try
                        Mcom.CommandText = "DELETE FROM jadwal_penanggungjawabrak WHERE (NIK,HARI,KODE_MODIS,NORAK) IN (SELECT NIK,HARI,KODE_MODIS,NORAK FROM TEMP_HAPUS_JADWAL_PJR);"
                        Mcom.ExecuteNonQuery()

                    Catch ex As Exception

                    End Try

                    Mcom.CommandText = "INSERT IGNORE INTO jadwal_penanggungjawabrak SELECT *,'' FROM temp_jadwal_penanggungjawabrak"
                    Mcom.ExecuteNonQuery()
                    Try
                        Mcom.CommandText = "DELETE FROM TEMP_HAPUS_JADWAL_PJR"
                        Mcom.ExecuteNonQuery()

                    Catch ex As Exception

                    End Try
                    MessageBox.Show("Berhasil Proses Simpan !", "Berhasil", MessageBoxButtons.OK)
                    Me.Close()
                    Rtn = True
                End If
                Label10.Visible = False
            Catch ex As Exception
                IDM.Fungsi.TraceLog("Error WDCP_SimpanJadwalPJR " & ex.Message & ex.StackTrace)
            Finally
                Conn.Close()
            End Try

        Else

        End If
        Label8.Visible = False
        btnSimpanPJR.Enabled = True

    End Sub

    'Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
    '    ProgressBar1.Increment(1)

    '    If ProgressBar1.Value = 100 Then
    '        Timer1.Enabled = False
    '    End If
    'End Sub

    Private Sub cmbNorak_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbNorak.SelectedIndexChanged
        Dim ClsPJR As New ClsPJRController

        Dim rtn As Boolean

        Dim NomorRak As String = ""
        Dim shelf_awal As String = ""
        Dim shelf_akhir As String = ""

        If cmbNorak.SelectedIndex <> -1 Then

            TextBoxFrom.Text = ""
            TextBoxTo.Text = ""
            'txtNamaModis.Text = ""

            Try
                rtn = ClsPJR.ambilNoshelf(cmbModis.Text, cmbNorak.Text, shelf_awal, shelf_akhir)



                TextBoxFrom.Text = shelf_awal

                TextBoxTo.Text = shelf_akhir

                btnTambahPJR.Enabled = True


            Catch ex As Exception
                MsgBox(ex.Message & ex.StackTrace)
                Exit Sub
            End Try
        End If
    End Sub


    'Private Sub dgvJadwalPJR_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvJadwalPJR.CellDoubleClick
    '    Dim cPJR As New ClsPJRController
    '    'HARI,NIK,MODIS,NORAK,SHELFING
    '    dgvJadwalPJR.CurrentRow.Selected = True

    '    If dgvJadwalPJR.Rows(e.RowIndex).Cells("NIK").Value.ToString <> "" Then
    '        cmbNIK.Text = dgvJadwalPJR.Rows(e.RowIndex).Cells("NIK").Value.ToString

    '        'txtNamaPersonil.Text = dgvJadwalPJR.Rows(e.RowIndex).Cells("NAMA").Value.ToString
    '        cmbHari.Text = dgvJadwalPJR.Rows(e.RowIndex).Cells("HARI").Value.ToString

    '        'cmbModis.Text = dgvJadwalPJR.Rows(e.RowIndex).Cells("MODIS").Value.ToString

    '        txtNamaModis.Text = dgvJadwalPJR.Rows(e.RowIndex).Cells("MODIS").Value.ToString

    '        cmbModis.Text = cPJR.AmbilKodeModis(cmbNIK.Text, txtNamaModis.Text)

    '        cmbNorak.Text = dgvJadwalPJR.Rows(e.RowIndex).Cells("NORAK").Value.ToString
    '        If dgvJadwalPJR.Rows(e.RowIndex).Cells("SHELFING").Value.ToString <> "" Then
    '            TextBoxFrom.Text = dgvJadwalPJR.Rows(e.RowIndex).Cells("SHELFING").Value.ToString.Split("-")(0)
    '            TextBoxTo.Text = dgvJadwalPJR.Rows(e.RowIndex).Cells("SHELFING").Value.ToString.Split("-")(1)
    '        End If
    '        btnHapusJadwal.Enabled = True
    '    End If

    'End Sub

    Private Sub btnHapusJadwal_Click(sender As Object, e As EventArgs) Handles btnHapusJadwal.Click
        Dim cpjr As New ClsPJRController
        Dim result As Boolean
        Dim dt As New DataTable
        Dim DtUser As New DataTable
        'Console.WriteLine(cmbHari.Text.Split(",")(1).Trim)
        Dim confirm As DialogResult


        confirm = MessageBox.Show("Apakah Anda yakin akan Menghapus Jadwal PJR tersebut?", "Hapus Jadwal PJR", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If confirm = Windows.Forms.DialogResult.Yes Then

            result = cpjr.HapusPersonilPJR_Temp(cmbNIK.Text, cmbHari.Text,
                               cmbModis.Text, cmbNorak.Text)

            dt = cpjr.GETJADWAL

            dgvJadwalPJR.DataSource = dt
            dgvJadwalPJR.Columns(0).Width = 117
            dgvJadwalPJR.Columns(1).Width = 160
            dgvJadwalPJR.Columns(2).Width = 50
            dgvJadwalPJR.Columns(3).Width = 70
            dgvJadwalPJR.Columns(4).Width = 73
            dgvJadwalPJR.Columns(5).Width = 160
            dgvJadwalPJR.Columns(6).Width = 45
            dgvJadwalPJR.Columns(7).Width = 45
            dgvJadwalPJR.ReadOnly = True
            dgvJadwalPJR.Refresh()

            cmbNIK.Text = ""

            cmbHari.Text = ""
            cmbModis.Text = ""

            txtNamaPersonil.Text = ""
            txtNamaModis.Text = ""
            cmbNorak.Text = ""
            TextBoxFrom.Text = ""
            TextBoxTo.Text = ""
            cmbHari.Enabled = False
            cmbModis.Enabled = False
            cmbNorak.Enabled = False

            btnHapusJadwal.Enabled = False
            btnTambahPJR.Enabled = False

        End If


    End Sub

    Private Sub dgvJadwalPJR_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvJadwalPJR.CellClick
        Dim cPJR As New ClsPJRController
        'HARI,NIK,MODIS,NORAK,SHELFING
        dgvJadwalPJR.CurrentRow.Selected = True
        'If e.RowIndex >= 0 And e.ColumnIndex >= 0 Then
        Try

            If dgvJadwalPJR.Rows(e.RowIndex).Cells("NIK").Value.ToString <> "" Then
                cmbNIK.Text = dgvJadwalPJR.Rows(e.RowIndex).Cells("NIK").Value.ToString
                txtNamaPersonil.Text = dgvJadwalPJR.Rows(e.RowIndex).Cells("NAMA").Value.ToString

                'txtNamaPersonil.Text = dgvJadwalPJR.Rows(e.RowIndex).Cells("NAMA").Value.ToString
                cmbHari.Text = dgvJadwalPJR.Rows(e.RowIndex).Cells("HARI").Value.ToString

                'cmbModis.Text = dgvJadwalPJR.Rows(e.RowIndex).Cells("MODIS").Value.ToString

                txtNamaModis.Text = dgvJadwalPJR.Rows(e.RowIndex).Cells("NAMA_MODIS").Value.ToString

                cmbModis.Text = dgvJadwalPJR.Rows(e.RowIndex).Cells("KODE_MODIS").Value.ToString
                'cmbModis.Text = cPJR.AmbilKodeModis(cmbNIK.Text, txtNamaModis.Text)

                cmbNorak.Text = dgvJadwalPJR.Rows(e.RowIndex).Cells("NORAK").Value.ToString
                If dgvJadwalPJR.Rows(e.RowIndex).Cells("SHELF").Value.ToString <> "" Then
                    TextBoxFrom.Text = dgvJadwalPJR.Rows(e.RowIndex).Cells("SHELF").Value.ToString.Split("-")(0)
                    TextBoxTo.Text = dgvJadwalPJR.Rows(e.RowIndex).Cells("SHELF").Value.ToString.Split("-")(1)
                End If
                btnHapusJadwal.Enabled = True
            End If
            'End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub dgvJadwalPJR_KeyUp(sender As Object, e As KeyEventArgs) Handles dgvJadwalPJR.KeyUp
        If e.KeyCode = Keys.Down Or e.KeyCode = Keys.Up Then

            dgvJadwalPJR.CurrentRow.Selected = True
            'If e.RowIndex >= 0 And e.ColumnIndex >= 0 Then
            Try
                Dim row As DataGridViewRow
                row = dgvJadwalPJR.CurrentRow

                If row.Cells("NIK").Value.ToString <> "" Then
                    cmbNIK.Text = row.Cells("NIK").Value.ToString
                    txtNamaPersonil.Text = row.Cells("NAMA").Value.ToString
                    'txtNamaPersonil.Text = dgvJadwalPJR.Rows(e.RowIndex).Cells("NAMA").Value.ToString
                    cmbHari.Text = row.Cells("HARI").Value.ToString

                    'cmbModis.Text = dgvJadwalPJR.Rows(e.RowIndex).Cells("MODIS").Value.ToString

                    txtNamaModis.Text = row.Cells("NAMA_MODIS").Value.ToString

                    cmbModis.Text = row.Cells("KODE_MODIS").Value.ToString
                    'cmbModis.Text = cPJR.AmbilKodeModis(cmbNIK.Text, txtNamaModis.Text)

                    cmbNorak.Text = row.Cells("NORAK").Value.ToString
                    If row.Cells("SHELF").Value.ToString <> "" Then
                        TextBoxFrom.Text = row.Cells("SHELF").Value.ToString.Split("-")(0)
                        TextBoxTo.Text = row.Cells("SHELF").Value.ToString.Split("-")(1)
                    End If
                    btnHapusJadwal.Enabled = True
                End If
                'End If
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub cmbNIK_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cmbNIK.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
End Class