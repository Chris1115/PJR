Imports System.Reflection
Imports System.ComponentModel
Imports PJR.ClsFungsi
Imports System.Threading
Imports PJR.clsFinger
Imports PJR.ClsPJRController
Imports IDM.Fungsi
Imports PJR.ClsHandheldController

Public Class FormMain

    Private ObjConn As New ClsConnection
    Public Shared panggil As New ProcessStartInfo
    Public Shared ckstatus As Boolean = True

    Private wifi As Boolean = ClsInterfaceWifi.initialInterfaceWifi
    Private port As String = ClsInterfaceWifi.getPortWDCP

    Private WithEvents _socketManager As New ShadowMud.Sockets.AsyncSocketController(port)

    Private listClient(9) As ClsClient
    Private maxDevice As Integer = 9

    Private lokasi_so As String = "" 'mode run lokasi scan (toko/gudang)
    Public Shared tabel_name As String = ""
    Private mode_run As String = "" 'mode run SO (baru/edit)
    Public Shared jenis_so As String = "" 'BIC/Tahunan/BPB/AT(Aktiva)
    Public Shared Toko As New ClsToko

    'Memo 1230
    Private FormatBaru1230 As Boolean = False
    Private mainCBRSOIC As Boolean = False
    Private mainTTL3 As Boolean = False
    Private isItemBKL As Boolean = False

    'Memo1314 SOED
    Private mainCBRSOED As Boolean = False

    Private parameter_form As String = ""
    Private parameter_supco As String = ""
    Private parameter_docno As String = ""

    Private parameter_noPO As String = ""
    Public Shared jenis_laporan As String = ""
    Public Shared cmbmodisText As String = ""

    Public Shared namauser As String = ""
    Public Shared isPJR As Boolean = False
    Public Shared norak_pjr As String = ""

    Private DtBPBBKL As New DataTable
    Private DtBPBBKL_DCP As New DataTable
    Private DtBPBNPS As New DataTable

    Private tmpDtLihatSo As New DataTable
    Private DtBPBCabang_Docno As New DataTable

    Private DtBPB_DCP As New DataTable
    Private DtRakSOBIC As New DataTable
    Private DtBazar As New DataTable
    Private DtED As New DataTable

    'revisi Memo No. 208 - CPS - 20
    '13/10/2020
    Private mAktivaList As New List(Of ClsAktiva)
    Private indeks As Integer

    Public Shared isPengganti As Boolean = False
    Public Shared isBazar As Boolean = False
    Public Shared isExpiredDate As Boolean = False
    'revisi lepas flag CBR SO PRODUK KHUSUS
    '09/09/2021
    Public Shared isFlagCBR As Boolean = False

    Private isMulaiScan As Boolean = False
    Private isLTF As Boolean = False
    Private MsgLTF As String = ""
    Private DescBPB As String = ""
    Private NoShelfStr As String = ""
    Private NoRakStr As String = ""

    Private CountNext As Integer = 0
    Private lCabang As New List(Of String)()
    Private lScanBox As New List(Of String)
    'bkl
    Private lSupplier As New List(Of String)()
    Private lScanBKL As New List(Of String)

    Private WorkersSO() As BackgroundWorker
    Private NumWorkersSO = 0
    Private WorkersSkip() As BackgroundWorker
    Private NumWorkersSkip = 0

    Public CmdArg As String() = Nothing

    Public Shared cbHariBukaToko As String = ""

    Dim cUtility As New Utility

    Private Delegate Sub VisiblePanelDelegate(ByVal pnl As Panel, ByVal value As Boolean)
    Private Delegate Sub DisplayGridBPBDelegate(ByVal DtPB As DataTable)
    Private Delegate Sub DisplayGridBPBBKLDelegate(ByVal DtPB As DataTable)
    Private Delegate Sub DisplayGridBPBNPSDelegate(ByVal DtPB As DataTable)

    Public Shared edit_so_nik_pemegang_shift As String = ""
    Public Shared edit_so_nama_pemegang_shift As String = ""

    '240/cps/23
    'docno jadi group_docno
    Public Shared main_groupdocno As Boolean = True
    Private user_aktiva As String = ""

    'LISTING  DISPLAY 1392/CPS/20
    Public Shared isCekDisplay As Boolean = False
    Public Shared user_cekdisplay As String = ""
    Public Shared nama_cekdisplay As String = ""
    Public Shared jabatan_cekdisplay As String = ""

    Private Sub btnRegisPersonil_Click(sender As Object, e As EventArgs) Handles btnRegisPersonil.Click
        isPengganti = False

        Dim cb As New FrmCbBox
        cb.ShowDialog()

        If cbHariBukaToko = "" Then
            MessageBox.Show("Harap Mengisi Jumlah Hari Buka Toko", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Else
            Dim a As New FrmRegistPJR
            a.lblHeaderPJR.Text = "Registrasi Personil PJR"

            a.Show()
        End If
    End Sub

    'Private Sub btnApprovalPJR_Click(sender As Object, e As EventArgs) Handles btnApprovalPJR.Click
    '    'PnlLoading.Visible = True
    '    'LblProses.Text = "Harap tunggu sedang proses data..."
    '    Application.DoEvents()
    '    Dim a As New frmApprovePJR

    '    a.Show()
    '    'PnlLoading.Visible = False
    '    'LblProses.Text = "Loading..."
    '    Application.DoEvents()
    'End Sub

    'Private Sub btnPelaksanaanPJR_Click(sender As Object, e As EventArgs) Handles btnPelaksanaanPJR.Click
    '    Try
    '        Dim cUser As New ClsUserController
    '        Dim cPJR As New ClsPJRController
    '        Dim DtRak As New DataTable
    '        Dim NikToko As String = ""
    '        Dim notif1 As String = ""
    '        Dim notif2 As String = ""
    '        Dim tes As String() = {"a", "b", "c"}

    '        Dim result1 As String = ""
    '        Dim result2 As String = ""

    '        If cUser.CekUserSO = False Then
    '            MessageBox.Show("User toko belum ada", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        Else


    '            notif1 = "Ada Penambahan atau Pengurangan Jumlah Karyawan di Toko Idm. dikarenakan Adanya Mutasi Karyawan, Diharapkan Chief Of Store Melakukan Penjadwalan Ulang PJR !"
    '            notif2 = "Ada Penambahan atau Pengurangan Data Mo.Disp. di Toko Idm., Diharapkan Chief Of Store Melakukan Penjadwalan Ulang PJR !"

    '            If Date.Now.Day = 1 Or Date.Now.Day = 2 Or Date.Now.Day = 3 Or Date.Now.Day = 4 Or Date.Now.Day = 5 Then

    '                result1 = cPJR.notif_cekJadwal_personil()
    '                If result1 = False Then
    '                    MessageBox.Show(notif1, "Notifikasi", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '                    Dim k As New FrmListMutasi_PJR
    '                    k.lblJudul.Text = "LIST PERUBAHAN MUTASI PERSONIL TOKO"
    '                    k.Show()

    '                End If
    '                result2 = cPJR.notif_cekJadwal_Modis()
    '                If result2 = False Then
    '                    MessageBox.Show(notif2, "Notifikasi", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '                    Dim k2 As New FrmListMutasi_PJR
    '                    k2.lblJudul.Text = "LIST PERUBAHAN DATA MODISP TOKO"
    '                    k2.Show()
    '                End If
    '            ElseIf Date.Now.Day > 5 And Date.Now.Day < 16 Then
    '                result1 = cPJR.notif_cekJadwal_personil()
    '                If result1 = False Then
    '                    MessageBox.Show("Maaf, Pelaksanaan PJR Tidak Dapat Dilakukan Dikarenakan Penjadwalan Ulang PJR Belum Dilakukan !", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                    Dim k As New FrmListMutasi_PJR
    '                    k.lblJudul.Text = "LIST PERUBAHAN MUTASI PERSONIL TOKO"
    '                    k.Show()
    '                    result2 = cPJR.notif_cekJadwal_Modis()
    '                    If result2 = False Then
    '                        'MessageBox.Show("Maaf, Pelaksanaan PJR Tidak Dapat Dilakukan Dikarenakan Penjadwalan Ulang PJR Belum Dilakukan !", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                        Dim k2 As New FrmListMutasi_PJR
    '                        k2.lblJudul.Text = "LIST PERUBAHAN DATA MODISP TOKO"
    '                        k2.Show()
    '                        Exit Try

    '                    End If
    '                    Exit Try
    '                Else
    '                    result2 = cPJR.notif_cekJadwal_Modis()
    '                    If result2 = False Then
    '                        MessageBox.Show("Maaf, Pelaksanaan PJR Tidak Dapat Dilakukan Dikarenakan Penjadwalan Ulang PJR Belum Dilakukan !", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                        Dim k2 As New FrmListMutasi_PJR
    '                        k2.lblJudul.Text = "LIST PERUBAHAN DATA MODISP TOKO"
    '                        k2.Show()
    '                        Exit Try
    '                    End If

    '                End If

    '            End If

    '            If Date.Now.Day = 16 Or Date.Now.Day = 17 Or Date.Now.Day = 18 Or Date.Now.Day = 19 Or Date.Now.Day = 20 Then
    '                result1 = cPJR.notif_cekJadwal_personil()
    '                If result1 = False Then
    '                    MessageBox.Show(notif1, "Notifikasi", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '                    Dim k As New FrmListMutasi_PJR
    '                    k.lblJudul.Text = "LIST PERUBAHAN MUTASI PERSONIL TOKO"
    '                    k.Show()
    '                End If
    '                result2 = cPJR.notif_cekJadwal_Modis()
    '                If result2 = False Then
    '                    MessageBox.Show(notif2, "Notifikasi", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '                    Dim k2 As New FrmListMutasi_PJR
    '                    k2.lblJudul.Text = "LIST PERUBAHAN DATA MODISP TOKO"
    '                    k2.Show()
    '                End If
    '            ElseIf Date.Now.Day > 20 Then
    '                result1 = cPJR.notif_cekJadwal_personil()
    '                If result1 = False Then
    '                    MessageBox.Show("Maaf, Pelaksanaan PJR Tidak Dapat Dilakukan Dikarenakan Penjadwalan Ulang PJR Belum Dilakukan !", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                    Dim k As New FrmListMutasi_PJR
    '                    k.lblJudul.Text = "LIST PERUBAHAN MUTASI PERSONIL TOKO"
    '                    k.Show()

    '                    result2 = cPJR.notif_cekJadwal_Modis()
    '                    If result2 = False Then
    '                        'MessageBox.Show("Maaf, Pelaksanaan PJR Tidak Dapat Dilakukan Dikarenakan Penjadwalan Ulang PJR Belum Dilakukan !", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                        Dim k2 As New FrmListMutasi_PJR
    '                        k2.lblJudul.Text = "LIST PERUBAHAN DATA MODISP TOKO"
    '                        k2.Show()
    '                        Exit Try
    '                    End If
    '                    Exit Try
    '                Else
    '                    result2 = cPJR.notif_cekJadwal_Modis()
    '                    If result2 = False Then
    '                        MessageBox.Show("Maaf, Pelaksanaan PJR Tidak Dapat Dilakukan Dikarenakan Penjadwalan Ulang PJR Belum Dilakukan !", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                        Dim k2 As New FrmListMutasi_PJR
    '                        k2.lblJudul.Text = "LIST PERUBAHAN DATA MODISP TOKO"
    '                        k2.Show()
    '                        Exit Try
    '                    End If

    '                End If

    '            End If

    '            tes = Panggil_CekFingerprintV3("frmSO", "WDCP_PJR")
    '            If tes(0) = "" And tes(1) = "" And tes(2) = "" Then
    '                MsgBox("Maaf, Validasi scanfinger tidak berhasil !", MsgBoxStyle.Exclamation)
    '                'btnOk.Enabled = False
    '                'SettingDisplay("SO")

    '                Debug.WriteLine("Password Salah")
    '            Else
    '                ConstNIKPJR(tes(2).Split("|")(0).Trim)

    '                'tabel_name = "CekPJR"
    '                'jenis_so = "CekPJR"
    '                'SettingDisplay("CekPJR")
    '                cPJR.CekTablePJR()

    '                NikToko = cPJR.getConstNIKPJR
    '                cPJR.ReloadJadwal(True)
    '                cPJR.insertTempJadwalPJR(NikToko)

    '                DtRak = cPJR.GetNamaRak(Date.Now.ToString, NikToko)
    '                'cmbModis.Items.Clear()

    '                If DtRak.Rows.Count > 0 Then
    '                    For Each Dr As DataRow In DtRak.Rows
    '                        'cmbModis.Items.Add(Dr(0))
    '                    Next
    '                Else
    '                    MsgBox("Maaf, jadwal PJR Anda belum terdaftar atau belum di-Approve !", MsgBoxStyle.Exclamation)
    '                    'btnOk.Enabled = False
    '                    'SettingDisplay("SO")
    '                End If
    '            End If

    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message & vbCrLf & ex.StackTrace)
    '    End Try
    'End Sub

    'Private Sub btnTIndakLanjutLBTD_Click(sender As Object, e As EventArgs) Handles btnTIndakLanjutLBTD.Click
    '    Try
    '        Dim cUser As New ClsUserController
    '        Dim cPJR As New ClsPJRController
    '        Dim DtRak As New DataTable
    '        Dim rakpjr As String = ""
    '        Dim tes As String() = {"a", "b", "c"}
    '        Dim NikToko As String = ""
    '        If cUser.CekUserSO = False Then
    '            MessageBox.Show("User toko belum ada", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        Else
    '            tes = Panggil_CekFingerprintV3("frm", "WDCP_PJR")

    '            If tes(0) = "" And tes(1) = "" And tes(2) = "" Then
    '                MsgBox("Maaf, Validasi scanfinger tidak berhasil !", MsgBoxStyle.Exclamation)
    '                'btnOk.Enabled = False
    '                'SettingDisplay("SO")

    '                Debug.WriteLine("Password Salah")
    '            Else
    '                ConstNIKPJR(tes(2).Split("|")(0).Trim)

    '                'tabel_name = "TINDAKLBTD"
    '                'jenis_so = "TINDAKLBTD"
    '                'SettingDisplay("TINDAKLBTD")
    '                cPJR.CekTableTINDAKLBTD()
    '                rakpjr = cPJR.getConstRakPJR
    '                NikToko = cPJR.getConstNIKPJR

    '                DtRak = cPJR.GetNamaRakLBTD(Date.Now.ToString("yyyy-MM-dd"), NikToko, "", rakpjr)
    '                'cmbModis.Items.Clear()
    '                'cmbModis.Enabled = True
    '                'If DtRak.Rows.Count > 0 Then
    '                '    For Each Dr As DataRow In DtRak.Rows
    '                '        cmbModis.Items.Add(Dr(0))
    '                '    Next

    '                '    InitialClient()

    '                '    btnOk.Enabled = True
    '                'Else
    '                '    btnOk.Enabled = False
    '                'End If
    '            End If

    '        End If


    '    Catch ex As Exception
    '        MsgBox(ex.Message & vbCrLf & ex.StackTrace)
    '    End Try
    'End Sub

    'Private Sub btnTIndakLBTDBAAS_Click(sender As Object, e As EventArgs) Handles btnTIndakLBTDBAAS.Click
    '    Try
    '        Dim confirm As DialogResult

    '        Dim cUser As New ClsUserController
    '        Dim cPJR As New ClsPJRController
    '        Dim DtRak As New DataTable
    '        Dim tanggal As String
    '        Dim rakpjr As String = ""
    '        Dim norak As String = ""
    '        Dim nik As String
    '        Dim tes As String() = {"a", "b", "c"}
    '        Dim NikToko As String = ""
    '        If cUser.CekUserSO = False Then
    '            MessageBox.Show("User toko belum ada", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        Else
    '            Dim DtSO As New DataTable
    '            Dim cProduk As New ClsProdukController

    '            confirm = MessageBox.Show("Akan mulai Scan Finger AS/AM, Silahkan klik Yes utk melanjutkan.", "Tindak LBTD BA PJR", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

    '            If confirm = Windows.Forms.DialogResult.Yes Then

    '                tes = Panggil_CekFingerprintV3("frm", "WDCP")

    '                If tes(0) = "" And tes(1) = "" And tes(2) = "" Then
    '                    MsgBox("Maaf, Validasi scanfinger AS/AM tidak berhasil !", MsgBoxStyle.Exclamation)
    '                    'btnOk.Enabled = False
    '                    'SettingDisplay("SO")

    '                    Debug.WriteLine("Password Salah")
    '                Else
    '                    nik = cPJR.getConstNIKPJR
    '                    rakpjr = cPJR.getConstRakPJR.Split("-")(0).Trim
    '                    tanggal = cPJR.getConstRakPJR.Split("-")(1).Trim
    '                    norak = cPJR.getConstRakPJR.Split("-")(2).Trim

    '                    cPJR.updateITT_ADJUST(tanggal.Replace("/", "-"), nik, rakpjr, norak)

    '                    Dim Frm As New frmTampilLBTD
    '                    frmTampilLBTD.jenis = "TINDAKLBTD_BAPJR"
    '                    frmTampilLBTD.nik_as = "user"
    '                    Frm.Show()

    '                    'tabel_name = "TINDAKLBTD_BAPJR"
    '                    'jenis_so = "TINDAKLBTD_BAPJR"
    '                    'SettingDisplay("TINDAKLBTD_BAPJR")
    '                    cPJR.CekTableTINDAKLBTD_BAPJR()
    '                    'cmbModis.Items.Clear()
    '                    'cmbModis.Enabled = True

    '                    'cmbModis.Items.Add("AKUMULASI ITT")
    '                    Label6.Text = "Jumlah"
    '                    Label4.Text = "Tanggal"

    '                    'txtNamaModis.Text = cPJR.GetTanggalAkumulasiLBTD_BAPJR
    '                    'If txtNamaModis.Text <> "" Then
    '                    '    InitialClient()
    '                    '    btnOk.Enabled = True
    '                    'Else
    '                    '    btnOk.Enabled = False
    '                    'End If
    '                End If


    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message & vbCrLf & ex.StackTrace)
    '    End Try
    'End Sub

    'Private Sub btnCetakLaporanPJR_Click(sender As Object, e As EventArgs) Handles btnCetakLaporanPJR.Click
    '    Dim Frm As New frmCPJR
    '    Frm.ShowDialog()
    'End Sub

End Class
