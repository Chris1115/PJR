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

    Private Sub FormMain_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        _socketManager.StopServer()
        Try
            Dim processName As String = "interfaceWIFI"
            Dim processes() As Process = Process.GetProcessesByName(processName)
            For Each proc As Process In processes
                proc.Kill()
            Next
        Catch ex As Exception

        End Try
    End Sub

#Region "Socket"
    Private Sub _socketManager_onConnectionAccept(ByVal SocketID As String, ByVal IpAddress As String) Handles _socketManager.onConnectionAccept
        Try
            Dim mCount As Integer = 0
            Dim mPosition As Integer = 0
            Dim Log As String = ""

            Dim util As New Utility
            util.TraceLogTxt("_socketManager_onConnectionAccept " & vbCrLf & "IP:" & IpAddress & " | SocketId:" & SocketID & " | ListClient:" & listClient.Length.ToString)

            For i As Integer = 0 To listClient.Length - 1
                If Not IsNothing(listClient(i)) Then
                    mCount += 1
                End If
            Next
            For i As Integer = 0 To listClient.Length - 1
                If Not IsNothing(listClient(i)) Then
                    If listClient(i).IpAddress.Trim = IpAddress.Split(":")(0).Trim Then
                        mPosition = i
                        listClient(mPosition).SocketID = SocketID
                        mainWindow.Invoke(ShowIpAddress, New Object() {mPosition, listClient(mPosition).IpAddress})
                        delaytime(100)
                        util.TraceLogTxt("_socketManager_onConnectionAccept " & vbCrLf & listClient(i).IpAddress.Trim & " = " & IpAddress.Split(":")(0).Trim & "IP:" & listClient(mPosition).IpAddress & " | SocketId:" & listClient(mPosition).SocketID & " | ListClient:" & listClient.Length.ToString)
                        Exit Sub
                    Else
                        Log &= "Cek: " & listClient(i).IpAddress.Trim & "=" & IpAddress.Split(":")(0).Trim & "| i: " & i & vbCrLf
                    End If
                Else
                    mPosition = i
                    Log &= "IsNothing(" & i & "): " & IsNothing(listClient(i)) & "| i: " & i & vbCrLf
                    Exit For
                End If
            Next

            If mCount > maxDevice Then
                _socketManager.Close(SocketID)
                Log &= "_socketManager.Close| SocketID: " & SocketID & vbCrLf
            Else
                Dim client As New ClsClient
                client.SocketID = SocketID
                client.IpAddress = IpAddress.Split(":")(0)
                client.Login = New ClsLogin
                client.Login.User = New ClsUser
                client.SO = New ClsSo

                listClient(mPosition) = client

                If jenis_so.ToLower = "tahunan" And mode_run = "E" Then
                    'Mencatat Device pada Tabel
                    Dim clsHandheld As New ClsHandheld
                    Dim clsHandheldController As New ClsHandheldController
                    clsHandheld.ipAddress = IpAddress.Split(":")(0)
                    clsHandheld.socketID = SocketID
                    clsHandheldController.addDevice(clsHandheld)
                End If

                mainWindow.Invoke(ShowIpAddress, New Object() {mPosition, client.IpAddress})
                Log &= "NewClientUser| SocketID: " & client.SocketID & " | IpAddress: " & client.IpAddress & vbCrLf
                util.TraceLogTxt("_socketManager_onConnectionAccept " & vbCrLf & Log)
            End If
            delaytime(100)
        Catch ex As Exception
            Dim util As New Utility
            util.TraceLogTxt("_socketManager_onConnectionAccept " & ex.Message & vbCrLf & ex.StackTrace & " IP:" & IpAddress)
        End Try

    End Sub

    Private Sub _socketManager_onSocketDisconnected(ByVal SocketID As String, ByVal IpAddress As String) Handles _socketManager.onSocketDisconnected
        Dim util As New Utility
        util.TraceLogTxt("onSocketDisconnected " & vbCrLf & "IP:" & IpAddress)
    End Sub

    Private Sub _socketManager_on(ByVal SocketID As String, ByVal SocketData As String, ByVal IpAddress As String) Handles _socketManager.onDataArrival
        Try
            Dim Util As New Utility
            Dim IdxArrival As Integer = 0
            Dim ClientArrival As New ClsClient
            Dim Server_Display As String = ""
            Dim HH_Display As String = ""
            Dim Log As String = ""

            Console.WriteLine("Socket Data: " & SocketData)

            Log = "onDataArrival (START) " & vbCrLf & "IP:" & IpAddress & " | SocketId:" & SocketID & " | SocketData:" & SocketData
            Log &= vbCrLf & "1" & vbCrLf
            For Each mClient As ClsClient In listClient
                If Not IsNothing(mClient) Then
                    Log &= mClient.IpAddress & " | " & mClient.SocketID & vbCrLf
                End If
            Next

            Dim Idx As Integer = 0
            For Each mClient As ClsClient In listClient
                If Not IsNothing(mClient) Then
                    If mClient.SocketID = SocketID Then
                        ClientArrival = mClient
                        IdxArrival = Idx

                        Log &= "===================== " & IdxArrival & " =============================" & vbCrLf
                        Log &= "IPAddress: " & ClientArrival.IpAddress & " - " & ClientArrival.SocketID & vbCrLf
                        Log &= "===================== " & Idx & " =============================" & vbCrLf

                        Exit For
                    End If
                End If
                Idx += 1
            Next
            delaytime(100)

            If SocketData.ToString = "LOGIN" Then
                'Jika jenis SO = Aktiva/Monitoring Price Tag, maka menu handheld langsung siap scan (tidak perlu login)
                If jenis_so = "AT" Then
                    HH_Display = TampilDeskripsiClient("", "", "", "", "")
                    Server_Display = TampilDeskripsiServer("", "", "", "", "")
                ElseIf jenis_so = "MPC" Then 'Revisi Memo No 296/CPS/23 Monitoring Price Tag by Kukuh 16 Mei 2023
                    HH_Display = TampilMonitoringPriceTagClient("", "", "1")
                    Server_Display = TampilMonitoringPriceTagServer("", "", "1")
                Else
                    HH_Display = TampilLoginClient("", "", "", "")
                    Server_Display = TampilLoginServer("", "", "")
                End If

            ElseIf Strings.Right(SocketData, 1) = "U" Then 'Jika sesudah memasukkan username
                Dim mLogin As ClsLogin
                Dim cLogin As New ClsUserController
                Dim Status As String = ""

                ClientArrival.Login.User = New ClsUser
                ClientArrival.Login.User.ID = SocketData.Substring(0, SocketData.Length - 1)

                mLogin = cLogin.GetUserLogin(ClientArrival.Login.User.ID, ClientArrival.Login.User.Password, ClientArrival.IpAddress, jenis_so)
                If Not IsNothing(mLogin.User) Then
                    If mLogin.User.Status = "5" Then
                        Status = mLogin.User.Status
                    End If
                End If

                HH_Display = TampilLoginClient(ClientArrival.Login.User.ID, "", Status, "")
                Server_Display = TampilLoginServer(ClientArrival.Login.User.ID, "", Status)

            ElseIf (Strings.Right(SocketData, 1) = "P") Then 'Jika sesudah memasukkan password
                Dim mLogin As ClsLogin
                Dim cLogin As New ClsUserController

                ClientArrival.Login.User.Password = SocketData.Substring(0, SocketData.Length - 1)

                mLogin = cLogin.GetUserLogin(ClientArrival.Login.User.ID, ClientArrival.Login.User.Password, ClientArrival.IpAddress, jenis_so)
                ClientArrival.Login = mLogin
                If mLogin.Status = "1" Then 'Login Sukses
                    If tabel_name.StartsWith("dcp_boxplu") Then
                        If IsNothing(ClientArrival.SO.NoContainer) Then ClientArrival.SO.NoContainer = ""
                        'Sukses Login Terima Barang
                        If ClientArrival.SO.NoContainer = "" Then 'Jika box terima barang belum di setting maka tampil pilhan lokasi
                            HH_Display = TampilContainerClient(ClientArrival.SO.NoContainer, DescBPB)
                            Server_Display = TampilContainerServer(False, ClientArrival.SO.NoContainer, "", DescBPB)
                        Else 'jika sudah ada lokasi maka langsung proses scan
                            ClientArrival.SO = New ClsSo
                            HH_Display = TampilDeskripsiClient("", "", "", "", "")
                            Server_Display = TampilDeskripsiServer("", "", "", "", "")
                        End If
                    ElseIf tabel_name.StartsWith("CekPlanogram") Then
                        'Sukses Login Planogram
                        HH_Display = TampilPlanoClient("", "", "", "")
                        Server_Display = TampilPlanoServer("", "", "", "")
                    ElseIf tabel_name.StartsWith("CekKesegaran") Then
                        'Sukses Login Planogram
                        HH_Display = TampilKesegaranClient("", "", "", "")
                        Server_Display = TampilKesegaranServer("", "", "", "")

                    ElseIf tabel_name.StartsWith("BPBBKL_WDCP") Then
                        'Sukses Login Terima Barang
                        namauser = mLogin.User.Nama
                        ClientArrival.SO = New ClsSo
                        HH_Display = TampilDeskripsiBKL("", "", "", "", "", "", "", "", "")
                        Server_Display = TampilDeskripsiServer("", "", "", "", "")

                    ElseIf tabel_name.ToLower.StartsWith("bpbnps_wdcp") Then
                        'Sukses Login Terima Barang
                        namauser = mLogin.User.Nama
                        ClientArrival.SO = New ClsSo
                        HH_Display = TampilDeskripsiNPS("", "", "", "", "", "", "", "", "")
                        Server_Display = TampilDeskripsiServer("", "", "", "", "")

                    ElseIf tabel_name.StartsWith("ptag_wdcp") Then
                        'Sukses Login Planogram
                        HH_Display = TampilPriceTagClient("", "", "", "")
                        Server_Display = TampilPriceTagServer("", "", "", "")

                    ElseIf tabel_name.StartsWith("SP") Then
                        ClientArrival.SO = New ClsSo
                        HH_Display = TampilDeskripsiClient("", "", "", "", "")
                        Server_Display = TampilDeskripsiServer("", "", "", "", "")
                        lokasi_so = "Toko"
                        If isMulaiScan = False Then
                            'Jika sudah ada yang login(pertama) maka menampilkan menu SO selesai dan cek SO
                            isMulaiScan = True
                        End If

                    ElseIf tabel_name.StartsWith("SZ") Then 'SO Bazar
                        HH_Display = TampilDeskripsiBazar("", "", "", "", "", "", "", "")
                        Server_Display = TampilDeskripsiServer("", "", "", "", "")
                        lokasi_so = "Toko"

                    ElseIf tabel_name.StartsWith("SE") Then 'SO Expired Date
                        HH_Display = TampilDeskripsiExpired("", "", "", "", "", "")
                        Server_Display = TampilDeskripsiExpiredServer("", "", "", "", "", "")
                        lokasi_so = "Toko"

                    ElseIf tabel_name.StartsWith("CekPJR") Or tabel_name.StartsWith("TINDAKLBTD") Or tabel_name.StartsWith("TINDAKLBTD_BAPJR") Then
                        'Sukses Login Planogram
                        HH_Display = TampilPJRClient("", "", "", "")
                        Server_Display = TampilPJRServer("", "", "", "")

                        '1392/CPS/20
                    ElseIf tabel_name.StartsWith("CekDisplay") Then
                        Dim cLDP As New ClsCekDisplayController
                        user_cekdisplay = ClientArrival.Login.User.ID
                        nama_cekdisplay = ClientArrival.Login.User.Nama
                        jabatan_cekdisplay = cLDP.GetJabatan(user_cekdisplay)
                        'Sukses Login Planogram
                        HH_Display = TampilCekDisplayClient("", "", "")
                        Server_Display = TampilCekDisplayServer("", "", "")

                    ElseIf tabel_name.StartsWith("SN") Then
                        'Untuk Setiap device dapat memilih lokasinya masing-masing
                        If (mode_run = "B" And lokasi_so = "") Or mode_run = "E" Then 'Untuk Setiap device dapat memilih lokasinya masing-masing
                            HH_Display = TampilLokasiClient()
                            Server_Display = TampilLokasiServer()
                        Else 'jika sudah ada lokasi maka langsung proses scan
                            ClientArrival.SO = New ClsSo
                            Dim clsHandheld As New ClsHandheld
                            Dim clsHandheldController As New ClsHandheldController

                            clsHandheld.lokasi_so = lokasi_so
                            clsHandheld.ipAddress = IpAddress.Split(":")(0)
                            clsHandheld.socketID = SocketID
                            clsHandheld.jenis_so = "tahunan"
                            clsHandheldController.addSO(clsHandheld)

                            HH_Display = TampilDeskripsiClient("", "", "", "", "")
                            Server_Display = TampilDeskripsiServer("", "", "", "", "")
                        End If
                    Else
                        'Sukses login SO
                        If lokasi_so = "" Then 'Jika lokasi SO belum di setting maka tampil pilhan lokasi
                            HH_Display = TampilLokasiClient()
                            Server_Display = TampilLokasiServer()
                        Else 'jika sudah ada lokasi maka langsung proses scan
                            ClientArrival.SO = New ClsSo
                            HH_Display = TampilDeskripsiClient("", "", "", "", "")
                            Server_Display = TampilDeskripsiServer("", "", "", "", "")
                        End If
                    End If
                ElseIf mLogin.Status = "2" Then 'ID tdk terdaftar 
                    HH_Display = TampilLoginClient("", "", mLogin.Status, "")
                    Server_Display = TampilLoginServer("", "", mLogin.Status)
                ElseIf mLogin.Status = "3" Then 'IP sudah dipakai
                    HH_Display = TampilLoginClient("", "", mLogin.Status, "")
                    Server_Display = TampilLoginServer("", "", mLogin.Status)
                ElseIf mLogin.Status = "4" Then 'Password salah
                    HH_Display = TampilLoginClient("", "", mLogin.Status, "")
                    Server_Display = TampilLoginServer("", "", mLogin.Status)
                End If

            ElseIf (Strings.Right(SocketData, 1) = "X") Then 'Jika memasukkan UPDATE password baru
                Dim cLogin As New ClsUserController

                ClientArrival.Login.User.Password = SocketData.Substring(0, SocketData.Length - 1)
                If cLogin.UpdatePassUserSO(ClientArrival.Login.User.ID, ClientArrival.Login.User.Password) Then
                    ClientArrival.Login.User.Status = "6"
                Else
                    ClientArrival.Login.User.Status = "7"
                End If
                HH_Display = TampilLoginClient(ClientArrival.Login.User.ID, "", ClientArrival.Login.User.Status, "")
                Server_Display = TampilLoginServer(ClientArrival.Login.User.ID, "", ClientArrival.Login.User.Status)

            ElseIf Strings.Right(SocketData, 1) = "L" Then 'jika memasukkan pilihan lokasi '1' atau '2' atau '3'
                Dim cUser As New ClsUserController
                ClientArrival.SO = New ClsSo
                Dim clsHandheld As New ClsHandheld
                Dim clsHandheldController As New ClsHandheldController

                'Mencatat SO pada Tabel
                If Strings.Left(SocketData, 1) = "1" Then '1 = SO TOKO
                    lokasi_so = "Toko"
                    clsHandheld.lokasi_so = "Toko"
                ElseIf Strings.Left(SocketData, 1) = "2" Then '2 = SO GUDANG
                    lokasi_so = "Gudang"
                    clsHandheld.lokasi_so = "Gudang"
                ElseIf Strings.Left(SocketData, 1) = "3" Then '3 = SO BARANG RUSAK
                    lokasi_so = "Barang Rusak"
                    clsHandheld.lokasi_so = "Barang Rusak"
                End If

                If jenis_so.ToLower = "tahunan" Then
                    clsHandheld.ipAddress = IpAddress.Split(":")(0)
                    clsHandheld.socketID = SocketID
                    clsHandheld.jenis_so = "tahunan"
                    clsHandheldController.addSO(clsHandheld)

                    If lokasi_so = "Toko" Then
                        If Not cUser.IsTutupHarian() Then
                            HH_Display = TampilLokasiClient("Blm closing harian!")
                            Server_Display = TampilLokasiServer("Blm closing harian!")
                        Else
                            HH_Display = TampilDeskripsiClient("", "", "", "", "")
                            Server_Display = TampilDeskripsiServer("", "", "", "", "")

                            If isMulaiScan = False Then
                                'Jika sudah ada yang login(pertama) maka menampilkan menu SO selesai dan cek SO
                                isMulaiScan = True
                            End If
                        End If
                    Else
                        HH_Display = TampilDeskripsiClient("", "", "", "", "")
                        Server_Display = TampilDeskripsiServer("", "", "", "", "")
                        If isMulaiScan = False Then
                            'Jika sudah ada yang login(pertama) maka menampilkan menu SO selesai dan cek SO
                            isMulaiScan = True
                        End If
                    End If

                    If mode_run = "E" Then
                        lokasi_so = ""
                    End If

                Else
                    HH_Display = TampilDeskripsiClient("", "", "", "", "")
                    Server_Display = TampilDeskripsiServer("", "", "", "", "")
                    If isMulaiScan = False Then
                        'Jika sudah ada yang login(pertama) maka menampilkan menu SO selesai dan cek SO

                        isMulaiScan = True
                    End If
                End If

            ElseIf Strings.Right(SocketData, 1) = "S" Or Strings.Right(SocketData, 1) = "B" Then 'Jika melakukan scan barcode
                Dim mSo As New ClsSo
                Dim mBPB As New ClsBPB
                Dim mBKL As New ClsBPBBKL
                Dim mNPS As New ClsBPBNPS
                Dim mPlano As New ClsPlanogram
                Dim mAktiva As New ClsAktiva
                Dim cSo As New ClsProdukController
                Dim cBPB As New ClsBPBController
                Dim cPlano As New ClsPlanoController
                Dim cAktiva As New ClsAktivaController
                Dim mPriceTag As New ClsPriceTag
                Dim cPriceTag As New ClsPriceTagController
                Dim mMonitoringPriceTag As New ClsMonitoringPriceTag 'UPDATE MEMO 296/CPS/23 by Kukuh
                Dim cMonitoringPriceTag As New ClsMonitoringPriceTagController 'UPDATE MEMO 296/CPS/23 by Kukuh
                Dim mKesegaran As New ClsKesegaran
                Dim cKesegaran As New ClsKesegaranController
                Dim mPJR As New ClsPJRProduk
                Dim cPJR As New ClsPJRController

                Dim cCekDisplay As New ClsCekDisplayController
                Dim mCekDisplay As New ClsCekDisplay

                Dim mED As New ClsSOED
                Dim cED As New ClsSOEDController

                ClientArrival.SO.BarcodePlu = SocketData.Substring(0, SocketData.Length - 1)
                If ClientArrival.SO.BarcodePlu.Length = 12 Then
                    ClientArrival.SO.BarcodePlu = "0" & ClientArrival.SO.BarcodePlu
                End If

                If tabel_name = "dcp_boxplu" Then
                    If ClientArrival.SO.BarcodePlu = "999" Then 'Input barcode container
                        ClientArrival.SO.NoContainer = ""
                        DescBPB = ""
                        DtBPB_DCP = cBPB.GetTableCekPB("")
                        DisplayGridBPB(DtBPB_DCP)
                        HH_Display = TampilContainerClient(ClientArrival.SO.NoContainer, DescBPB)
                        Server_Display = TampilContainerServer(False, ClientArrival.SO.NoContainer, "", DescBPB)
                    Else
                        Dim i As Integer = 0
                        'mBPB = cSo.GetDeskripsiProdukBPB(tabel_name, container_no, client.SO.BarcodePlu)
                        mBPB = cBPB.CekPluBarang(ClientArrival.SO.BarcodePlu, ClientArrival.SO.NoContainer, CbxKodeGudang.SelectedValue)
                        If mBPB.Desc.Trim <> "" Then
                            Dim DrRemove As DataRow = Nothing
                            Dim AddBox As Boolean = True
                            For Each Dr As DataRow In DtBPB_DCP.Rows
                                If Dr("PRDCD").ToString.Trim = mBPB.Prdcd Then
                                    DrRemove = Dr
                                    'Add list string dus_no
                                    For Each Box As String In lScanBox
                                        If Box.Trim = Dr("Dus_NO").ToString.Trim Then
                                            AddBox = False
                                            Exit For
                                        End If
                                    Next
                                    If AddBox Then
                                        lScanBox.Add(Dr("Dus_NO").ToString.Trim)
                                    End If
                                End If
                            Next
                            If Not IsNothing(DrRemove) Then
                                DtBPB_DCP.Rows.Remove(DrRemove)
                                DtBPB_DCP.AcceptChanges()
                                DisplayGridBPB(DtBPB_DCP)
                            End If
                        End If
                        HH_Display = TampilDeskripsiClient(mBPB.Prdcd, mBPB.Desc, "", mBPB.Qty, "")
                        Server_Display = TampilDeskripsiServer(mBPB.Prdcd, mBPB.Desc, "", mBPB.Qty, "")
                    End If
                ElseIf tabel_name = "BPBBKL_WDCP" Then 'Input barcode BKL
                    If ClientArrival.SO.BarcodePlu = "999" Then 'Input barcode container
                        ClientArrival.SO.NoContainer = ""
                        DescBPB = ""
                        DtBPBBKL_DCP = cBPB.GetTableCekPBBKL("", "")
                        DisplayGridBPBBKL(DtBPBBKL_DCP)
                        HH_Display = TampilContainerBKLClient(ClientArrival.SO.NoContainer, DescBPB)
                        Server_Display = TampilContainerBKLServer(False, ClientArrival.SO.NoContainer, "", DescBPB)
                    Else
                        Dim i As Integer = 0
                        mBKL = cBPB.CekPluBarangBKL(ClientArrival.SO.BarcodePlu, CbxKodeGudang.SelectedValue)
                        If mBKL.BKL.Desc.Trim <> "" Then
                            Dim DrRemove As DataRow = Nothing
                            Dim AddBox As Boolean = True
                            For Each Dr As DataRow In DtBPB_DCP.Rows
                                If Dr("PRDCD").ToString.Trim = mBKL.BKL.Prdcd Then
                                    DrRemove = Dr
                                    'Add list string dus_no
                                    For Each Box As String In lScanBox
                                        If Box.Trim = Dr("Dus_NO").ToString.Trim Then
                                            AddBox = False
                                            Exit For
                                        End If
                                    Next
                                    If AddBox Then
                                        lScanBKL.Add(Dr("Dus_NO").ToString.Trim)
                                    End If
                                End If
                            Next
                            If Not IsNothing(DrRemove) Then
                                DtBPBBKL_DCP.Rows.Remove(DrRemove)
                                DtBPBBKL_DCP.AcceptChanges()
                                DisplayGridBPBBKL(DtBPBBKL_DCP)
                            End If
                        End If
                        ClientArrival.SO.Tgl_exp = ""
                        If mBKL.StatusDesc = "1" Then
                            HH_Display = TampilDeskripsiBKL(mBKL.BKL.Prdcd, mBKL.BKL.Desc, "", "", "", mBKL.StatusDesc, "", "")
                            Server_Display = TampilDeskripsiServer(mBKL.BKL.Prdcd, mBKL.BKL.Desc, "", "", "")
                        Else
                            HH_Display = TampilDeskripsiBKL(mBKL.BKL.Prdcd, mBKL.BKL.Desc, "", "", "", "", "", "",, mBKL.BKL.fraction_pcs)
                            Server_Display = TampilDeskripsiServer(mBKL.BKL.Prdcd, mBKL.BKL.Desc, "", "", "", mBKL.BKL.fraction_pcs)
                        End If

                    End If
                ElseIf tabel_name.ToLower = "bpbnps_wdcp" Then 'Input barcode NPS

                    Dim i As Integer = 0
                    mNPS = cBPB.CekPluBarangNPS(ClientArrival.SO.BarcodePlu, parameter_noPO)
                    ClientArrival.SO.Tgl_exp = ""
                    If mNPS.StatusDesc = "1" Then
                        HH_Display = TampilDeskripsiNPS(mNPS.NPS.Prdcd, mNPS.NPS.Desc, "", "", "", mNPS.StatusDesc, "", "")
                        Server_Display = TampilDeskripsiServer(mNPS.NPS.Prdcd, mNPS.NPS.Desc, "", "", "")

                    Else
                        HH_Display = TampilDeskripsiNPS(mNPS.NPS.Prdcd, mNPS.NPS.Desc, "", "", "", "", "", "")
                        Server_Display = TampilDeskripsiServer(mNPS.NPS.Prdcd, mNPS.NPS.Desc, "", "", "")
                    End If

                ElseIf tabel_name = "CekPlanogram" Then 'Input barcode planogram
                    mPlano = cPlano.GetDeskripsiPlanogram(tabel_name, ClientArrival.SO.BarcodePlu, cmbModis.Text, NoShelfStr, ClientArrival.Login.User)
                    HH_Display = TampilPlanoClient(mPlano.Prdcd, mPlano.Desc, mPlano.MaxRet, mPlano.Price)
                    Server_Display = TampilPlanoServer(mPlano.Prdcd, mPlano.Desc, mPlano.MaxRet, mPlano.Price)

                ElseIf tabel_name = "CekPJR" Or tabel_name = "TINDAKLBTD" Then
                    mPJR = cPJR.GetDeskripsiPJR(tabel_name, ClientArrival.SO.BarcodePlu, cmbModis.Text.Split("-")(0).Trim, NoShelfStr, ClientArrival.Login.User)
                    HH_Display = TampilPJRClient(ClientArrival.SO.BarcodePlu, mPJR.Desc, mPJR.MaxRet, mPJR.Price)
                    Server_Display = TampilPJRServer(ClientArrival.SO.BarcodePlu, mPJR.Desc, mPJR.MaxRet, mPJR.Price)

                ElseIf tabel_name = "TINDAKLBTD_BAPJR" Then
                    mPJR = cPJR.GetDeskripsiPJR_LBTD_BA_PJR(tabel_name, ClientArrival.SO.BarcodePlu, ClientArrival.Login.User)
                    HH_Display = TampilPJRClient(ClientArrival.SO.BarcodePlu, mPJR.Desc, mPJR.MaxRet, mPJR.Price)
                    Server_Display = TampilPJRServer(ClientArrival.SO.BarcodePlu, mPJR.Desc, mPJR.MaxRet, mPJR.Price)

                ElseIf tabel_name = "CekKesegaran" Then 'Input barcode kesegaran
                    mKesegaran = cKesegaran.GetDeskripsiKesegaran(tabel_name, ClientArrival.SO.BarcodePlu, cmbModis.Text, ClientArrival.Login.User, "")
                    HH_Display = TampilKesegaranClient(mKesegaran.Prdcd, mKesegaran.Desc, mKesegaran.MaxRet, "")
                    Server_Display = TampilKesegaranServer(mKesegaran.Prdcd, mKesegaran.Desc, mKesegaran.MaxRet, "")

                ElseIf tabel_name.ToUpper.Contains("OA") Then 'Input barcode Aktiva
                    'revisi Memo No. 208 - CPS - 20
                    '13/10/2020
                    If Strings.Right(SocketData, 1) = "B" Then
                        If ClientArrival.SO.BarcodePlu.Length >= 6 Then
                            indeks = 0
                            mAktivaList = cAktiva.GetListAktiva(tabel_name, ClientArrival.SO.BarcodePlu, Toko)
                            If mAktivaList.Count = 0 Then
                                HH_Display = TampilDeskripsiClient("", "Tidak Ditemukan", "", "", "")
                                Server_Display = TampilDeskripsiServer("", "Tidak Ditemukan", "", "", "")
                            Else
                                'DATA ADA
                                HH_Display = TampilDeskripsiListClient(mAktivaList, indeks)
                                Server_Display = TampilDeskripsiListServer(mAktivaList, indeks)
                                'revisi Memo No. 208 - CPS - 20
                                '13/10/2020
                                ClientArrival.SO.statusBarcode = "I"
                            End If
                        Else
                            HH_Display = TampilDeskripsiClient("", "Tidak Ditemukan", "", "", "")
                            Server_Display = TampilDeskripsiServer("", "Tidak Ditemukan", "", "", "")
                        End If
                    Else
                        mAktiva = cAktiva.GetDeskripsiAktiva(tabel_name, ClientArrival.SO.BarcodePlu, Toko)
                        'Set deskripsi & max qty Aktiva
                        ClientArrival.SO.Deskripsi = mAktiva.Deskripsi
                        If mAktiva.Deskripsi2.Contains("AT baru") Then
                            ClientArrival.SO.Deskripsi = mAktiva.Deskripsi2
                        End If
                        ClientArrival.SO.QTYCom = mAktiva.QtyMax
                        'revisi Memo No. 208 - CPS - 20
                        '13/10/2020
                        ClientArrival.SO.statusBarcode = "S"
                        HH_Display = TampilDeskripsiClient(mAktiva.NSeri, mAktiva.Deskripsi, "", "", "", mAktiva.Deskripsi2)
                        Server_Display = TampilDeskripsiServer(mAktiva.NSeri, mAktiva.Deskripsi, "", "", "", mAktiva.Deskripsi2)
                    End If
                ElseIf tabel_name = "ptag_wdcp" Then
                    mPriceTag = cPriceTag.GetDeskripsiPriceTag(tabel_name, ClientArrival.SO.BarcodePlu, "")
                    HH_Display = TampilPriceTagClient(mPriceTag.Prdcd, mPriceTag.Desc, "", mPriceTag.Keterangan)
                    Server_Display = TampilPriceTagServer(mPriceTag.Prdcd, mPriceTag.Desc, "", mPriceTag.Keterangan)

                ElseIf tabel_name = "monitoring_wdcp_ptag" Then
                    mMonitoringPriceTag = cMonitoringPriceTag.GetDeskripsiMonitoringPriceTag(tabel_name, ClientArrival.SO.BarcodePlu, Strings.Right(SocketData, 1))
                    HH_Display = TampilMonitoringPriceTagClient(mMonitoringPriceTag.barcode, mMonitoringPriceTag.keterangan, mMonitoringPriceTag.setMenu)
                    Server_Display = TampilMonitoringPriceTagServer(mMonitoringPriceTag.barcode, mMonitoringPriceTag.keterangan, mMonitoringPriceTag.setMenu)

                ElseIf tabel_name.ToUpper.StartsWith("SP") Then
                    mSo = cSo.GetDeskripsiProdukSOKhusus(tabel_name, ClientArrival.SO.BarcodePlu)
                    ClientArrival.SO = mSo
                    CountNext = 0
                    DtRakSOBIC = New DataTable
                    If ClientArrival.SO.TotalRak > 1 Then
                        DtRakSOBIC = cSo.GetListRakProdukSO(ClientArrival.SO.BarcodePlu)
                    End If
                    lokasi_so = "Toko"
                    If mSo.Deskripsi = "Tidak Ditemukan" Then
                        HH_Display = TampilDeskripsiClient(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko, ClientArrival.SO.QTYCom, "", ClientArrival.SO.TotalRak)
                        Server_Display = TampilDeskripsiServer(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko, ClientArrival.SO.QTYToko)
                    Else
                        'revisi lepas flag CBR SO PRODUK KHUSUS
                        '09/09/2021
                        If isFlagCBR = False Then
                            If mode_run = "E" Then
                                If ClientArrival.SO.QTYToko = 0 Then
                                    ClientArrival.SO.QTYInput = ClientArrival.SO.qtyTTL1_OLD
                                    HH_Display = TampilDeskripsiClient(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.qtyTTL1_OLD, ClientArrival.SO.QTYCom, "", ClientArrival.SO.TotalRak)
                                    Server_Display = TampilDeskripsiServer(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.qtyTTL1_OLD, ClientArrival.SO.QTYToko)
                                Else
                                    HH_Display = TampilDeskripsiClient(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko, ClientArrival.SO.QTYCom, "", ClientArrival.SO.TotalRak)
                                    Server_Display = TampilDeskripsiServer(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko, ClientArrival.SO.QTYToko)
                                End If
                            Else
                                HH_Display = TampilDeskripsiClient(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko, ClientArrival.SO.QTYCom, "", ClientArrival.SO.TotalRak)
                                Server_Display = TampilDeskripsiServer(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko, ClientArrival.SO.QTYToko)
                            End If
                        Else
                            If mSo.statusBarcode = "CBRY" Then
                                If mode_run = "E" Then
                                    If ClientArrival.SO.QTYToko = 0 Then
                                        ClientArrival.SO.QTYInput = ClientArrival.SO.qtyTTL1_OLD
                                        HH_Display = TampilDeskripsiClient(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.qtyTTL1_OLD, ClientArrival.SO.QTYCom, "", ClientArrival.SO.TotalRak)
                                        Server_Display = TampilDeskripsiServer(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.qtyTTL1_OLD, ClientArrival.SO.QTYToko)
                                    Else
                                        HH_Display = TampilDeskripsiClient(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko, ClientArrival.SO.QTYCom, "", ClientArrival.SO.TotalRak)
                                        Server_Display = TampilDeskripsiServer(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko, ClientArrival.SO.QTYToko)
                                    End If
                                Else
                                    HH_Display = TampilDeskripsiClient(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko, ClientArrival.SO.QTYCom, "", ClientArrival.SO.TotalRak)
                                    Server_Display = TampilDeskripsiServer(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko, ClientArrival.SO.QTYToko)
                                End If
                            Else
                                If mode_run = "E" Then
                                    If ClientArrival.SO.QTYToko = 0 Then
                                        ClientArrival.SO.QTYInput = ClientArrival.SO.qtyTTL1_OLD
                                        HH_Display = TampilDeskripsiClient(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.qtyTTL1_OLD, ClientArrival.SO.QTYCom, "", ClientArrival.SO.TotalRak)
                                        Server_Display = TampilDeskripsiServer(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.qtyTTL1_OLD, ClientArrival.SO.QTYCom)
                                    Else
                                        HH_Display = TampilDeskripsiClient(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko, ClientArrival.SO.QTYCom, "", ClientArrival.SO.TotalRak)
                                        Server_Display = TampilDeskripsiServer(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko, ClientArrival.SO.QTYCom)
                                    End If

                                Else
                                    If Strings.Right(SocketData, 1) = "B" Then
                                        HH_Display = TampilDeskripsiClient("", "", "", "", "")
                                        Server_Display = TampilDeskripsiServer("", "", "", "", "")
                                    Else
                                        'untuk flagprod selain CBR=Y maka lgsung dianggap qty input = 1
                                        ClientArrival.SO.QTYInput = 1
                                        ClientArrival.SO.statusBarcode = "CBRN"
                                        HH_Display = TampilDeskripsiClient(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko + 1, ClientArrival.SO.QTYCom, "CBRN", ClientArrival.SO.TotalRak)
                                        Server_Display = TampilDeskripsiServer(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko + 1, ClientArrival.SO.QTYCom, "CBRNE")
                                    End If
                                End If
                            End If
                        End If
                    End If
                ElseIf tabel_name.StartsWith("SZ") Then 'Input Barcode Bazar
                    mSo = cSo.GetDeskripsiProdukBazar(tabel_name, ClientArrival.SO.BarcodePlu)
                    ClientArrival.SO = mSo
                    CountNext = 0
                    DtED = New DataTable
                    HH_Display = TampilDeskripsiBazar(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.QTYTotal, ClientArrival.SO.Tgl_exp, ClientArrival.SO.QTYTotal, "", "", "")
                    Server_Display = TampilDeskripsiBKLServer(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.QTYTotal, ClientArrival.SO.Tgl_exp, "")
                    '1392/CPS/20
                ElseIf tabel_name.StartsWith("SE") Then 'Input Barcode SOED
                    mED = cED.getDeskripsiExpiredDate(tabel_name, ClientArrival.SO.BarcodePlu, "", mainCBRSOED)

                    If mED.StatusBarcode = "CBRN" And Strings.Right(SocketData, 1) = "B" Then
                        HH_Display = TampilDeskripsiExpired("", "Data Wajib di Scan", "", "", "", "0")
                        Server_Display = TampilDeskripsiExpiredServer("", "Data Wajib di Scan", "", "", "", "0")
                    Else
                        ClientArrival.SOED = mED
                        HH_Display = TampilDeskripsiExpired(ClientArrival.SOED.PRDCD, ClientArrival.SOED.Deskripsi,
                                                        ClientArrival.SOED.Lokasi, "", "", ClientArrival.SOED.Feedback)
                        Server_Display = TampilDeskripsiExpiredServer(ClientArrival.SOED.PRDCD, ClientArrival.SOED.Deskripsi,
                                                            ClientArrival.SOED.Lokasi, "", "", ClientArrival.SOED.Feedback)
                    End If

                ElseIf tabel_name = "CekDisplay" Then
                    mCekDisplay = cCekDisplay.GetDeskripsiListingDisplay(tabel_name, ClientArrival.SO.BarcodePlu, cmbModis.Text, ClientArrival.Login.User)
                    HH_Display = TampilCekDisplayClient(mCekDisplay.Prdcd, mCekDisplay.Desc, "")
                    Server_Display = TampilCekDisplayServer(mCekDisplay.Prdcd, mCekDisplay.Desc, "")
                Else
                    If jenis_so.ToLower = "tahunan" Then
                        mSo = cSo.GetDeskripsiProdukSO(tabel_name, ClientArrival.SO.BarcodePlu, Strings.Right(SocketData, 1), False)

                        Dim clsHandheld As New ClsHandheld
                        Dim clsHandheldController As New ClsHandheldController
                        clsHandheld.ipAddress = IpAddress.Split(":")(0)
                        lokasi_so = clsHandheldController.getLokasiSO(clsHandheld)

                    Else
                        mSo = cSo.GetDeskripsiProdukSO(tabel_name, ClientArrival.SO.BarcodePlu, Strings.Right(SocketData, 1), True, mainCBRSOIC, mainTTL3, lokasi_so)
                    End If

                    ClientArrival.SO = mSo
                    CountNext = 0
                    DtRakSOBIC = New DataTable

                    If ClientArrival.SO.TotalRak > 1 Then
                        DtRakSOBIC = cSo.GetListRakProdukSO(ClientArrival.SO.BarcodePlu)
                    End If

                    If mSo.statusBarcode = "CBRN" And Strings.Right(SocketData, 1) = "B" Then
                        ClientArrival.SO.Deskripsi = "Data Wajib di Scan"
                    End If

                    If lokasi_so = "Toko" Then
                        HH_Display = TampilDeskripsiClient(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko, ClientArrival.SO.QTYCom, mSo.statusBarcode, ClientArrival.SO.TotalRak)
                        Server_Display = TampilDeskripsiServer(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko, ClientArrival.SO.QTYCom, mSo.statusBarcode)
                    ElseIf lokasi_so = "Gudang" Then
                        HH_Display = TampilDeskripsiClient(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYGudang, ClientArrival.SO.QTYCom, mSo.statusBarcode, ClientArrival.SO.TotalRak)
                        Server_Display = TampilDeskripsiServer(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYGudang, ClientArrival.SO.QTYCom, mSo.statusBarcode)
                    Else
                        HH_Display = TampilClientSOICTTL3(ClientArrival.SO.PRDCD, ClientArrival.SO.Deskripsi, "", "", "")
                        Server_Display = TampilServerSOICTTL3(ClientArrival.SO.PRDCD, ClientArrival.SO.Deskripsi, "", "", "")
                    End If
                End If

                Log &= vbCrLf & "S/B - Scan Barang" & vbCrLf
                Log &= ClientArrival.IpAddress & " | " & ClientArrival.SocketID & " | PRDCD: " & ClientArrival.SO.PRDCD & " | QTY " & ClientArrival.SO.QTYInput & vbCrLf

            ElseIf Strings.Right(SocketData, 1) = "Q" Then 'Jika melakukan input Qty
                Dim mBPB As New ClsBPB
                Dim mBKL As New ClsBPBBKL

                Dim mNPS As New CLSNPS
                Dim mBPBNPS As New ClsBPBNPS

                Dim mBZR As New ClsBazar

                Dim mED As New ClsSOED
                Dim cED As New ClsSOEDController

                Dim parameter_docno As String = ""

                If CmdArg.Length > 1 And tabel_name.ToLower = "bpbbkl_wdcp" Then
                    parameter_docno = CmdArg(3)
                    parameter_docno = parameter_docno.Substring(1, parameter_docno.Length - 1)
                End If

                Dim mAktiva As New ClsAktiva
                Dim cBPB As New ClsBPBController
                Dim cAktiva As New ClsAktivaController
                Dim StatusQTY As String = ""

                Dim mKesegaran As New ClsKesegaran
                Dim cKesegaran As New ClsKesegaranController
                '1392/CPS/20
                Dim mCekDisplay As New ClsCekDisplay
                Dim cCekDisplay As New ClsCekDisplayController

                ClientArrival.SO.QTYInput = SocketData.Substring(0, SocketData.Length - 1)

                Log &= vbCrLf & "Q - Input QTY - 1" & vbCrLf
                Log &= ClientArrival.IpAddress & " | " & ClientArrival.SocketID & " | PRDCD: " & ClientArrival.SO.PRDCD & " | QTY " & ClientArrival.SO.QTYInput & vbCrLf

                Console.WriteLine("Q" & tabel_name & " | " & mainTTL3 & " | " & lokasi_so & " | " & ClientArrival.SO.isBADraft & " | " & ClientArrival.SO.isWtran)

                If tabel_name = "dcp_boxplu" Then 'Penerimaan Barang
                    mBPB = cBPB.RevisiQtyBarang(ClientArrival.SO.BarcodePlu, ClientArrival.SO.NoContainer, CbxKodeGudang.SelectedValue, ClientArrival.SO.QTYInput)
                    HH_Display = TampilDeskripsiClient(mBPB.Prdcd, mBPB.Desc, "", mBPB.Qty, "")
                    Server_Display = TampilDeskripsiServer(mBPB.Prdcd, mBPB.Desc, "", mBPB.Qty, "")

                ElseIf tabel_name.ToUpper.Contains("BPBBKL_WDCP") Then 'Penerimaan Barang BKL
                    mBKL = cBPB.cekQTYBKL(parameter_docno, ClientArrival.SO.QTYInput, ClientArrival.SO.BarcodePlu, CbxKodeGudang.SelectedValue, ClientArrival.SO.Tgl_exp, namauser)

                    StatusQTY = mBKL.StatusQTY
                    If StatusQTY = 1 Then ' Tidak ditemukan
                        HH_Display = TampilDeskripsiBKL(mBKL.BKL.Prdcd, mBKL.BKL.Desc, "", ClientArrival.SO.Tgl_exp, "", "", StatusQTY, "",, mBKL.BKL.fraction_pcs)
                        Server_Display = TampilDeskripsiBKLServer(mBKL.BKL.Prdcd, mBKL.BKL.Desc, "", ClientArrival.SO.Tgl_exp, mBKL.Feedback, mBKL.BKL.fraction_pcs)
                    ElseIf StatusQTY = 3 Then 'QTY melebihi
                        HH_Display = TampilDeskripsiBKL(mBKL.BKL.Prdcd, mBKL.BKL.Desc, "", ClientArrival.SO.Tgl_exp, "", "", StatusQTY, "", mBKL.BKL.sjQty, mBKL.BKL.fraction_pcs)
                        Server_Display = TampilDeskripsiBKLServer(mBKL.BKL.Prdcd, mBKL.BKL.Desc, "", "", mBKL.Feedback, mBKL.BKL.fraction_pcs)
                    ElseIf StatusQTY = 4 Then
                        HH_Display = TampilDeskripsiBKL(mBKL.BKL.Prdcd, mBKL.BKL.Desc, "", ClientArrival.SO.Tgl_exp, "", "", StatusQTY, "",, mBKL.BKL.fraction_pcs)
                        Server_Display = TampilDeskripsiBKLServer(mBKL.BKL.Prdcd, mBKL.BKL.Desc, "", "", mBKL.Feedback, mBKL.BKL.fraction_pcs)
                    ElseIf StatusQTY = 5 Then
                        HH_Display = TampilDeskripsiBKL(mBKL.BKL.Prdcd, mBKL.BKL.Desc, "", ClientArrival.SO.Tgl_exp, "", "", StatusQTY, "",, mBKL.BKL.fraction_pcs)
                        Server_Display = TampilDeskripsiBKLServer(mBKL.BKL.Prdcd, mBKL.BKL.Desc, "", "", mBKL.Feedback, mBKL.BKL.fraction_pcs)
                    Else 'Berhasil
                        HH_Display = TampilDeskripsiBKL(mBKL.BKL.Prdcd, mBKL.BKL.Desc, ClientArrival.SO.QTYInput, ClientArrival.SO.Tgl_exp, "Berhasil qty= " & mBKL.BKL.totalqty, "UPDATE QTY", StatusQTY, "",, mBKL.BKL.fraction_pcs)
                        Server_Display = TampilDeskripsiBKLServer(mBKL.BKL.Prdcd, mBKL.BKL.Desc, ClientArrival.SO.QTYInput, ClientArrival.SO.Tgl_exp, "Berhasil qty= " & mBKL.BKL.totalqty, mBKL.BKL.fraction_pcs)
                    End If
                    'DIRECT SHIPMENT MEMO 1029
                ElseIf tabel_name.ToUpper.Contains("BPBNPS_WDCP") Then 'Penerimaan Barang NPS
                    mBPBNPS = cBPB.cekQTYNPS(ClientArrival.SO.QTYInput, ClientArrival.SO.BarcodePlu, CbxKodeGudang.SelectedValue, ClientArrival.SO.Tgl_exp, namauser)

                    StatusQTY = mBPBNPS.StatusQTY

                    If StatusQTY = 1 Then ' Tidak ditemukan
                        HH_Display = TampilDeskripsiNPS(mBPBNPS.NPS.Prdcd, mBPBNPS.NPS.Desc, "", ClientArrival.SO.Tgl_exp, "", "", StatusQTY, "")
                        Server_Display = TampilDeskripsiBKLServer(mBPBNPS.NPS.Prdcd, mBPBNPS.NPS.Desc, "", ClientArrival.SO.Tgl_exp, mBPBNPS.Feedback)
                    ElseIf StatusQTY = 3 Then 'QTY melebihi
                        HH_Display = TampilDeskripsiNPS(mBPBNPS.NPS.Prdcd, mBPBNPS.NPS.Desc, "", ClientArrival.SO.Tgl_exp, "", "", StatusQTY, "", mBPBNPS.NPS.sjQty)
                        Server_Display = TampilDeskripsiBKLServer(mBPBNPS.NPS.Prdcd, mBPBNPS.NPS.Desc, "", "", mBPBNPS.Feedback)
                    ElseIf StatusQTY = 4 Then
                        HH_Display = TampilDeskripsiNPS(mBPBNPS.NPS.Prdcd, mBPBNPS.NPS.Desc, "", ClientArrival.SO.Tgl_exp, "", "", StatusQTY, "")
                        Server_Display = TampilDeskripsiBKLServer(mBPBNPS.NPS.Prdcd, mBPBNPS.NPS.Desc, "", "", mBPBNPS.Feedback)
                    ElseIf StatusQTY = 5 Then
                        HH_Display = TampilDeskripsiNPS(mBPBNPS.NPS.Prdcd, mBPBNPS.NPS.Desc, "", ClientArrival.SO.Tgl_exp, "", "", StatusQTY, "")
                        Server_Display = TampilDeskripsiBKLServer(mBPBNPS.NPS.Prdcd, mBPBNPS.NPS.Desc, "", "", mBPBNPS.Feedback)
                    Else 'Berhasil
                        HH_Display = TampilDeskripsiNPS(mBPBNPS.NPS.Prdcd, mBPBNPS.NPS.Desc, ClientArrival.SO.QTYInput, ClientArrival.SO.Tgl_exp, "Berhasil qty= " & mBPBNPS.NPS.totalqty, "UPDATE QTY", StatusQTY, "")
                        Server_Display = TampilDeskripsiBKLServer(mBPBNPS.NPS.Prdcd, mBPBNPS.NPS.Desc, ClientArrival.SO.QTYInput, ClientArrival.SO.Tgl_exp, "Berhasil qty= " & mBPBNPS.NPS.totalqty)
                    End If

                    'revisi Memo No. 959 - CPS - 20
                    '13/10/2020
                    'Cek Kesegaran
                ElseIf tabel_name = "CekKesegaran" Then

                    mKesegaran = cKesegaran.GetDeskripsiKesegaran(tabel_name, ClientArrival.SO.BarcodePlu, cmbModis.Text, ClientArrival.Login.User, ClientArrival.SO.QTYInput)
                    If mKesegaran.Desc = "QTY melebihi LPP!" Then
                        HH_Display = TampilKesegaranClient("", mKesegaran.Desc, "", "")
                        Server_Display = TampilKesegaranServer("", mKesegaran.Desc, "", "")
                    Else

                        HH_Display = TampilKesegaranClient("", "", "", "")
                        Server_Display = TampilKesegaranServer("", "", "", "")
                    End If

                ElseIf tabel_name.ToUpper.Contains("OA") Then 'Stock Opname Aktiva
                    'revisi Memo No. 208 - CPS - 20
                    '13/10/2020
                    mAktiva = cAktiva.GetDeskripsiAktiva(tabel_name, ClientArrival.SO.BarcodePlu, Toko, ClientArrival.SO.QTYInput)

                    If mAktiva.statusQty = True And Not mAktiva.Deskripsi2.Contains("AT baru !") Then
                        mAktiva = cAktiva.UpdateQtyProduk(tabel_name, ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.QTYInput, 0, ClientArrival.SO.statusBarcode, Toko)
                        HH_Display = TampilDeskripsiClient(mAktiva.NSeri, mAktiva.Deskripsi, "", "", "", mAktiva.Deskripsi2)
                        Server_Display = TampilDeskripsiServer(mAktiva.NSeri, mAktiva.Deskripsi, "", "", "", mAktiva.Deskripsi2)
                    Else
                        If mAktiva.Deskripsi2.Contains("AT baru !") Then
                            mAktiva.statusQty = False
                        End If
                        If ClientArrival.SO.QTYInput > mAktiva.QtyMax And mAktiva.statusQty = True Then
                            mAktiva = cAktiva.UpdateQtyProduk(tabel_name, ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.QTYInput, 0, ClientArrival.SO.statusBarcode, Toko)
                            HH_Display = TampilDeskripsiClient(mAktiva.NSeri, mAktiva.Deskripsi, "", "", "", mAktiva.Deskripsi2)
                            Server_Display = TampilDeskripsiServer(mAktiva.NSeri, mAktiva.Deskripsi, "", "", "", mAktiva.Deskripsi2)
                        Else
                            HH_Display = TampilDeskripsiClient(mAktiva.NSeri, mAktiva.Deskripsi, mAktiva.statusQty.ToString, ClientArrival.SO.QTYInput, "", mAktiva.Deskripsi2)
                            Server_Display = TampilDeskripsiServer(mAktiva.NSeri, mAktiva.Deskripsi, mAktiva.statusQty.ToString, ClientArrival.SO.QTYInput, "", mAktiva.Deskripsi2)
                        End If
                    End If

                ElseIf tabel_name.StartsWith("SZ") Then 'SO Bazar
                    Dim statusExp As String = ""
                    statusExp = mBZR.StatusExp

                    HH_Display = TampilDeskripsiBazar("", "", "", "", "", "", "", "")
                    Server_Display = TampilDeskripsiBKLServer("", "", "", "", "")

                ElseIf tabel_name.StartsWith("SE") Then 'SO Expired Date
                    Dim qtyInput As Integer = SocketData.Substring(0, SocketData.Length - 1)

                    mED = cED.SimpanQty_ED(tabel_name, ClientArrival.SOED.PRDCD, ClientArrival.SOED.ExpDate, qtyInput, ClientArrival.SOED.noPropED)

                    If ClientArrival.SOED.Feedback = "4" Then
                        HH_Display = TampilDeskripsiExpired("", "", "", "", "", mED.Feedback)
                        Server_Display = TampilDeskripsiExpiredServer("", "", "", "", "", mED.Feedback)
                    Else
                        HH_Display = TampilDeskripsiExpired("", mED.Deskripsi, "", "", "", mED.Feedback)
                        Server_Display = TampilDeskripsiExpiredServer("", mED.Deskripsi, "", "", "", mED.Feedback)
                    End If

                ElseIf tabel_name = "CekDisplay" Then '1392/CPS/20
                    cCekDisplay.SimpanQTYCekDisplay("CekDisplay", ClientArrival.SO.BarcodePlu, ClientArrival.SO.QTYInput, cmbModis.Text, ClientArrival.Login.User.ID, nama_cekdisplay, jabatan_cekdisplay)
                    HH_Display = TampilCekDisplayClient("", "", "", "1")
                    Server_Display = TampilCekDisplayServer("", "", "")

                Else 'Stock Opname 
                    Dim temp_ttlQTY As Double

                    If ClientArrival.SO.QTYInput = 0 Then
                        ClientArrival.SO.QTYToko = 0
                        ClientArrival.SO.QTYGudang = 0
                        ClientArrival.SO.QTYTotal = 0
                    End If

                    If jenis_so.ToLower = "tahunan" And mode_run = "E" Then
                        'Setiap Device bisa beda lokasi
                        Dim clsHandheld As New ClsHandheld
                        Dim clsHandheldController As New ClsHandheldController

                        clsHandheld.ipAddress = IpAddress.Split(":")(0)
                        clsHandheld.socketID = SocketID
                        lokasi_so = clsHandheldController.getLokasiSO(clsHandheld)
                    End If

                    'Revisi MEMO 1230 SOIC
                    If tabel_name.ToUpper.Contains("SB") And mainTTL3 And lokasi_so = "Barang Rusak" Then
                        Dim clsSOIC As New ClsSOICController
                        Dim alasan As String
                        Dim tabelTTL3 As String

                        If ClientArrival.SO.isBADraft Or ClientArrival.SO.isWtran Then
                            If ClientArrival.SO.qtyReturExpired = "" Then
                                ClientArrival.SO.qtyReturExpired = ClientArrival.SO.QTYInput
                                alasan = "Barang Expired"
                            ElseIf ClientArrival.SO.qtyReturKemasan = "" Then
                                ClientArrival.SO.qtyReturKemasan = ClientArrival.SO.QTYInput
                                alasan = "Kemasan Rusak"
                            Else
                                ClientArrival.SO.qtyReturDigigit = ClientArrival.SO.QTYInput
                                alasan = "Digigit Tikus/Serangga"
                            End If

                            If ClientArrival.SO.isBADraft Then
                                tabelTTL3 = "ba_draft"
                            Else
                                tabelTTL3 = "wtran"
                            End If

                            clsSOIC.insertOrUpdate_barangRusak(ClientArrival.SO.PRDCD, ClientArrival.SO.QTYInput, tabelTTL3, alasan)

                            If ClientArrival.SO.qtyReturExpired = "" Or ClientArrival.SO.qtyReturKemasan = "" Or ClientArrival.SO.qtyReturDigigit = "" Then
                                GoTo NEXT_QTY
                            Else
                                temp_ttlQTY = Double.Parse(ClientArrival.SO.qtyReturExpired) + Double.Parse(ClientArrival.SO.qtyReturKemasan) + Double.Parse(ClientArrival.SO.qtyReturDigigit)
                                ClientArrival.SO.QTYInput = temp_ttlQTY.ToString
                            End If
                        End If
                    End If

                    'jika input qty 0 tampilkan deskripsi sebelumnya
                    If ClientArrival.SO.QTYInput = "0" Then
NEXT_QTY:
                        If lokasi_so = "Toko" Then
                            If tabel_name.ToUpper.Contains("SP") Then
                                HH_Display = TampilDeskripsiClient(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko, ClientArrival.SO.QTYCom)
                                Server_Display = TampilDeskripsiServer(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko, ClientArrival.SO.QTYCom)
                            Else
                                HH_Display = TampilDeskripsiClient(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko, ClientArrival.SO.QTYCom)
                                Server_Display = TampilDeskripsiServer(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, "", ClientArrival.SO.QTYCom)
                            End If
                        ElseIf lokasi_so = "Gudang" Then
                            HH_Display = TampilDeskripsiClient(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYGudang, ClientArrival.SO.QTYCom)
                            Server_Display = TampilDeskripsiServer(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, "", ClientArrival.SO.QTYCom)
                        Else
                            HH_Display = TampilClientSOICTTL3(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.qtyReturExpired, ClientArrival.SO.qtyReturKemasan, ClientArrival.SO.qtyReturDigigit)
                            Server_Display = TampilServerSOICTTL3(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.qtyReturExpired, ClientArrival.SO.qtyReturKemasan, ClientArrival.SO.qtyReturDigigit)
                        End If
                    Else
                        If tabel_name.ToUpper.Contains("SP") Then
                            HH_Display = TampilDeskripsiClient("", "", "", "", "")
                            Server_Display = TampilDeskripsiServer("", "", "", "", "")
                        Else
                            HH_Display = TampilDeskripsiClient("", "", "", "", "")
                            Server_Display = TampilDeskripsiServer("", "", "", "", "")
                        End If
                    End If
                End If
                Log &= vbCrLf & "Q - Input QTY - 2" & vbCrLf
                Log &= ClientArrival.IpAddress & " | " & ClientArrival.SocketID & " | PRDCD: " & ClientArrival.SO.PRDCD & " | QTY " & ClientArrival.SO.QTYInput & vbCrLf
            ElseIf Strings.Right(SocketData, 1) = "C" Then 'Jika melakukan input scan barcode container
                DescBPB = ""
                Dim cBPB As New ClsBPBController
                Dim Box1, Box2 As Integer
                Dim IsContainer As Boolean = True
                Dim BarcodeBox As String = cUtility.BrcdKontainerOrBronjong(SocketData.Substring(0, SocketData.Length - 1))
                Box1 = 0
                Box2 = 0
                DtBPB_DCP = cBPB.CekPluContainer(BarcodeBox, CbxKodeGudang.SelectedValue, Box1, Box2)
                If Box1 > 0 Then
                    ClientArrival.SO = New ClsSo
                    ClientArrival.SO.NoContainer = BarcodeBox
                    DisplayGridBPB(DtBPB_DCP)
                    DescBPB = "Mulai scan barang"
                    For Each KdCbng As String In lCabang
                        If ClientArrival.SO.NoContainer.Contains(KdCbng) Then
                            IsContainer = False
                            Exit For
                        End If
                    Next
                    HH_Display = TampilContainerClient2(IsContainer, ClientArrival.SO.NoContainer, Box2 & "/" & Box1 & "Record", DescBPB)
                    Server_Display = TampilContainerServer(IsContainer, ClientArrival.SO.NoContainer, Box2 & "/" & Box1 & "Record", DescBPB)
                    txtRec.Text = Box2 & "/" & Box1
                Else
                    ClientArrival.SO.NoContainer = ""
                    DescBPB = "Barcode tdk trdaftar"
                    HH_Display = TampilContainerClient(ClientArrival.SO.NoContainer, DescBPB)
                    Server_Display = TampilContainerServer(False, ClientArrival.SO.NoContainer, "", DescBPB)
                End If

            ElseIf Strings.Right(SocketData, 1) = "R" Then
                If (jenis_so = "BIC" Or jenis_so = "Kasus") And SocketData.Substring(0, SocketData.Length - 1) = 0 Then
                    CountNext = 0
                    If lokasi_so = "Toko" Then
                        HH_Display = TampilDeskripsiClient(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko, ClientArrival.SO.QTYCom, "", ClientArrival.SO.TotalRak)
                        Server_Display = TampilDeskripsiServer(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYToko, ClientArrival.SO.QTYCom)
                    Else
                        HH_Display = TampilDeskripsiClient(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYGudang, ClientArrival.SO.QTYCom, "", ClientArrival.SO.TotalRak)
                        Server_Display = TampilDeskripsiServer(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.Rak, ClientArrival.SO.QTYGudang, ClientArrival.SO.QTYCom)
                    End If
                Else
                    GoTo NEXT_BIC
                End If
                'revisi Memo No. 208 - CPS - 20
                '13/10/2020
            ElseIf Strings.Right(SocketData, 1) = "K" Then 'Jika melakukan input Qty Rusak (OA)
                Dim mAktiva As New ClsAktiva
                Dim cAktiva As New ClsAktivaController

                ClientArrival.SO.QTYInputRusak = SocketData.Substring(0, SocketData.Length - 1)
                ClientArrival.SO.QTYInputBaik = ClientArrival.SO.QTYInput
                If tabel_name.ToUpper.Contains("OA") Then 'Stock Opname Aktiva
                    'revisi Memo No. 208 - CPS - 20
                    '13/10/2020
                    mAktiva = cAktiva.UpdateQtyProduk(tabel_name, ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, ClientArrival.SO.QTYInputBaik, ClientArrival.SO.QTYInputRusak, ClientArrival.SO.statusBarcode, Toko)

                    HH_Display = TampilDeskripsiClient(mAktiva.NSeri, mAktiva.Deskripsi, "", "", "", mAktiva.Deskripsi2)
                    Server_Display = TampilDeskripsiServer(mAktiva.NSeri, mAktiva.Deskripsi, "", "", "", mAktiva.Deskripsi2)
                End If
                'revisi Memo No. 208 - CPS - 20
                '13/10/2020
            ElseIf (Strings.Right(SocketData, 1) = "I") Then 'Tampil list indeks
                Dim mAktiva As New ClsAktiva
                Dim cAktiva As New ClsAktivaController
                ClientArrival.SO.indeksList = SocketData.Substring(0, SocketData.Length - 1) ' nomor indeks

                If ClientArrival.SO.indeksList.StartsWith("0") Or ClientArrival.SO.indeksList = "0" Then
                    ' pilih = 0 (NEXT)
                    If indeks + 1 > mAktivaList.Count Then
                        If indeks + 1 > mAktivaList.Count Then
                            indeks -= 1
                            HH_Display = TampilDeskripsiListClient(mAktivaList, indeks)
                            Server_Display = TampilDeskripsiListServer(mAktivaList, indeks)
                        Else

                            HH_Display = TampilDeskripsiListClient(mAktivaList, indeks)
                            Server_Display = TampilDeskripsiListServer(mAktivaList, indeks)

                        End If
                    ElseIf indeks + 1 = mAktivaList.Count Then
                        indeks = 0
                        HH_Display = TampilDeskripsiListClient(mAktivaList, indeks)
                        Server_Display = TampilDeskripsiListServer(mAktivaList, indeks)

                    Else
                        'mAktivaList = cAktiva.GetListAktiva(tabel_name, ClientArrival.SO.BarcodePlu, Toko)
                        indeks += 1
                        HH_Display = TampilDeskripsiListClient(mAktivaList, indeks)
                        Server_Display = TampilDeskripsiListServer(mAktivaList, indeks)

                    End If

                Else
                    If ClientArrival.SO.indeksList > mAktivaList.Count Or mAktivaList.Count = 0 Then
                        HH_Display = TampilDeskripsiClient("", "Tidak Ditemukan", "", "", "")
                        Server_Display = TampilDeskripsiServer("", "Tidak Ditemukan", "", "", "")
                    Else
                        mAktiva = cAktiva.GetDeskripsiAktiva(tabel_name, mAktivaList(ClientArrival.SO.indeksList - 1).NSeri, Toko)

                        'Set deskripsi & max qty Aktiva
                        ClientArrival.SO.Deskripsi = mAktiva.Deskripsi
                        If mAktiva.Deskripsi2.Contains("AT baru") Then
                            ClientArrival.SO.Deskripsi = mAktiva.Deskripsi2
                        End If
                        ClientArrival.SO.QTYCom = mAktiva.QtyMax
                        ClientArrival.SO.BarcodePlu = mAktiva.NSeri
                        HH_Display = TampilDeskripsiClient(mAktiva.NSeri, mAktiva.Deskripsi, "", "", "", mAktiva.Deskripsi2)
                        Server_Display = TampilDeskripsiServer(mAktiva.NSeri, mAktiva.Deskripsi, "", "", "", mAktiva.Deskripsi2)
                    End If

                End If

            ElseIf Strings.Right(SocketData, 1) = "T" And SocketData <> "NEXT" Then 'Jika melakukan input TGL EXP
                Dim mBPB As New ClsBKL
                Dim mBKL As New ClsBPBBKL

                Dim mNPS As New CLSNPS
                Dim mBPBNPS As New ClsBPBNPS

                Dim mBZR As New ClsBazar
                Dim cBPB As New ClsBPBController

                Dim mED As New ClsSOED
                Dim cED As New ClsSOEDController

                Dim StatusEXP As String = ""
                Dim feedback As String = ""

                ClientArrival.SO.Tgl_exp = SocketData.Substring(0, SocketData.Length - 1)
                ClientArrival.SOED.ExpDateInput = SocketData.Substring(0, SocketData.Length - 1)

                Log &= vbCrLf & "T - Input TGL EXP" & vbCrLf
                Log &= ClientArrival.IpAddress & " | " & ClientArrival.SocketID & " | PRDCD: " & ClientArrival.SO.PRDCD & " | QTY " & ClientArrival.SO.Tgl_exp & vbCrLf

                Console.WriteLine(ClientArrival.IpAddress & " | " & ClientArrival.SocketID & " | PRDCD: " & ClientArrival.SOED.PRDCD & " | ExpDateInput " & ClientArrival.SOED.ExpDateInput & vbCrLf)

                If tabel_name.Trim.ToLower = "bpbbkl_wdcp" Then

                    mBKL = cBPB.inputTGL_EPW(ClientArrival.SO.BarcodePlu, CbxKodeGudang.SelectedValue, ClientArrival.SO.QTYInput, ClientArrival.SO.Tgl_exp)

                    StatusEXP = mBKL.StatusExp

                    ClientArrival.SO.Tgl_exp = mBKL.BKL.TglEXP

                    If mBKL.StatusExp = "1" Or mBKL.StatusExp = "3" Then

                        HH_Display = TampilDeskripsiBKL(mBKL.BKL.Prdcd, mBKL.BKL.Desc, "", ClientArrival.SO.Tgl_exp, "", "", "", StatusEXP,, mBKL.BKL.fraction_pcs)

                        Server_Display = TampilDeskripsiBKLServer(mBKL.BKL.Prdcd, mBKL.BKL.Desc, "", "", mBKL.Feedback, mBKL.BKL.fraction_pcs)

                    Else
                        HH_Display = TampilDeskripsiBKL(mBKL.BKL.Prdcd, mBKL.BKL.Desc, "", ClientArrival.SO.Tgl_exp, "TGL EXP OKE", "", "", StatusEXP,, mBKL.BKL.fraction_pcs)

                        Server_Display = TampilDeskripsiBKLServer(mBKL.BKL.Prdcd, mBKL.BKL.Desc, "", ClientArrival.SO.Tgl_exp, mBKL.Feedback, mBKL.BKL.fraction_pcs)
                    End If

                    'DIRECT SHIPMENT MEMO 1029
                ElseIf tabel_name.Trim.ToLower = "bpbnps_wdcp" Then

                    mBPBNPS = cBPB.inputTGL_EPW_NPS(ClientArrival.SO.BarcodePlu, CbxKodeGudang.SelectedValue, ClientArrival.SO.QTYInput, ClientArrival.SO.Tgl_exp)

                    StatusEXP = mBPBNPS.StatusExp

                    ClientArrival.SO.Tgl_exp = mBPBNPS.NPS.TglEXP

                    If mBPBNPS.StatusExp = "1" Or mBPBNPS.StatusExp = "3" Then
                        'NOK
                        HH_Display = TampilDeskripsiBKL(mBPBNPS.NPS.Prdcd, mBPBNPS.NPS.Desc, "", ClientArrival.SO.Tgl_exp, "", "", "", StatusEXP)
                        Server_Display = TampilDeskripsiBKLServer(mBPBNPS.NPS.Prdcd, mBPBNPS.NPS.Desc, "", "", mBPBNPS.Feedback)

                    Else
                        HH_Display = TampilDeskripsiBKL(mBPBNPS.NPS.Prdcd, mBPBNPS.NPS.Desc, "", ClientArrival.SO.Tgl_exp, "TGL EXP OKE", "", "", StatusEXP)
                        Server_Display = TampilDeskripsiBKLServer(mBPBNPS.NPS.Prdcd, mBPBNPS.NPS.Desc, "", ClientArrival.SO.Tgl_exp, mBPBNPS.Feedback)
                    End If
                ElseIf tabel_name.StartsWith("SZ") Then 'so Bazar
                    mBZR = cBPB.inputTGL_EXP_Bazar(tabel_name, ClientArrival.SO.BarcodePlu, ClientArrival.SO.Tgl_exp)
                    StatusEXP = mBZR.StatusExp

                    If mBZR.StatusExp = "1" Or mBZR.StatusExp = "3" Then
                        HH_Display = TampilDeskripsiBazar(mBZR.BZR.PRDCD, mBZR.BZR.Deskripsi, mBZR.BZR.QTYInput, ClientArrival.SO.Tgl_exp, mBZR.BZR.QTYTotal, "", "", StatusEXP)
                        Server_Display = TampilDeskripsiBKLServer(mBZR.BZR.BarcodePlu, mBZR.BZR.Deskripsi, mBZR.BZR.QTYInput, ClientArrival.SO.Tgl_exp, mBZR.BZR.Feedback)
                    Else
                        HH_Display = TampilDeskripsiBazar(mBZR.BZR.PRDCD, mBZR.BZR.Deskripsi, mBZR.BZR.QTYInput, ClientArrival.SO.Tgl_exp, mBZR.BZR.QTYTotal, "Tgl EXP valid", "", StatusEXP)
                        Server_Display = TampilDeskripsiBKLServer(mBZR.BZR.BarcodePlu, mBZR.BZR.Deskripsi, mBZR.BZR.QTYInput, ClientArrival.SO.Tgl_exp, mBZR.BZR.Feedback)
                    End If
                ElseIf tabel_name.StartsWith("SE") Then 'SO Expired Date
                    mED = cED.inputTglExp(tabel_name, ClientArrival.SOED.PRDCD, ClientArrival.SOED.ExpDateInput, mainCBRSOED)
                    ClientArrival.SOED.noPropED = mED.noPropED
                    ClientArrival.SOED.ExpDate = mED.ExpDate

                    HH_Display = TampilDeskripsiExpired(mED.PRDCD, mED.Deskripsi, mED.Lokasi, mED.ExpDate, "", mED.Feedback)
                    Server_Display = TampilDeskripsiExpiredServer(mED.PRDCD, mED.Deskripsi, mED.Lokasi, mED.ExpDate, "", mED.Feedback)
                End If
            ElseIf (Strings.Right(SocketData, 1) = "F") Then 'Jika berhasil
                Dim mSo As New ClsSo
                Dim mPriceTag As New ClsPriceTag
                Dim cPriceTag As New ClsPriceTagController
                Dim mMonitoringPriceTag As New ClsMonitoringPriceTag 'UPDATE MEMO 296/CPS/23 by Kukuh
                Dim cMonitoringPriceTag As New ClsMonitoringPriceTagController 'UPDATE MEMO 296/CPS/23 by Kukuh

                If tabel_name = "ptag_wdcp" Then
                    ClientArrival.SO.Deskripsi = SocketData.Substring(0, SocketData.Length - 1)

                    mPriceTag = cPriceTag.GetDeskripsiPriceTag(tabel_name, ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi)
                    If mPriceTag.Keterangan = "INSERT" Or mPriceTag.Keterangan = "HAPUS" Then
                        HH_Display = TampilPriceTagClient(mPriceTag.Prdcd, mPriceTag.Desc, "", mPriceTag.Keterangan)
                        Server_Display = TampilPriceTagServer(mPriceTag.Prdcd, mPriceTag.Desc, "", mPriceTag.Keterangan)
                    Else
                        HH_Display = TampilPriceTagClient("", "", "", "")
                        Server_Display = TampilPriceTagServer("", "", "", "")
                    End If
                ElseIf tabel_name = "monitoring_wdcp_ptag" Then 'Revisi Memo No 296/CPS/23 Monitoring Price Tag by Kukuh 16 Mei 2023
                    If Strings.Left(SocketData, 1) = "1" Then
                        mMonitoringPriceTag = cMonitoringPriceTag.GetDeskripsiMonitoringPriceTag(tabel_name, ClientArrival.SO.BarcodePlu, Strings.Right(SocketData, 1))
                        Dim isSuccess As Boolean

                        If mMonitoringPriceTag.setMenu = "3" Then
                            HH_Display = TampilMonitoringPriceTagClient("", mMonitoringPriceTag.keterangan, "1")
                            Server_Display = TampilMonitoringPriceTagServer("", mMonitoringPriceTag.keterangan, "1")
                        Else
                            isSuccess = cMonitoringPriceTag.UpdateDataMPC(tabel_name, mMonitoringPriceTag.barcode)
                            If isSuccess Then
                                HH_Display = TampilMonitoringPriceTagClient("", "Berhasil Update", "1")
                                Server_Display = TampilMonitoringPriceTagServer("", "Berhasil Update", "1")
                            Else
                                HH_Display = TampilMonitoringPriceTagClient(ClientArrival.SO.BarcodePlu, "Simpan Ulang", "2")
                                Server_Display = TampilMonitoringPriceTagServer(ClientArrival.SO.BarcodePlu, "Simpan Ulang", "2")
                            End If
                        End If
                    Else
                        HH_Display = TampilMonitoringPriceTagClient("", "", "1")
                        Server_Display = TampilMonitoringPriceTagServer("", "", "1")
                    End If
                Else
                    ClientArrival.SO = New ClsSo
                    HH_Display = TampilDeskripsiBKL("", "", "", "", "", "", "", "")
                    Server_Display = TampilDeskripsiServer("", "", "", "", "")
                End If

            ElseIf Strings.Right(SocketData, 4) = "NEXT" Then
NEXT_BIC:
                Dim MaxRowPage As Integer = 2 'Batas baris info Rak di WDCP
                Dim DtRak As New DataTable

                If DtRakSOBIC.Rows.Count > 0 Then
                    Dim NoFirst As Integer = (CountNext * MaxRowPage) + 1
                    Dim NoLast As Integer = (NoFirst + MaxRowPage) - 1

                    DtRak.Columns.Add("NO", GetType(Integer))
                    DtRak.Columns.Add("TIPERAK", GetType(String))
                    DtRak.Columns.Add("NORAK", GetType(String))
                    DtRak.Columns.Add("NOSHELF", GetType(String))

                    For iNo As Integer = NoFirst To NoLast
                        For Each Dr As DataRow In DtRakSOBIC.Rows
                            If iNo = CInt(Dr("NO")) Then
                                Dim NRow As DataRow = DtRakSOBIC.NewRow
                                DtRak.Rows.Add(New Object() {Dr("NO"), Dr("TIPERAK"), Dr("NORAK"), Dr("NOSHELF")})
                            End If
                        Next
                    Next
                    CountNext += 1
                End If

                HH_Display = TampilNextSOBIC_Client(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, CountNext, (DtRakSOBIC.Rows.Count Mod 2) + 1, DtRak)
                Server_Display = TampilNextSOBIC_Server(ClientArrival.SO.BarcodePlu, ClientArrival.SO.Deskripsi, CountNext, (DtRakSOBIC.Rows.Count Mod 2) + 1, DtRak)

                If CountNext = (DtRakSOBIC.Rows.Count Mod 2) + 1 Then
                    CountNext = 0
                End If
            End If

            Log &= vbCrLf & "Set Client " & vbCrLf
            For i As Integer = 0 To listClient.Length - 1
                If listClient(i).SocketID = ClientArrival.SocketID Then
                    Log &= "IF: " & listClient(i).SocketID & "=" & ClientArrival.SocketID & vbCrLf

                    listClient(i) = ClientArrival

                    Log &= "listClient(" & i & "): " & listClient(i).IpAddress & "|" & listClient(i).SocketID & vbCrLf
                    Log &= "Client: " & ClientArrival.IpAddress & "|" & ClientArrival.SocketID & vbCrLf
                    Exit For
                End If
            Next

            Log &= vbCrLf & "onDataArrival (FINISH) " & vbCrLf
            For Each mClient As ClsClient In listClient
                If Not IsNothing(mClient) Then
                    Log &= mClient.IpAddress & " | " & mClient.SocketID & vbCrLf
                End If
            Next
            Util.TraceLogTxt(Log)

            SendData(SocketID, HH_Display)
            mainWindow.Invoke(ShowDisplayClient, New Object() {IdxArrival, Server_Display})
        Catch ex As Exception
            Dim util As New Utility
            util.TraceLogTxt("_socketManager_onDataArrival " & ex.Message & vbCrLf & ex.StackTrace & " IP:" & IpAddress)
            TraceLog("_socketManager_onDataArrival " & ex.Message & vbCrLf & ex.StackTrace & " IP:" & IpAddress)
        End Try

    End Sub

    Private Sub DisplayGridBPB(ByVal DtBPB As DataTable)
        If Me.DgBPB.InvokeRequired Then
            Dim d As New DisplayGridBPBDelegate(AddressOf DisplayGridBPB)
            Me.Invoke(d, New Object() {DtBPB})
        Else
            DtBPB.TableName = "boxplu"
            Me.DgBPB.DataSource = DtBPB
        End If
    End Sub

    Private Sub DisplayGridBPBBKL(ByVal DtBPBBKL As DataTable)
        If Me.DgBPB.InvokeRequired Then
            Dim d As New DisplayGridBPBBKLDelegate(AddressOf DisplayGridBPBBKL)
            Me.Invoke(d, New Object() {DtBPBBKL})
        Else
            DtBPBBKL.TableName = "BPBBKL"
            Me.DgBPB.DataSource = DtBPBBKL

        End If
    End Sub
    Private Sub DisplayGridBPBNPS(ByVal DtBPBNPS As DataTable)
        If Me.DgBPB.InvokeRequired Then
            Dim d As New DisplayGridBPBNPSDelegate(AddressOf DisplayGridBPBNPS)
            Me.Invoke(d, New Object() {DtBPBNPS})
        Else
            DtBPBNPS.TableName = "BPBNPS"
            Me.DgBPB.DataSource = DtBPBNPS

        End If
    End Sub
    Public Sub SendData(ByVal SocketID As String, ByVal tmpData As String)
        Try
            'Dim util As New Utility
            'util.TraceLogTxt("SendData (Start) " & vbCrLf & "Data:" & tmpData & " | SocketId:" & SocketID & " | ListClient:" & listClient.Length.ToString)

            If SocketID <> "" Then
                _socketManager.Send(SocketID, tmpData, False)
            Else
                Dim I As Integer
                For I = 0 To _socketManager.Count - 1
                    _socketManager.Send(_socketManager.ItembyIndex(I).SocketID, tmpData, False)
                Next
            End If
        Catch ex As Exception
            Dim util As New Utility
            util.TraceLogTxt("SendData " & ex.Message & vbCrLf & ex.StackTrace & " Data:" & tmpData)
        End Try

    End Sub
#End Region

    Private mainWindow As Form

    Private Delegate Sub IPClient(ByVal client_no As Integer, ByVal IPAddress As String)
    Private ShowIpAddress As IPClient

    Private Delegate Sub DisplayClient(ByVal client_no As Integer, ByVal display As String)
    Private ShowDisplayClient As DisplayClient

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        mainWindow = Me
        listClient = New ClsClient() {Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing}

    End Sub

    Public Sub InitialClient()
        If jenis_so = "BPB" Or jenis_so = "BPBBKL" Or jenis_so = "BPBNPS" Then
            ShowIpAddress = AddressOf showIpClient2
            ShowDisplayClient = AddressOf showDisplayBPB
        ElseIf jenis_so = "Planogram" Or jenis_so = "Kesegaran" Or jenis_so = "CekPJR" Or jenis_so = "TINDAKLBTD" Or jenis_so = "TINDAKLBTD_BAPJR" Or jenis_so = "CekDisplay" Then
            ShowIpAddress = AddressOf showIpClient3
            ShowDisplayClient = AddressOf showDisplayPlano
        Else
            ShowIpAddress = AddressOf showIpClient
            ShowDisplayClient = AddressOf showDisplay
        End If
    End Sub

#Region "Display"

    Public client1display As String
    Public client2display As String
    Public client3display As String
    Public client4display As String
    Public client5display As String
    Public client6display As String
    Public client7display As String
    Public client8display As String
    Public client9display As String
    Public rejectdisplay As String

    Public Sub showDisplay(ByVal client_no As Integer, ByVal display As String)
        Try
            If client_no = 0 Then
                Handheld1_box.Text = display
            ElseIf client_no = 1 Then
                Handheld2_box.Text = display
            ElseIf client_no = 2 Then
                Handheld3_box.Text = display
            ElseIf client_no = 3 Then
                Handheld4_box.Text = display
            ElseIf client_no = 4 Then
                Handheld5_box.Text = display
            ElseIf client_no = 5 Then
                Handheld6_box.Text = display
            ElseIf client_no = 6 Then
                Handheld7_Box.Text = display
            ElseIf client_no = 7 Then
                Handheld8_box.Text = display
            Else
                Handheld9_box.Text = display
            End If
        Catch ex As Exception
            ShowError("Error showdisplay", ex)
        End Try
    End Sub

    Public Sub showDisplayBPB(ByVal client_no As Integer, ByVal display As String)
        Try
            'Dim util As New Utility
            'util.TraceLogTxt("showDisplayBPB " & vbCrLf & "client_no:" & client_no & " | display:" & display)

            'tambahan 6/10/20
            If jenis_so = "BPBBKL" Or jenis_so = "BPBNPS" Then
                If client_no = 0 Then
                    DCP1_box.Text = display
                ElseIf client_no = 1 Then
                    DCP2_box.Text = display
                ElseIf client_no = 2 Then
                    DCP3_box.Text = display
                Else
                    DCP4_box.Text = display
                End If
            Else
                If client_no = 0 Then
                    DCP1_box.Text = display
                Else
                    DCP2_box.Text = display
                End If
            End If
        Catch ex As Exception
            ShowError("Error showdisplay", ex)
        End Try
    End Sub

    Public Sub showDisplayPlano(ByVal client_no As Integer, ByVal display As String)
        Try
            If client_no = 0 Then
                Plano1_Box.Text = display
            Else
                Plano2_Box.Text = display
            End If
        Catch ex As Exception
            ShowError("Error showdisplay", ex)
        End Try
    End Sub

    Public Function modeldisplay(ByVal client_no As Integer)
        Dim display As String = ""

        Try
            If client_no = 0 Then
                display = client1display
            ElseIf client_no = 1 Then
                display = client2display
            ElseIf client_no = 2 Then
                display = client3display
            ElseIf client_no = 3 Then
                display = client4display
            ElseIf client_no = 4 Then
                display = client5display
            ElseIf client_no = 5 Then
                display = client6display
            ElseIf client_no = 6 Then
                display = client7display
            ElseIf client_no = 7 Then
                display = client8display
            ElseIf client_no = 8 Then
                display = client9display
            Else
                display = rejectdisplay
            End If
        Catch ex As Exception
            ShowError("Error modeldisplay", ex)
        End Try

        Return display
    End Function

    Public Function modeldisplay2(ByVal client_no As Integer)
        Dim display As String = ""

        Try
            If client_no = 0 Then
                display = client1display
            ElseIf client_no = 1 Then
                display = client2display
            Else
                display = rejectdisplay
            End If
        Catch ex As Exception
            ShowError("Error modeldisplay", ex)
        End Try

        Return display
    End Function

    Public Sub showIpClient(ByVal client_no As Integer, ByVal ipadrs As String)
        Try
            If client_no = 0 Then
                Handheld1.Text = ipadrs
            ElseIf client_no = 1 Then
                Handheld2.Text = ipadrs
            ElseIf client_no = 2 Then
                Handheld3.Text = ipadrs
            ElseIf client_no = 3 Then
                Handheld4.Text = ipadrs
            ElseIf client_no = 4 Then
                Handheld5.Text = ipadrs
            ElseIf client_no = 5 Then
                Handheld6.Text = ipadrs
            ElseIf client_no = 6 Then
                Handheld7.Text = ipadrs
            ElseIf client_no = 7 Then
                Handheld8.Text = ipadrs
            Else
                Handheld9.Text = ipadrs
            End If
        Catch ex As Exception
            ShowError("Error ShowIPClient", ex)
        End Try

    End Sub

    Public Sub showIpClient2(ByVal client_no As Integer, ByVal ipadrs As String)
        Try
            'tambahan 6/10/20
            If jenis_so = "BPBBKL" Or jenis_so = "BPBNPS" Then
                If client_no = 0 Then
                    DCP1.Text = ipadrs
                ElseIf client_no = 1 Then
                    DCP2.Text = ipadrs
                ElseIf client_no = 2 Then
                    DCP3.Text = ipadrs
                Else
                    DCP4.Text = ipadrs
                End If
            Else
                If client_no = 0 Then
                    DCP1.Text = ipadrs
                Else
                    DCP2.Text = ipadrs
                End If
            End If
        Catch ex As Exception
            ShowError("Error ShowIPClient", ex)
        End Try

    End Sub

    Public Sub showIpClient3(ByVal client_no As Integer, ByVal ipadrs As String)
        Try
            If client_no = 0 Then
                Plano1.Text = ipadrs
            Else
                Plano2.Text = ipadrs
            End If
        Catch ex As Exception
            ShowError("Error ShowIPClient", ex)
        End Try

    End Sub

    Sub ShowError(ByVal MsgStr As String, ByVal excp As Exception)
        Dim util As New Utility
        Dim strmsg As String = MsgStr & " : " & vbCrLf & excp.Message
        MessageBox.Show(strmsg, "Error Program", MessageBoxButtons.OK, MessageBoxIcon.Error)
        util.TraceLogTxt(MsgStr & " " & vbCrLf & excp.Message & excp.StackTrace)
    End Sub

#Region "Tampil Login"
    Private Function TampilLoginClient(ByVal username As String, ByVal password As String, ByVal status As String, ByVal Keterangan As String) As String
        Dim data As String = ""
        Try
            If tabel_name = "dcp_boxplu" Then
                data += Padcenter("BPB Toko", 20) + Chr(2)
            ElseIf tabel_name = "CekPlanogram" Then
                data += Padcenter(jenis_so, 20) + Chr(2)
            ElseIf tabel_name = "ptag_wdcp" Then
                data += Padcenter("Cetak Price Tag ", 20) + Chr(2)
            ElseIf tabel_name.ToLower = "bpbnps_wdcp" Then
                data += Padcenter("BPB NPS Toko", 20) + Chr(2)
            ElseIf tabel_name = "CekKesegaran" Then
                data += Padcenter(jenis_so, 20) + Chr(2)

            Else
                data += Padcenter("SO " & jenis_so, 20) + Chr(2)
            End If
            data += Padcenter("********************", 20) + Chr(2)
            data += ("User:" & username).PadRight(20) + Chr(2)
            data += "".PadRight(20) + Chr(2)
            If status = "5" Then
                data += "NewPassw:" & password.PadRight(20) + Chr(2)
            Else
                data += "Password:" & password.PadRight(20) + Chr(2)
            End If

            If status = "2" Then
                data += "ID Tdk Terdaftar".PadRight(20) + Chr(2)
            ElseIf status = "3" Then
                data += "IP Dipakai".PadRight(20) + Chr(2)
            ElseIf status = "4" Then
                data += "Password salah".PadRight(20) + Chr(2)
            ElseIf status = "5" Then
                data += "Masukan password baru".PadRight(20) + Chr(2)
            ElseIf status = "6" Then
                data += "Sukses, login kembali".PadRight(20) + Chr(2)
            ElseIf status = "7" Then
                data += "Gagal UPDATE password".PadRight(20) + Chr(2)
            Else
                data += "".PadRight(20) + Chr(2)
            End If
            If Keterangan.Trim = "" Then
                data += "".PadRight(20)
                data += "".PadRight(20)
            Else
                data += Keterangan.PadRight(20)
                data += "".PadRight(20)
            End If
            data += Chr(3)

            data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "U"
            If username = "" Then
                data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "U"
            ElseIf status = "5" Then
                data += "1" + "e" + "J" + Chr((15 * 16) + 0) + "X"
            ElseIf status = "7" Then
                data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "U"
            Else
                data += "1" + "e" + "J" + Chr((15 * 16) + 0) + "P"
            End If

        Catch ex As Exception
            ShowError("Error set TampilLogin", ex)
        End Try

        Return data
    End Function

    Private Function TampilLoginServer(ByVal username As String, ByVal password As String, ByVal status As String) As String
        Dim data_result As String = ""
        Dim header As String = ""
        If tabel_name = "dcp_boxplu" Then
            header = "BPB Toko"
        ElseIf tabel_name = "CekPlanogram" Then
            header = jenis_so
        ElseIf tabel_name = "bpbnps_wdcp" Then
            header = "BPB NPS Toko"
        ElseIf tabel_name = "CekKesegaran" Then
            header = jenis_so
        Else
            header = "SO " & jenis_so
        End If
        Dim model() = {Padcenter(header, 20), "====================", "", "", "", "", "", ""}
        Dim i As Integer = 0

        Try
            If status = "" Then
                model(2) = "User:" & username
                model(4) = "Password:" & password
            ElseIf status = "2" Then
                model(2) = "User:"
                model(4) = "Password:"
                model(5) = "ID Tdk Terdaftar"
            ElseIf status = "3" Then
                model(2) = "User:"
                model(4) = "Password:"
                model(5) = "ID Dipakai"
            ElseIf status = "4" Then
                model(2) = "User:" & username
                model(4) = "Password:"
                model(5) = "Password salah"
            ElseIf status = "5" Then
                model(2) = "User:" & username
                model(4) = "NewPassw:"
                model(5) = "Masukan password baru"
            ElseIf status = "6" Then
                model(2) = "User:" & username
                model(4) = "Password:" & password
                model(5) = "Sukses, login kembali"
            ElseIf status = "7" Then
                model(2) = "User:" & username
                model(4) = "Password:" & password
                model(5) = "Gagal UPDATE password"
            End If

            For i = 0 To 6
                data_result += model(i) + Chr(10)
            Next
            data_result += model(7)

        Catch ex As Exception
            ShowError("Error set TampilLoginServer", ex)
        End Try

        Return data_result
    End Function

#End Region

#Region "Tampil Lokasi"
    Private Function TampilLokasiClient() As String
        Dim data As String = ""
        Try
            data += Padcenter("SO " & jenis_so, 20) + Chr(2)
            data += Padcenter("********************", 20) + Chr(2)

            data += "1. Toko".PadRight(20) + Chr(2)
            data += "2. Gudang".PadRight(20) + Chr(2)

            If mainTTL3 Then
                data += "3. Barang Rusak".PadRight(20) + Chr(2)
            Else
                data += "".PadRight(20) + Chr(2)
            End If

            data += "Pilih:".PadRight(20) + Chr(2)
            data += "".PadRight(20) + Chr(2)
            data += "".PadRight(20)
            data += Chr(3)

            data += "1" + "f" + "G" + Chr((1 * 16) + 0) + "L"
            data += "1" + "f" + "G" + Chr((1 * 16) + 0) + "L"
        Catch ex As Exception
            ShowError("Error set TampilLokasi", ex)
        End Try

        Return data
    End Function

    Private Function TampilLokasiServer() As String
        Dim data_result As String = ""
        Dim model() = {Padcenter("SO " & jenis_so, 20), "====================", "", "", "", "", "", ""}
        Dim i As Integer = 0

        Try
            model(2) = "1. Toko"
            model(3) = "2. Gudang"
            If mainTTL3 Then
                model(4) = "3. Barang Rusak"
            End If
            model(5) = "Pilih:"

            For i = 0 To 6
                data_result += model(i) + Chr(10)
            Next

            data_result += model(7)

        Catch ex As Exception
            ShowError("Error set TampilLoginServer", ex)
        End Try

        Return data_result
    End Function
#End Region

#Region "Tampil Deskripsi"
    Private Function TampilDeskripsiClient(ByVal barcode_plu As String, ByVal deskripsi As String,
                                           ByVal rak As String, ByVal qty_total As String,
                                           ByVal qty_com As String, Optional ByVal statusBarcode As String = "",
                                           Optional ByVal total_rak As Integer = 0) As String
        Dim data As String = ""

        Try
            If tabel_name = "dcp_boxplu" Then
                'Terima Barang
                data += Padcenter("BPB Toko", 20) + Chr(2)
                data += Padcenter("********************", 20) + Chr(2)

                If barcode_plu = "" Then
                    data += "Prod:".PadRight(20) + Chr(2)
                Else
                    barcode_plu = "Prod:" & barcode_plu
                    data += barcode_plu.PadRight(20) + Chr(2)
                End If

                data += "Desc:".PadRight(20) + Chr(2)
                If deskripsi = "" Then
                    data += "".PadRight(20) + Chr(2)
                Else
                    data += deskripsi.PadRight(20) + Chr(2)
                End If

                data += ""
                data += "QTY:" + qty_total.PadRight(20) + Chr(2)
                data += ""
                data += "REV:" + Chr(2)
                data += Chr(3)

                If barcode_plu <> "" And deskripsi <> "Tidak Ditemukan" Then 'ada data
                    data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                    data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"
                    data += "0" + "H" + "E" + Chr((8 * 16) + 0) + "Q"
                    data += "0" + "H" + "E" + Chr((8 * 16) + 0) + "Q"
                    data += "1" + "H" + "E" + Chr((8 * 16) + 0) + "Q"
                    data += "0"
                ElseIf barcode_plu = "" And deskripsi = "Tidak Ditemukan" Then 'data tidak ditemukan
                    data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                    data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"
                    data += "0"
                    data += "0"
                    data += "0"
                    data += "0"
                Else 'belum ada data
                    data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                    data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"
                    data += "0"
                    data += "0"
                    data += "0"
                    data += "0"
                End If
            ElseIf tabel_name.ToUpper.Contains("OA") Then
                'revisi Memo No. 208 - CPS - 20
                '13/10/2020
                'SO Aktiva
                data += Padcenter(" Aktiva", 20) + Chr(2)
                data += Padcenter("********************", 20) + Chr(2)

                If barcode_plu = "" Then
                    data += "Prod:".PadRight(20) + Chr(2)
                Else
                    barcode_plu = "Prod:" & barcode_plu
                    data += barcode_plu.PadRight(20) + Chr(2)
                End If
                data += "Desc:".PadRight(20) + Chr(2)
                If deskripsi = "" Then
                    data += "".PadRight(20) + Chr(2)
                ElseIf deskripsi = "Tidak Ditemukan" Then
                    data += "Tidak Ditemukan".PadRight(20) + Chr(2)
                Else
                    If deskripsi.Trim.Length > 20 Then
                        deskripsi = deskripsi.Substring(0, 20)
                    End If
                    data += deskripsi.PadRight(20) + Chr(2)
                End If

                If qty_total = "" Then
                    data += "QTY Baik:".PadRight(20) + Chr(2)
                Else
                    qty_total = "QTY Baik:" & qty_total
                    data += qty_total.PadRight(20) + Chr(2)
                End If
                'revisi 18/11/2020
                If rak = "False" Then
                    data += "QTY Rusak:".PadRight(20) + Chr(2)
                Else
                    data += "".PadRight(20) + Chr(2)
                End If

                data += statusBarcode + Chr(2)
                data += Chr(3)

                If barcode_plu <> "" And deskripsi.Length > 0 Then 'ada data
                    data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                    data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"

                    If qty_total = "" Then
                        data += "1" + "f" + "J" + Chr((3 * 16) + 0) + "Q"
                    ElseIf qty_com = "" And qty_total <> "" And rak = "False" Then
                        data += "1" + "g" + "K" + Chr((3 * 16) + 0) + "K"
                    Else
                        data += "0"
                    End If

                    data += "0"
                    data += "0"
                Else
                    If statusBarcode.Length > 0 And statusBarcode.Contains("AT baru") Then
                        data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "S"
                        If qty_total = "" Then
                            data += "1" + "f" + "J" + Chr((3 * 16) + 0) + "Q" 'REVISI
                        ElseIf qty_com = "" And qty_total <> "" And rak = "False" Then
                            data += "1" + "g" + "K" + Chr((3 * 16) + 0) + "K"
                        Else
                            data += "0"
                        End If

                        data += "0"
                        data += "0"
                    Else
                        data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                        data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"
                        data += "0"
                        data += "0"
                        data += "0"
                        data += "0"
                    End If
                End If

            Else
                'SO

                'A
                If jenis_so = "Khusus" Then
                    data += Padcenter("SO Produk " & jenis_so, 20) + Chr(2)
                Else
                    data += Padcenter("SO " & jenis_so & " - " & lokasi_so, 20) + Chr(2)
                End If

                'B
                data += Padcenter("********************", 20) + Chr(2)

                'C
                If barcode_plu = "" Then
                    data += "Prod:".PadRight(20) + Chr(2)
                Else
                    barcode_plu = "Prod:" & barcode_plu
                    data += barcode_plu.PadRight(20) + Chr(2)
                End If

                'D
                data += "Desc:".PadRight(20) + Chr(2)

                'E
                If deskripsi = "" Then
                    data += "".PadRight(20) + Chr(2)
                ElseIf deskripsi = "Tidak Ditemukan" Then
                    data += "Tidak Ditemukan".PadRight(20) + Chr(2)
                Else
                    data += deskripsi.PadRight(20) + Chr(2)
                End If

                'F
                If rak = "" Then
                    data += "Rak/Shelf:".PadRight(20) + Chr(2)
                Else
                    If total_rak > 1 Then
                        rak = "Rak/Shelf:" & rak & "*"
                    Else
                        rak = "Rak/Shelf:" & rak
                    End If
                    data += rak.PadRight(20) + Chr(2)
                End If

                'G
                qty_total = "TTL:" + qty_total

                If mode_run = "B" Or mode_run = "E" Then
                    If lokasi_so.ToLower = "toko" Then
                        If isLTF = True Then
                            qty_com = "COM:" + qty_com
                        Else
                            qty_com = ""
                        End If
                    Else
                        qty_com = ""
                    End If
                Else
                    qty_com = ""
                End If

                data += qty_total.PadRight(10) + qty_com.PadRight(9) + Chr(2)

                'revisi lepas flag CBR SO PRODUK KHUSUS
                '09/09/2021
                TraceLog("isMain 1230 Client: " & FormatBaru1230)
                If FormatBaru1230 Then
                    If isFlagCBR = False And statusBarcode <> "CBRN" Then
                        data += "QTY:" + Chr(2)
                    Else
                        data += "" + Chr(2)
                    End If
                Else
                    If isFlagCBR = False Then
                        data += "QTY:" + Chr(2)
                    Else
                        If statusBarcode = "CBRN" And mode_run = "B" Then
                            data += "" + Chr(2)
                        Else
                            data += "QTY:" + Chr(2)
                        End If
                    End If
                End If

                data += Chr(3)

                If barcode_plu = "" And deskripsi = "" Then 'belum ada data
                    data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                    data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"

                    data += "0" + "A" + "A" + Chr(16) + " "
                    data += "0" + "A" + "A" + Chr(16) + " "
                    data += "0" + "A" + "A" + Chr(16) + " "
                    data += "0" + "A" + "A" + Chr(16) + " "
                ElseIf barcode_plu <> "" And deskripsi <> "Tidak Ditemukan" Then 'ada data
                    'Revisi 20 November 2019 (Memo 1081/CPS/19)
                    'Hitung lokasi RAK untuk item, Jika ada lebih dari 1 lokasi aktifkan fitur NEXT WDCP
                    If statusBarcode = "CBRN" Then
                        data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                        data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"
                        data += "0" + "A" + "A" + Chr(16) + " "
                        data += "0" + "A" + "A" + Chr(16) + " "
                        data += "0" + "A" + "A" + Chr(16) + " "
                        data += "0" + "A" + "A" + Chr(16) + " "
                    Else
                        If total_rak > 1 Then
                            data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "S"
                            data += "1" + "h" + "E" + Chr((8 * 16) + 0) + "Q"

                            'yg tidak diaktifkan tetap 5 byte; contoh seperti di bawah; sebenarnya 5 byte-nya nya tidak harus seperti itu, yg penting byte1 "0" + 4 byte sembarangan saja asal tidak chr(0)
                            data += "0" + "A" + "A" + Chr(16) + " "
                            data += "0" + "A" + "A" + Chr(16) + " "
                            data += "0" + "A" + "A" + Chr(16) + " "
                            data += "1"
                        Else
                            data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "S"
                            data += "1" + "h" + "E" + Chr((8 * 16) + 0) + "Q"
                            data += "0" + "A" + "A" + Chr(16) + " "
                            data += "0" + "A" + "A" + Chr(16) + " "
                            data += "0" + "A" + "A" + Chr(16) + " "
                            data += "0" + "A" + "A" + Chr(16) + " "
                        End If

                    End If
                ElseIf barcode_plu <> "" And deskripsi = "Tidak Ditemukan" Then 'data tidak ditemukan
                    data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "S"
                    data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"

                    data += "0" + "A" + "A" + Chr(16) + " "
                    data += "0" + "A" + "A" + Chr(16) + " "
                    data += "0" + "A" + "A" + Chr(16) + " """
                    data += "0" + "A" + "A" + Chr(16) + " "
                End If
            End If

        Catch ex As Exception
            ShowError("Error set TampilDeskripsiClient", ex)
        End Try

        Return data
    End Function

    Private Function TampilDeskripsiServer(ByVal barcode_plu As String, ByVal deskripsi As String, ByVal rak As String,
                                           ByVal qty_total As String, ByVal qty_com As String,
                                           Optional ByVal statusBarcode As String = "") As String
        Dim data_result As String = ""
        Dim header As String = ""
        If tabel_name = "dcp_boxplu" Then
            header = "BPB Toko"
        ElseIf tabel_name.ToUpper.StartsWith("BPBBKL_WDCP") Then
            header = "BPB BKL Toko"
        ElseIf tabel_name.ToUpper.StartsWith("BPBNPS_WDCP") Then
            header = "BPB NPS Toko"
        Else
            If jenis_so = "Khusus" Then
                header = "SO Produk " & jenis_so
            Else
                header = "SO " & jenis_so & " - " & lokasi_so

            End If
        End If

        Dim model() = {Padcenter(header, 20), "====================", "", "", "", "", "", ""}
        Dim i As Integer = 0

        Try
            If tabel_name = "dcp_boxplu" Then
                model(0) = Padcenter(header, 20)
                If barcode_plu = "" Then
                    model(2) = "Prod:"
                Else
                    model(2) = "Prod:" & barcode_plu
                End If

                model(3) = "Desc:"
                If deskripsi = "" Then
                    model(4) = ""
                Else
                    model(4) = deskripsi
                End If
                model(5) = "QTY:" + qty_total.PadRight(20)

                For i = 0 To 6
                    data_result += model(i) + Chr(10)
                Next
                data_result += model(7)
            ElseIf tabel_name.ToUpper.Contains("BPBBKL_WDCP") Then
                model(0) = Padcenter(header, 20)
                If barcode_plu = "" Then
                    model(2) = "Prod:"
                Else
                    model(2) = "Prod:" & barcode_plu
                End If
                If deskripsi <> "" Then
                    model(3) = "Desc:" & deskripsi
                Else
                    model(3) = "Desc:"
                End If
                model(4) = "EXP:"
                model(5) = "QTY:"
                model(6) = ""
                For i = 0 To 6
                    data_result += model(i) + Chr(10)
                Next
                data_result += model(7)
            ElseIf tabel_name.ToUpper.Contains("BPBNPS_WDCP") Then
                model(0) = Padcenter(header, 20)
                If barcode_plu = "" Then
                    model(2) = "Prod:"
                Else
                    model(2) = "Prod:" & barcode_plu
                End If
                If deskripsi <> "" Then
                    model(3) = "Desc:" & deskripsi
                Else
                    model(3) = "Desc:"
                End If
                model(4) = "EXP:"
                model(5) = "QTY:"
                model(6) = ""
                For i = 0 To 6
                    data_result += model(i) + Chr(10)
                Next
                data_result += model(7)
            ElseIf tabel_name.ToUpper.Contains("OA") Then
                model(0) = Padcenter(header, 20)
                If barcode_plu = "" Then
                    model(2) = "Prod:"
                Else
                    model(2) = "Prod:" & barcode_plu
                End If

                model(3) = "Desc:"
                If deskripsi = "" Then
                    model(4) = ""
                Else
                    If deskripsi.Trim.Length > 20 Then
                        deskripsi = deskripsi.Substring(0, 20)
                    End If
                    model(4) = deskripsi
                End If
                model(5) = "QTY BAIK:" + qty_total.PadRight(20)
                If rak = "False" Then
                    model(6) = "QTY RUSAK" + qty_com.PadRight(20)
                Else
                    model(6) = ""
                End If
                model(7) = statusBarcode.PadRight(20)

                For i = 0 To 7
                    data_result += model(i) + Chr(10)
                Next
            Else
                model(0) = Padcenter(header, 20)
                If barcode_plu = "" Then
                    model(2) = "Prod:"
                Else
                    model(2) = "Prod:" & barcode_plu
                End If

                model(3) = "Desc:"
                If deskripsi = "" Then
                    model(4) = ""
                Else
                    model(4) = deskripsi
                End If

                If rak = "" Then
                    model(5) = "Rak/Shelf:"
                Else
                    model(5) = "Rak/Shelf:" & rak
                End If

                If mode_run = "B" Or mode_run = "E" Then
                    If lokasi_so.ToLower = "toko" Then
                        If isLTF = True Then
                            model(6) = "TTL:" + qty_total.PadRight(6) + "COM:" + qty_com.PadRight(6)
                        Else
                            model(6) = "TTL:" + qty_total.PadRight(20)
                        End If
                    Else
                        model(6) = "TTL:" + qty_total.PadRight(20)
                    End If
                Else
                    model(6) = "TTL:" + qty_total.PadRight(20)
                End If

                TraceLog("isMain 1230 Server: " & FormatBaru1230)
                If FormatBaru1230 Then
                    If isFlagCBR = False And statusBarcode <> "CBRN" Then
                        model(7) = "QTY:"
                    Else
                        model(7) = ""
                    End If
                Else
                    If isFlagCBR = False Then
                        model(7) = "QTY:"
                    Else
                        If statusBarcode = "CBRNE" Then
                            model(7) = ""
                        Else
                            model(7) = "QTY:"
                        End If
                    End If
                End If

                For i = 0 To 6
                    data_result += model(i) + Chr(10)
                Next

                data_result += model(7)
            End If

        Catch ex As Exception
            ShowError("Error set TampilDeskripsiServer", ex)
        End Try

        Return data_result
    End Function

    Private Function TampilDeskripsiBKL(ByVal barcode_plu As String, ByVal deskripsi As String, ByVal qty As String, ByVal tgl_exp As String, ByVal feedback As String, ByVal statusDesc As String, ByVal statusQTY As String, ByVal statusEXP As String, Optional ByVal deskripsi2 As String = "", Optional ByVal fraction_pcs As String = "") As String
        Dim data As String = ""

        Try
            'Terima Barang
            data = ""
            data += Padcenter("BPB BKL Toko", 20) + Chr(2) 'a
            If tabel_name.Trim.ToLower = "bpbbkl_wdcp" Then
                If fraction_pcs = "" Then
                    data += "Fraction = ".PadRight(20) + Chr(2) 'b

                Else
                    fraction_pcs = "Fraction = " & fraction_pcs
                    data += fraction_pcs.PadRight(20) + Chr(2)

                End If
            Else
                data += Padcenter("********************", 20) + Chr(2) 'b

            End If



            'c
            If barcode_plu = "" Then
                data += "Prod:".PadRight(20) + Chr(2)
            Else
                barcode_plu = "Prod:" & barcode_plu
                data += barcode_plu.PadRight(20) + Chr(2)
            End If


            If statusDesc = "1" Or statusDesc = "3" Then
                data += "Desc:".PadRight(20) + Chr(2)
            Else
                If deskripsi = "" Then
                    data += "Desc:".PadRight(20) + Chr(2)
                Else
                    deskripsi = "Desc:" & Strings.Left(deskripsi, 13)
                    data += deskripsi.PadRight(20) + Chr(2)
                End If
            End If

            If statusEXP = "1" Or statusEXP = "3" Then
                data += "EXP(ddMMyy):".PadRight(20) + Chr(2) 'g
            Else
                If tgl_exp = "" Then
                    data += "EXP(ddMMyy):".PadRight(20) + Chr(2) 'g
                Else
                    tgl_exp = "EXP:" & tgl_exp
                    data += tgl_exp.PadRight(20) + Chr(2) 'g

                End If
            End If

            If statusQTY = "1" Or statusQTY = "3" Or statusQTY = "4" Or statusQTY = "5" Then
                data += "QTY(In Fraction):".PadRight(20) + Chr(2) 'e
            Else
                If qty = "" Then
                    data += "QTY(In Fraction):".PadRight(20) + Chr(2) 'e
                Else
                    qty = "QTY(In Fraction):" & qty
                    data += qty.PadRight(20) + Chr(2)
                End If
            End If



            If statusQTY = "1" Then
                feedback = "Kesalahan data"
                data += feedback.PadRight(20) + Chr(3)
            ElseIf statusQTY = "3" Then
                feedback = "QTY(" & deskripsi2 & ") melebihi"
                data += feedback.PadRight(20) + Chr(3)
            ElseIf statusQTY = "4" Then
                feedback = "QTY 0"
                data += feedback.PadRight(20) + Chr(3)
            ElseIf statusQTY = "5" Then
                feedback = "QTY tdk sesuai minor"
                data += feedback.PadRight(20) + Chr(3)
            ElseIf statusEXP = "1" Then
                feedback = "Format EXP salah"
                data += feedback.PadRight(20) + Chr(3)
            ElseIf statusEXP = "3" Then
                feedback = "Tgl EXP tdk sesuai"
                data += feedback.PadRight(20) + Chr(3)
            ElseIf statusDesc = "1" Then
                feedback = "Tdk Terdaftar"
                data += feedback.PadRight(20) + Chr(3)
                'ElseIf statusDesc = "3" Then
                '    feedback = "Min.Or < Fraction"
                '    data += feedback.PadRight(20) + Chr(3)
            Else
                data += feedback.PadRight(20) + Chr(3)
            End If
            'End If

            If barcode_plu <> "" And deskripsi <> "Tidak Ditemukan" Then 'ada data

                data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"


                If tgl_exp = "" Then
                    data += "1" + "e" + "M" + Chr((8 * 16) + 0) + "T"

                ElseIf qty = "" And tgl_exp <> "" And (statusEXP = "2" Or statusEXP = "") Then
                    data += "1" + "f" + "R" + Chr((3 * 16) + 0) + "Q"

                ElseIf feedback.Length > 1 Then
                    If statusEXP = "1" Or statusEXP = "3" Then
                        data += "1" + "e" + "M" + Chr((8 * 16) + 0) + "T"
                    ElseIf statusQTY = "1" Or statusQTY = "3" Or statusQTY = "5" Then
                        data += "1" + "f" + "R" + Chr((3 * 16) + 0) + "Q"

                    Else
                        data += "1" + "A" + "J" + Chr((15 * 16) + 0) + "F"
                    End If
                Else
                    data += "0"
                    'Else
                    '    data += "1" + "h" + "J" + Chr((15 * 16) + 0) + "F"

                End If

            ElseIf barcode_plu = "" And deskripsi = "Tidak Ditemukan" Then 'data tidak ditemukan
                data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"
                data += "0"
                data += "0"
                data += "0"
                data += "0"
            Else 'belum ada data
                data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"
                data += "0"
                data += "0"
                data += "0"
                data += "0"
            End If



        Catch ex As Exception
            ShowError("Error set TampilDeskripsiBKL", ex)
        End Try

        Return data
    End Function

    Private Function TampilDeskripsiNPS(ByVal barcode_plu As String, ByVal deskripsi As String, ByVal qty As String, ByVal tgl_exp As String, ByVal feedback As String, ByVal statusDesc As String, ByVal statusQTY As String, ByVal statusEXP As String, Optional ByVal deskripsi2 As String = "", Optional ByVal total_rak As Integer = 0) As String
        Dim data As String = ""

        Try
            'Terima Barang
            data = ""
            data += Padcenter("BPB NPS Toko", 20) + Chr(2) 'a
            data += Padcenter("********************", 20) + Chr(2) 'b

            'c
            If barcode_plu = "" Then
                data += "Prod:".PadRight(20) + Chr(2)
            Else
                barcode_plu = "Prod:" & barcode_plu
                data += barcode_plu.PadRight(20) + Chr(2)
            End If


            If statusDesc = "1" Or statusDesc = "3" Then
                data += "Desc:".PadRight(20) + Chr(2)
            Else
                If deskripsi = "" Then
                    data += "Desc:".PadRight(20) + Chr(2)
                Else
                    deskripsi = "Desc:" & Strings.Left(deskripsi, 13)
                    data += deskripsi.PadRight(20) + Chr(2)
                End If
            End If

            If statusEXP = "1" Or statusEXP = "3" Then
                data += "EXP(ddMMyy):".PadRight(20) + Chr(2) 'g
            Else
                If tgl_exp = "" Then
                    data += "EXP(ddMMyy):".PadRight(20) + Chr(2) 'g
                Else
                    tgl_exp = "EXP:" & tgl_exp
                    data += tgl_exp.PadRight(20) + Chr(2) 'g

                End If
            End If

            If statusQTY = "1" Or statusQTY = "3" Or statusQTY = "4" Or statusQTY = "5" Then
                data += "QTY(In Fraction):".PadRight(20) + Chr(2) 'e
            Else
                If qty = "" Then
                    data += "QTY(In Fraction):".PadRight(20) + Chr(2) 'e
                Else
                    qty = "QTY(In Fraction):" & qty
                    data += qty.PadRight(20) + Chr(2)
                End If
            End If

            If statusQTY = "1" Then
                feedback = "Kesalahan data"
                data += feedback.PadRight(20) + Chr(3)
            ElseIf statusQTY = "3" Then
                feedback = "QTY(" & deskripsi2 & ") melebihi"
                data += feedback.PadRight(20) + Chr(3)
            ElseIf statusQTY = "4" Then
                feedback = "QTY 0"
                data += feedback.PadRight(20) + Chr(3)
            ElseIf statusQTY = "5" Then
                feedback = "QTY tdk sesuai minor"
                data += feedback.PadRight(20) + Chr(3)
            ElseIf statusEXP = "1" Then
                feedback = "Format EXP salah"
                data += feedback.PadRight(20) + Chr(3)
            ElseIf statusEXP = "3" Then
                feedback = "Tgl EXP tdk sesuai"
                data += feedback.PadRight(20) + Chr(3)
            ElseIf statusDesc = "1" Then
                feedback = "Tdk Terdaftar"
                data += feedback.PadRight(20) + Chr(3)
                'ElseIf statusDesc = "3" Then
                '    feedback = "Min.Or < Fraction"
                '    data += feedback.PadRight(20) + Chr(3)
            Else
                data += feedback.PadRight(20) + Chr(3)
            End If
            'End If

            If barcode_plu <> "" And deskripsi <> "Tidak Ditemukan" Then 'ada data

                data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"


                If tgl_exp = "" Then
                    data += "1" + "e" + "M" + Chr((8 * 16) + 0) + "T"

                ElseIf qty = "" And tgl_exp <> "" And (statusEXP = "2" Or statusEXP = "") Then
                    data += "1" + "f" + "R" + Chr((3 * 16) + 0) + "Q"

                ElseIf feedback.Length > 1 Then
                    If statusEXP = "1" Or statusEXP = "3" Then
                        data += "1" + "e" + "M" + Chr((8 * 16) + 0) + "T"
                    ElseIf statusQTY = "1" Or statusQTY = "3" Or statusQTY = "5" Then
                        data += "1" + "f" + "R" + Chr((3 * 16) + 0) + "Q"

                    Else
                        data += "1" + "A" + "J" + Chr((15 * 16) + 0) + "F"
                    End If
                Else
                    data += "0"
                    'Else
                    '    data += "1" + "h" + "J" + Chr((15 * 16) + 0) + "F"

                End If

            ElseIf barcode_plu = "" And deskripsi = "Tidak Ditemukan" Then 'data tidak ditemukan
                data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"
                data += "0"
                data += "0"
                data += "0"
                data += "0"
            Else 'belum ada data
                data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"
                data += "0"
                data += "0"
                data += "0"
                data += "0"
            End If

        Catch ex As Exception
            ShowError("Error set TampilDeskripsiBKL", ex)
        End Try

        Return data
    End Function

    'revisi Memo No. 208 - CPS - 20
    '13/10/2020

    Private Function TampilDeskripsiListClient(ByVal list As List(Of ClsAktiva), ByVal indeks As Integer) As String
        Dim data As String = ""
        Dim count As Integer = list.Count
        Dim nseri As String = ""
        Dim desc As String = ""
        Dim lgth As Integer = 0
        Try
            data += Padcenter("List SO Aktiva", 20) + Chr(2)
            data += Padcenter("********************", 20) + Chr(2)

            'nseri
            nseri = indeks + 1 & ". " & list(indeks).NSeri
            data += nseri.PadRight(20) + Chr(2)

            lgth = list(indeks).Deskripsi.Length

            If lgth > 40 Then
                lgth = 40
            End If

            data += "Desc:".PadRight(20) + Chr(2)

            If lgth < 20 Then
                desc = list(indeks).Deskripsi
            Else
                desc = list(indeks).Deskripsi.Substring(0, 20)
            End If

            desc = desc.Trim
            data += desc.PadRight(20) + Chr(2)

            If lgth >= 20 Then
                desc = list(indeks).Deskripsi.Substring(20, lgth - 40 + 20)
                desc = desc.Trim
                data += desc.PadRight(20) + Chr(2)
            Else
                data += "".PadRight(20) + Chr(2)
            End If

            If count > 1 Then
                data += "0. NEXT".PadRight(20) + Chr(2)
            Else
                data += "".PadRight(20) + Chr(2)

            End If
            data += "Pilih:".PadRight(20) + Chr(2)
            data += Chr(3)
            data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"
            data += "1" + "h" + "G" + Chr((15 * 16) + 0) + "I"
            data += "0"
            data += "0"
        Catch ex As Exception
            ShowError("Error set TampilDeskripsiClient" & ex.Message & ex.StackTrace, ex)
        End Try

        Return data
    End Function
    'revisi Memo No. 208 - CPS - 20
    '13/10/2020

    Private Function TampilDeskripsiListServer(ByVal list As List(Of ClsAktiva), ByVal indeks As Integer) As String
        Dim data_result As String = ""
        Dim header As String = ""
        Dim lgth As Integer
        Dim desc As String = ""
        tabel_name.ToUpper.Contains("OA")
        header = "List SO Aktiva"

        Dim model() = {Padcenter(header, 20), "====================", "", "", "", "", "", ""}
        Dim i As Integer = 0
        model(0) = Padcenter(header, 20)
        If list(indeks).NSeri = "" Then
            model(2) = "Prod:"
        Else
            model(2) = indeks + 1 & ". " & list(indeks).NSeri
        End If

        model(3) = "Desc:"

        If list(indeks).Deskripsi = "" Then
            model(4) = ""
        Else
            lgth = list(indeks).Deskripsi.Trim.Length
            If lgth < 20 Then
                lgth = list(indeks).Deskripsi.Trim.Length
            Else
                lgth = 20
            End If
            desc = list(indeks).Deskripsi.Substring(0, lgth - 20 + 20)

            model(4) = desc
        End If
        If list(indeks).Deskripsi = "" Then
            model(5) = ""
        Else
            lgth = list(indeks).Deskripsi.Trim.Length
            If lgth > 40 Then
                lgth = 40
            End If
            If list(indeks).Deskripsi.Trim.Length > 20 Then
                desc = list(indeks).Deskripsi.Substring(20, lgth - 40 + 20)
            End If
            model(5) = desc
        End If

        If list.Count > 1 Then
            model(6) += "0. NEXT"
        Else
            model(6) += ""

        End If

        'model(6) = "0. NEXT"
        model(7) = "Pilih:" & indeks

        For i = 0 To 7
            data_result += model(i) + Chr(10)
        Next
        Return data_result

    End Function

    Private Function TampilDeskripsiKhususClient(ByVal barcode_plu As String, ByVal deskripsi As String,
                                           ByVal rak As String, ByVal qty_total As String,
                                           ByVal qty_com As String, Optional ByVal deskripsi2 As String = "",
                                           Optional ByVal total_rak As Integer = 0, Optional ByVal qty_ttl As String = "") As String
        Dim data As String = ""
        Dim qty_ttl1 As String = ""
        Try

            'SO
            If jenis_so = "Khusus" Then
                data += Padcenter("SO Produk " & jenis_so, 20) + Chr(2)
            Else
                data += Padcenter("SO " & jenis_so & " - " & lokasi_so, 20) + Chr(2)
            End If
            data += Padcenter("********************", 20) + Chr(2)

            If barcode_plu = "" Then
                data += "Prod:".PadRight(20) + Chr(2)
            Else
                barcode_plu = "Prod:" & barcode_plu
                data += barcode_plu.PadRight(20) + Chr(2)
            End If

            data += "Desc:".PadRight(20) + Chr(2)
            If deskripsi = "" Then
                data += "".PadRight(20) + Chr(2)
            ElseIf deskripsi = "Tidak Ditemukan" Then
                data += "Tidak Ditemukan".PadRight(20) + Chr(2)
            Else
                data += deskripsi.PadRight(20) + Chr(2)
            End If

            If rak = "" Then
                data += "Rak/Shelf:".PadRight(20) + Chr(2)
            Else
                If total_rak > 1 Then
                    rak = "Rak/Shelf:" & rak & "*"
                Else
                    rak = "Rak/Shelf:" & rak
                End If
                data += rak.PadRight(20) + Chr(2)
            End If

            qty_total = "TTLOLD:" + qty_total
            qty_ttl1 = "TTL:" + qty_ttl
            data += qty_total.PadRight(10) + qty_ttl1.PadRight(9) + Chr(2)
            If deskripsi2 = "CBRN" Then
                'data += "QTY:" + qty_ttl.PadRight(20) + Chr(2)
                data += "" + Chr(2)

            Else
                data += "QTY:" + Chr(2)
            End If
            data += Chr(3)

            If barcode_plu = "" And deskripsi = "" Then 'belum ada data
                data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"

                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
            ElseIf barcode_plu <> "" And deskripsi <> "Tidak Ditemukan" Then 'ada data
                'Revisi 20 November 2019 (Memo 1081/CPS/19)
                'Hitung lokasi RAK untuk item, Jika ada lebih dari 1 lokasi aktifkan fitur NEXT WDCP
                If total_rak > 1 Then
                    data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "S"
                    If deskripsi2 = "CBRN" Then
                        data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"
                    Else
                        data += "1" + "H" + "E" + Chr((8 * 16) + 0) + "Q"
                    End If
                    'yg tidak diaktifkan tetap 5 byte; contoh seperti di bawah; sebenarnya 5 byte-nya nya tidak harus seperti itu, yg penting byte1 "0" + 4 byte sembarangan saja asal tidak chr(0)
                    data += "0" + "A" + "A" + Chr(16) + " "
                    data += "0" + "A" + "A" + Chr(16) + " "
                    data += "0" + "A" + "A" + Chr(16) + " "
                    data += "1"
                Else
                    data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "S"
                    If deskripsi2 = "CBRN" Then
                        data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"
                    Else
                        data += "1" + "h" + "E" + Chr((8 * 16) + 0) + "Q"

                    End If

                    data += "0" + "A" + "A" + Chr(16) + " "
                    data += "0" + "A" + "A" + Chr(16) + " "
                    data += "0" + "A" + "A" + Chr(16) + " "
                    data += "0" + "A" + "A" + Chr(16) + " "
                End If
            ElseIf barcode_plu <> "" And deskripsi = "Tidak Ditemukan" Then 'data tidak ditemukan
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"

                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " """
                data += "0" + "A" + "A" + Chr(16) + " "
            End If

        Catch ex As Exception
            ShowError("Error set TampilDeskripsiClient", ex)
        End Try

        Return data
    End Function

    Private Function TampilDeskripsiKhususServer(ByVal barcode_plu As String, ByVal deskripsi As String, ByVal rak As String,
                                           ByVal qty_total As String, ByVal qty_com As String,
                                           Optional ByVal deskripsi2 As String = "") As String
        Dim data_result As String = ""
        Dim header As String = ""

        If jenis_so = "Khusus" Then
            header = "SO Produk " & jenis_so
        Else
            header = "SO " & jenis_so & " - " & lokasi_so

        End If
        Dim model() = {Padcenter(header, 20), "====================", "", "", "", "", "", ""}
        Dim i As Integer = 0
        Try

            model(0) = Padcenter(header, 20)
            If barcode_plu = "" Then
                model(2) = "Prod:"
            Else
                model(2) = "Prod:" & barcode_plu
            End If

            model(3) = "Desc:"
            If deskripsi = "" Then
                model(4) = ""
            Else
                model(4) = deskripsi
            End If

            If rak = "" Then
                model(5) = "Rak/Shelf:"
            Else
                model(5) = "Rak/Shelf:" & rak
            End If

            model(6) = "TTLOLD:" + qty_total.PadRight(6) + "TTL:" + qty_com.PadRight(6)


            model(7) = "QTY:"

            For i = 0 To 6
                data_result += model(i) + Chr(10)
            Next
            data_result += model(7)


        Catch ex As Exception
            ShowError("Error set TampilDeskripsiServer", ex)
        End Try

        Return data_result
    End Function

    Private Function TampilNextSOBIC_Client(ByVal barcode_plu As String, ByVal deskripsi As String,
                                            ByVal Page As Integer, ByVal TotalPage As Integer,
                                            ByVal dtRak As DataTable) As String
        Dim data As String = ""
        Try
            'SO
            data += Padcenter("Lokasi Rak (" & Page & "/" & TotalPage & ")", 20) + Chr(2)
            data += Padcenter("********************", 20) + Chr(2)

            If barcode_plu = "" Then
                data += "Prod:".PadRight(20) + Chr(2)
            Else
                barcode_plu = "Prod:" & barcode_plu
                data += barcode_plu.PadRight(20) + Chr(2)
            End If

            data += "Desc:".PadRight(20) + Chr(2)
            If deskripsi = "" Then
                data += "".PadRight(20) + Chr(2)
            ElseIf deskripsi = "Tidak Ditemukan" Then
                data += "Tidak Ditemukan".PadRight(20) + Chr(2)
            Else
                data += deskripsi.PadRight(20) + Chr(2)
            End If
            data += "Rak/Shelf:".PadRight(20) + Chr(2)
            If dtRak.Rows.Count > 0 Then
                For Each Dr As DataRow In dtRak.Rows
                    Dim rak As String = ""
                    Dim no_rak As String = ""
                    Dim no_shelf As String = ""
                    If Dr("NORAK") & "" = "" Then
                        data += "-".PadRight(20) + Chr(2)
                    Else
                        no_rak = CInt(Dr("NORAK"))
                        no_rak = no_rak.PadLeft(3, "0")
                        no_shelf = CInt(Dr("NOSHELF"))
                        no_shelf = no_shelf.PadLeft(3, "0")

                        rak = Dr("NO") & ". " & no_rak & "/" & no_shelf
                        data += rak.PadRight(20) + Chr(2)
                    End If
                Next
            Else
                data += "-".PadRight(20) + Chr(2)
            End If
            data += Chr(3)

            data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "S"
            'data += "0" + "A" + "A" + Chr(16) + " "
            data += "1" + "A" + "A" + Chr(16) + "R"
            'yg tidak diaktifkan tetap 5 byte; contoh seperti di bawah; sebenarnya 5 byte-nya nya tidak harus seperti itu, yg penting byte1 "0" + 4 byte sembarangan saja asal tidak chr(0)
            data += "0" + "A" + "A" + Chr(16) + " "
            data += "0" + "A" + "A" + Chr(16) + " "
            data += "0" + "A" + "A" + Chr(16) + " "
            data += "1"

        Catch ex As Exception
            ShowError("Error set TampilNextSOBIC_Client", ex)
        End Try

        Return data
    End Function

    Private Function TampilDeskripsiBazar(ByVal barcode_plu As String, ByVal deskripsi As String, ByVal qty As String, ByVal tgl_exp As String, ByVal qty_total As String,
                                          ByVal feedback As String, ByVal statusDesc As String, ByVal statusEXP As String, Optional ByVal deskripsi2 As String = "") As String
        Dim data As String = ""

        Try
            'Terima Barang
            data = ""
            data += Padcenter("SO Bazar", 20) + Chr(2) 'a
            data += Padcenter("********************", 20) + Chr(2) 'b

            'c
            If barcode_plu = "" Then
                data += "Prod:".PadRight(20) + Chr(2)
            Else
                barcode_plu = "Prod:" & barcode_plu
                data += barcode_plu.PadRight(20) + Chr(2)
            End If

            'd
            data += "Desc:".PadRight(20) + Chr(2)
            If deskripsi = "" Then
                data += "".PadRight(20) + Chr(2)
            Else
                data += deskripsi.PadRight(20) + Chr(2)
            End If

            'f
            If tgl_exp = "" Then
                data += "EXP(MMyyyy):".PadRight(20) + Chr(2)
            Else
                tgl_exp = "EXP(MMyyyy):" & tgl_exp
                data += tgl_exp.PadRight(20) + Chr(2)
            End If

            'g
            If qty = "" Then
                data += "QTY:".PadRight(20) + Chr(2) 'e
            Else
                data += "QTY:" + Chr(2)
            End If

            'h
            If statusEXP = "1" Then
                feedback = "Format EXP salah"
                data += feedback.PadRight(20) + Chr(3)
            ElseIf statusEXP = "3" Then
                feedback = "Tgl EXP tdk sesuai"
                data += feedback.PadRight(20) + Chr(3)
            ElseIf statusDesc = "1" Then
                feedback = "Tdk Terdaftar"
                data += feedback.PadRight(20) + Chr(3)
            Else
                data += feedback.PadRight(20) + Chr(3)
            End If

            If barcode_plu <> "" And deskripsi <> "Tidak Ditemukan" Then 'ada data

                data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"

                If tgl_exp = "" Then
                    data += "1" + "f" + "M" + Chr((6 * 16) + 0) + "T"

                ElseIf qty = "" And tgl_exp <> "" And (statusEXP = "2" Or statusEXP = "") Then
                    data += "1" + "g" + "E" + Chr((3 * 16) + 0) + "Q"

                ElseIf feedback.Length > 1 Then
                    If statusEXP = "1" Or statusEXP = "3" Then
                        data += "1" + "f" + "M" + Chr((8 * 16) + 0) + "T"
                    Else
                        data += "1" + "A" + "J" + Chr((15 * 16) + 0) + "F"
                    End If
                Else
                    data += "0"
                    'Else
                    '    data += "1" + "h" + "J" + Chr((15 * 16) + 0) + "F"
                End If

            ElseIf barcode_plu = "" And deskripsi = "Tidak Ditemukan" Then 'data tidak ditemukan
                data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
            Else 'belum ada data
                data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"

                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
            End If
        Catch ex As Exception
            ShowError("Error set TampilDeskripsiBazar", ex)
        End Try

        Return data
    End Function

    Private Function TampilNextSOBIC_Server(ByVal barcode_plu As String, ByVal deskripsi As String,
                                            ByVal Page As Integer, ByVal TotalPage As Integer,
                                            ByVal dtRak As DataTable) As String
        Dim data_result As String = ""
        Dim header As String = ""
        Try
            header = "Lokasi Rak (" & Page & "/" & TotalPage & ")"
            Dim model() = {Padcenter(header, 20), "====================", "", "", "", "", "", ""}
            Dim i As Integer = 0

            model(0) = Padcenter(header, 20)
            If barcode_plu = "" Then
                model(2) = "Prod:"
            Else
                model(2) = "Prod:" & barcode_plu
            End If

            model(3) = "Desc:"
            If deskripsi = "" Then
                model(4) = ""
            Else
                model(4) = deskripsi
            End If

            model(5) = "Rak/Shelf:"
            Dim IdxModel As Integer = 5
            If dtRak.Rows.Count > 0 Then
                For Each Dr As DataRow In dtRak.Rows
                    IdxModel += 1
                    Dim rak As String = ""
                    Dim no_rak As String = ""
                    Dim no_shelf As String = ""
                    If Dr("NORAK") & "" = "" Then
                        model(IdxModel) += "-".PadRight(20)
                    Else
                        no_rak = CInt(Dr("NORAK"))
                        no_rak = no_rak.PadLeft(3, "0")
                        no_shelf = CInt(Dr("NOSHELF"))
                        no_shelf = no_shelf.PadLeft(3, "0")

                        rak = Dr("NO") & ". " & no_rak & "/" & no_shelf
                        model(IdxModel) += rak.PadRight(20)
                    End If
                Next
            Else
                IdxModel += 1
                model(IdxModel) += "-".PadRight(20)
            End If

            For i = 0 To IdxModel
                data_result += model(i) + Chr(10)
            Next
            'data_result += model(7)

        Catch ex As Exception
            ShowError("Error set TampilNextSOBIC_Server", ex)
        End Try

        Return data_result
    End Function
#End Region

#Region "Tampil SO IC"
    Private Function TampilClientSOICTTL3(ByVal barcode_plu As String, ByVal deskripsi As String,
                                          ByVal qty_expired As String, ByVal qty_kemasan As String,
                                          ByVal qty_digigit As String) As String
        Dim data As String = ""

        Try
            'A
            data += Padcenter("SO " & jenis_so & " - " & lokasi_so, 20) + Chr(2)

            'B
            data += Padcenter("********************", 20) + Chr(2)

            'C
            If barcode_plu = "" Then
                data += "Prod:".PadRight(20) + Chr(2)
            Else
                barcode_plu = "Prod:" & barcode_plu
                data += barcode_plu.PadRight(20) + Chr(2)
            End If

            'D
            data += "Desc:".PadRight(20) + Chr(2)

            'E
            If deskripsi = "" Then
                data += "".PadRight(20) + Chr(2)
            ElseIf deskripsi = "Tidak Ditemukan" Then
                data += "Tidak Ditemukan".PadRight(20) + Chr(2)
            Else
                data += deskripsi.PadRight(20) + Chr(2)
            End If

            If barcode_plu <> "" Then
                'F
                If qty_expired = "" Then
                    data += "QTY Expired:".PadRight(20) & Chr(2) 'M
                Else
                    qty_expired = "QTY Expired:" & qty_expired
                    data += qty_expired.PadRight(20) + Chr(2)
                End If

                'G
                If qty_kemasan = "" Then
                    data += "QTY Kemasan:".PadRight(20) & Chr(2) 'M
                Else
                    qty_kemasan = "QTY Kemasan:" & qty_kemasan
                    data += qty_kemasan.PadRight(20) + Chr(2)
                End If

                'H
                If qty_digigit = "" Then
                    data += "QTY Digigit:".PadRight(20) & Chr(3) 'M
                Else
                    qty_digigit = "QTY Digigit:" & qty_digigit
                    data += qty_digigit.PadRight(20) + Chr(3)
                End If
            End If

            If barcode_plu = "" And deskripsi = "" Then 'belum ada data
                data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"

                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
            ElseIf barcode_plu <> "" And deskripsi <> "" Then 'ada data
                data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"

                If qty_expired = "" Then
                    data += "1" + "f" + "M" + Chr((3 * 16) + 0) + "Q"
                End If

                If qty_expired <> "" And qty_kemasan = "" Then
                    data += "1" + "g" + "M" + Chr((3 * 16) + 0) + "Q"
                End If

                If qty_expired <> "" And qty_kemasan <> "" And qty_digigit = "" Then
                    data += "1" + "h" + "M" + Chr((3 * 16) + 0) + "Q"
                End If
            ElseIf deskripsi = "Tidak Ditemukan" Then 'data tidak ditemukan
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"

                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
            End If

        Catch ex As Exception
            ShowError("Error Set TampilClientSOICTTL3", ex)
            TraceLog("Error TampilClientSOICTTL3: " & ex.ToString)
        End Try

        Return data
    End Function

    Private Function TampilServerSOICTTL3(ByVal barcode_plu As String, ByVal deskripsi As String,
                                          ByVal qty_expired As String, ByVal qty_kemasan As String,
                                          ByVal qty_digigit As String) As String
        Dim data_result As String = ""
        Dim header As String = ""

        Try
            header = "SO " & jenis_so & " - " & lokasi_so

            Dim model() = {Padcenter(header, 20), "====================", "", "", "", "", "", ""}
            Dim i As Integer = 0

            model(0) = Padcenter(header, 20)
            If barcode_plu = "" Then
                model(2) = "Prod:"
            Else
                model(2) = "Prod:" & barcode_plu
            End If

            model(3) = "Desc:"
            If deskripsi = "" Then
                model(4) = ""
            Else
                model(4) = deskripsi
            End If

            If barcode_plu <> "" Then
                If qty_expired = "" Then
                    model(5) = "QTY Expired:"
                Else
                    model(5) = "QTY Expired:" & qty_expired
                End If

                If qty_kemasan = "" Then
                    model(6) = "QTY Kemasan:"
                Else
                    model(6) = "QTY Kemasan:" & qty_kemasan
                End If

                If qty_digigit = "" Then
                    model(7) = "QTY Digigit:"
                Else
                    model(7) = "QTY Digigit:" & qty_digigit
                End If
            End If

            For i = 0 To 6
                data_result += model(i) + Chr(10)
            Next

            data_result += model(7)

        Catch ex As Exception
            ShowError("Error Set TampilServerSOICTTL3", ex)
            TraceLog("Error TampilServerSOICTTL3: " & ex.ToString)
        End Try

        Return data_result
    End Function

#End Region

#Region "Tampil Expired Date"
    Private Function TampilDeskripsiExpired(ByVal barcode_plu As String, ByVal deskripsi As String,
                                            ByVal lokasi As String, ByVal tgl_exp As String,
                                            ByVal qty As String, ByVal feedback As String) As String

        Dim data As String = ""

        'Contoh Tampilan
        '********SOED********
        'PRDCD:1000019
        'DESC:
        'INDOMILK SKM PTH 370
        'LOKASI:21-10-01
        'EXP(MMyyyy):
        'QTY:
        '

        Try
            'A
            data += Padcenter("********SOED********", 20) + Chr(2)

            'B
            If barcode_plu = "" Then
                data += "Prod:".PadRight(20) + Chr(2)
            Else
                barcode_plu = "Prod:" & barcode_plu
                data += barcode_plu.PadRight(20) + Chr(2)
            End If

            'C-D
            data += "Desc:".PadRight(20) + Chr(2)
            If deskripsi <> "" Then
                data += deskripsi.PadRight(20) + Chr(2)
            Else
                data += "".PadRight(20) + Chr(2)
            End If


            If feedback = "1" Then
                'E
                If lokasi <> "" Then
                    lokasi = "Lokasi:" & lokasi
                    data += lokasi.PadRight(20) + Chr(2)
                Else
                    data += "".PadRight(20) + Chr(2)
                End If

                'F
                If feedback = "1" And tgl_exp = "" Then
                    data += "EXP(MMyyyy):".PadRight(20) + Chr(2)
                Else
                    tgl_exp = "EXP(MMyyyy):" & tgl_exp
                    data += tgl_exp.PadRight(20) + Chr(2)
                End If

                'G
                If tgl_exp <> "" And qty = "" Then
                    data += "QTY:".PadRight(20) + Chr(2)
                Else
                    data += "QTY:" + Chr(2)
                End If

            End If

            'H
            If feedback = "2" Then
                data += "Format EXP salah".PadRight(20) + Chr(3)
            ElseIf "Format EXP salah" = "3" Then
                data += "Tgl EXP tdk sesuai".PadRight(20) + Chr(3)
            ElseIf feedback = "4" Then
                data += "Selesai Simpan QTY".PadRight(20) + Chr(3)
            Else
                data += "".PadRight(20) + Chr(3)
            End If

            If barcode_plu <> "" And (feedback = "1" Or feedback = "2") Then 'ada data
                data += "1" + "b" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "B" + "F" + Chr((15 * 16) + 0) + "B"

                If tgl_exp = "" Then
                    data += "1" + "f" + "M" + Chr((6 * 16) + 0) + "T"
                End If

                If (tgl_exp <> "") And (qty = "") Then
                    data += "1" + "g" + "E" + Chr((3 * 16) + 0) + "Q"
                End If

                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
            ElseIf barcode_plu = "" And feedback = "0" Then 'data tidak ditemukan
                data += "1" + "b" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "B" + "F" + Chr((15 * 16) + 0) + "B"
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
            Else 'belum ada data
                data += "1" + "b" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "B" + "F" + Chr((15 * 16) + 0) + "B"
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
            End If

        Catch ex As Exception
            MsgBox("TampilDeskripsiExpired Error: " & ex.ToString, MsgBoxStyle.OkOnly, "TampilDeskripsiExpired Error")
            TraceLog("TampilDeskripsiExpired Error: " & ex.ToString)
        End Try

        Return data
    End Function

    Private Function TampilDeskripsiExpiredServer(ByVal barcode_plu As String, ByVal deskripsi As String,
                                            ByVal lokasi As String, ByVal tgl_exp As String,
                                            ByVal qty As String, ByVal feedback As String) As String

        Dim Data As String = ""
        Dim header As String = "SO ED"

        Dim model() = {Padcenter(header, 20), "====================", "", "", "", "", "", ""}
        Dim i As Integer = 0
        Try

            model(0) = Padcenter(header, 20)

            If barcode_plu = "" Then
                model(1) = "Prod:"
            Else
                model(1) = "Prod:" & barcode_plu
            End If

            If deskripsi <> "" Then
                model(2) = "Desc:"
                model(3) = Strings.Right(deskripsi, 13)
            Else
                model(2) = "Desc:"
            End If

            If lokasi <> "" Then
                model(4) = "Lokasi:" & lokasi
            Else
                model(4) = "Lokasi:"
            End If

            If feedback = "1" Then
                model(4) = "Lokasi:" & lokasi

                If tgl_exp <> "" Then
                    model(5) += "EXP(MMyyyy):"
                Else
                    model(5) += "EXP(MMyyyy):" & tgl_exp
                End If
            End If

            If feedback = "2" Then
                model(6) += "Format EXP salah"
            ElseIf "Format EXP salah" = "3" Then
                model(6) += "Tgl EXP tdk sesuai"
            ElseIf feedback = "4" Then
                model(6) += "Selesai Simpan QTY"
            Else
                model(6) += ""
            End If

            For i = 0 To 7
                Data += model(i) + Chr(10)
            Next

            Data += model(7)

        Catch ex As Exception
            ShowError("Error set TampilDeskripsiServer", ex)
        End Try

        Return Data
    End Function
#End Region

#Region "Tampil BPB"
    Private Function TampilContainerClient(ByVal container As String, ByVal deskripsi As String) As String
        Dim data As String = ""
        Try

            data += Padcenter("BPB Toko", 20) + Chr(2)
            data += Padcenter("********************", 20) + Chr(2)
            data += "Scan Cont/Bron:".PadRight(20) + Chr(2)
            If deskripsi.Trim.Length > 0 Then
                data += container.PadRight(20) + Chr(2)
            Else
                data += "".PadRight(20) + Chr(2)
            End If
            data += "".PadRight(20) + Chr(2)
            data += "".PadRight(20) + Chr(2)
            data += deskripsi.PadRight(20) + Chr(2)
            data += "".PadRight(20) + Chr(2)
            data += Chr(3)

            If container = "" And deskripsi = "" Then 'belum ada data
                data += "1" + "d" + "A" + Chr((15 * 16) + 0) + "C"
                data += "1" + "D" + "A" + Chr((15 * 16) + 0) + "C"
                data += "0"
                data += "0"
                data += "0"
                data += "0"
            ElseIf container = "" And deskripsi.ToLower = "barcode tdk trdaftar" Then 'belum ada data
                data += "1" + "d" + "A" + Chr((15 * 16) + 0) + "C"
                data += "1" + "D" + "A" + Chr((15 * 16) + 0) + "C"
                data += "0"
                data += "0"
                data += "0"
                data += "0"
            Else  'ada data
                data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "C"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "C"
                data += "0" + "H" + "E" + Chr((8 * 16) + 0) + "Q"
                data += "0" + "H" + "E" + Chr((8 * 16) + 0) + "Q"
                data += "1" + "H" + "E" + Chr((8 * 16) + 0) + "Q"
                data += "0"
            End If

        Catch ex As Exception
            ShowError("Error set TampilContainer", ex)
        End Try

        Return data
    End Function

    Private Function TampilContainerClient2(ByVal IsContainer As Boolean, ByVal container As String,
                                            ByVal deskripsi As String, ByVal deskripsi2 As String) As String
        Dim data As String = ""
        Try
            data += Padcenter("BPB Toko", 20) + Chr(2)
            data += Padcenter("********************", 20) + Chr(2)
            If IsContainer Then
                data += "Container:".PadRight(20) + Chr(2)
            Else
                data += "Bronjong:".PadRight(20) + Chr(2)
            End If
            'data += "Container:".PadRight(20) + Chr(2)
            data += container.PadRight(20) + Chr(2)
            data += deskripsi.PadRight(20) + Chr(2)
            data += "Prod:".PadRight(20) + Chr(2)
            data += deskripsi2.PadRight(20) + Chr(2)
            data += "".PadRight(20) + Chr(2)
            data += Chr(3)

            data += "1" + "f" + "F" + Chr((15 * 16) + 0) + "S"
            data += "1" + "F" + "F" + Chr((15 * 16) + 0) + "B"
            data += "0"
            data += "1" + "f" + "F" + Chr((8 * 16) + 0) + "B"
            data += "0"
            data += "0"

        Catch ex As Exception
            ShowError("Error set TampilContainer", ex)
        End Try

        Return data
    End Function

    Private Function TampilContainerServer(ByVal IsContainer As Boolean, ByVal container As String,
                                           ByVal deskripsi As String, ByVal deskripsi2 As String) As String
        Dim data_result As String = ""
        Dim model() = {Padcenter("BPB Toko", 20), "====================", "", "", "", "", "", ""}
        Dim i As Integer = 0

        Try
            If container.Trim = "" Then
                model(2) = "Scan Cont/Bron:"
            Else
                If IsContainer Then
                    model(2) = "Container:"
                Else
                    model(2) = "Bronjong:"
                End If
            End If
            model(3) = container
            model(4) = deskripsi
            model(5) = ""
            model(6) = deskripsi2
            model(7) = ""

            For i = 0 To 6
                data_result += model(i) + Chr(10)
            Next
            data_result += model(7)

        Catch ex As Exception
            ShowError("Error set TampilContainer", ex)
        End Try

        Return data_result
    End Function
#End Region

#Region "Tampil Planogram"
    Private Function TampilPlanoClient(ByVal barcode_plu As String, ByVal deskripsi As String,
                                       ByVal retur As String, ByVal price As String) As String
        Dim data As String = ""

        Try
            data += Padcenter(jenis_so, 20) + Chr(2)
            data += Padcenter("********************", 20) + Chr(2)

            If barcode_plu = "" Then
                data += "Prod:".PadRight(20) + Chr(2)
            Else
                barcode_plu = "Prod:" & barcode_plu
                data += barcode_plu.PadRight(20) + Chr(2)
            End If

            data += "Desc:".PadRight(20) + Chr(2)
            If deskripsi = "" Then
                data += "".PadRight(20) + Chr(2)
                data += "".PadRight(20) + Chr(2)
            ElseIf deskripsi = "Tidak Ditemukan" Then
                data += "Tidak Ditemukan".PadRight(20) + Chr(2)
                data += "".PadRight(20) + Chr(2)
            Else
                data += deskripsi.Substring(0, 20) + Chr(2)
                data += deskripsi.Substring(20, 20) + Chr(2)
            End If

            If retur = "" Then
                data += "Tgl Exp < ".PadRight(20) + Chr(2)
            Else
                retur = " Tgl Exp < " & retur
                data += retur.PadRight(20) + Chr(2)
            End If

            If price = "" Then
                data += " Price:".PadRight(20) + Chr(2)
            Else
                price = " Price:" & price
                data += price.PadRight(20) + Chr(2)
            End If

            data += ""
            data += Chr(3)

            data += " 1" + " C" + " F" + Chr((15 * 16) + 0) + " S"
            data += " 1" + " C" + " F" + Chr((15 * 16) + 0) + " B"
            data += " 0"
            data += " 0"
            data += " 0"
            data += " 0"

        Catch ex As Exception
            ShowError(" Error set TampilPlanoClient", ex)
        End Try

        Return data
    End Function

    Private Function TampilPlanoServer(ByVal barcode_plu As String, ByVal deskripsi As String,
                                       ByVal retur As String, ByVal price As String) As String
        Dim data_result As String = ""
        Dim header As String = ""

        header = jenis_so
        Dim model() = {Padcenter(header, 20), "====================", "", "", "", "", "", ""}
        Dim i As Integer = 0
        Try
            model(0) = Padcenter(header, 20)
            If barcode_plu = "" Then
                model(2) = "Prod:"
            Else
                model(2) = "Prod:" & barcode_plu
            End If

            model(3) = "Desc:"
            If deskripsi = "" Then
                model(4) = ""
            Else
                model(4) = deskripsi
            End If

            If retur = "" Then
                model(5) = "Tgl Exp < "
            Else
                model(5) = "Tgl Exp < " & retur
            End If

            If price = "" Then
                model(6) = "Price:"
            Else
                model(6) = "Price:" & price
            End If

            For i = 0 To 6
                data_result += model(i) + Chr(10)
            Next
            data_result += model(7)

        Catch ex As Exception
            ShowError("Error set TampilPlanoServer", ex)
        End Try

        Return data_result
    End Function

#End Region

#Region "Tampil PJR"
    Private Function TampilPJRClient(ByVal barcode_plu As String, ByVal deskripsi As String,
                                       ByVal retur As String, ByVal price As String) As String
        Dim data As String = ""

        Try
            If jenis_so = "TINDAKLBTD" Then
                data += Padcenter("Tindak LBTD", 20) + Chr(2)
            ElseIf jenis_so.ToUpper = "TINDAKLBTD_BAPJR" Then
                data += Padcenter("Tindak LBTD BA PJR", 20) + Chr(2)

            Else
                data += Padcenter("PJR", 20) + Chr(2)

            End If
            data += Padcenter("********************", 20) + Chr(2)

            If barcode_plu = "" Then
                data += "Prod:".PadRight(20) + Chr(2)
            Else
                barcode_plu = "Prod:" & barcode_plu
                data += barcode_plu.PadRight(20) + Chr(2)
            End If

            data += "Desc:".PadRight(20) + Chr(2)
            If deskripsi = "" Then
                data += "".PadRight(20) + Chr(2)
                data += "".PadRight(20) + Chr(2)
            ElseIf deskripsi = "Tidak Ditemukan" Then
                data += "Tidak Ditemukan".PadRight(20) + Chr(2)
                data += "".PadRight(20) + Chr(2)
            ElseIf deskripsi.ToLower.Trim = "tolak" Then
                data += "Tolak Input Plu".PadRight(20) + Chr(2)
                data += "".PadRight(20) + Chr(2)
            Else
                If deskripsi.Length <= 20 Then
                    data += deskripsi.Substring(0, deskripsi.Length) + Chr(2)
                Else
                    data += deskripsi.Substring(0, 20) + Chr(2)
                    data += deskripsi.Substring(20, 20) + Chr(2)

                End If
            End If

            If retur = "" Then
                data += "Tgl Exp < ".PadRight(20) + Chr(2)
            Else
                retur = "Tgl Exp < " & retur
                data += retur.PadRight(20) + Chr(2)
            End If

            If price = "" Then
                data += "Price:".PadRight(20) + Chr(2)
            Else
                price = "Price:" & price
                data += price.PadRight(20) + Chr(2)
            End If

            data += ""
            data += Chr(3)

            data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "S"
            data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"
            data += "0"
            data += "0"
            data += "0"
            data += "0"

        Catch ex As Exception
            ShowError("Error set TampilPJRClient", ex)
        End Try

        Return data
    End Function

    Private Function TampilPJRServer(ByVal barcode_plu As String, ByVal deskripsi As String,
                                       ByVal retur As String, ByVal price As String) As String
        Dim data_result As String = ""
        Dim header As String = ""

        If jenis_so = "CekPJR" Then
            header = "PJR"

        ElseIf jenis_so = "TINDAKLBTD" Then
            header = "LBTD"
        ElseIf jenis_so = "TINDAKLBTD_BAPJR" Then
            header = "LBTD BA PJR"
        End If
        Dim model() = {Padcenter(header, 20), "====================", "", "", "", "", "", ""}
        Dim i As Integer = 0
        Try
            model(0) = Padcenter(header, 20)
            If barcode_plu = "" Then
                model(2) = "Prod:"
            Else
                model(2) = "Prod:" & barcode_plu
            End If

            model(3) = "Desc:"
            If deskripsi = "" Then
                model(4) = ""
            Else
                If deskripsi.ToLower.Trim = "tolak" Then
                    deskripsi = "Tolak Input Plu"
                End If
                model(4) = deskripsi
            End If

            If retur = "" Then
                model(5) = "Tgl Exp < "
            Else
                model(5) = "Tgl Exp < " & retur
            End If

            If price = "" Then
                model(6) = "Price:"
            Else
                model(6) = "Price:" & price
            End If

            For i = 0 To 6
                data_result += model(i) + Chr(10)
            Next
            data_result += model(7)

        Catch ex As Exception
            ShowError("Error set TampilPJRServer", ex)
        End Try

        Return data_result
    End Function

#End Region

#Region "Tampil Cek Display"
    Private Function TampilCekDisplayClient(ByVal barcode_plu As String, ByVal deskripsi As String,
                                       ByVal qty As String, Optional ByVal feedback As String = "") As String
        Dim data As String = ""

        Try
            data += Padcenter(jenis_so, 20) + Chr(2)
            data += Padcenter("********************", 20) + Chr(2)

            If barcode_plu = "" Then
                data += "Prod:".PadRight(20) + Chr(2)
            Else
                barcode_plu = "Prod:" & barcode_plu
                data += barcode_plu.PadRight(20) + Chr(2)
            End If

            data += "Desc:".PadRight(20) + Chr(2)
            If deskripsi = "" Then
                data += "".PadRight(20) + Chr(2)
                data += "".PadRight(20) + Chr(2)
            ElseIf deskripsi = "Tidak Ditemukan" Then
                data += "Tidak Ditemukan".PadRight(20) + Chr(2)
                data += "".PadRight(20) + Chr(2)
            Else

                data += deskripsi.PadRight(20) + Chr(2)
                data += "".PadRight(20) + Chr(2)

            End If
            data += "QTY:" + qty.PadRight(20) + Chr(2)
            data += Chr(3)



            If barcode_plu <> "" And deskripsi <> "Tidak Ditemukan" Then 'ada data
                data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"

                data += "1" + "g" + "E" + Chr((8 * 16) + 0) + "Q"

                data += "0" + "A" + "E" + Chr((8 * 16) + 0) + "Q"
                data += "0" + "A" + "E" + Chr((8 * 16) + 0) + "Q"
                data += "0"
            ElseIf barcode_plu = "" And deskripsi = "Tidak Ditemukan" Then 'data tidak ditemukan
                data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"
                data += "0"
                data += "0"
                data += "0"
                data += "0"
            Else 'belum ada data
                data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"
                data += "0"
                data += "0"
                data += "0"
                data += "0"
            End If

        Catch ex As Exception
            ShowError("Error set TampilCekDisplayClient", ex)
        End Try

        Return data
    End Function

    Private Function TampilCekDisplayServer(ByVal barcode_plu As String, ByVal deskripsi As String,
                                       ByVal qty As String) As String
        Dim data_result As String = ""
        Dim header As String = ""

        header = jenis_so
        Dim model() = {Padcenter(header, 20), "====================", "", "", "", "", "", ""}
        Dim i As Integer = 0
        Try
            model(0) = Padcenter(header, 20)
            If barcode_plu = "" Then
                model(2) = "Prod:"
            Else
                model(2) = "Prod:" & barcode_plu
            End If

            model(3) = "Desc:"
            If deskripsi = "" Then
                model(4) = ""
            Else
                model(4) = deskripsi
            End If

            If qty = "" Then
                model(5) = "QTY "
            Else
                model(5) = "QTY: " & qty
            End If

            model(6) = ""


            For i = 0 To 6
                data_result += model(i) + Chr(10)
            Next
            data_result += model(7)

        Catch ex As Exception
            ShowError("Error set TampilCekDisplayServer", ex)
        End Try

        Return data_result
    End Function

#End Region

#Region "Tampil TampilKesegaran"
    Private Function TampilKesegaranClient(ByVal barcode_plu As String, ByVal deskripsi As String,
                                       ByVal retur As String, ByVal qty As String) As String
        Dim data As String = ""
        Dim dataretur As String()
        Try
            data += Padcenter(jenis_so, 20) + Chr(2)
            data += Padcenter("********************", 20) + Chr(2)

            'If barcode_plu = "" Then
            '    data += "Prod:".PadRight(20) + Chr(2)
            'Else
            '    barcode_plu = "Prod:" & barcode_plu
            '    data += barcode_plu.PadRight(20) + Chr(2)
            'End If

            'data += "Desc:".PadRight(20) + Chr(2)
            If deskripsi = "" Then
                data += "".PadRight(20) + Chr(2)
                data += "".PadRight(20) + Chr(2)
            ElseIf deskripsi = "Tidak Ditemukan" Or deskripsi = "QTY melebihi LPP!" Then
                data += deskripsi.PadRight(20) + Chr(2)
                data += "".PadRight(20) + Chr(2)
            Else
                data += deskripsi.Substring(0, 20) + Chr(2)
                data += deskripsi.Substring(20, 20) + Chr(2)
            End If

            If retur = "" Then
                data += "Tgl Exp <= (".PadRight(20) + Chr(2)
                data += "".PadRight(20) + Chr(2)
                data += "".PadRight(20) + Chr(2)

            Else
                data += "Tgl Exp <= (".PadRight(20) + Chr(2)
                If retur.Length > 5 Then
                    dataretur = retur.Split(",")
                    If dataretur.Length = 1 Then
                        data += dataretur(0) & ")".PadRight(20) + Chr(2)
                        data += ""
                    ElseIf dataretur.Length = 2 Then
                        retur = dataretur(0) & "," & dataretur(1) & ")"
                        data += retur.PadRight(20) + Chr(2)
                        data += "".PadRight(20) + Chr(2)
                    ElseIf dataretur.Length = 3 Then
                        retur = dataretur(0) & "," & dataretur(1)
                        data += retur.PadRight(20) + Chr(2)
                        retur = dataretur(2) & ")"
                        data += retur.PadRight(20) + Chr(2)

                    ElseIf dataretur.Length = 4 Then
                        retur = dataretur(0) & "," & dataretur(1)
                        data += retur.PadRight(20) + Chr(2)
                        retur = dataretur(2) & "," & dataretur(3) & ")"
                        data += retur.PadRight(20) + Chr(2)
                    End If
                End If

            End If
            If qty = "" Then
                data += "QTY:".PadRight(20) + Chr(2)
            Else
                qty = "QTY:" & qty
                data += qty.PadRight(20) + Chr(2)
            End If

            data += ""
            data += Chr(3)

            If barcode_plu <> "" And deskripsi <> "Tidak Ditemukan" Then 'ada data
                data += "1" + "c" + "A" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "A" + Chr((15 * 16) + 0) + "B"
                data += "1" + "h" + "E" + Chr((8 * 16) + 0) + "Q"
                data += "0"
                data += "0"
                data += "0"
            ElseIf barcode_plu = "" And (deskripsi = "Tidak Ditemukan" Or deskripsi = "QTY melebihi LPP!") Then 'data tidak ditemukan
                data += "1" + "c" + "A" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "A" + Chr((15 * 16) + 0) + "B"
                data += "0"
                data += "0"
                data += "0"
                data += "0"
            Else 'belum ada data
                data += "1" + "c" + "A" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "A" + Chr((15 * 16) + 0) + "B"
                data += "0"
                data += "0"
                data += "0"
                data += "0"
            End If


        Catch ex As Exception
            ShowError("Error set TampilKesegaranClient", ex)
        End Try

        Return data
    End Function

    Private Function TampilKesegaranServer(ByVal barcode_plu As String, ByVal deskripsi As String,
                                       ByVal retur As String, ByVal qty As String) As String
        Dim data_result As String = ""
        Dim header As String = ""

        header = jenis_so
        Dim model() = {Padcenter(header, 20), "====================", "", "", "", "", "", ""}
        Dim i As Integer = 0
        Try
            model(0) = Padcenter(header, 20)
            If barcode_plu = "" Then
                model(2) = ""
            Else
                model(2) = deskripsi
            End If

            model(3) = "Tgl Exp <= ("

            If retur = "" Then
                model(4) = ""
            Else
                model(4) = retur
            End If

            If qty = "" Then
                model(5) = "QTY:"
            Else
                model(5) = "QTY:" & qty
            End If

            model(6) = ""


            For i = 0 To 6
                data_result += model(i) + Chr(10)
            Next
            data_result += model(7)

        Catch ex As Exception
            ShowError("Error set TampilKesegaranServer", ex)
        End Try

        Return data_result
    End Function

#End Region

#Region "TampilBPBBKL"
    Private Function TampilContainerBKLClient(ByVal container As String, ByVal deskripsi As String) As String
        Dim data As String = ""
        Try

            data += Padcenter("BPB BKL Toko", 20) + Chr(2)
            data += Padcenter("********************", 20) + Chr(2)
            data += "Scan Cont/Bron:".PadRight(20) + Chr(2)
            If deskripsi.Trim.Length > 0 Then
                data += container.PadRight(20) + Chr(2)
            Else
                data += "".PadRight(20) + Chr(2)
            End If
            data += "".PadRight(20) + Chr(2)
            data += "".PadRight(20) + Chr(2)
            data += deskripsi.PadRight(20) + Chr(2)
            data += "".PadRight(20) + Chr(2)
            data += Chr(3)

            If container = "" And deskripsi = "" Then 'belum ada data
                data += "1" + "d" + "A" + Chr((15 * 16) + 0) + "C"
                data += "1" + "D" + "A" + Chr((15 * 16) + 0) + "C"
                data += "0"
                data += "0"
                data += "0"
                data += "0"
            ElseIf container = "" And deskripsi.ToLower = "barcode tdk trdaftar" Then 'belum ada data
                data += "1" + "d" + "A" + Chr((15 * 16) + 0) + "C"
                data += "1" + "D" + "A" + Chr((15 * 16) + 0) + "C"
                data += "0"
                data += "0"
                data += "0"
                data += "0"
            Else  'ada data
                data += "1" + "c" + "F" + Chr((15 * 16) + 0) + "C"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "C"
                data += "0" + "H" + "E" + Chr((8 * 16) + 0) + "Q"
                data += "0" + "H" + "E" + Chr((8 * 16) + 0) + "Q"
                data += "1" + "H" + "E" + Chr((8 * 16) + 0) + "Q"
                data += "0"
            End If

        Catch ex As Exception
            ShowError("Error set TampilContainer BKL", ex)
        End Try

        Return data
    End Function

    Private Function TampilDeskripsiBKLServer(ByVal barcode_plu As String, ByVal deskripsi As String, ByVal qty As String, ByVal tgl_exp As String, ByVal feedback As String, Optional ByVal deskripsi2 As String = "") As String
        Dim data_result As String = ""
        Dim header As String = ""
        If tabel_name = "dcp_boxplu" Then
            header = "BPB Toko"
        ElseIf tabel_name = "BPBBKL_WDCP" Then
            header = "BPB BKL Toko"
        ElseIf tabel_name.ToUpper = "BPBNPS_WDCP" Then
            header = "BPB NPS Toko"
        ElseIf tabel_name.ToUpper.Contains("OA") Then
            header = "SO Aktiva"
        Else
            header = "SO " & jenis_so & " - " & lokasi_so
        End If
        Dim model() = {Padcenter(header, 20), "====================", "", "", "", "", "", ""}
        Dim i As Integer = 0
        Try

            model(0) = Padcenter(header, 20)
            If tabel_name.Trim.ToLower = "bpbbkl_wdcp" Then
                If deskripsi2 = "" Then
                    model(1) = "Fraction PCS:"
                Else
                    model(1) = "Fraction PCS:" & deskripsi2

                End If
            End If


            If barcode_plu = "" Then
                model(2) = "Prod:"
            Else
                model(2) = "Prod:" & barcode_plu
            End If
            If deskripsi <> "" Then
                model(3) = "Desc:" & Strings.Right(deskripsi, 13)
            Else
                model(3) = "Desc:"
            End If
            If tgl_exp <> "" Then
                model(4) = "EXP:" & tgl_exp

            Else
                model(4) = "EXP:"

            End If
            If qty <> "" Then
                model(5) = "QTY:" & qty
            Else
                model(5) = "QTY:"

            End If

            If feedback <> "" Then
                model(6) = feedback
            Else
                model(6) = ""

            End If
            For i = 0 To 6
                data_result += model(i) + Chr(10)
            Next
            data_result += model(7)

        Catch ex As Exception
            ShowError("Error set TampilDeskripsiServer", ex)
        End Try

        Return data_result
    End Function

    Private Function TampilContainerBKLServer(ByVal IsContainer As Boolean, ByVal container As String, ByVal deskripsi As String, ByVal deskripsi2 As String) As String
        Dim data_result As String = ""
        Dim model() = {Padcenter("BPB BKL Toko", 20), "====================", "", "", "", "", "", ""}
        Dim i As Integer = 0

        Try
            If container.Trim = "" Then
                model(2) = "Scan Cont/Bron:"
            Else
                If IsContainer Then
                    model(2) = "Container:"
                Else
                    model(2) = "Bronjong:"
                End If
            End If
            model(3) = container
            model(4) = deskripsi
            model(5) = ""
            model(6) = deskripsi2
            model(7) = ""

            For i = 0 To 6
                data_result += model(i) + Chr(10)
            Next
            data_result += model(7)

        Catch ex As Exception
            ShowError("Error set TampilContainer", ex)
        End Try

        Return data_result
    End Function

#End Region

#Region "Cetak Price Tag"
    Private Function TampilPriceTagClient(ByVal barcode_plu As String, ByVal deskripsi As String, ByVal konfirm As String, ByVal keterangan As String) As String
        Dim data As String = ""

        Try
            data += Padcenter("Cetak Price Tag", 20) + Chr(2)
            data += Padcenter("********************", 20) + Chr(2)

            If barcode_plu = "" Then
                data += "Prod:".PadRight(20) + Chr(2)
            Else
                barcode_plu = "Prod:" & barcode_plu
                data += barcode_plu.PadRight(20) + Chr(2)
            End If

            data += "Desc:".PadRight(20) + Chr(2)
            If deskripsi = "" Then
                data += "".PadRight(20) + Chr(2)
                data += "".PadRight(20) + Chr(2)
            ElseIf deskripsi = "Tidak Ditemukan" Then
                data += "Tidak Ditemukan".PadRight(20) + Chr(2)
                data += "".PadRight(20) + Chr(2)
            Else
                If deskripsi.Length < 20 Then
                    data += deskripsi.Substring(0, deskripsi.Length) + Chr(2)
                    data += "".PadRight(20) + Chr(2)
                Else
                    data += deskripsi.Substring(0, 20) + Chr(2)

                    If deskripsi.Length < 40 Then
                        data += deskripsi.Substring(20, deskripsi.Length - 20) + Chr(2)
                    Else
                        data += deskripsi.Substring(20, 20) + Chr(2)
                    End If

                End If
            End If
            If keterangan = "INSERT" Or keterangan = "" Then
                data += "(1.Yes, 2.No)".PadRight(20) + Chr(2)

                data += "Input(1/2):".PadRight(20) + Chr(2)
            ElseIf keterangan = "HAPUS" Then
                data += "(1.Hapus, 2.Batal)".PadRight(20) + Chr(2)

                data += "Input(1/2):".PadRight(20) + Chr(2)

            End If

            data += ""
            data += Chr(3)

            If konfirm = "" And keterangan = "" Then 'tampilan awal
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "

            ElseIf konfirm = "" And (keterangan = "INSERT" Or keterangan = "HAPUS") Then 'input konfirmasi (INSERT/HAPUS) 1/2
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "S"
                'data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"
                data += "1" + "h" + "L" + Chr((15 * 16) + 0) + "F"
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
            ElseIf konfirm <> "" And (keterangan = "INSERT" Or keterangan = "HAPUS") Then 'tampilan awal stlah berhasil input konfirmasi
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
            End If


        Catch ex As Exception
            ShowError("Error set TampilPriceTagClient", ex)
        End Try

        Return data
    End Function

    Private Function TampilPriceTagServer(ByVal barcode_plu As String, ByVal deskripsi As String, ByVal konfirm As String, ByVal keterangan As String) As String
        Dim data_result As String = ""
        Dim header As String = ""

        header = "Cetak Price Tag"
        Dim model() = {Padcenter(header, 20), "====================", "", "", "", "", "", ""}
        Dim i As Integer = 0
        Try
            model(0) = Padcenter(header, 20)
            If barcode_plu = "" Then
                model(2) = "Prod:"
            Else
                model(2) = "Prod:" & barcode_plu
            End If

            model(3) = "Desc:"
            If deskripsi = "" Then
                model(4) = ""
            Else
                model(4) = deskripsi
            End If


            If keterangan = "INSERT" Or keterangan = "" Then
                model(5) = "(1.Yes, 2.No)"
            ElseIf keterangan = "HAPUS" Then
                model(5) = "(1.Hapus, 2.Batal)"
            End If


            model(6) = "Input(1/2):"

            For i = 0 To 6
                data_result += model(i) + Chr(10)
            Next
            data_result += model(7)

        Catch ex As Exception
            ShowError("Error set TampilPriceTagServer", ex)
        End Try

        Return data_result
    End Function
#End Region

#Region "Monitoring Price Tag"
    Private Function TampilMonitoringPriceTagClient(ByVal barcode_plu As String, ByVal deskripsi As String, ByVal menu As String) As String
        'Revisi Memo No 296/CPS/23 Monitoring Price Tag by Kukuh 16 Mei 2023
        Dim data As String = ""

        Try
            data += Padcenter("Monitoring Price Tag", 20) + Chr(2)
            data += Padcenter("********************", 20) + Chr(2)

            If barcode_plu = "" Then
                data += "Prod:".PadRight(20) + Chr(2)
            Else
                barcode_plu = "Prod:" & barcode_plu
                data += barcode_plu.PadRight(20) + Chr(2)
            End If

            data += "Desc:".PadRight(20) + Chr(2)
            If deskripsi = "" Then
                data += "".PadRight(20) + Chr(2)
                data += "".PadRight(20) + Chr(2)
            ElseIf deskripsi = "Tidak Ditemukan" Then
                data += "Tidak Ditemukan".PadRight(20) + Chr(2)
                data += "".PadRight(20) + Chr(2)
            Else
                If deskripsi.Length < 20 Then
                    data += deskripsi.Substring(0, deskripsi.Length) + Chr(2)
                    data += "".PadRight(20) + Chr(2)
                Else
                    data += deskripsi.Substring(0, 20) + Chr(2)

                    If deskripsi.Length < 60 And deskripsi.Length > 40 Then
                        data += deskripsi.Substring(20, 20) + Chr(2)
                        data += deskripsi.Substring(40, deskripsi.Length - 40) + Chr(2)
                    ElseIf deskripsi.Length < 40 Then
                        data += deskripsi.Substring(20, deskripsi.Length - 20) + Chr(2)
                    Else
                        data += deskripsi.Substring(20, 20) + Chr(2)
                    End If

                End If
            End If

            If menu = "1" Then
                data += "Silahkan Scan atau".PadRight(20) + Chr(2)
                data += "Input Barcode".PadRight(20) + Chr(2)
                data += ""
                data += Chr(3)
            ElseIf menu = "2" Then
                data += "(1. Simpan, 2. Ulang)".PadRight(20) + Chr(2)
                data += "Input(1/2):".PadRight(20) + Chr(2)
                data += ""
                data += Chr(3)
            Else
                data += "".PadRight(20) + Chr(2)
                data += "".PadRight(20) + Chr(2)
                data += ""
                data += Chr(3)
            End If

            If menu = "1" Or menu = "3" Then
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "B"
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
            ElseIf menu = "2" Then
                data += "1" + "C" + "F" + Chr((15 * 16) + 0) + "S"
                data += "1" + "h" + "L" + Chr((15 * 16) + 0) + "F"
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
                data += "0" + "A" + "A" + Chr(16) + " "
            End If

        Catch ex As Exception
            ShowError("Error set TampilPriceTagClient", ex)
        End Try

        Return data
    End Function

    Private Function TampilMonitoringPriceTagServer(ByVal barcode_plu As String, ByVal deskripsi As String, ByVal menu As String) As String
        Dim data_result As String = ""
        Dim header As String = ""

        header = "Monitoring Price Tag"
        Dim model() = {Padcenter(header, 20), "====================", "", "", "", "", "", ""}
        Dim i As Integer = 0
        Try
            model(0) = Padcenter(header, 20)
            If barcode_plu = "" Then
                model(2) = "Prod:"
            Else
                model(2) = "Prod:" & barcode_plu
            End If

            model(3) = "Desc:"
            If deskripsi = "" Then
                model(4) = ""
            Else
                model(4) = deskripsi
            End If

            If menu = "1" Then
                model(5) = "Silahkan Scan atau"
                model(6) = "Input Barcode"
            ElseIf menu = "2" Then
                model(5) = "(1. Simpan, 2. Ulang)"
                model(6) = "Input(1/2):"
            End If

            For i = 0 To 6
                data_result += model(i) + Chr(10)
            Next
            data_result += model(7)

        Catch ex As Exception
            ShowError("Error set TampilPriceTagServer", ex)
        End Try

        Return data_result
    End Function

#End Region

#Region "Utility Display"


    ''' <summary>
    ''' fungsi untuk men delay response, dll
    ''' </summary>
    ''' <param name="time"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function delaytime(ByVal time As Integer)
        Dim due As DateTime

        due = Now.AddMilliseconds(time)
        Do While Now < due
        Loop

        Return Nothing
    End Function

    ''' <summary>
    ''' fungsi untuk membuat text center
    ''' </summary>
    ''' <param name="kalimat"></param>
    ''' <param name="length"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Padcenter(ByVal kalimat, ByVal length) As String
        kalimat = kalimat.PadLeft((length + kalimat.Length) \ 2).PadRight(length)
        Return kalimat
    End Function

#End Region

#End Region

    ''' <summary>
    ''' Prosedur untuk memulai cek PB
    ''' </summary>
    ''' <remarks></remarks>
    ''' <history>
    ''' wisnu b 25/5/2015 - wisanggeni 23/9/2016
    ''' </history>
    Private Sub MulaiCekPB(ByVal kodeDC As String)
        Dim cBPB As New ClsBPBController
        Dim Brg, Brg2, Dus, Dus2 As Integer
        Dim docno As String = ""
        Try
            Dim VarTbl = cUtility.ExecuteScalar("SHOW TABLES LIKE '" & tabel_name & "';")
            If IsNothing(VarTbl) Or IsDBNull(VarTbl) Then
                MsgBox("Table " & tabel_name & " Tidak Tersedia!!")

                BtnCekPB.Text = "MULAI"
                CbxKodeGudang.Enabled = True

                BtnCariPlu.Enabled = False
                BtnSplitCtn.Enabled = False
                txtLokasi.Text = ""
                txtPlu.Text = ""
                Exit Sub
            End If

            If tabel_name = "dcp_boxplu" Then
                'PROGRAM CEK BARANG BISA PILIH PER DOCNO (NO NPB) - 438/03-23/E/PMO
                docno = kodeDC.Split("-")(1).ToString
                kodeDC = kodeDC.Split("-")(0).ToString
                DtBPB_DCP = cBPB.GetTableCekPB(kodeDC, docno)

                'DtBPB_DCP = cBPB.GetTableCekPB(kodeDC)
                DtBPB_DCP.TableName = "boxplu"
                DgBPB.DataSource = DtBPB_DCP

                If DtBPB_DCP.Rows.Count = 0 Then
                    MsgBox("Tidak Tersedia Data untuk Cek BPB" & vbCrLf & "atau Cek BPB hari ini sudah selesai", MsgBoxStyle.Information)
                    BtnCekPB.Text = "MULAI"
                    CbxKodeGudang.Enabled = True

                    BtnCariPlu.Enabled = False
                    BtnSplitCtn.Enabled = False
                    txtLokasi.Text = ""
                    txtPlu.Text = ""
                    Exit Sub
                End If
                'PROGRAM CEK BARANG BISA PILIH PER DOCNO (NO NPB) - 438/03-23/E/PMO
                Brg = Integer.Parse(cUtility.ExecuteScalar("SELECT Count(*) FROM Dcp_Boxplu " _
                                                      & "WHERE Recid <> '1' " _
                                                      & "AND KIRIM = '" & kodeDC & "' AND DOCNO = '" & docno & "'"))
                'PROGRAM CEK BARANG BISA PILIH PER DOCNO (NO NPB) - 438/03-23/E/PMO

                Brg2 = Integer.Parse(cUtility.ExecuteScalar("SELECT Count(*) FROM Dcp_Boxplu " _
                                                           & "WHERE KIRIM = '" & kodeDC & "' AND DOCNO = '" & docno & "'"))
                'PROGRAM CEK BARANG BISA PILIH PER DOCNO (NO NPB) - 438/03-23/E/PMO

                Dus = Integer.Parse(cUtility.ExecuteScalar("SELECT Count(distinct dus_no) FROM Dcp_Boxplu " _
                                                           & "WHERE Recid <> '1' " _
                                                           & "AND KIRIM = '" & kodeDC & "' AND DOCNO = '" & docno & "'"))
                'PROGRAM CEK BARANG BISA PILIH PER DOCNO (NO NPB) - 438/03-23/E/PMO

                Dus2 = Integer.Parse(cUtility.ExecuteScalar("SELECT Count(distinct dus_no) FROM Dcp_Boxplu " _
                                                           & "WHERE KIRIM = '" & kodeDC & "' AND DOCNO = '" & docno & "'"))

                ''PROGRAM CEK BARANG BISA PILIH PER DOCNO (NO NPB) - 438/03-23/E/PMO
                'Brg = Integer.Parse(cUtility.ExecuteScalar("SELECT Count(*) FROM Dcp_Boxplu " _
                '                                      & "WHERE Recid <> '1' " _
                '                                      & "AND KIRIM = '" & kodeDC & "'"))
                ''PROGRAM CEK BARANG BISA PILIH PER DOCNO (NO NPB) - 438/03-23/E/PMO

                'Brg2 = Integer.Parse(cUtility.ExecuteScalar("SELECT Count(*) FROM Dcp_Boxplu " _
                '                                           & "WHERE KIRIM = '" & kodeDC & "'"))
                ''PROGRAM CEK BARANG BISA PILIH PER DOCNO (NO NPB) - 438/03-23/E/PMO

                'Dus = Integer.Parse(cUtility.ExecuteScalar("SELECT Count(distinct dus_no) FROM Dcp_Boxplu " _
                '                                           & "WHERE Recid <> '1' " _
                '                                           & "AND KIRIM = '" & kodeDC & "'"))
                ''PROGRAM CEK BARANG BISA PILIH PER DOCNO (NO NPB) - 438/03-23/E/PMO

                'Dus2 = Integer.Parse(cUtility.ExecuteScalar("SELECT Count(distinct dus_no) FROM Dcp_Boxplu " _
                '                                           & "WHERE KIRIM = '" & kodeDC & "'"))

                txtRec.Text = Brg & "/" & Brg2
                txtCont.Text = Dus & "/" & Dus2

            ElseIf tabel_name = "BPBBKL_WDCP" Then
                DtBPBBKL_DCP = cBPB.GetTableCekPBBKL(kodeDC, parameter_docno)
                DtBPBBKL_DCP.TableName = "BPBBKL"

                DgBPB.DataSource = DtBPBBKL_DCP
                If DtBPBBKL_DCP.Rows.Count = 0 Then
                    MsgBox("Tidak Tersedia Data untuk Cek BPB BKL" & vbCrLf & "atau Cek BPB BKL hari ini sudah selesai", MsgBoxStyle.Information)
                    BtnCekPB.Text = "MULAI"
                    CbxKodeGudang.Enabled = True

                    BtnCariPlu.Enabled = False
                    BtnSplitCtn.Enabled = False
                    txtLokasi.Text = ""
                    txtPlu.Text = ""
                    Exit Sub
                End If

                Brg = Integer.Parse(cUtility.ExecuteScalar("SELECT Count(*) FROM BPBBKL_WDCP " _
                                                     & "WHERE tgl_bpbw IS NULL AND FINISHW IS NULL"))
                Brg2 = Integer.Parse(cUtility.ExecuteScalar("SELECT Count(*) FROM BPBBKL_WDCP "))

                txtRec.Text = Brg & "/" & Brg2
                txtCont.Text = "-"
                'DIRECT SHIPMENT
            ElseIf tabel_name.ToLower = "bpbnps_wdcp" Then
                DtBPBNPS = cBPB.GetTableCekPBNPS(parameter_noPO)
                DtBPBNPS.TableName = "BPBNPS"

                DgBPB.DataSource = DtBPBNPS
                If DtBPBNPS.Rows.Count = 0 Then
                    MsgBox("Tidak Tersedia Data untuk Cek BPB NPS" & vbCrLf & "atau Cek BPB NPS hari ini sudah selesai", MsgBoxStyle.Information)
                    BtnCekPB.Text = "MULAI"
                    CbxKodeGudang.Enabled = True

                    BtnCariPlu.Enabled = False
                    BtnSplitCtn.Enabled = False
                    txtLokasi.Text = ""
                    txtPlu.Text = ""
                    Exit Sub
                End If

                Brg = Integer.Parse(cUtility.ExecuteScalar("SELECT Count(*) FROM BPBNPS_WDCP " _
                                                     & "WHERE tgl_bpbw IS NULL AND FINISHW IS NULL"))
                Brg2 = Integer.Parse(cUtility.ExecuteScalar("SELECT Count(*) FROM BPBNPS_WDCP "))

                txtRec.Text = Brg & "/" & Brg2
                txtCont.Text = "-"
            End If

            _socketManager.Start()
            DCP1_box.Enabled = True
            DCP2_box.Enabled = True
            '6/10/20
            If jenis_so = "BPBBKL" Or jenis_so = "BPBNPS" Then
                DCP3_box.Enabled = True
                DCP4_box.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    Private Sub SettingDisplay(ByVal menu As String)
        Me.Height = 640
        lblIpServer.Top = 580
        lblAppVersion.Top = 580

        'tambahan 6/10/20 KHUSUS BPB BKL
        DCP3_box.Visible = False
        DCP4_box.Visible = False
        DCP3.Visible = False
        DCP4.Visible = False
        PictureBox1.Location = New Point(393, 52)
        Label2.Location = New Point(454, 284)
        txtRec.Location = New Point(349, 284)
        Me.Width = 559
        gbxDCP.Width = 508

        Label5.Visible = True
        cmbShelfTo.Visible = True

        Label9.Visible = False
        Label8.Visible = False
        cmbShelfFrom2.Visible = False
        cmbShelfTo2.Visible = False

        Label6.Text = "No.Shelf"

        If menu = "SO" Then
            GbxHandheld.Visible = True
            gbxLihatSO.Visible = False
            gbxDCP.Visible = False
            GbxPilihGudang.Visible = False
            gbxPlanogram.Visible = False
        ElseIf menu = "CekPJR" Then
            GbxHandheld.Visible = False
            gbxLihatSO.Visible = False
            gbxDCP.Visible = False
            GbxPilihGudang.Visible = False
            gbxPlanogram.Visible = True
            gbxPlanogram.Location = New Point(21, 102)

            cmbModis.SelectedIndex = -1
            cmbShelfFrom.SelectedIndex = -1
            cmbShelfTo.SelectedIndex = -1
            cmbShelfFrom.Enabled = False
            cmbShelfTo.Enabled = False
            txtNamaModis.Text = ""

            Label9.Visible = True
            Label8.Visible = True
            cmbShelfFrom2.Visible = True
            cmbShelfTo2.Visible = True

            Label5.Visible = False
            cmbShelfTo.Visible = False

            Label6.Text = "No.Rak"
            Label9.Text = "No.Shelf"


            DCP1.Text = "IP="
            DCP2.Text = "IP="
            Plano1_Box.Enabled = False
            Plano2_Box.Enabled = False
        ElseIf menu = "TINDAKLBTD" Then
            GbxHandheld.Visible = False
            gbxLihatSO.Visible = False
            gbxDCP.Visible = False
            GbxPilihGudang.Visible = False
            gbxPlanogram.Visible = True
            gbxPlanogram.Location = New Point(21, 102)

            cmbModis.SelectedIndex = -1
            cmbShelfFrom.SelectedIndex = -1
            cmbShelfTo.SelectedIndex = -1
            cmbShelfFrom.Enabled = False
            cmbShelfTo.Enabled = False
            txtNamaModis.Text = ""

            Label9.Visible = True
            Label8.Visible = True
            cmbShelfFrom2.Visible = True
            cmbShelfTo2.Visible = True

            Label5.Visible = False
            cmbShelfTo.Visible = False

            Label6.Text = "No.Rak"
            Label9.Text = "No.Shelf"


            DCP1.Text = "IP="
            DCP2.Text = "IP="
            Plano1_Box.Enabled = False
            Plano2_Box.Enabled = False

        ElseIf menu = "TINDAKLBTD_BAPJR" Then
            GbxHandheld.Visible = False
            gbxLihatSO.Visible = False
            gbxDCP.Visible = False
            GbxPilihGudang.Visible = False
            gbxPlanogram.Visible = True
            gbxPlanogram.Location = New Point(21, 102)

            cmbModis.SelectedIndex = -1
            cmbShelfFrom.SelectedIndex = -1
            cmbShelfTo.SelectedIndex = -1
            cmbShelfFrom.Visible = False
            cmbShelfTo.Visible = False
            txtNamaModis.Text = ""
            Label6.Text = "Jumlah"
            Label9.Visible = False
            Label8.Visible = False
            cmbShelfFrom2.Visible = False
            cmbShelfTo2.Visible = False

            Label5.Visible = False
            cmbShelfTo.Visible = False

            Label6.Text = "No.Rak"
            Label9.Text = "No.Shelf"


            DCP1.Text = "IP="
            DCP2.Text = "IP="
            Plano1_Box.Enabled = False
            Plano2_Box.Enabled = False
        End If

    End Sub
    Private Sub cmbModis_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbModis.SelectedIndexChanged
        Dim ClsPlano As New ClsPlanoController
        Dim ClsKesegaran As New ClsKesegaranController
        Dim ClsPJR As New ClsPJRController
        Dim ClsCekDisplay As New ClsCekDisplayController
        Dim DtModis As New DataTable
        Dim CountModis As Integer = 0
        Dim NamaModis As String = ""
        Dim Modis As Boolean
        Dim jumlahItem As String = ""
        Dim tanggal As String = ""

        If cmbModis.SelectedIndex <> -1 Then
            cmbShelfFrom.Enabled = True
            cmbShelfTo.Enabled = True
            cmbShelfFrom.Items.Clear()
            cmbShelfTo.Items.Clear()
            Try
                If jenis_so = "CekPJR" Or jenis_so = "TINDAKLBTD" Then
                    cmbShelfFrom2.Items.Clear()
                    cmbShelfTo2.Items.Clear()
                    cmbShelfFrom2.Text = ""
                    cmbShelfTo2.Text = ""
                    cmbShelfFrom.Text = ""
                    cmbShelfTo.Text = ""
                    Dim NikToko As String = ""
                    NikToko = ClsPJR.getConstNIKPJR

                    DtModis = ClsPJR.CekModis(cmbModis.Text.Split("-")(0).Trim, cmbModis.Text.Split("-")(1).Trim.Replace("/", "-"), NikToko, CountModis, NamaModis, If(jenis_so = "CekPJR", True, False))
                    'DtModis = ClsPJR.CekModis(cmbModis.Text.Split("-")(0).Trim, Now.ToString("dd-MM-yyyy"), CountModis, NamaModis, True)

                    If CountModis = 0 Then
                        MsgBox("Modis ini sudah selesai diproses!")
                        cmbShelfFrom.SelectedIndex = -1
                        cmbShelfTo.SelectedIndex = -1
                        cmbShelfFrom.Enabled = False
                        cmbShelfTo.Enabled = False

                        cmbShelfFrom2.SelectedIndex = -1
                        cmbShelfTo2.SelectedIndex = -1
                        cmbShelfFrom2.Enabled = False
                        cmbShelfTo2.Enabled = False

                    Else


                        txtNamaModis.Text = NamaModis
                        'cmbShelfFrom.Text = DtModis.Rows(0)("norak").ToString

                        'cmbShelfTo2.Text = DtModis.Rows(DtModis.Rows.Count - 1)("noshelf").ToString
                        cmbShelfFrom.Enabled = True
                        cmbShelfTo.Enabled = False

                        If DtModis.Rows.Count > 0 Then
                            For Each Dr As DataRow In DtModis.Rows
                                cmbShelfFrom.Items.Add(Dr(0))
                            Next
                            'InitialClient()
                            'btnOk.Enabled = True

                        End If

                    End If

                ElseIf jenis_so = "Kesegaran" Then
                    If cmbModis.SelectedIndex <> -1 Then
                        cmbShelfFrom.Enabled = True
                        cmbShelfTo.Enabled = True
                        cmbShelfFrom.Items.Clear()
                        cmbShelfTo.Items.Clear()
                        Try
                            Modis = ClsKesegaran.CekModis(cmbModis.Text, jumlahItem, NamaModis)
                            If Modis = False Then
                                MsgBox("Modis ini sudah selesai diproses!")
                                cmbShelfFrom.SelectedIndex = -1
                                cmbShelfTo.SelectedIndex = -1
                                cmbShelfFrom.Enabled = False
                                cmbShelfTo.Enabled = False
                            Else
                                txtNamaModis.Text = NamaModis
                                JmlItemKesegaran.Text = jumlahItem
                            End If
                        Catch ex As Exception
                            ShowError("Error load noshelf!", ex)
                            Exit Sub
                        End Try
                    End If
                ElseIf jenis_so = "TINDAKLBTD_BAPJR" Then

                    If ClsPJR.CekListBAPJR = False Then
                        MsgBox("Modis ini sudah selesai diproses!")

                    Else
                        ClsPJR.GetTanggalAkumulasiLBTD_BAPJR(tanggal, jumlahItem)
                        txtNamaModis.Text = tanggal
                        JmlItemKesegaran.Text = jumlahItem
                        If tanggal <> "" Then
                            InitialClient()
                            btnOk.Enabled = True
                        Else
                            btnOk.Enabled = False
                        End If



                    End If

                ElseIf jenis_so = "CekDisplay" Then
                    txtNamaModis.Text = ClsCekDisplay.CekModis(cmbModis.Text)


                Else
                    DtModis = ClsPlano.CekModis(cmbModis.Text, CountModis, NamaModis)
                    If CountModis = 0 Then
                        MsgBox("Modis ini sudah selesai diproses!")
                        cmbShelfFrom.SelectedIndex = -1
                        cmbShelfTo.SelectedIndex = -1
                        cmbShelfFrom.Enabled = False
                        cmbShelfTo.Enabled = False
                    Else
                        txtNamaModis.Text = NamaModis
                        If DtModis.Rows.Count > 0 Then
                            For Each Dr As DataRow In DtModis.Rows
                                cmbShelfFrom.Items.Add(Dr("noshelf"))
                                cmbShelfTo.Items.Add(Dr("noshelf"))
                            Next
                        End If
                    End If
                End If
            Catch ex As Exception
                ShowError("Error load noshelf!", ex)
                Exit Sub
            End Try
        End If
    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click

        Dim cPlano As New ClsPlanoController
        Dim cKesegaran As New ClsKesegaranController
        Dim cPJR As New ClsPJRController
        Dim cCekDisplay As New ClsCekDisplayController

        Dim Client As New ClsClient
        Dim util As New Utility
        Try
            If jenis_so = "Planogram" Then
                If btnOk.Text.ToUpper.Contains("MULAI") Then
                    If cmbModis.Text = "" Or cmbModis.SelectedIndex = -1 Then
                        MsgBox("Pilih Modis Terlebih Dahulu")
                        Exit Sub
                    End If
                    If cmbShelfFrom.Text = "" Or cmbShelfFrom.SelectedIndex = -1 Then
                        MsgBox("Pilih No.Shelf Terlebih Dahulu")
                        Exit Sub
                    End If
                    If cmbShelfTo.Text = "" Or cmbShelfTo.SelectedIndex = -1 Then
                        MsgBox("Pilih No.Shelf Terlebih Dahulu")
                        Exit Sub
                    End If

                    _socketManager.Start()
                    Plano1_Box.Enabled = True
                    Plano2_Box.Enabled = True

                    NoShelfStr = ""
                    For i As Integer = cmbShelfFrom.Text To cmbShelfTo.Text
                        NoShelfStr = NoShelfStr & "'" & i & "',"
                    Next
                    If NoShelfStr.Trim.Length > 1 Then
                        NoShelfStr = NoShelfStr.Remove(NoShelfStr.Length - 1, 1)
                    End If

                    cmbShelfFrom.Enabled = False
                    cmbShelfTo.Enabled = False

                    'btnOk.Text = "Selesai Scan"
                    btnOk.Text = "Batal"
                    btnSimpan.Visible = True

                ElseIf btnOk.Text.ToUpper.Contains("BATAL") Then
                    Dim CekPlano As String = cPlano.CekPlano(tabel_name, cmbModis.Text, NoShelfStr)
                    If CekPlano = "" Then
                        MessageBox.Show("Batal Proses Scan item" & vbCrLf, "Cek Planogram" & "", MessageBoxButtons.OK,
                                        MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        btnOk.Text = "Mulai Scan"
                        cmbModis.SelectedIndex = -1
                        NoShelfStr = ""
                        txtNamaModis.Text = ""
                        cmbShelfFrom.Text = ""
                        cmbShelfTo.Text = ""
                        cmbShelfFrom.Enabled = False
                        cmbShelfTo.Enabled = False
                        btnSimpan.Visible = False
                    Else
                        MessageBox.Show("Proses Batal belum dapat dilakukan...! Silahkan menyelesaikan Proses Scan item terlebih dahulu", "Cek Planogram", MessageBoxButtons.OK,
                                        MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    End If

                End If

            ElseIf jenis_so = "Kesegaran" Then
                If btnOk.Text.ToUpper.Contains("MULAI") Then
                    If cmbModis.Text = "" Or cmbModis.SelectedIndex = -1 Then
                        MsgBox("Pilih Modis Terlebih Dahulu")
                        Exit Sub
                    End If

                    If cKesegaran.MulaiCekKesegaran(cmbModis.Text) = True Then
                        _socketManager.Start()
                        Plano1_Box.Enabled = True
                        Plano2_Box.Enabled = True
                        btnOk.Text = "Batal"
                        btnSimpan.Visible = True
                        'btnOk.Text = "Selesai Scan"
                    Else
                        MsgBox("Modis " & cmbModis.Text & " telah Selesai diproses", MsgBoxStyle.Critical)

                    End If
                    cmbModis.Enabled = False
                ElseIf btnOk.Text.ToUpper.Contains("BATAL") Then
                    Dim CekPlano As String = cKesegaran.CekKesegaran(cmbModis.Text, NoShelfStr)
                    If CekPlano = "" Then
                        MessageBox.Show("Batal Proses Scan item" & vbCrLf, "Cek Kesegaran" & "", MessageBoxButtons.OK,
                                        MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        btnOk.Text = "Mulai Scan"
                        cmbModis.SelectedIndex = -1
                        NoShelfStr = ""
                        txtNamaModis.Text = ""
                        cmbShelfFrom.Text = ""
                        cmbShelfTo.Text = ""
                        cmbShelfFrom.Enabled = False
                        cmbShelfTo.Enabled = False
                        'btnSimpan.Visible = False
                        cmbModis.Enabled = True
                        JmlItemKesegaran.Text = ""

                        btnSimpan.Visible = False

                    Else
                        MessageBox.Show("Proses Batal belum dapat dilakukan...! Silahkan menyelesaikan Proses Scan item terlebih dahulu", "Cek Kesegaran", MessageBoxButtons.OK,
                                        MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    End If

                End If

            ElseIf jenis_so = "CekPJR" Or jenis_so = "TINDAKLBTD" Then
                If btnOk.Text.ToUpper.Contains("MULAI") Then
                    If cmbModis.Text = "" Or cmbModis.SelectedIndex = -1 Then
                        MsgBox("Pilih Modis Terlebih Dahulu")
                        Exit Sub
                    End If
                    ReloadCekPJR(cmbModis.Text)

                    _socketManager.Start()
                    Plano1_Box.Enabled = True
                    Plano2_Box.Enabled = True

                    NoRakStr = ""

                    For i As Integer = cmbShelfFrom2.Text To cmbShelfTo2.Text
                        NoShelfStr = NoShelfStr & "'" & i & "',"
                    Next

                    If NoShelfStr.Trim.Length > 1 Then
                        NoShelfStr = NoShelfStr.Remove(NoShelfStr.Length - 1, 1)
                    End If

                    cmbShelfFrom.Enabled = False
                    cmbShelfTo.Enabled = False
                    cmbShelfFrom2.Enabled = False
                    cmbShelfTo2.Enabled = False

                    'btnOk.Text = "Selesai Scan"
                    btnOk.Text = "Batal"
                    btnSimpan.Visible = True

                    norak_pjr = cmbShelfFrom.Text

                ElseIf btnOk.Text.ToUpper.Contains("BATAL") Then
                    Dim CekPJR As String = cPJR.CekPJR(tabel_name, cmbModis.Text.Split("-")(0).Trim, NoShelfStr)
                    If CekPJR = "" Then
                        MessageBox.Show("Batal Proses Scan item" & vbCrLf, "Cek Planogram" & "", MessageBoxButtons.OK,
                                        MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        btnOk.Text = "Mulai Scan"
                        cmbModis.SelectedIndex = -1
                        NoShelfStr = ""
                        txtNamaModis.Text = ""
                        cmbShelfFrom.Text = ""
                        cmbShelfTo.Text = ""
                        cmbShelfFrom2.Text = ""
                        cmbShelfTo2.Text = ""
                        cmbShelfFrom.Enabled = False
                        cmbShelfTo.Enabled = False
                        cmbShelfFrom2.Enabled = False
                        cmbShelfTo2.Enabled = False
                        btnSimpan.Visible = False
                    Else
                        MessageBox.Show("Proses Batal belum dapat dilakukan...! Silahkan menyelesaikan Proses Scan item terlebih dahulu", "Cek Planogram", MessageBoxButtons.OK,
                                        MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    End If


                End If

            ElseIf jenis_so = "TINDAKLBTD_BAPJR" Then
                If btnOk.Text.ToUpper.Contains("MULAI") Then
                    If cmbModis.Text = "" Or cmbModis.SelectedIndex = -1 Then
                        MsgBox("Pilih Modis Terlebih Dahulu")
                        Exit Sub
                    End If

                    _socketManager.Start()
                    Plano1_Box.Enabled = True
                    Plano2_Box.Enabled = True

                    NoRakStr = ""

                    'NoRakStr = cmbShelfFrom.Text
                    cmbShelfFrom.Enabled = False
                    cmbShelfTo.Enabled = False
                    cmbShelfFrom2.Enabled = False
                    cmbShelfTo2.Enabled = False

                    'btnOk.Text = "Selesai Scan"
                    btnOk.Text = "Batal"
                    btnSimpan.Visible = True

                ElseIf btnOk.Text.ToUpper.Contains("BATAL") Then
                    Dim CekPJR As String = cPJR.CekPJR(tabel_name, cmbModis.Text.Split("-")(0).Trim, NoShelfStr)
                    If CekPJR = "" Then
                        MessageBox.Show("Batal Proses Scan item" & vbCrLf, "Cek Planogram" & "", MessageBoxButtons.OK,
                                        MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        btnOk.Text = "Mulai Scan"
                        cmbModis.SelectedIndex = -1
                        NoShelfStr = ""
                        txtNamaModis.Text = ""
                        cmbShelfFrom.Text = ""
                        cmbShelfTo.Text = ""
                        cmbShelfFrom2.Text = ""
                        cmbShelfTo2.Text = ""
                        cmbShelfFrom.Enabled = False
                        cmbShelfTo.Enabled = False
                        cmbShelfFrom2.Enabled = False
                        cmbShelfTo2.Enabled = False
                        btnSimpan.Visible = False
                    Else
                        MessageBox.Show("Proses Batal belum dapat dilakukan...! Silahkan menyelesaikan Proses Scan item terlebih dahulu", "Cek Planogram", MessageBoxButtons.OK,
                                        MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    End If


                End If

            ElseIf jenis_so = "CekDisplay" Then
                If btnOk.Text.ToUpper.Contains("MULAI") Then
                    If cmbModis.Text = "" Or cmbModis.SelectedIndex = -1 Then
                        MsgBox("Pilih Modis Terlebih Dahulu")
                        Exit Sub
                    End If
                    Dim CekPlano As String = cCekDisplay.DeleteTempCekDisplay(tabel_name)

                    _socketManager.Start()
                    Plano1_Box.Enabled = True
                    Plano2_Box.Enabled = True

                    NoRakStr = ""

                    'NoRakStr = cmbShelfFrom.Text
                    cmbShelfFrom.Enabled = False
                    cmbShelfTo.Enabled = False
                    cmbShelfFrom2.Enabled = False
                    cmbShelfTo2.Enabled = False

                    'btnOk.Text = "Selesai Scan"
                    btnOk.Text = "Batal"
                    btnSimpan.Visible = True

                ElseIf btnOk.Text.ToUpper.Contains("BATAL") Then
                    Dim CekPlano As String = cCekDisplay.CekPlano(tabel_name, cmbModis.Text, NoShelfStr)
                    If CekPlano = "" Then
                        MessageBox.Show("Batal Proses Scan item" & vbCrLf, " Cek Display" & "", MessageBoxButtons.OK,
                                        MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        btnOk.Text = "Mulai Scan"
                        cmbModis.SelectedIndex = -1
                        NoShelfStr = ""
                        txtNamaModis.Text = ""
                        cmbShelfFrom.Text = ""
                        cmbShelfTo.Text = ""
                        cmbShelfFrom.Enabled = False
                        cmbShelfTo.Enabled = False
                        btnSimpan.Visible = False
                    Else
                        MessageBox.Show("Proses Batal belum dapat dilakukan...! Silahkan menyelesaikan Proses Scan item terlebih dahulu", "Cek Display", MessageBoxButtons.OK,
                                        MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    End If


                End If

            End If


        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub
    Private Sub btnSimpan_Click(sender As Object, e As EventArgs) Handles btnSimpan.Click
        Dim cKesegaran As New ClsKesegaranController
        Dim cPlano As New ClsPlanoController
        Dim cPJR As New ClsPJRController

        Dim Client As New ClsClient
        Dim util As New Utility
        Dim confirm As DialogResult
        Dim Result As DialogResult
        Dim DtRak As New DataTable

        Try
            For Each mClient As ClsClient In listClient
                If Not IsNothing(mClient) Then
                    Client = mClient
                    Exit For
                End If
            Next
            If jenis_so = "Planogram" Then
                If Not IsNothing(Client.Login) Then
                    confirm = MessageBox.Show("Apakah Anda yakin akan Proses Selesai Scan?", "Cek Planogram", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If confirm = Windows.Forms.DialogResult.Yes Then
                        Dim CekPlano As Boolean = cPlano.SelesaiCekPlano(tabel_name, cmbModis.Text, NoShelfStr, Client.Login.User)
                        If CekPlano = False Then
                            MessageBox.Show("Proses Scan sudah selesai...!" & vbCrLf &
                                            "Tidak ada barang yang tidak terdisplay!", "Cek Planogram" & "", MessageBoxButtons.OK,
                                            MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

                        Else
                            MessageBox.Show("Proses Scan sudah selesai...!", "Cek Planogram", MessageBoxButtons.OK,
                                            MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

                            Dim Frm As New frmRptCP
                            Frm.ShowDialog()
                        End If
                        btnOk.Text = "Mulai Scan"
                        cmbModis.SelectedIndex = -1
                        NoShelfStr = ""
                        txtNamaModis.Text = ""
                        cmbShelfFrom.Text = ""
                        cmbShelfTo.Text = ""
                        cmbShelfFrom.Enabled = False
                        cmbShelfTo.Enabled = False
                        btnSimpan.Visible = False
                    Else
                    End If

                Else
                    MessageBox.Show("Object class client WDCP is nothing!", "Cek Planogram", MessageBoxButtons.OK,
                                     MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    btnOk.Text = "Mulai Scan"
                    cmbModis.SelectedIndex = -1
                    NoShelfStr = ""
                    txtNamaModis.Text = ""
                    cmbShelfFrom.Text = ""
                    cmbShelfTo.Text = ""
                    cmbShelfFrom.Enabled = False
                    cmbShelfTo.Enabled = False
                    btnSimpan.Visible = False

                End If
            ElseIf jenis_so = "Kesegaran" Then
                'If Not IsNothing(Client.Login) Then
                Result = MessageBox.Show("Apakah Anda ingin mengakhiri dan menyimpan data Scan Untuk Modis " & cmbModis.Text & "?", "Scan Selesai..",
                                         MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                If Result = Windows.Forms.DialogResult.Yes Then
                    Dim CekPlano As Boolean = cKesegaran.SelesaiCekKesegaran(tabel_name, cmbModis.Text)
                    If CekPlano = False Then
                        MessageBox.Show("Proses Scan sudah selesai...!" & vbCrLf & "Tidak ada data", "Selesai" & "", MessageBoxButtons.OK,
                                                MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Else
                        MessageBox.Show("Proses Scan sudah selesai...!" & vbCrLf & "", "Selesai" & "", MessageBoxButtons.OK,
                                                MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    End If
                    'Else
                    'End If
                    cmbModis.Items.Clear()
                    DtRak.Clear()
                    DtRak = cKesegaran.GetNamaRak_byKesegaran
                    If DtRak.Rows.Count > 0 Then
                        For Each Dr As DataRow In DtRak.Rows
                            cmbModis.Items.Add(Dr(0))
                        Next
                        InitialClient()
                        btnOk.Enabled = True
                    Else
                        btnOk.Enabled = False
                    End If

                    btnOk.Text = "Mulai Scan"
                    cmbModis.SelectedIndex = -1
                    cmbModis.Text = ""
                    NoShelfStr = ""
                    txtNamaModis.Text = ""
                    cmbShelfFrom.Text = ""
                    cmbShelfTo.Text = ""
                    JmlItemKesegaran.Text = ""
                    cmbShelfFrom.Enabled = False
                    cmbShelfTo.Enabled = False
                    cmbModis.Enabled = True
                    _socketManager.StopServer()
                    Plano1_Box.Text = "IP="
                    Plano1_Box.Text = ""
                    Plano2_Box.Text = "IP="
                    Plano2_Box.Text = ""
                    Plano1_Box.Enabled = False
                    Plano2_Box.Enabled = False
                    btnSimpan.Visible = False
                End If
            ElseIf jenis_so = "CekPJR" Or jenis_so = "TINDAKLBTD" Then
                If Not IsNothing(Client.Login) Then
                    If jenis_so = "TINDAKLBTD" Then
                        confirm = MessageBox.Show("Apakah barang sudah dicek dengan benar? Apakah Anda sudah memastikan bahwa fisik barang tidak ada di Toko?", "Cek " & jenis_so & "", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                        If confirm = Windows.Forms.DialogResult.Yes Then
                            Dim CekPlano As Boolean = cPJR.SelesaiCekPJR(tabel_name, cmbModis.Text.Split("-")(0).Trim, NoShelfStr, Client.Login.User, cmbModis.Text.Split("-")(1).Trim, cmbShelfFrom.Text)
                            If CekPlano = False Then
                                MessageBox.Show("Proses Scan sudah selesai...!" & vbCrLf &
                                                "Tidak ada barang yang tidak terdisplay!", "Selesai" & " LGG2194", MessageBoxButtons.OK,
                                                MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

                                MessageBox.Show("Proses Scan Tindak Lanjut LBTD sudah selesai! Silahkan lanjut Proses Approval Pemegang Shift!", "Selesai" & " LGG2194", MessageBoxButtons.OK,
                                            MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                btnOk.Text = "Mulai Scan"
                                cmbModis.SelectedIndex = -1
                                cmbModis.Text = ""
                                NoShelfStr = ""
                                txtNamaModis.Text = ""
                                cmbShelfFrom.Text = ""
                                cmbShelfTo.Text = ""
                                cmbShelfFrom.Enabled = False
                                cmbShelfTo.Enabled = False
                                cmbModis.Enabled = True
                                _socketManager.StopServer()
                                Plano1_Box.Text = "IP="
                                Plano1_Box.Text = ""
                                Plano2_Box.Text = "IP="
                                Plano2_Box.Text = ""
                                Plano1_Box.Enabled = False
                                Plano2_Box.Enabled = False
                                btnSimpan.Visible = False

                                Dim Frm As New frmLBTD
                                Frm.ShowDialog()
                            Else
                                isPJR = True
                                'SIMPAN RAK PJR UTK PROSES SELANJUTNYA
                                cPJR.ConstRakPJR(cmbModis.Text & " - " & cmbShelfFrom.Text.Trim)

                                cPJR.buatTabelItemSO_BA_AS()
                                cPJR.insertDataBAPJR_AS(Client.Login.User.ID, cmbModis.Text.Split("-")(0).Trim, cmbShelfFrom.Text.Trim)

                                btnOk.Text = "Mulai Scan"
                                cmbModis.SelectedIndex = -1
                                cmbModis.Text = ""
                                cmbModis.Items.Clear()
                                NoShelfStr = ""
                                txtNamaModis.Text = ""
                                cmbShelfFrom.Text = ""
                                cmbShelfTo.Text = ""
                                cmbShelfFrom.Enabled = False
                                cmbShelfTo.Enabled = False
                                cmbModis.Enabled = True
                                _socketManager.StopServer()
                                Plano1_Box.Text = "IP="
                                Plano1_Box.Text = ""
                                Plano2_Box.Text = "IP="
                                Plano2_Box.Text = ""
                                Plano1_Box.Enabled = False
                                Plano2_Box.Enabled = False
                                btnSimpan.Visible = False

                                If jenis_so = "TINDAKLBTD" Then
                                    MessageBox.Show("Proses Scan Tindak Lanjut LBTD sudah selesai! Silahkan lanjut Proses Approval Pemegang Shift!", "Selesai" & " LGG2194", MessageBoxButtons.OK,
                                                MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    Dim Frm As New frmLBTD
                                    Frm.ShowDialog()
                                End If

                            End If
                            SettingDisplay("SO")
                        End If
                    Else
                        confirm = MessageBox.Show("Apakah Anda yakin akan Proses Selesai Scan?", "Cek " & jenis_so & "", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                        If confirm = Windows.Forms.DialogResult.Yes Then
                            Dim CekPlano As Boolean = cPJR.SelesaiCekPJR(tabel_name, cmbModis.Text.Split("-")(0).Trim, NoShelfStr, Client.Login.User, cmbModis.Text.Split("-")(1).Trim, cmbShelfFrom.Text)
                            If CekPlano = False Then
                                MessageBox.Show("Proses Scan sudah selesai...!" & vbCrLf &
                                                "Tidak ada barang yang tidak terdisplay!", "Selesai" & " LGG2194", MessageBoxButtons.OK,
                                                MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                SettingDisplay("SO")
                            Else
                                isPJR = True
                                'SIMPAN RAK PJR UTK PROSES SELANJUTNYA
                                cPJR.ConstRakPJR(cmbModis.Text & " - " & cmbShelfFrom.Text)

                                If jenis_so = "CekPJR" Then
                                    MessageBox.Show("Proses Scan sudah selesai! Silahkan untuk melanjutkan PROSES Tindak Lanjut LBTD!", "Selesai" & " LGG2194", MessageBoxButtons.OK,
                                                MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    MessageBox.Show("Harap Gunakan Printer Besar!", "Selesai" & " LGG2194", MessageBoxButtons.OK,
                                                MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

                                    Dim Frm As New frmTampilLBTD
                                    frmTampilLBTD.jenis = jenis_so
                                    Frm.Show()
                                End If
                            End If
                            _socketManager.StopServer()

                            btnOk.Text = "Mulai Scan"
                            cmbModis.Text = ""
                            Dim NikToko As String = ""
                            NoShelfStr = ""
                            txtNamaModis.Text = ""
                            cmbShelfFrom.Text = ""
                            cmbShelfTo.Text = ""
                            cmbShelfFrom.Enabled = False
                            cmbShelfTo.Enabled = False
                            cmbShelfFrom2.Text = ""
                            cmbShelfTo2.Text = ""
                            cmbShelfFrom2.Enabled = False
                            cmbShelfTo2.Enabled = False
                            btnSimpan.Visible = False

                            Plano1.Text = "IP="
                            Plano1_Box.Text = ""
                            Plano2.Text = "IP="
                            Plano2_Box.Text = ""

                            cmbModis.Enabled = False
                            SettingDisplay("SO")
                        Else

                        End If
                    End If
                Else
                    MessageBox.Show("Object class client WDCP is nothing!", "Selesai" & " LGG2194", MessageBoxButtons.OK,
                                 MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    btnOk.Text = "Mulai Scan"
                    cmbModis.SelectedIndex = -1
                    NoShelfStr = ""
                    txtNamaModis.Text = ""
                    cmbShelfFrom.Text = ""
                    cmbShelfTo.Text = ""
                    cmbShelfFrom.Enabled = False
                    cmbShelfTo.Enabled = False
                    btnSimpan.Visible = False
                    Plano1.Text = "IP="
                    Plano1_Box.Text = ""
                    Plano2.Text = "IP="
                    Plano2_Box.Text = ""
                End If

            ElseIf jenis_so = "TINDAKLBTD_BAPJR" Then
                If Not IsNothing(Client.Login) Then
                    confirm = MessageBox.Show("Apakah barang sudah dicek dengan benar? Apakah Anda sudah memastikan bahwa fisik barang tidak ada di Toko?", "Cek " & jenis_so & "", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If confirm = Windows.Forms.DialogResult.Yes Then
                        Dim CekPlano As Boolean = cPJR.SelesaiCekLBTD_BAPJR(tabel_name, Client.Login.User)
                        If CekPlano = False Then
                            MessageBox.Show("Proses Scan sudah selesai...!" & vbCrLf &
                                            "Tidak ada barang yang tidak terdisplay!", "Selesai" & " LGG2194", MessageBoxButtons.OK,
                                            MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

                            btnOk.Text = "Mulai Scan"
                            cmbModis.SelectedIndex = -1
                            cmbModis.Text = ""
                            'cmbModis.Items.Clear()
                            NoShelfStr = ""
                            txtNamaModis.Text = ""
                            cmbShelfFrom.Text = ""
                            cmbShelfTo.Text = ""
                            cmbShelfFrom.Enabled = False
                            cmbShelfTo.Enabled = False
                            cmbModis.Enabled = True
                            _socketManager.StopServer()
                            Plano1_Box.Text = "IP="
                            Plano1_Box.Text = ""
                            Plano2_Box.Text = "IP="
                            Plano2_Box.Text = ""
                            Plano1_Box.Enabled = False
                            Plano2_Box.Enabled = False
                            btnSimpan.Visible = False
                        Else
                            cPJR.buatTabelBA()
                            cPJR.loadDataBAPJR_AS()

                            Dim frm As New frmBAPJR
                            frm.Show()
                        End If
                    End If
                Else

                    MessageBox.Show("Object class client WDCP is nothing!", "Selesai" & " LGG2194", MessageBoxButtons.OK,
                                 MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    btnOk.Text = "Mulai Scan"
                    cmbModis.SelectedIndex = -1
                    NoShelfStr = ""
                    txtNamaModis.Text = ""
                    cmbShelfFrom.Text = ""
                    cmbShelfTo.Text = ""
                    cmbShelfFrom.Enabled = False
                    cmbShelfTo.Enabled = False
                    btnSimpan.Visible = False
                    Plano1.Text = "IP="
                    Plano1_Box.Text = ""
                    Plano2.Text = "IP="
                    Plano2_Box.Text = ""
                End If
            ElseIf jenis_so = "CekDisplay" Then
                If Not IsNothing(Client.Login) Then
                    confirm = MessageBox.Show("Apakah Anda yakin akan Proses Selesai Scan?", "Cek Display", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If confirm = Windows.Forms.DialogResult.Yes Then
                        Dim Frm As New FrmRptDCP
                        isCekDisplay = True
                        Frm.modis = cmbModis.Text
                        Frm.ShowDialog()
                        MessageBox.Show("Proses Scan sudah selesai...!", "Cek Display", MessageBoxButtons.OK,
                                                MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)


                        btnOk.Text = "Mulai Scan"
                        cmbModis.SelectedIndex = -1
                        NoShelfStr = ""
                        txtNamaModis.Text = ""
                        cmbShelfFrom.Text = ""
                        cmbShelfTo.Text = ""
                        cmbShelfFrom.Enabled = False
                        cmbShelfTo.Enabled = False
                        btnSimpan.Visible = False

                        _socketManager.StopServer()
                        Plano1.Text = "IP="
                        Plano2.Text = "IP="
                        Plano1_Box.Text = ""
                        Plano2_Box.Text = ""
                        Plano1_Box.Enabled = False
                        Plano2_Box.Enabled = False
                    End If
                Else
                    MessageBox.Show("Object class client WDCP is nothing!", "Cek Planogram", MessageBoxButtons.OK,
                                     MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    btnOk.Text = "Mulai Scan"
                    cmbModis.SelectedIndex = -1
                    NoShelfStr = ""
                    txtNamaModis.Text = ""
                    cmbShelfFrom.Text = ""
                    cmbShelfTo.Text = ""
                    cmbShelfFrom.Enabled = False
                    cmbShelfTo.Enabled = False
                    btnSimpan.Visible = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    Private Sub cmbShelfFrom_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbShelfFrom.SelectedIndexChanged
        Dim ClsPlano As New ClsPlanoController
        Dim ClsKesegaran As New ClsKesegaranController
        Dim ClsPJR As New ClsPJRController
        Dim DtModis As New DataTable
        Dim jumlahItem As String = ""

        Dim shelf_awal As String = ""
        Dim shelf_akhir As String = ""


        Try
            If jenis_so = "CekPJR" Or jenis_so = "TINDAKLBTD" Then
                Try


                    DtModis = ClsPJR.AmbilNoRak(cmbModis.Text.Split("-")(0).Trim, cmbModis.Text.Split("-")(1).Trim.Replace("/", "-"), cmbShelfFrom.Text)

                    cmbShelfFrom.Enabled = True
                    cmbShelfTo.Enabled = False

                    If DtModis.Rows.Count > 0 Then
                        shelf_awal = DtModis.Rows(0)("SHELFING").ToString.Split("-")(0)
                        shelf_akhir = DtModis.Rows(0)("SHELFING").ToString.Split("-")(1)

                        cmbShelfFrom2.Text = shelf_awal

                        cmbShelfTo2.Text = shelf_akhir
                        InitialClient()
                        btnOk.Enabled = True

                    End If
                Catch ex As Exception

                End Try
            End If
        Catch ex As Exception
            ShowError("Error load noshelf!", ex)
            Exit Sub
        End Try
    End Sub

    Private Sub btnRegisPersonil_Click_1(sender As Object, e As EventArgs) Handles btnRegisPersonil.Click
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

    Private Sub btnApprovalPJR_Click(sender As Object, e As EventArgs) Handles btnApprovalPJR.Click
        PnlLoading.Visible = True
        LblProses.Text = "Harap tunggu sedang proses data..."
        Application.DoEvents()
        Dim a As New frmApprovePJR

        a.Show()
        PnlLoading.Visible = False
        LblProses.Text = "Loading..."
        Application.DoEvents()
    End Sub

    Private Sub btnTIndakLBTDBAAS_Click(sender As Object, e As EventArgs) Handles btnTIndakLBTDBAAS.Click
        Try
            Dim confirm As DialogResult

            Dim cUser As New ClsUserController
            Dim cPJR As New ClsPJRController
            Dim DtRak As New DataTable
            Dim tanggal As String
            Dim rakpjr As String = ""
            Dim norak As String = ""
            Dim nik As String
            Dim tes As String() = {"a", "b", "c"}
            Dim NikToko As String = ""
            If cUser.CekUserSO = False Then
                MessageBox.Show("User toko belum ada", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Dim DtSO As New DataTable
                Dim cProduk As New ClsProdukController

                confirm = MessageBox.Show("Akan mulai Scan Finger AS/AM, Silahkan klik Yes utk melanjutkan.", "Tindak LBTD BA PJR", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                If confirm = Windows.Forms.DialogResult.Yes Then

                    tes = Panggil_CekFingerprintV3("frm", "WDCP")

                    If tes(0) = "" And tes(1) = "" And tes(2) = "" Then
                        MsgBox("Maaf, Validasi scanfinger AS/AM tidak berhasil !", MsgBoxStyle.Exclamation)
                        btnOk.Enabled = False
                        SettingDisplay("SO")

                        Debug.WriteLine("Password Salah")
                    Else
                        nik = cPJR.getConstNIKPJR
                        rakpjr = cPJR.getConstRakPJR.Split("-")(0).Trim
                        tanggal = cPJR.getConstRakPJR.Split("-")(1).Trim
                        norak = cPJR.getConstRakPJR.Split("-")(2).Trim

                        cPJR.updateITT_ADJUST(tanggal.Replace("/", "-"), nik, rakpjr, norak)

                        Dim Frm As New frmTampilLBTD
                        frmTampilLBTD.jenis = "TINDAKLBTD_BAPJR"
                        frmTampilLBTD.nik_as = "user"
                        Frm.Show()

                        tabel_name = "TINDAKLBTD_BAPJR"
                        jenis_so = "TINDAKLBTD_BAPJR"
                        SettingDisplay("TINDAKLBTD_BAPJR")
                        cPJR.CekTableTINDAKLBTD_BAPJR()
                        cmbModis.Items.Clear()
                        cmbModis.Enabled = True

                        cmbModis.Items.Add("AKUMULASI ITT")
                        Label6.Text = "Jumlah"
                        Label4.Text = "Tanggal"

                        'txtNamaModis.Text = cPJR.GetTanggalAkumulasiLBTD_BAPJR
                        'If txtNamaModis.Text <> "" Then
                        '    InitialClient()
                        '    btnOk.Enabled = True
                        'Else
                        '    btnOk.Enabled = False
                        'End If
                    End If


                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
        End Try

    End Sub

    Private Sub btnCetakLaporanPJR_Click(sender As Object, e As EventArgs) Handles btnCetakLaporanPJR.Click
        Dim Frm As New frmCPJR
        Frm.ShowDialog()
    End Sub

    Private Sub btnTIndakLanjutLBTD_Click(sender As Object, e As EventArgs) Handles btnTIndakLanjutLBTD.Click
        Try
            Dim cUser As New ClsUserController
            Dim cPJR As New ClsPJRController
            Dim DtRak As New DataTable
            Dim rakpjr As String = ""
            Dim tes As String() = {"a", "b", "c"}
            Dim NikToko As String = ""
            If cUser.CekUserSO = False Then
                MessageBox.Show("User toko belum ada", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                tes = Panggil_CekFingerprintV3("frm", "WDCP_PJR")

                If tes(0) = "" And tes(1) = "" And tes(2) = "" Then
                    MsgBox("Maaf, Validasi scanfinger tidak berhasil !", MsgBoxStyle.Exclamation)
                    btnOk.Enabled = False
                    SettingDisplay("SO")

                    Debug.WriteLine("Password Salah")
                Else
                    ConstNIKPJR(tes(2).Split("|")(0).Trim)

                    tabel_name = "TINDAKLBTD"
                    jenis_so = "TINDAKLBTD"
                    SettingDisplay("TINDAKLBTD")
                    cPJR.CekTableTINDAKLBTD()
                    rakpjr = cPJR.getConstRakPJR
                    NikToko = cPJR.getConstNIKPJR

                    DtRak = cPJR.GetNamaRakLBTD(Date.Now.ToString("yyyy-MM-dd"), NikToko, "", rakpjr)
                    cmbModis.Items.Clear()
                    cmbModis.Enabled = True
                    If DtRak.Rows.Count > 0 Then
                        For Each Dr As DataRow In DtRak.Rows
                            cmbModis.Items.Add(Dr(0))
                        Next

                        InitialClient()

                        btnOk.Enabled = True
                    Else
                        btnOk.Enabled = False
                    End If
                End If

            End If


        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    Private Sub btnPelaksanaanPJR_Click(sender As Object, e As EventArgs) Handles btnPelaksanaanPJR.Click
        Try
            Dim cUser As New ClsUserController
            Dim cPJR As New ClsPJRController
            Dim DtRak As New DataTable
            Dim NikToko As String = ""
            Dim notif1 As String = ""
            Dim notif2 As String = ""
            Dim tes As String() = {"a", "b", "c"}

            Dim result1 As String = ""
            Dim result2 As String = ""

            If cUser.CekUserSO = False Then
                MessageBox.Show("User toko belum ada", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else


                notif1 = "Ada Penambahan atau Pengurangan Jumlah Karyawan di Toko Idm. dikarenakan Adanya Mutasi Karyawan, Diharapkan Chief Of Store Melakukan Penjadwalan Ulang PJR !"
                notif2 = "Ada Penambahan atau Pengurangan Data Mo.Disp. di Toko Idm., Diharapkan Chief Of Store Melakukan Penjadwalan Ulang PJR !"

                If Date.Now.Day = 1 Or Date.Now.Day = 2 Or Date.Now.Day = 3 Or Date.Now.Day = 4 Or Date.Now.Day = 5 Then

                    result1 = cPJR.notif_cekJadwal_personil()
                    If result1 = False Then
                        MessageBox.Show(notif1, "Notifikasi", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Dim k As New FrmListMutasi_PJR
                        k.lblJudul.Text = "LIST PERUBAHAN MUTASI PERSONIL TOKO"
                        k.Show()

                    End If
                    result2 = cPJR.notif_cekJadwal_Modis()
                    If result2 = False Then
                        MessageBox.Show(notif2, "Notifikasi", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Dim k2 As New FrmListMutasi_PJR
                        k2.lblJudul.Text = "LIST PERUBAHAN DATA MODISP TOKO"
                        k2.Show()
                    End If
                ElseIf Date.Now.Day > 5 And Date.Now.Day < 16 Then
                    result1 = cPJR.notif_cekJadwal_personil()
                    If result1 = False Then
                        MessageBox.Show("Maaf, Pelaksanaan PJR Tidak Dapat Dilakukan Dikarenakan Penjadwalan Ulang PJR Belum Dilakukan !", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Dim k As New FrmListMutasi_PJR
                        k.lblJudul.Text = "LIST PERUBAHAN MUTASI PERSONIL TOKO"
                        k.Show()
                        result2 = cPJR.notif_cekJadwal_Modis()
                        If result2 = False Then
                            'MessageBox.Show("Maaf, Pelaksanaan PJR Tidak Dapat Dilakukan Dikarenakan Penjadwalan Ulang PJR Belum Dilakukan !", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Dim k2 As New FrmListMutasi_PJR
                            k2.lblJudul.Text = "LIST PERUBAHAN DATA MODISP TOKO"
                            k2.Show()
                            Exit Try

                        End If
                        Exit Try
                    Else
                        result2 = cPJR.notif_cekJadwal_Modis()
                        If result2 = False Then
                            MessageBox.Show("Maaf, Pelaksanaan PJR Tidak Dapat Dilakukan Dikarenakan Penjadwalan Ulang PJR Belum Dilakukan !", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Dim k2 As New FrmListMutasi_PJR
                            k2.lblJudul.Text = "LIST PERUBAHAN DATA MODISP TOKO"
                            k2.Show()
                            Exit Try
                        End If

                    End If

                End If

                If Date.Now.Day = 16 Or Date.Now.Day = 17 Or Date.Now.Day = 18 Or Date.Now.Day = 19 Or Date.Now.Day = 20 Then
                    result1 = cPJR.notif_cekJadwal_personil()
                    If result1 = False Then
                        MessageBox.Show(notif1, "Notifikasi", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Dim k As New FrmListMutasi_PJR
                        k.lblJudul.Text = "LIST PERUBAHAN MUTASI PERSONIL TOKO"
                        k.Show()
                    End If
                    result2 = cPJR.notif_cekJadwal_Modis()
                    If result2 = False Then
                        MessageBox.Show(notif2, "Notifikasi", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Dim k2 As New FrmListMutasi_PJR
                        k2.lblJudul.Text = "LIST PERUBAHAN DATA MODISP TOKO"
                        k2.Show()
                    End If
                ElseIf Date.Now.Day > 20 Then
                    result1 = cPJR.notif_cekJadwal_personil()
                    If result1 = False Then
                        MessageBox.Show("Maaf, Pelaksanaan PJR Tidak Dapat Dilakukan Dikarenakan Penjadwalan Ulang PJR Belum Dilakukan !", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Dim k As New FrmListMutasi_PJR
                        k.lblJudul.Text = "LIST PERUBAHAN MUTASI PERSONIL TOKO"
                        k.Show()

                        result2 = cPJR.notif_cekJadwal_Modis()
                        If result2 = False Then
                            'MessageBox.Show("Maaf, Pelaksanaan PJR Tidak Dapat Dilakukan Dikarenakan Penjadwalan Ulang PJR Belum Dilakukan !", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Dim k2 As New FrmListMutasi_PJR
                            k2.lblJudul.Text = "LIST PERUBAHAN DATA MODISP TOKO"
                            k2.Show()
                            Exit Try
                        End If
                        Exit Try
                    Else
                        result2 = cPJR.notif_cekJadwal_Modis()
                        If result2 = False Then
                            MessageBox.Show("Maaf, Pelaksanaan PJR Tidak Dapat Dilakukan Dikarenakan Penjadwalan Ulang PJR Belum Dilakukan !", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Dim k2 As New FrmListMutasi_PJR
                            k2.lblJudul.Text = "LIST PERUBAHAN DATA MODISP TOKO"
                            k2.Show()
                            Exit Try
                        End If

                    End If

                End If

                tes = Panggil_CekFingerprintV3("frmSO", "WDCP_PJR")
                If tes(0) = "" And tes(1) = "" And tes(2) = "" Then
                    MsgBox("Maaf, Validasi scanfinger tidak berhasil !", MsgBoxStyle.Exclamation)
                    btnOk.Enabled = False
                    SettingDisplay("SO")

                    Debug.WriteLine("Password Salah")
                Else
                    ConstNIKPJR(tes(2).Split("|")(0).Trim)

                    tabel_name = "CekPJR"
                    jenis_so = "CekPJR"
                    SettingDisplay("CekPJR")
                    cPJR.CekTablePJR()

                    NikToko = cPJR.getConstNIKPJR
                    cPJR.ReloadJadwal(True)
                    cPJR.insertTempJadwalPJR(NikToko)

                    DtRak = cPJR.GetNamaRak(Date.Now.ToString, NikToko)
                    cmbModis.Items.Clear()

                    If DtRak.Rows.Count > 0 Then
                        For Each Dr As DataRow In DtRak.Rows
                            cmbModis.Items.Add(Dr(0))
                        Next
                    Else
                        MsgBox("Maaf, jadwal PJR Anda belum terdaftar atau belum di-Approve !", MsgBoxStyle.Exclamation)
                        btnOk.Enabled = False
                        SettingDisplay("SO")
                    End If
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub
End Class
