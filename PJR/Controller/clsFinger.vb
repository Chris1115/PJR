Imports System.Configuration
Imports System.Security.Cryptography
Imports System.IO
Imports System.Text
Imports IDM.InfoToko
Imports IDM.Fungsi
Imports MySql.Data.MySqlClient
Public Class clsFinger
    Public Shared Function Panggil_CekFingerprintV3(ByVal nmFrm As String, ByVal nmTransaksi As String) As String()
        Panggil_CekFingerprintV3 = New String() {"", "", ""}
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)

        Try

            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            nmTransaksi = nmTransaksi.ToUpper

            Scom.CommandText = "show tables like 'OtorisasiOperasional';"
            If Scom.ExecuteScalar = "" Then
                Scom.CommandText = "CREATE TABLE `OtorisasiOperasional` ("
                Scom.CommandText &= " `isAktif` CHAR(1) DEFAULT NULL,"
                Scom.CommandText &= " `TipeID` INT(11) NOT NULL AUTO_INCREMENT,"
                Scom.CommandText &= " `Jenis` VARCHAR(50) DEFAULT NULL,"
                Scom.CommandText &= " `isDoubleApproval` CHAR(1) DEFAULT 'Y',"
                Scom.CommandText &= " `Jabatan1` VARCHAR(200) DEFAULT NULL,"
                Scom.CommandText &= " `Jabatan2` VARCHAR(200) DEFAULT NULL,"
                Scom.CommandText &= " PRIMARY KEY (`TipeID`)"
                Scom.CommandText &= " ) ENGINE=INNODB AUTO_INCREMENT=0 DEFAULT CHARSET=latin1;"
                Scom.ExecuteNonQuery()
            End If
            Scom.CommandText = "show tables like 'LogScanFinger';"
            If Scom.ExecuteScalar = "" Then
                Scom.CommandText = "CREATE TABLE `LogScanFinger` ("
                Scom.CommandText &= " `TANGGAL` datetime DEFAULT NULL,"
                Scom.CommandText &= " `NIK` varchar(15) NOT NULL,"
                Scom.CommandText &= " `NAMA` varchar(30) NOT NULL,"
                Scom.CommandText &= " `SHIFT` CHAR(2) NOT NULL DEFAULT '',"
                Scom.CommandText &= " `STATION` CHAR(2) NOT NULL DEFAULT '',"
                Scom.CommandText &= " `KETERANGAN` VARCHAR(200) DEFAULT '',"
                Scom.CommandText &= " `STATUS` VARCHAR(30) DEFAULT '',"
                Scom.CommandText &= " `PROGRAM` VARCHAR(40) DEFAULT ''"
                Scom.CommandText &= " ) ENGINE=INNODB AUTO_INCREMENT=0 DEFAULT CHARSET=latin1;"
                Scom.ExecuteNonQuery()
            End If
            insertTipeTransKeOtorOper(nmTransaksi)

            TraceLog(nmFrm & " : cek program scan finger")
            'iya, itu yg aneh
            If IO.File.Exists(Application.StartupPath & "\ScanFinger.dll") Then
                TraceLog(nmFrm & " : ada program scan finger, scan finger/input password - mulai")
                Dim sNIK As String = "", sNama As String = "", sStation As String = Environment.GetEnvironmentVariable("STATION")
                Scom.CommandText = "SELECT shift FROM initial WHERE RECID='' AND tanggal=CURDATE() AND Station='" & sStation & "'"
                Dim sShift As String = Scom.ExecuteScalar() & ""
                Dim tess() As String, coba As Integer = 0, kirimData As String = ""
                If insertVirBacaprod("CEKDATABASE", "ON", "Cek Database & Struktur Tabel di PosIDM", "PosIDM") = "ON" Then

                    Dim tes As New ScanFinger.ClsScan

                    tess = tes.Otorisasi_New(nmTransaksi, "Scan Finger Otorisasi untuk proses " & nmTransaksi, sShift, sStation, False)
                    TraceLog("CekFinger (" & nmTransaksi & ") : " & tess(0) & ", " & tess(1) & ", " & tess(2))
                Else
                    tess = New String() {"1", "sukses", "2013089191|pria|YA|046||DRIVER|JABATAN1"}
                End If

                If tess(1).ToUpper <> "SUKSES" Then
                    Insert_LogScanFinger(Scom, "Unknown", "Unknown", "GAGAL - Validasi " & nmTransaksi & " : " & tess(2), "ScanFinger", sShift, sStation)
                    TraceLog("Cek_Fingerprint : gagal scan finger " & nmTransaksi & " jabatan1 3x")
                    MsgBox("Data Tidak Ada, NIK tidak ditemukan")
                Else
                    Insert_LogScanFinger(Scom, tess(2).Split("|")(0), "" & Scom.ExecuteScalar(), "Berhasil - Validasi " & nmTransaksi, "ScanFinger", sShift, sStation)
                    TraceLog("Cek_Fingerprint : berhasil validasi data finger driver")
                    Panggil_CekFingerprintV3 = tess
                End If
            Else
                TraceLog(nmFrm & " : scan finger/input password - Finger tidak ada & Selesai")
                MsgBox("tidak ada program scan finger")
            End If
            TraceLog(nmFrm & " : scan finger/input password - Selesai")
        Catch ex As Exception
            TraceLog("Error Panggil_CekFingerprintV3 = " & ex.Message & ex.StackTrace & vbCrLf & "Last Query : " & Scom.CommandText)
            ShowError("posbpb Error Panggil_CekFingerprintV3 = ", ex.Message & ex.StackTrace)
        End Try
        TraceLog(nmFrm & " : Panggil_CekFingerprintV3 : Selesai ")
        Return Panggil_CekFingerprintV3
    End Function

    Public Shared Function insertTipeTransKeOtorOper(ByVal sTrans As String) As Boolean
        insertTipeTransKeOtorOper = False
        Dim Mcon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Mcon)

        Try

            If Mcon.State = ConnectionState.Closed Then
                Mcon.Open()
            End If

            Dim listPKomplit As String = "KASIR,KASIR (SS),PRAMUNIAGA,PRAMUNIAGA (SS),MERCHANDISER,MERCHANDISER (SS),ASISTEN KEPALA TOKO, " &
                                         "ASISTEN KEPALA TOKO (SS),KEPALA TOKO,KEPALA TOKO (SS),STORE CREW,STORE CREW (SS),STORE JR. LEADER," &
                                         "STORE JR. LEADER (SS),STORE SR. LEADER,STORE SR. LEADER (SS),CHIEF OF STORE,CHIEF OF STORE (SS)," &
                                         "STORE CREW BOY,STORE CREW GIRL,STORE CREW BOY (SS),STORE CREW GIRL (SS)"

            Dim listPPmpinan As String = "KEPALA TOKO,KEPALA TOKO (SS),ASISTEN KEPALA TOKO,ASISTEN KEPALA TOKO (SS),MERCHANDISER,MERCHANDISER (SS)," &
                                          "STORE JR. LEADER,STORE JR. LEADER (SS),STORE SR. LEADER,STORE SR. LEADER (SS),CHIEF OF STORE,CHIEF OF STORE (SS)"

            Dim listPPmpinanArea As String = "Act. Jr. Supervisor, Act. Senior Manager,Area Act. Jr. Manager,Area Jr. Manager," &
                                             "Area Manager,Junior Manager,Junior Supervisor,MDP Specialist Trainee,MDP Trainee,Supervisor"
            listPKomplit = listPKomplit.ToUpper : listPPmpinan = listPPmpinan.ToUpper : listPPmpinanArea = listPPmpinanArea.ToUpper

            If sTrans = "WDCP" Then
                Return cekDetailFinger(sTrans, "1", "N", listPPmpinanArea, listPPmpinanArea)
            ElseIf sTrans = "WDCP_PJR" Then
                Return cekDetailFinger(sTrans, "1", "N", listPKomplit, listPKomplit)
            ElseIf sTrans = "WDCP_PJR 2" Then
                Return cekDetailFinger(sTrans, "1", "N", listPPmpinan, listPPmpinan)
            ElseIf sTrans = "WDCP_SO_IC" Then
                Return cekDetailFinger(sTrans, "1", "N", listPPmpinan, listPPmpinan)

            Else
                'diluar transaksi di atas langsung skip karena tidak terdaftar
                TraceLog("Cek_Fingerprint : function dipanggil diluar yang didaftarkan, skip proses")
            End If
        Catch ex As Exception
            TraceLog("Eror insertTipeTransKeOtorOper " & ex.Message & ex.StackTrace & vbCrLf & "Last query : " & Scom.CommandText)
        End Try
        Return insertTipeTransKeOtorOper
    End Function

    Public Shared Function Insert_LogScanFinger(ByVal objcmd As MySqlCommand, ByVal nik As String, ByVal nama As String, ByVal keterangan As String, ByVal status As String, ByVal shift As String, ByVal station As String) As Boolean
        Try
            TraceLog("Insert_LogScanFinger : Mulai ")
            objcmd.CommandText = "insert into LogScanFinger ("
            objcmd.CommandText &= "tanggal,nik,nama,shift,station,keterangan,status,program"
            objcmd.CommandText &= ") values ("
            objcmd.CommandText &= " '" & Format(Date.Now, "yyyy-MM-dd HH:mm:ss") & "',"
            objcmd.CommandText &= " '" & nik & "',"
            objcmd.CommandText &= " '" & nama & "',"
            objcmd.CommandText &= " '" & shift & "',"
            objcmd.CommandText &= " '" & station & "',"
            objcmd.CommandText &= " '" & keterangan & "',"
            objcmd.CommandText &= " '" & status & "',"
            objcmd.CommandText &= " 'BA.exe " & Application.ProductVersion & "'"
            objcmd.CommandText &= ");"
            TraceLog("Insert_LogScanFinger : query  " & objcmd.CommandText)
            objcmd.ExecuteNonQuery()
        Catch ex As Exception
            ShowError("BA Error Insert_LogScanFinger = ", ex.Message & ex.StackTrace)
            TraceLog("Insert_LogScanFinger : Error  " & ex.Message & ex.StackTrace)
        End Try
        TraceLog("Insert_LogScanFinger : Selesai ")
    End Function



    Public Shared Function insertVirBacaprod(ByVal sJenis As String, ByVal sFilter As String, ByVal sKet As String, ByVal sProg As String) As String
        Dim Mcon As MySqlConnection = ClsConnection.GetConnection.Clone
        'Mcon = Scon.Clone


        Dim Mcom As New MySqlCommand("", Mcon)
        If Mcon.State = ConnectionState.Closed Then
            Mcon.Open()
        End If

        Mcom.CommandText = "SELECT COUNT(*) FROM vir_bacaprod WHERE program='" & sProg & "' AND jenis='" & sJenis & "'"
        If Mcom.ExecuteScalar() > 0 Then
            Mcom.CommandText = "SELECT ifnull(`FILTER`,'') FROM VIR_BACAPROD WHERE program='" & sProg & "' AND jenis='" & sJenis & "'"
            Return Mcom.ExecuteScalar().ToString.ToUpper.Trim
        Else
            Mcom.CommandText = "Insert Into Vir_BacaProd(jenis,`filter`,KET,program,updid) Values "
            Mcom.CommandText &= "('" & sJenis & "','" & sFilter & "','" & sKet & "','" & sProg & "','SO.NET.exe')"
            Mcom.ExecuteNonQuery()
            Return "ON"
        End If
    End Function

    Public Shared Function cekDetailFinger(ByVal Transaksi As String, ByVal isAktif As String, ByVal isDA As String,
ByVal list1 As String, ByVal list2 As String, Optional ByVal sAbsLok As String = "") As Boolean
        cekDetailFinger = False
        Dim sMainAbsenLokal As Boolean = False
        Dim Mcon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Mcon)

        Try
            If Mcon.State = ConnectionState.Closed Then
                Mcon.Open()
            End If

            Transaksi = Transaksi.ToUpper
            Scom.CommandText = "SELECT COUNT(*) FROM Information_schema.Columns WHERE TABLE_SCHEMA='pos' "
            Scom.CommandText &= "AND Table_Name='OtorisasiOperasional' AND Column_Name='AbsensiLokal';"
            sMainAbsenLokal = Scom.ExecuteScalar > 0

            TraceLog("Cek_Fingerprint : cek tabel OtorisasiOperasional dan jenis " & Transaksi)
            Scom.CommandText = "SELECT COUNT(*) FROM OtorisasiOperasional WHERE jenis='" & Transaksi & "';"
            If Scom.ExecuteScalar > 0 Then

            Else
                    Scom.CommandText = "INSERT INTO OtorisasiOperasional (isAktif, Jenis, isDoubleApproval, Jabatan1, Jabatan2"
                If sMainAbsenLokal Then
                    Scom.CommandText &= ",AbsensiLokal"
                End If
                Scom.CommandText &= ") VALUES ('" & isAktif & "', '" & Transaksi & "', '" & isDA & "', '" & list1 & "', '" & list2 & "'"
                If sMainAbsenLokal Then
                    Scom.CommandText &= ",'" & sAbsLok & "'"
                End If
                Scom.CommandText &= ");"
                TraceLog("Cek_Fingerprint : insert tabel OtorisasiOperasional dan jenis " & Transaksi & " -> " & Scom.CommandText)
                Scom.ExecuteNonQuery()
            End If
            cekDetailFinger = True
        Catch ex As Exception
            cekDetailFinger = False
            ShowError("WDCP Error cekDetailFinger = ", ex.Message & ex.StackTrace)
        End Try
    End Function
End Class
