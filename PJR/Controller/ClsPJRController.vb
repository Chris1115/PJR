Imports MySql.Data.MySqlClient
Imports IDM.Fungsi
Imports IDM.InfoToko
Public Class ClsPJRController
    Public Function GetNamaRak(ByVal tanggal As String, Optional ByVal nik As String = "", Optional ByVal hari As String = "") As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable
        Dim listtanggal As String = ""
        Dim Mcom As New MySqlCommand("", Conn)
        Dim temp_tanggal As Date
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            If nik = "" And hari = "" Then
                TraceLog("GetNamaRak - nik & tgl kosong")
                Madp.SelectCommand.CommandText = "select distinct nama_rak from rak where nama_Rak NOT IN (SELECT KODE_MODIS FROM TEMP_JADWAL_PJR WHERE TANGGAL = '" & tanggal.Trim & "');"
                TraceLog("GetNamaRak - Get Data TEMP_JADWAL_PJR: " & Madp.SelectCommand.CommandText)
                Madp.Fill(Rtn)
            Else
                TraceLog("GetNamaRak - nik & tgl tidak kosong")
                temp_tanggal = tanggal
                tanggal = temp_tanggal.ToString("yyyy-MM-dd")

                Madp.SelectCommand.CommandText = "select  CONCAT(CAST(kode_modis AS CHAR(20)),' - ',
                                                    CAST(REPLACE(tanggal,'-','/') AS CHAR(15))) as Modis 
                                                    from temp_jadwal_pjr 
                                                    where nik = '" & nik & "' 
                                                    AND tanggal = curdate() AND (STATUSAPPROVAL = 'Y' OR STATUSAPPROVAL = 'P')
                                                    AND RECID = ''
                                                    GROUP BY MODIS
                                                  "
                TraceLog("GetNamaRak - Get Data TEMP_JADWAL_PJR: " & Madp.SelectCommand.CommandText)
                Rtn.Clear()
                Madp.Fill(Rtn)
            End If
        Catch ex As Exception
            TraceLog("Error WDCP_GetNamaRak : " & ex.Message & ex.StackTrace)
            Rtn = New DataTable
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function GetNamaRak_2() As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable
        Dim listtanggal As String = ""
        Dim Mcom As New MySqlCommand("", Conn)
        Dim cekJumlah As String = ""
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            cekJumlah = cekPerbandinganPersonilVSModis(FormMain.cbHariBukaToko)

            If cekJumlah = True Then
                Madp.SelectCommand.CommandText = "SELECT KODE_MODIS FROM temp_jadwal_penanggungjawabrak GROUP BY KODE_MODIS"

            Else
                Madp.SelectCommand.CommandText = "SELECT KODE_MODIS FROM temp_jadwal_penanggungjawabrak  WHERE shelfing = '' GROUP BY KODE_MODIS"

            End If

            ''Console.Writeline(Madp.SelectCommand.CommandText)
            Madp.Fill(Rtn)

        Catch ex As Exception
            TraceLog("Error WDCP_GetNamaRak : " & ex.Message & ex.StackTrace)
            Rtn = New DataTable
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function GetNamaRak_Pengganti(ByVal nik As String, ByVal hari As String) As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable
        Dim listtanggal As String = ""
        Dim Mcom As New MySqlCommand("", Conn)

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            If FormMain.isPengganti = False Then

            Else
                If hari.Length > 0 Then
                    hari = hari.Split(",")(0).Trim

                End If
            End If
            Madp.SelectCommand.CommandText = "SELECT DISTINCT kode_modis FROM jadwal_penanggungjawabrak WHERE nik = '" & nik & "' AND HARI = '" & hari & "'
                                        AND (nik,kode_modis,norak) NOT IN (SELECT nik,kode_modis,norak FROM temp_jadwal_penanggungjawabrak_pengganti
                                        WHERE nik = '" & nik & "');"

            Madp.Fill(Rtn)

        Catch ex As Exception
            TraceLog("Error WDCP_GetNamaRak : " & ex.Message & ex.StackTrace)
            Rtn = New DataTable
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function


    Public Function GetNamaRakLBTD(ByVal tanggal As String, Optional ByVal nik As String = "", Optional ByVal hari As String = "", Optional ByVal rak As String = "") As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            If nik = "" And hari = "" Then
                Madp.SelectCommand.CommandText = "select distinct nama_rak from rak where nama_Rak NOT IN (SELECT KODE_MODIS FROM TEMP_JADWAL_PJR WHERE TANGGAL = '" & tanggal.Trim & "');"
                Madp.Fill(Rtn)
            Else
                Madp.SelectCommand.CommandText = "select  CONCAT(CAST(kode_modis AS CHAR(20)),' - ',
                                                    CAST(REPLACE(tanggal,'-','/') AS CHAR(15))) as Modis 
                                                    from temp_jadwal_pjr 
                                                    where nik = '" & nik & "' 
                                                    AND tanggal = '" & tanggal & "'  
                                                    AND RECID = 'P' AND KODE_MODIS = '" & rak.Split("-")(0).Trim & "' AND NORAK = '" & rak.Split("-")(2).Trim & "'"
                Console.WriteLine(Madp.SelectCommand.CommandText)
                Madp.Fill(Rtn)
            End If
        Catch ex As Exception
            TraceLog("Error WDCP_GetNamaRak : " & ex.Message & ex.StackTrace)
            Rtn = New DataTable
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function GetTanggalAkumulasiLBTD_BAPJR(ByRef tanggal As String, ByRef jumlah As String) As String
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New Boolean

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            Mcom.CommandText = "  SELECT CONCAT(CAST(MIN(TANGGAL) AS CHAR),' / ',CAST(MAX(TANGGAL)AS CHAR))   
                                                    FROM ITEMSO_PJR_BA_AS 
                                                    WHERE RECID = ''"
            Console.WriteLine(Mcom.CommandText)
            tanggal = Mcom.ExecuteScalar

            Mcom.CommandText = "  SELECT COUNT(1)   
                                                    FROM ITEMSO_PJR_BA_AS 
                                                    WHERE RECID = ''"
            Console.WriteLine(Mcom.CommandText)
            jumlah = Mcom.ExecuteScalar



        Catch ex As Exception
            TraceLog("Error WDCP_CekModis : " & ex.Message & ex.StackTrace)

            Rtn = Nothing
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function



    Public Function GetPersonil(Optional ByVal nik As String = "") As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable
        Dim Mcom As New MySqlCommand("", Conn)
        Dim dt As New DataTable
        Dim jabatan As String = ""
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            jabatan = getJabatanVirbacaprod()

            If FormMain.isPengganti = False Then

                Mcom.CommandText = "SET @@lc_time_names = 'id_ID';"
                Mcom.ExecuteNonQuery()

                Madp.SelectCommand.CommandText = "SELECT DISTINCT nik FROM SOPPAGENT.ABSSETTINGSHIFT a LEFT JOIN `soppagent`.`abspegawaimst` b ON a.nik = b.menoin
					                            WHERE b.jabatan IN (" & jabatan & ")
                                              
                                              AND DAYNAME(a.tanggal) <> 'Minggu' 
                                              AND pinjaman = 0" '18/1/2023, hasil simulasi toko, exclude karyaman pinjaman

                TraceLog("WDCP_GetPersonil : " & Madp.SelectCommand.CommandText)
                Madp.Fill(Rtn)
                Mcom.CommandText = "SET @@lc_time_names = 'en_US';"
                Mcom.ExecuteNonQuery()
            Else

                'AMBIL DATA SHIFT DALAM MINGGU TERKAIT

                Madp.SelectCommand.CommandText = "SELECT a.NIK,a.TANGGAL FROM SOPPAGENT.ABSSETTINGSHIFT a  LEFT JOIN `soppagent`.`abspegawaimst` b ON a.nik = b.menoin
					                            WHERE b.jabatan IN (" & jabatan & ") AND 
                                               TANGGAL >= DATE_ADD(CAST(CURDATE() AS DATE), INTERVAL(2-DAYOFWEEK(CAST(CURDATE() AS DATE ))) DAY) 
                                               AND
                                               TANGGAL <= DATE_ADD(CAST(CURDATE() AS DATE), INTERVAL(8-DAYOFWEEK(CAST(CURDATE() AS DATE ))) DAY) 
                                               AND pinjaman = 0 "
                dt.Clear()
                Madp.Fill(dt)
                Dim menoin As String = ""
                Dim tanggal As Date
                Dim tanggal2 As String = ""
                Dim concat As String = ""
                Dim jum1 As String = ""
                Dim jum2 As String = ""

                'looping utk cek absen
                For i As Integer = 0 To dt.Rows.Count - 1
                    menoin = dt.Rows(i)("nik").ToString
                    Console.WriteLine(menoin)
                    tanggal = dt.Rows(i)("tanggal")
                    Console.WriteLine(tanggal)
                    tanggal2 = tanggal.ToString("yyyy-MM-dd")
                    Console.WriteLine(tanggal2)
                    Console.WriteLine(Date.Now)
                    If tanggal2 <= Date.Now.ToString("yyyy-MM-dd") Then
                        Mcom.CommandText = "SELECT RECID FROM TEMP_JADWAL_PJR WHERE RECID = '1' AND NIK = '" & menoin & "' AND TANGGAL = '" & tanggal2 & "' 
                                            AND TANGGAL >= DATE_ADD(CAST(CURDATE() AS DATE), INTERVAL(2-DAYOFWEEK(CAST(CURDATE() AS DATE ))) DAY) 
                                            AND TANGGAL <= DATE_ADD(CAST(CURDATE() AS DATE), INTERVAL(8-DAYOFWEEK(CAST(CURDATE() AS DATE ))) DAY)                                             
                                            AND TANGGAL <= CURDATE() "
                        If Mcom.ExecuteScalar = "1" Then
                            GoTo lanjut
                        End If

                        Mcom.CommandText = "SET @@lc_time_names = 'id_ID';"
                        Mcom.ExecuteNonQuery()

                        Mcom.CommandText = "SELECT HARI FROM JADWAL_PENANGGUNGJAWABRAK WHERE NIK = '" & menoin & "' AND HARI = DAYNAME('" & tanggal2 & "')"
                        Console.WriteLine(Mcom.CommandText)
                        If Mcom.ExecuteScalar = "" Then
                            GoTo lanjut
                        End If

                        Mcom.CommandText = "SELECT HARI FROM TEMP_JADWAL_PJR WHERE RECID = '1' AND NIK = '" & menoin & "' AND TANGGAL = '" & tanggal2 & "'
                                            AND TANGGAL >= DATE_ADD(CAST(CURDATE() AS DATE), INTERVAL(2-DAYOFWEEK(CAST(CURDATE() AS DATE ))) DAY) 
                                            AND TANGGAL <= DATE_ADD(CAST(CURDATE() AS DATE), INTERVAL(8-DAYOFWEEK(CAST(CURDATE() AS DATE ))) DAY)                                             
                                            AND TANGGAL <= CURDATE() AND (nik,kode_modis,norak) NOT IN
                                            (SELECT nik,kode_modis,norak FROM temp_jadwal_penanggungjawabrak_pengganti)"
                        If Mcom.ExecuteScalar <> "" Then
                            GoTo lanjut
                        End If

                        'jika tidak ada / tidak ketemu
                        If Mcom.ExecuteScalar = "" Then

                            concat &= "'" & menoin & "',"


                        End If
                    End If
lanjut:
                Next
                If concat.Length <> 0 Then
                    concat = concat.Substring(0, concat.Length - 1)
                Else
                    concat = "''"
                End If
                Mcom.CommandText = "SET @@lc_time_names = 'en_US';"
                Mcom.ExecuteNonQuery()

                Madp.SelectCommand.CommandText = "SELECT DISTINCT NIK FROM SOPPAGENT.ABSSETTINGSHIFT WHERE nik IN(" & concat & ")"
                Console.WriteLine(Madp.SelectCommand.CommandText)
                Madp.Fill(Rtn)
            End If

        Catch ex As Exception
            TraceLog("Error WDCP GetPersonil_PJR: " & ex.Message & ex.StackTrace)
            Mcom.CommandText = "SET @@lc_time_names = 'en_US';"
            Mcom.ExecuteNonQuery()
            Rtn = New DataTable
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function GetPersonilCetak(Optional ByVal nik As String = "") As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable
        Dim Mcom As New MySqlCommand("", Conn)
        Dim dt As New DataTable
        Dim jabatan As String = ""
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            jabatan = getJabatanVirbacaprod()

            Mcom.CommandText = "SET @@lc_time_names = 'id_ID';"
            Mcom.ExecuteNonQuery()
            Madp.SelectCommand.CommandText = "SELECT DISTINCT nik FROM SOPPAGENT.ABSSETTINGSHIFT a LEFT JOIN `soppagent`.`abspegawaimst` b ON a.nik = b.menoin
					                            WHERE b.jabatan IN (" & jabatan & ")
                                              AND CONCAT(a.nik,'-',CAST(DAYNAME(a.tanggal) AS CHAR) ) IN (SELECT CONCAT(NIK,'-',hari) FROM temp_jadwal_penanggungjawabrak)
                                                AND pinjaman = 0" 'uat

            TraceLog("WDCP_GetPersonilJadwal : " & Madp.SelectCommand.CommandText)
            Madp.Fill(Rtn)
            Mcom.CommandText = "SET @@lc_time_names = 'en_US';"
            Mcom.ExecuteNonQuery()


        Catch ex As Exception
            TraceLog("Error WDCP GetPersonil_PJR: " & ex.Message & ex.StackTrace)

            Rtn = New DataTable
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function



    Public Function createTabelJadwal() As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable
        Dim Mcom As New MySqlCommand("", Conn)
        Dim dt As New DataTable
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            'tabel jadwal pjr
            Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`jadwal_penanggungjawabrak` (
                                `NIK` VARCHAR(12), 
                                `NAMA` VARCHAR(99), 
                                `JABATAN` VARCHAR(50), 
                                `HARI` VARCHAR(30),
                                `Kode_Modis` VARCHAR(20),
                                `MODIS` VARCHAR(99),
                                `Norak` VARCHAR(10),
                                `Shelfing` VARCHAR(10),
                                `Addtime` DATE,
                                `Totalitem` Varchar(10),
                                `TotalEstimasi` Varchar(10),
                                `StatusApproval` Varchar(2),
                                Primary Key(NIK,HARI, Kode_Modis,Norak)
                                ) ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci;"
            Mcom.ExecuteNonQuery()
            'tabel temporary jadwal pjr
            Mcom.CommandText = "DROP TABLE IF EXISTS `pos`.`temp_jadwal_penanggungjawabrak`"
            Mcom.ExecuteNonQuery()

            Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`temp_jadwal_penanggungjawabrak` (
                                `NIK` VARCHAR(12), 
                                `NAMA` VARCHAR(99), 
                                `JABATAN` VARCHAR(50), 
                                `HARI` VARCHAR(30),
                                `Kode_Modis` VARCHAR(20),
                                `MODIS` VARCHAR(99),
                                `Norak` VARCHAR(10),
                                `Shelfing` VARCHAR(10),
                                `Addtime` DATE,
                                `Totalitem` Varchar(10),
                                `TotalEstimasi` Varchar(10),
                                Primary Key(NIK,HARI, Kode_Modis,Norak)
                                ) ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci;"
            Mcom.ExecuteNonQuery()

            'tabel tampungan untuk pjr
            Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`temp_jadwal_pjr` (
                                `RECID` VARCHAR(2) DEFAULT '', 
                                `NIK` VARCHAR(12), 
                                `NAMA` VARCHAR(99), 
                                `JABATAN` VARCHAR(50), 
                                `HARI` VARCHAR(30),
                                `TANGGAL` DATE,
                                `Kode_Modis` VARCHAR(20),
                                `MODIS` VARCHAR(99),
                                `Shelfing` VARCHAR(10),
                                `norak` VARCHAR(10),

                                `Addtime` DATE,
                                `Totalitem` Varchar(10),
                                `TotalEstimasi` Varchar(10),
                                `ITT` Varchar(10),
                                `FisikAda` Varchar(10),
                                `FisikTidakAda` Varchar(10),
                                `KetMinggu` Varchar(10),
                                `StatusApproval` Varchar(2)
                                ) ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci;"
            Mcom.ExecuteNonQuery()
            Try
                Mcom.CommandText = "ALTER TABLE jadwal_penanggungjawabrak DROP PRIMARY KEY,ADD PRIMARY KEY(NIK,HARI, Kode_Modis,Norak)"
                Mcom.ExecuteNonQuery()

            Catch ex As Exception

            End Try
            Try
                Mcom.CommandText = "ALTER TABLE TEMP_jadwal_penanggungjawabrak DROP PRIMARY KEY,ADD PRIMARY KEY(NIK,HARI, Kode_Modis,Norak)"
                Mcom.ExecuteNonQuery()
            Catch ex As Exception

            End Try

            Mcom.CommandText = "INSERT IGNORE INTO temp_jadwal_penanggungjawabrak SELECT  `NIK`,`NAMA` , `JABATAN` , `HARI`,`Kode_Modis` ,`MODIS`,`Norak` ,`Shelfing` ,`Addtime` ,`Totalitem` ,`TotalEstimasi` FROM jadwal_penanggungjawabrak"
            Mcom.ExecuteNonQuery()

            'Madp.SelectCommand.CommandText = "select distinct nama_rak,ket_rak from rak where nama_Rak NOT IN (SELECT KODE_MODIS FROM temp_jadwal_penanggungjawabrak) and nama_rak <> '' AND KET_RAK <> '';"
            ''Console.Writeline(Madp.SelectCommand.CommandText)
            'Rtn.Clear()
            'Madp.Fill(Rtn)

            'Mcom.CommandText = "INSERT IGNORE INTO `temp_jadwal_penanggungjawabrak` SELECT '','','','', kodemodis,ket_rak,'','',NULL,'0','0' FROM RAK where (kodemodis,norak) NOT IN (SELECT KODE_MODIS,norak FROM temp_jadwal_penanggungjawabrak) and nama_rak <> '' AND KET_RAK <> '' group by nama_rak;"

            'Memo 447/cps/23
            'PJR hnya FJP=Y

            Mcom.CommandText = "INSERT IGNORE INTO `temp_jadwal_penanggungjawabrak` SELECT '','','','', a.kodemodis,a.ket_rak,b.no_rak,'',NULL,'0','0' FROM rak a LEFT JOIN bracket b ON a.kodemodis = b. modisp
WHERE (b.modisp,b.no_rak) NOT IN (SELECT KODE_MODIS,norak 
FROM temp_jadwal_penanggungjawabrak) AND a.nama_rak <> '' AND a.KET_RAK <> '' AND a.flagprod LIKE '%FJP=Y%' GROUP BY b.modisp,b.no_rak"

            '            Mcom.CommandText = "INSERT IGNORE INTO `temp_jadwal_penanggungjawabrak` SELECT '','','','', a.kodemodis,a.ket_rak,b.no_rak,'',NULL,'0','0' FROM rak a LEFT JOIN bracket b ON a.kodemodis = b. modisp
            'WHERE (b.modisp,b.no_rak) NOT IN (SELECT KODE_MODIS,norak 
            'FROM temp_jadwal_penanggungjawabrak) AND a.nama_rak <> '' AND a.KET_RAK <> '' AND a.flagprod NOT LIKE '%FJP=N%' GROUP BY b.modisp,b.no_rak"


            Mcom.ExecuteNonQuery()

            'Memo 447/cps/23
            'PJR hnya FJP=Y
            Mcom.CommandText = "INSERT IGNORE INTO temp_jadwal_penanggungjawabrak 
                                SELECT '','','','', kodemodis,ket_rak,'1','',NULL,'0','0' FROM rak 
                                WHERE kodemodis NOT IN (SELECT kode_modis FROM temp_jadwal_penanggungjawabrak)
                                AND flagprod  LIKE '%FJP=Y%' GROUP BY KODEMODIS
                                "
            'Mcom.CommandText = "INSERT IGNORE INTO temp_jadwal_penanggungjawabrak 
            '                    SELECT '','','','', kodemodis,ket_rak,'1','',NULL,'0','0' FROM rak 
            '                    WHERE kodemodis NOT IN (SELECT kode_modis FROM temp_jadwal_penanggungjawabrak)
            '                    AND flagprod NOT LIKE '%FJP=N%' GROUP BY KODEMODIS
            '                    "
            Mcom.ExecuteNonQuery()

            'For i As Integer = 0 To Rtn.Rows.Count - 1

            '    Mcom.CommandText = "INSERT IGNORE INTO `temp_jadwal_penanggungjawabrak` VALUES ('', '',''
            '                    ,'','" & Rtn.Rows(i)("nama_rak") & "','" & Rtn.Rows(i)("ket_rak") & "',''
            '                    ,'',NULL,'0','0')"
            '    Mcom.ExecuteNonQuery()
            'Next




        Catch ex As Exception
            TraceLog("Error WDCP GetPersonil_PJR: " & ex.Message & ex.StackTrace)

            Rtn = New DataTable
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function CekPersonil(ByVal NIK As String, ByRef NamaPersonil As String) As String
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)


        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If


            Mcom.CommandText = "SELECT DISTINCT MENAME FROM `soppagent`.`abspegawaimst` where MENOIN = '" & NIK & "' AND pinjaman = 0;"
            NamaPersonil = Mcom.ExecuteScalar
            If NamaPersonil = "" Or IsDBNull(NamaPersonil) Then
                NamaPersonil = ""
            End If

        Catch ex As Exception
            TraceLog("Error WDCP CekPersonil_PJR: " & ex.Message & ex.StackTrace)
            NamaPersonil = ""
        Finally
            Conn.Close()
        End Try
        Return NamaPersonil
    End Function

    Public Function AmbilKodeModis(ByVal NIK As String, ByVal modis As String) As String
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim hasil As String = ""

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If


            Mcom.CommandText = "SELECT DISTINCT kode_modis FROM temp_jadwal_penanggungjawabrak where nik = '" & NIK & "' and modis = '" & modis & "';"
            hasil = Mcom.ExecuteScalar

        Catch ex As Exception
            TraceLog("Error WDCP ambilKodeModis: " & ex.Message & ex.StackTrace)
            hasil = ""
        Finally
            Conn.Close()
        End Try
        Return hasil
    End Function

    Public Function CekModis(ByVal Modis As String, ByVal tanggal As String, ByVal nik As String, ByRef CountModis As Integer, ByRef NamaModis As String, Optional ByVal isScan As Boolean = False) As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            'Mcom.CommandText = "select count(*) from rak r 
            '                    where nama_rak = '" & Modis & "' "
            Mcom.CommandText = "select count(*) from rak r 
                                where KODEMODIS = '" & Modis & "' "
            'And noshelf Not in (select SHELFING from temp_jadwal_pjr where KODE_MODIS = '" & Modis & "')"
            Mcom.CommandText &= "AND kodemodis IN(SELECT KODE_MODIS  FROM TEMP_JADWAL_PJR "

            If isScan = True Then
                Mcom.CommandText &= " WHERE KODE_MODIS = '" & Modis & "' AND TANGGAL = '" & tanggal & "' AND NIK = '" & nik & "') "
            Else
                Mcom.CommandText &= " WHERE KODE_MODIS = '" & Modis & "' AND TANGGAL = '" & tanggal & "' AND NIK = '" & nik & "' AND RECID = 'P') "

            End If

            Mcom.CommandText &= " order by noshelf asc;"
            Console.WriteLine(Mcom.CommandText)
            CountModis = Mcom.ExecuteScalar

            If CountModis = 0 Then
                Return Rtn
                Exit Function
            End If

            Mcom.CommandText = "select KET_RAK from rak where kodemodis = '" & Modis & "';"
            NamaModis = Mcom.ExecuteScalar

            'Revisi (15 April 2019)
            'Email: RE: Permasalahan  ITT Item ISMOD (Andry)
            'Khusus modis / rak Ecommerce tidak perlu dimasukan kedalam list yang harus dicek
            'Madp.SelectCommand.CommandText = "select distinct r.noshelf from rak r, prodmast p
            '                                    where r.prdcd = p.prdcd
            '                                    and r.nama_rak = '" & Modis & "'
            '                                   "
            Madp.SelectCommand.CommandText = "SELECT NORAK FROM TEMP_JADWAL_PJR"

            If isScan = True Then
                Madp.SelectCommand.CommandText &= " WHERE KODE_MODIS = '" & Modis & "' AND tanggal = '" & tanggal & "' AND NIK = '" & nik & "' and recid = '' "

                Console.WriteLine(Madp.SelectCommand.CommandText)
            Else
                Madp.SelectCommand.CommandText &= " WHERE KODE_MODIS = '" & Modis & "' AND tanggal = '" & tanggal & "' AND NIK = '" & nik & "' And RECID ='P'"

            End If
            Madp.SelectCommand.CommandText &= " order by abs(norak)"
            ''Console.Writeline(Madp.SelectCommand.CommandText)
            Madp.Fill(Rtn)
        Catch ex As Exception
            CountModis = 0
            TraceLog("Error WDCP_CekModis : " & ex.Message & ex.StackTrace)

            Rtn = Nothing
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function CekListBAPJR() As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New Boolean

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            Mcom.CommandText = "SELECT COUNT(1) FROM ITEMSO_PJR_BA_AS WHERE RECID = ''"
            If Mcom.ExecuteScalar <> 0 Then
                Rtn = True
            Else
                Rtn = False

            End If

        Catch ex As Exception
            TraceLog("Error WDCP_CekModis : " & ex.Message & ex.StackTrace)

            Rtn = False
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function AmbilNoRak(ByVal Modis As String, ByVal tanggal As String, ByVal norak As String) As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            Madp.SelectCommand.CommandText = "SELECT SHELFING FROM TEMP_JADWAL_PJR WHERE KODE_MODIS = '" & Modis & "' AND tanggal = '" & tanggal & "' AND NORAK = '" & norak & "' ORDER BY ABS(norak) "
            Madp.Fill(Rtn)


        Catch ex As Exception
            TraceLog("Error WDCP_CekModis : " & ex.Message & ex.StackTrace)

            Rtn = Nothing
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function AmbilNoRak_Pengganti(ByVal Modis As String, ByVal tanggal As String, ByVal norak As String) As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            Madp.SelectCommand.CommandText = "SELECT SHELFING FROM TEMP_JADWAL_PJR WHERE KODE_MODIS = '" & Modis & "' AND tanggal = '" & tanggal & "' AND NORAK = '" & norak & "' ORDER BY ABS(norak) "
            Madp.Fill(Rtn)


        Catch ex As Exception
            TraceLog("Error WDCP_CekModis : " & ex.Message & ex.StackTrace)

            Rtn = Nothing
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function


    Public Function CekModis(ByVal Modis As String, ByRef NamaModis As String, ByRef NomorRak As String, Optional ByVal hari As String = "", Optional ByVal nik As String = "") As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable
        Dim cekJumlah As String = ""
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            'Mcom.CommandText = "select KET_RAK from rak where nama_rak = '" & Modis & "';"
            'NamaModis = Mcom.ExecuteScalar
            Mcom.CommandText = "select KET_RAK from rak where kodemodis = '" & Modis & "';"
            Console.WriteLine(Mcom.CommandText)
            NamaModis = Mcom.ExecuteScalar

            'Mcom.CommandText = "select NORAK from rak where nama_rak = '" & Modis & "';"
            'NomorRak = Mcom.ExecuteScalar
            'Mcom.CommandText = "select distinct NO_RAK from bracket where modisp = '" & Modis & "';"
            'NomorRak = Mcom.ExecuteScalar

            If FormMain.isPengganti = False Then
                '   Madp.SelectCommand.CommandText = "select distinct r.noshelf from rak r, prodmast p
                '                                   where r.prdcd = p.prdcd
                '                                   and r.nama_rak = '" & Modis & "'

                'order by r.noshelf asc;"
                'Madp.SelectCommand.CommandText = "select distinct NO_RAK from bracket where modisp = '" & Modis & "' ;"

                cekJumlah = cekPerbandinganPersonilVSModis(FormMain.cbHariBukaToko)
                If cekJumlah = True Then
                    Madp.SelectCommand.CommandText = "SELECT DISTINCT norak FROM bracket a 
                                                  LEFT JOIN temp_jadwal_penanggungjawabrak b ON a.modisp = b.kode_modis 
                                                  WHERE modisp = '" & Modis & "' order by abs(norak)"
                Else
                    Madp.SelectCommand.CommandText = "SELECT DISTINCT norak FROM bracket a 
                                                  LEFT JOIN temp_jadwal_penanggungjawabrak b ON a.modisp = b.kode_modis 
                                                  WHERE modisp = '" & Modis & "' AND shelfing = '' order by abs(norak)"

                End If


            Else


                'Madp.SelectCommand.CommandText = "SELECT DISTINCT norak FROM bracket a 
                '                                  LEFT JOIN jadwal_penanggungjawabrak b ON a.modisp = b.kode_modis 
                '                                  WHERE modisp = '" & Modis & "' and hari = '" & hari & "' and nik = '" & nik & "'
                '                                  AND norak NOT IN (SELECT norak FROM temp_jadwal_penanggungjawabrak_pengganti WHERE kode_modis = '" & Modis & "')
                '                                  order by abs(norak)"
                Madp.SelectCommand.CommandText = "SELECT DISTINCT norak FROM jadwal_penanggungjawabrak
                                                  WHERE kode_modis = '" & Modis & "' and hari = '" & hari & "' and nik = '" & nik & "'
                                                  AND norak NOT IN (SELECT norak FROM temp_jadwal_penanggungjawabrak_pengganti WHERE kode_modis = '" & Modis & "')
                                                  order by abs(norak)"

            End If

            TraceLog(Madp.SelectCommand.CommandText)
            Madp.Fill(Rtn)
        Catch ex As Exception
            TraceLog("Error WDCP_CekModis : " & ex.Message & ex.StackTrace)

            Rtn = Nothing
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function


    Public Function ambilNoshelf(ByVal Modis As String, ByVal NomorRak As String, ByRef noshelf_awal As String, ByRef noshelf_akhir As String) As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As Boolean = False

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            Mcom.CommandText = "select distinct shelfing_awal from bracket where modisp = '" & Modis & "' and no_rak = '" & NomorRak & "' ;"
            Console.WriteLine(Mcom.CommandText)
            noshelf_awal = Mcom.ExecuteScalar

            Mcom.CommandText = "select distinct shelfing_akhir from bracket where modisp = '" & Modis & "' and no_rak = '" & NomorRak & "' ;"
            Console.WriteLine(Mcom.CommandText)

            noshelf_akhir = Mcom.ExecuteScalar

            'cek dri tabel rak jika tidak ketemu
            'permintaan OPR utk default norak = 1 18/10/22
            'Memo 10/22
            If noshelf_awal = "" Then
                Mcom.CommandText = "select min(noshelf) from rak where kodemodis = '" & Modis & "' ;"
                Console.WriteLine(Mcom.CommandText)
                noshelf_awal = Mcom.ExecuteScalar
            End If
            If noshelf_akhir = "" Then
                Mcom.CommandText = "select max(noshelf) from rak where kodemodis = '" & Modis & "' ;"
                Console.WriteLine(Mcom.CommandText)
                noshelf_akhir = Mcom.ExecuteScalar
            End If

            If noshelf_awal = "" Or noshelf_akhir = "" Then
                Rtn = True
            End If


        Catch ex As Exception
            TraceLog("Error WDCP_CekModis : " & ex.Message & ex.StackTrace)

        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function


    Public Sub getData(ByVal NIK As String)
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim result As New ClsPJR

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            Mcom.CommandText = "SELECT NAMA FROM TEMP_JADWAL_PJR WHERE NIK = '" & NIK & "'"
            result.NAMA = Mcom.ExecuteScalar
            Mcom.CommandText = "SELECT HARI FROM TEMP_JADWAL_PJR WHERE NIK = '" & NIK & "'"
            result.HARI = Mcom.ExecuteScalar

        Catch ex As Exception

        Finally
            Conn.Close()
        End Try
    End Sub

    Public Function GETJADWAL_BYNIK(Optional ByVal nik As String = "", Optional ByVal hari As String = "") As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable
        Dim Mcom As New MySqlCommand("", Conn)

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`temp_jadwal_pjr` (
                                `RECID` VARCHAR(2) DEFAULT '', 
                                `NIK` VARCHAR(12), 
                                `NAMA` VARCHAR(99), 
                                `JABATAN` VARCHAR(50), 
                                `HARI` VARCHAR(30),
                                `TANGGAL` DATE,
                                `Kode_Modis` VARCHAR(20),
                                `MODIS` VARCHAR(99),
                                `Shelfing` VARCHAR(10),
                                `Addtime` DATE,
                                `Totalitem` Varchar(10),
                                `TotalEstimasi` Varchar(10),
                                `ITT` Varchar(10),
                                `FisikAda` Varchar(10),
                                `FisikTidakAda` Varchar(10),
                                `KetMinggu` Varchar(10),
                                `StatusApproval` Varchar(2)
                                ) ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci;"
            Mcom.ExecuteNonQuery()

            'Mcom.CommandText = "Select COUNT(*) From Information_schema.Columns "
            'Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('temp_jadwal_pjr') "
            'Mcom.CommandText &= "And Column_Name='SOID' "
            ''Console.Writeline(Mcom.CommandText)
            'If Mcom.ExecuteScalar = 0 Then
            '    Mcom.CommandText = "ALTER TABLE `temp_jadwal_pjr` ADD COLUMN `SOID` VARCHAR(1) DEFAULT '' "
            '    Mcom.ExecuteNonQuery()
            'End If

            'Madp.SelectCommand.CommandText = "select hari,modis,shelfing from temp_jadwal_pjr where nik = '" & nik & "';"

            Madp.SelectCommand.CommandText = "select HARI,TANGGAL,MODIS,SHELFING from temp_jadwal_pjr where nik = '" & nik & "' 
                                               "

            Madp.Fill(Rtn)

        Catch ex As Exception

            Rtn = New DataTable
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function GETJADWAL(Optional ByVal nik As String = "", Optional ByVal hari As String = "") As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable
        Dim Mcom As New MySqlCommand("", Conn)
        Dim jabatan As String = ""
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            'Alur -> tabel jadwal_penanggungjawabrak dibackup ke tabel temp_jadwal_penanggungjawabrak untuk keperluan tambah atau update

            'tabel jadwal pjr
            Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`jadwal_penanggungjawabrak` (
                                `NIK` VARCHAR(12), 
                                `NAMA` VARCHAR(99), 
                                `JABATAN` VARCHAR(50), 
                                `HARI` VARCHAR(30),
                                `Kode_Modis` VARCHAR(20),
                                `MODIS` VARCHAR(99),
                                `Norak` VARCHAR(10),
                                `Shelfing` VARCHAR(10),
                                `Addtime` DATE,
                                `Totalitem` Varchar(10),
                                `TotalEstimasi` Varchar(10),
                                `StatusApproval` Varchar(2),
                                Primary Key(NIK,HARI, Kode_Modis)
                                ) ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci;"
            Mcom.ExecuteNonQuery()
            'tabel temporary jadwal pjr
            Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`temp_jadwal_penanggungjawabrak` (
                                `NIK` VARCHAR(12), 
                                `NAMA` VARCHAR(99), 
                                `JABATAN` VARCHAR(50), 
                                `HARI` VARCHAR(30),
                                `Kode_Modis` VARCHAR(20),
                                `MODIS` VARCHAR(99),
                                `Norak` VARCHAR(10),
                                `Shelfing` VARCHAR(10),
                                `Addtime` DATE,
                                `Totalitem` Varchar(10),
                                `TotalEstimasi` Varchar(10),
                                Primary Key(NIK,HARI, Kode_Modis)
                                ) ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci;"
            Mcom.ExecuteNonQuery()

            'tabel tampungan untuk pjr
            Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`temp_jadwal_pjr` (
                                `RECID` VARCHAR(2) DEFAULT '', 
                                `NIK` VARCHAR(12), 
                                `NAMA` VARCHAR(99), 
                                `JABATAN` VARCHAR(50), 
                                `HARI` VARCHAR(30),
                                `TANGGAL` DATE,
                                `Kode_Modis` VARCHAR(20),
                                `MODIS` VARCHAR(99),
                                `Shelfing` VARCHAR(10),
                                `norak` VARCHAR(10),
                                `Addtime` DATE,
                                `Totalitem` Varchar(10),
                                `TotalEstimasi` Varchar(10),
                                `ITT` Varchar(10),
                                `FisikAda` Varchar(10),
                                `FisikTidakAda` Varchar(10),
                                `KetMinggu` Varchar(10),
                                `StatusApproval` Varchar(2)
                                ) ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci;"
            Mcom.ExecuteNonQuery()

            'tabel temporary jadwal pjr pengganti
            Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`temp_jadwal_penanggungjawabrak_pengganti` (
                                `NIK` VARCHAR(12), 
                                `NAMA` VARCHAR(99), 
                                `JABATAN` VARCHAR(50), 
                                `HARI` VARCHAR(30),
                                `Kode_Modis` VARCHAR(20),
                                `MODIS` VARCHAR(99),
                                `Norak` VARCHAR(10),
                                `Shelfing` VARCHAR(10),
                                `Addtime` DATE,
                                `Totalitem` Varchar(10),
                                `TotalEstimasi` Varchar(10),
                                `HariPengganti` Varchar(30),

                                Primary Key(NIK,HARI, Kode_Modis)
                                ) ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci;"
            Mcom.ExecuteNonQuery()

            Try
                Mcom.CommandText = "ALTER TABLE `pos`.`temp_jadwal_penanggungjawabrak_pengganti` CHANGE `Norak` `Norak` VARCHAR(10) CHARSET latin1 COLLATE latin1_swedish_ci NOT NULL, DROP PRIMARY KEY, ADD PRIMARY KEY (`NIK`, `HARI`, `Kode_Modis`, `Norak`)"
                Mcom.ExecuteNonQuery()
            Catch ex As Exception

            End Try



            If FormMain.isPengganti = False Then
                Madp.SelectCommand.CommandText = "SELECT JABATAN,NAMA,HARI,NIK,KODE_MODIS,MODIS as NAMA_MODIS,NORAK,SHELFING as SHELF FROM temp_jadwal_penanggungjawabrak ORDER BY FIELD(JABATAN,'CHIEF OF STORE (SS)','STORE SR. LEADER (SS)',
                                                    'STORE JR. LEADER (SS)','Store Crew Boy (Ss)','Store Crew Girl (Ss)','') ,NIK desc, FIELD(hari,'Senin','Selasa','Rabu','Kamis','Jumat','Sabtu',''), KODE_MODIS,NORAK  "

                Madp.Fill(Rtn)
            Else
                Mcom.CommandText = "DELETE FROM temp_jadwal_penanggungjawabrak_pengganti WHERE TRIM(SUBSTRING_INDEX(HARIPENGGANTI,',',-1)) < DATE_ADD(CAST(CURDATE() AS DATE), INTERVAL(2-DAYOFWEEK(CAST(CURDATE() AS DATE ))) DAY)  "
                Mcom.ExecuteNonQuery()

                Madp.SelectCommand.CommandText = "select HARI,HARIPENGGANTI,NIK,MODIS,NORAK,SHELFING from temp_jadwal_penanggungjawabrak_PENGGANTI"

                Madp.Fill(Rtn)
            End If


        Catch ex As Exception
            TraceLog("Error WDCP_GETJADWAL :" & ex.Message & ex.StackTrace)
            Rtn = New DataTable
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function


    Public Function tambahPersonilPJR(ByVal NIK As String, ByVal nama As String,
                                      ByVal hari As String, ByVal tanggal As String, ByVal modis As String,
                                      ByVal nama_modis As String, ByVal shelfing As String) As Boolean
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Scon)
        Dim Madp As New MySqlDataAdapter("", Scon)
        Dim DtCP As New DataTable
        Dim Rtn As New Boolean
        Dim jabatan As String
        Dim dt1 As New DataTable
        Dim timesecond As Double = 0.0
        Dim jumlahitem As Integer = 0
        Dim totalitem As Integer = 0
        Dim totalestimasi As Double = 0.0
        Dim minggu As String = ""
        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            'Mcom.CommandText = "SHOW TABLES LIKE 'temp_PJR'"
            Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`temp_jadwal_pjr_detail` (
                                `RECID` VARCHAR(2) DEFAULT '', 
                                `NIK` VARCHAR(12), 
                                `NAMA` VARCHAR(99), 
                                `JABATAN` VARCHAR(50), 
                                `HARI` VARCHAR(30),
                                `TANGGAL` DATE,
                                `Kode_Modis` VARCHAR(20),
                                `MODIS` VARCHAR(99),
                                `Shelfing` VARCHAR(10),
                                `Addtime` DATE,
                                `Totalitem` Varchar(10),
                                `TotalEstimasi` Varchar(10),
                                `ITT` Varchar(10),
                                `FisikAda` Varchar(10),
                                `FisikTidakAda` Varchar(10),
                                `KetMinggu` Varchar(10),
                                `StatusApproval` Varchar(2)
                                ) ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci;"
            Mcom.ExecuteNonQuery()


            Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`temp_jadwal_pjr_estimasi_Detail` ( 
                                `NIK` VARCHAR(12), 
                                `NAMA` VARCHAR(99), 
                                `JABATAN` VARCHAR(50),
                                `HARI` VARCHAR(10),
                                `TANGGAL` DATE,
                                `Kode_Modis` VARCHAR(20),
                                `MODIS` VARCHAR(99),
                                `Shelfing` VARCHAR(10),
                                `Addtime` DATE,
                                `Cat_cod` Varchar(8),
                                `Kemasan` Varchar(5),
                                `Totalitem` Varchar(5),
                                `Timesecond` Varchar(5)
                                ) ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci;"
            Mcom.ExecuteNonQuery()
            'Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`temp_jadwal_pjr_estimasi` ( 
            '                    `NIK` VARCHAR(12), 
            '                    `NAMA` VARCHAR(99), 
            '                    `JABATAN` VARCHAR(50),
            '                    `HARI` VARCHAR(10),
            '                    `TANGGAL` VARCHAR(15),
            '                    `Kode_Modis` VARCHAR(20),
            '                    `MODIS` VARCHAR(99),
            '                    `Shelfing` VARCHAR(10),
            '                    `Addtime` DATE,
            '                    `Totalitem` Varchar(5),
            '                    `TotalEstimasi` Varchar(5)
            '                    ) ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci;"
            'Mcom.ExecuteNonQuery()

            Mcom.CommandText = "SELECT jabatan FROM soppagent.abspegawaimst a LEFT JOIN soppagent.abssettingshift b ON a.menoin = b.nik WHERE NIK = '" & NIK & "'"
            jabatan = Mcom.ExecuteScalar
            Mcom.CommandText = "SELECT (WEEK(CURDATE()) - WEEK(DATE_FORMAT(CURDATE(),'%Y-%m-01'))) +1  "
            minggu = Mcom.ExecuteScalar
            If hari = "Minggu" Then
                minggu += -1
            End If
            Mcom.CommandText = "INSERT IGNORE INTO TEMP_JADWAL_PJR VALUES(
                                '','" & NIK & "', '" & nama & "','" & jabatan & "'
                                ,'" & hari & "','" & tanggal & "','" & modis & "','" & nama_modis & "'
                                ,'" & shelfing & "'
                                ,NOW(),'0','0','','','','" & minggu & "','W')"

            ''Console.Writeline(Mcom.CommandText)
            Mcom.ExecuteNonQuery()

            'estimasi waktu
            Madp.SelectCommand.CommandText = "SELECT a.cat_cod, kemasan From prodmast a 
                                LEFT Join rak b ON a.prdcd = b.prdcd Where NAMA_RAK = '" & modis & "' GROUP BY CAT_COD,KEMASAn ORDER BY CAT_COD "
            ''Console.Writeline(Madp.SelectCommand.CommandText)
            dt1.Clear()
            Madp.Fill(dt1)


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
                timesecond = Mcom.ExecuteScalar
                ''Console.Writeline(timesecond)
                'ambil jumlah item per rak, cat_cod dan kemasan
                Mcom.CommandText = "SELECT COUNT(a.cat_cod) FROM prodmast a
                                    LEFT JOIN rak b ON a.prdcd = b.prdcd 
                                    WHERE NAMA_RAK = '" & modis & "' AND CAT_COD LIKE '%" & dt1.Rows(i)("cat_cod") & "%' 
                                    AND KEMASAN LIKE '%" & dt1.Rows(i)("kemasan") & "%' ORDER BY CAT_COD "
                ''Console.Writeline(Mcom.CommandText)

                jumlahitem = Mcom.ExecuteScalar
                ''Console.Writeline(jumlahitem)
                totalitem += jumlahitem
                Mcom.CommandText = "INSERT IGNORE INTO TEMP_JADWAL_PJR_estimasi_detail VALUES(
                                '" & NIK & "', '" & nama & "','" & jabatan & "'
                                ,'" & hari & "','" & tanggal & "','" & modis & "','" & nama_modis & "'
                                ,'" & shelfing & "'
                                ,NOW(),'" & dt1.Rows(i)("cat_cod") & "','" & dt1.Rows(i)("kemasan") & "'
                                ," & jumlahitem & ", " & timesecond & ")"
                ''Console.Writeline(Mcom.CommandText)
                Mcom.ExecuteNonQuery()

            Next
            Mcom.CommandText = "SELECT CEILING(SUM(TOTALITEM*timesecond)/60)  FROM temp_jadwal_pjr_estimasi_DETAIL 
                                WHERE NIK = '" & NIK & "' AND TANGGAL = '" & tanggal & "' AND KODE_MODIS = '" & modis & "';"
            totalestimasi = Mcom.ExecuteScalar
            Mcom.CommandText = "UPDATE TEMP_JADWAL_PJR SET `Totalitem` = '" & totalitem & "' , TotalEstimasi = '" & totalestimasi & "'
                                WHERE NIK = '" & NIK & "' AND TANGGAL = '" & tanggal & "' AND KODE_MODIS = '" & modis & "'; "
            'Mcom.CommandText = "INSERT IGNORE INTO TEMP_JADWAL_PJR_estimasi VALUES(
            '                    '" & NIK & "', '" & nama & "','" & jabatan & "'
            '                    ,'" & hari & "','" & tanggal & "','" & modis & "','" & nama_modis & "'
            '                    ,'" & shelfing & "'
            '                    ,NOW(),'" & totalitem & "','" & totalestimasi & "')"
            ''Console.Writeline(Mcom.CommandText)
            Mcom.ExecuteNonQuery()
            MsgBox("Data telah tersimpan!")

            Rtn = True

        Catch ex As Exception
            Rtn = False
            TraceLog("Error WDCP tambahPersonilPJR : " & ex.Message & ex.StackTrace)
        Finally
            Scon.Close()
        End Try
        Return Rtn
    End Function

    Public Function tambahPersonilPJR_Temp(ByVal NIK As String, ByVal nama As String,
                                      ByVal hari As String, ByVal modis As String,
                                      ByVal nama_modis As String, ByVal norak As String, ByVal shelfing As String, Optional ByVal hariPengganti As String = "") As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim DtCP As New DataTable
        Dim Rtn As New Boolean
        Dim jabatan As String
        Dim dt1 As New DataTable
        Dim timesecond As Double = 0.0
        Dim jumlahitem As Integer = 0
        Dim totalitem As Integer = 0
        Dim totalestimasi As Double = 0.0
        Dim minggu As String = ""
        Dim cekJumlah As String = ""
        Dim namamodis As String = ""
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            Mcom.CommandText = "SELECT jabatan FROM soppagent.abspegawaimst a LEFT JOIN soppagent.abssettingshift b ON a.menoin = b.nik WHERE NIK = '" & NIK & "'"
            jabatan = Mcom.ExecuteScalar
            If FormMain.isPengganti = False Then

                cekJumlah = cekPerbandinganPersonilVSModis(FormMain.cbHariBukaToko)
                If cekJumlah = True Then
                    Mcom.CommandText = "SELECT MODIS FROM temp_jadwal_penanggungjawabrak WHERE KODE_MODIS = '" & modis & "' and NORAK = '" & norak & "' AND shelfing = ''"
                    If Mcom.ExecuteScalar = "" Then
                        Mcom.CommandText = "SELECT MODIS FROM temp_jadwal_penanggungjawabrak WHERE KODE_MODIS = '" & modis & "'"
                        namamodis = Mcom.ExecuteScalar
                        Mcom.CommandText = "INSERT IGNORE INTO `pos`.`temp_jadwal_penanggungjawabrak` 
                                            (`NIK`, `NAMA`, `JABATAN`, `HARI`, `Kode_Modis`, `MODIS`, `Norak`, `Shelfing`, `Addtime`, `Totalitem`, `TotalEstimasi`) 
                                            VALUES ('" & NIK & "', '" & nama & "', '" & jabatan & "', '" & hari & "', '" & modis & "', '" & namamodis & "', '" & norak & "', '" & shelfing & "', NOW(), '0', '0') "

                    Else
                        Mcom.CommandText = "UPDATE `temp_jadwal_penanggungjawabrak` SET NIK = '" & NIK & "', NAMA = '" & nama & "',HARI = '" & hari & "', JABATAN  = '" & jabatan & "', NORAK = '" & norak & "', SHELFING = '" & shelfing & "', Addtime = NOW()
                                     WHERE KODE_MODIS = '" & modis & "' and NORAK = '" & norak & "'"
                    End If

                Else
                    Mcom.CommandText = "UPDATE `temp_jadwal_penanggungjawabrak` SET NIK = '" & NIK & "', NAMA = '" & nama & "',HARI = '" & hari & "', JABATAN  = '" & jabatan & "', NORAK = '" & norak & "', SHELFING = '" & shelfing & "', Addtime = NOW()
                                     WHERE KODE_MODIS = '" & modis & "' and NORAK = '" & norak & "'"
                End If


                Mcom.ExecuteNonQuery()
                Rtn = True

                'Mcom.CommandText = "INSERT IGNORE INTO `temp_jadwal_penanggungjawabrak` VALUES ('" & NIK & "', '" & nama & "','" & jabatan & "'
                '                ,'" & hari & "','" & modis & "','" & nama_modis & "','" & norak & "'
                '                ,'" & shelfing & "',NOW(),'0','0')"
                'Mcom.ExecuteNonQuery()
                'Rtn = True
            Else
                'tanggalPengganti = hariPengganti.Split(",")(1).Trim
                Mcom.CommandText = "SELECT Totalitem FROM jadwal_penanggungjawabrak WHERE NIK = '" & NIK & "' AND HARI = '" & hari & "' AND KODE_MODIS = '" & modis & "' AND NORAK = '" & norak & "'"
                Console.WriteLine(Mcom.CommandText)

                totalitem = Mcom.ExecuteScalar
                Mcom.CommandText = "SELECT totalestimasi FROM jadwal_penanggungjawabrak WHERE NIK = '" & NIK & "' AND HARI = '" & hari & "' AND KODE_MODIS = '" & modis & "' AND NORAK = '" & norak & "'"
                Console.WriteLine(Mcom.CommandText)

                totalestimasi = Mcom.ExecuteScalar
                'Mcom.CommandText = "UPDATE temp_jadwal_penanggungjawabrak_pengganti SET HARIPENGGANTI = '" & hariPengganti & "' WHERE NIK = '" & NIK & "' AND HARI = '" & hari & "' AND KODE_MODIS = '" & modis & "'"
                Mcom.CommandText = "INSERT IGNORE INTO `temp_jadwal_penanggungjawabrak_pengganti` VALUES ('" & NIK & "', '" & nama & "','" & jabatan & "'
                                ,'" & hari & "','" & modis & "','" & nama_modis & "','" & norak & "'
                                ,'" & shelfing & "',NOW(),'" & totalitem & "','" & totalestimasi & "','" & hariPengganti & "')"
                TraceLog(Mcom.CommandText)

                Mcom.ExecuteNonQuery()
                Rtn = True
            End If


        Catch ex As Exception
            Rtn = False
            TraceLog("Error WDCP tambahPersonilPJR : " & ex.Message & ex.StackTrace)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function


    Public Function HapusPersonilPJR_Temp(ByVal NIK As String,
                                      ByVal hari As String, ByVal modis As String,
                                      ByVal norak As String) As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim DtCP As New DataTable
        Dim Rtn As New Boolean

        Dim minggu As String = ""

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Mcom.CommandText = "SELECT COUNT(1) FROM RAK WHERE KODEMODIS = '" & modis & "'"
            If Mcom.ExecuteScalar > 0 Then
                'Mcom.CommandText = "DROP TABLE IF EXISTS temp_hapus_jadwal_pjr"
                'Mcom.ExecuteNonQuery()

                Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`temp_hapus_jadwal_pjr` ( 
                                `NIK` VARCHAR(12), 
                                `HARI` VARCHAR(10),
                                `Kode_Modis` VARCHAR(99),
                                `NORAK` VARCHAR(20),
                                `ADDTIME` DATETIME

                                ) ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci;"
                Mcom.ExecuteNonQuery()
                Mcom.CommandText = "INSERT IGNORE INTO temp_hapus_jadwal_pjr VALUES ('" & NIK & "', '" & hari & "', '" & modis & "', '" & norak & "',NOW())"
                Mcom.ExecuteNonQuery()

                Mcom.CommandText = "update  temp_jadwal_penanggungjawabrak set nik= '',nama= '',jabatan= '',hari= '',
                                    SHELFING = '', addtime= NULL,totalitem= 0,totalestimasi = 0  WHERE NIK = '" & NIK & "'  
                                    AND HARI = '" & hari & "' AND KODE_MODIS = '" & modis & "' AND NORAK = '" & norak & "' "
                Mcom.ExecuteNonQuery()
                'Mcom.CommandText = "update  jadwal_penanggungjawabrak set nik= '',nama= '',jabatan= '',hari= '',
                '                    SHELFING = '', addtime= NULL,totalitem= 0,totalestimasi = 0, statusapproval = ''  WHERE NIK = '" & NIK & "'  
                '                    AND HARI = '" & hari & "' AND KODE_MODIS = '" & modis & "' AND NORAK = '" & norak & "' "
                'Mcom.ExecuteNonQuery()
            Else
                Mcom.CommandText = "DELETE FROM temp_jadwal_penanggungjawabrak WHERE KODE_MODIS = '" & modis & "'"
                Mcom.ExecuteNonQuery()

            End If

            Rtn = True
        Catch ex As Exception
            Rtn = False
            TraceLog("Error WDCP hapusPersonilPJR : " & ex.Message & ex.StackTrace)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function


    Public Function insertTempJadwalPJR(Optional ByVal nik As String = "") As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim dt2 As New DataTable
        Dim Rtn As New Boolean
        Dim dt1 As New DataTable
        Dim hari As String = ""
        Dim minggu As String = ""
        Dim bulan As String = ""
        Dim tanggal As String = ""
        Dim tahun As String = ""
        Dim ketminggu As String = ""
        Dim tanggalPengganti As String
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            ReloadJadwal(True)

            hari = Date.Now.ToString("dddd", New Globalization.CultureInfo("id-ID"))
            tanggalPengganti = Date.Now.ToString("dddd, yyyy-MM-dd", New Globalization.CultureInfo("id-ID"))

            If hari.ToLower = "minggu" Then
                'belom didaftar

                Mcom.CommandText = "SELECT count(*) from temp_jadwal_penanggungjawabrak_pengganti where NIK = '" & nik & "' AND hariPengganti = '" & tanggalPengganti & "'"
                Console.WriteLine(Mcom.CommandText)

                If Mcom.ExecuteScalar = 0 Then
                    Madp.SelectCommand.CommandText = "SELECT nik,nama,jabatan,hari,kode_modis,modis,norak,shelfing,totalitem,totalestimasi FROM JADWAL_PENANGGUNGJAWABRAK WHERE NIK = '" & nik & "'
                                        AND HARI NOT IN(SELECT HARI FROM TEMP_JADWAL_PENANGGUNGJAWABRAK_PENGGANTI WHERE NIK = '" & nik & "')"
                    Console.WriteLine(Madp.SelectCommand.CommandText)

                    dt1.Clear()
                    Madp.Fill(dt1)
                    For i As Integer = 0 To dt1.Rows.Count - 1
                        Mcom.CommandText = "SELECT RECID FROM TEMP_JADWAL_PJR WHERE RECID = '1' AND HARI = '" & dt1.Rows(i)("hari") & "' AND NIK = '" & dt1.Rows(i)("nik") & "'"
                        Console.WriteLine(Mcom.CommandText)

                        If Mcom.ExecuteScalar = "" Then

                            'Mcom.CommandText = "UPDATE temp_jadwal_penanggungjawabrak_pengganti SET HARIPENGGANTI = '" & hariPengganti & "' WHERE NIK = '" & NIK & "' AND HARI = '" & hari & "' AND KODE_MODIS = '" & modis & "'"
                            Mcom.CommandText = "INSERT IGNORE INTO `temp_jadwal_penanggungjawabrak_pengganti` VALUES ('" & dt1.Rows(i)("nik") & "', '" & dt1.Rows(i)("nama") & "','" & dt1.Rows(i)("jabatan") & "'
                                ,'" & dt1.Rows(i)("hari") & "','" & dt1.Rows(i)("kode_modis") & "','" & dt1.Rows(i)("modis") & "','" & dt1.Rows(i)("norak") & "'
                                ,'" & dt1.Rows(i)("shelfing") & "',NOW(),'" & dt1.Rows(i)("totalitem") & "','" & dt1.Rows(i)("totalestimasi") & "','" & tanggalPengganti & "')"
                            Console.WriteLine(Mcom.CommandText)
                            Mcom.ExecuteNonQuery()
                        End If
                    Next
                    'ambil totalitem dimana jadwal belom didaftarkan
                    Mcom.CommandText = "SELECT (WEEK(CURDATE()) - WEEK(DATE_FORMAT(CURDATE(),'%Y-%m-01'))) +1 "
                    ketminggu = Mcom.ExecuteScalar
                    Mcom.CommandText = "INSERT IGNORE INTO TEMP_JADWAL_PJR SELECT '',nik,nama,jabatan,TRIM(SUBSTRING_INDEX(hariPengganti,',',1)) AS HARI,CURDATE() as tanggal ,kode_modis,modis,shelfing,norak,CURDATE(),Totalitem,TotalEstimasi,
                                        '','','','" & ketminggu & "','P' FROM temp_jadwal_penanggungjawabrak_pengganti 
                                        WHERE NIK = '" & nik & "' AND TRIM(SUBSTRING_INDEX(hariPengganti,',',-1)) = CURDATE() "
                    Console.WriteLine(Mcom.CommandText)
                    Mcom.ExecuteNonQuery()
                End If
            End If

            'cek di hari tsb apakah nik scanfinger ada jadwal
            If nik = "" Then
                Mcom.CommandText = "SELECT count(*) from jadwal_penanggungjawabrak where  hari = '" & hari & "'"
            Else
                Mcom.CommandText = "SELECT count(*) from jadwal_penanggungjawabrak where NIK = '" & nik & "' AND hari = '" & hari & "'"
            End If
            If Mcom.ExecuteScalar = 0 Then ' jika tdak ada
                Rtn = False
            Else
                If nik = "" Then
                    Mcom.CommandText = "SELECT COUNT(*) FROM TEMP_JADWAL_PJR WHERE  TANGGAL = CURDATE() and recid = ''" 'cek sudah pernah load atau belum
                Else
                    Mcom.CommandText = "SELECT COUNT(*) FROM TEMP_JADWAL_PJR WHERE NIK = '" & nik & "' AND TANGGAL = CURDATE() and recid = ''" 'cek sudah pernah load atau belum
                End If

                If Mcom.ExecuteScalar > 0 Then
                    Rtn = False
                Else
                    If hari.ToLower <> "minggu" Then
                        Mcom.CommandText = "SET @@lc_time_names = 'id_ID';"
                        Mcom.ExecuteNonQuery()
                        Mcom.CommandText = "SELECT (WEEK(CURDATE()) - WEEK(DATE_FORMAT(CURDATE(),'%Y-%m-01'))) +1 "
                        ketminggu = Mcom.ExecuteScalar
                        If nik = "" Then
                            Mcom.CommandText = "INSERT IGNORE INTO TEMP_JADWAL_PJR SELECT '',nik,nama,jabatan,hari,CURDATE() as tanggal ,kode_modis,modis,shelfing,norak,CURDATE(),Totalitem,TotalEstimasi,
                                        '','','','" & ketminggu & "',StatusApproval,'' FROM jadwal_penanggungjawabrak 
                                        WHERE  hari = DAYNAME(CURDATE()) AND STATUSAPPROVAL <> '' AND 
                                        (nik,hari,kode_modis,norak) NOT IN (SELECT nik,hari,kode_modis,norak FROM temp_jadwal_pjr WHERE recid <> ''  AND tanggal = CURDATE())  "
                            Mcom.ExecuteNonQuery()
                        Else
                            Mcom.CommandText = "INSERT IGNORE INTO TEMP_JADWAL_PJR SELECT '',nik,nama,jabatan,hari,CURDATE() as tanggal ,kode_modis,modis,shelfing,norak,CURDATE(),Totalitem,TotalEstimasi,
                                        '','','','" & ketminggu & "',StatusApproval,'' FROM jadwal_penanggungjawabrak 
                                        WHERE NIK = '" & nik & "' AND hari = DAYNAME(CURDATE()) and STATUSAPPROVAL <> '' AND 
                                        (nik,hari,kode_modis,norak) NOT IN (SELECT nik,hari,kode_modis,norak FROM temp_jadwal_pjr WHERE recid <> ''
                                        and NIK = '" & nik & "' AND tanggal = CURDATE())"
                            Mcom.ExecuteNonQuery()
                        End If


                        Mcom.CommandText = "SET @@lc_time_names = 'en_US';"
                        Mcom.ExecuteNonQuery()
                    End If

                End If
            End If
            If nik = "" Then
                Mcom.CommandText = "SELECT count(*) from temp_jadwal_penanggungjawabrak_pengganti where   hariPengganti = '" & tanggalPengganti & "'"
            Else
                Mcom.CommandText = "SELECT count(*) from temp_jadwal_penanggungjawabrak_pengganti where NIK = '" & nik & "' AND hariPengganti = '" & tanggalPengganti & "'"
            End If
            TraceLog(Mcom.CommandText)
            Console.WriteLine(Mcom.CommandText)
            If Mcom.ExecuteScalar = 0 Then ' jika tdak ada
                Rtn = False
            Else
                If nik = "" Then
                    Mcom.CommandText = "SELECT COUNT(*) FROM TEMP_JADWAL_PJR WHERE  (TANGGAL,kode_modis,norak) IN(SELECT TRIM(SUBSTRING_INDEX(hariPengganti,',',-1)),kode_modis,norak FROM temp_jadwal_penanggungjawabrak_pengganti WHERE   TRIM(SUBSTRING_INDEX(hariPengganti,',',-1)) = CURDATE())" 'cek sudah pernah load atau belum

                Else
                    Mcom.CommandText = "SELECT COUNT(*) FROM TEMP_JADWAL_PJR WHERE NIK = '" & nik & "' AND (TANGGAL,kode_modis,norak) IN(SELECT TRIM(SUBSTRING_INDEX(hariPengganti,',',-1)),kode_modis,norak FROM temp_jadwal_penanggungjawabrak_pengganti WHERE nik = '" & nik & "' AND TRIM(SUBSTRING_INDEX(hariPengganti,',',-1)) = CURDATE())" 'cek sudah pernah load atau belum

                End If
                TraceLog(Mcom.CommandText)
                Console.WriteLine(Mcom.CommandText)
                If Mcom.ExecuteScalar > 0 Then
                    Rtn = False

                Else
                    Mcom.CommandText = "SELECT (WEEK(CURDATE()) - WEEK(DATE_FORMAT(CURDATE(),'%Y-%m-01'))) +1 "
                    ketminggu = Mcom.ExecuteScalar
                    If nik = "" Then
                        Mcom.CommandText = "INSERT IGNORE INTO TEMP_JADWAL_PJR SELECT '',nik,nama,jabatan,TRIM(SUBSTRING_INDEX(hariPengganti,',',1)) AS HARI,CURDATE() as tanggal ,kode_modis,modis,shelfing,norak,CURDATE(),Totalitem,TotalEstimasi,
                                        '','','','" & ketminggu & "','P','' FROM temp_jadwal_penanggungjawabrak_pengganti 
                                        WHERE   TRIM(SUBSTRING_INDEX(hariPengganti,',',-1)) = CURDATE() "
                    Else
                        Mcom.CommandText = "INSERT IGNORE INTO TEMP_JADWAL_PJR SELECT '',nik,nama,jabatan,TRIM(SUBSTRING_INDEX(hariPengganti,',',1)) AS HARI,CURDATE() as tanggal ,kode_modis,modis,shelfing,norak,CURDATE(),Totalitem,TotalEstimasi,
                                        '','','','" & ketminggu & "','P','' FROM temp_jadwal_penanggungjawabrak_pengganti 
                                        WHERE NIK = '" & nik & "' AND TRIM(SUBSTRING_INDEX(hariPengganti,',',-1)) = CURDATE() "
                    End If

                    TraceLog(Mcom.CommandText)
                    Console.WriteLine(Mcom.CommandText)
                    Mcom.ExecuteNonQuery()


                End If
            End If
            Rtn = True

        Catch ex As Exception
            Rtn = False
            TraceLog("Error WDCP tambahPersonilPJR : " & ex.Message & ex.StackTrace)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function


    Public Function cariJadwal(ByVal tanggalawal As String) As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim DtCP As New DataTable
        Dim Rtn As New DataTable
        Dim sqltampung As String = ""

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            'Mcom.CommandText = "SELECT a.NIK, NAMA,JABATAN, 
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
            '                     >= (SELECT DATE_ADD(CAST('" & tanggalawal & "' AS DATE), INTERVAL(2-DAYOFWEEK(CAST('" & tanggalawal & "' AS DATE))) DAY))

            '                    AND CAST(CONCAT(SUBSTRING(TANGGAL,7,4),'-',SUBSTRING(TANGGAL,4,2),'-',SUBSTRING(TANGGAL,1,2)) AS DATE)
            '                     <= (SELECT DATE_ADD('" & tanggalawal & "', INTERVAL(8-DAYOFWEEK('" & tanggalawal & "')) DAY))"


            'Madp.SelectCommand.CommandText = Mcom.CommandText
            ''Console.Writeline(Madp.SelectCommand.CommandText)
            'Rtn.Clear()
            'Madp.Fill(Rtn)

            Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`temp_jadwal_pjr_estimasi_Detail` ( 
                                `NIK` VARCHAR(12), 
                                `NAMA` VARCHAR(99), 
                                `JABATAN` VARCHAR(50),
                                `HARI` VARCHAR(10),
                                `TANGGAL` DATE,
                                `Kode_Modis` VARCHAR(20),
                                `MODIS` VARCHAR(99),
                                `Shelfing` VARCHAR(10),
                                `Addtime` DATE,
                                `Cat_cod` Varchar(8),
                                `Kemasan` Varchar(5),
                                `Totalitem` Varchar(5),
                                `Timesecond` Varchar(5)
                                ) ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci;"
            Mcom.ExecuteNonQuery()





            'Madp.SelectCommand.CommandText = "SELECT DISTINCT NIK FROM TEMP_JADWAL_PJR"
            Madp.SelectCommand.CommandText = "SELECT  NIK,nama,hari,kode_modis FROM jadwal_penanggungjawabrak group by nik"
            Rtn.Clear()
            Madp.Fill(Rtn)

            For i As Integer = 1 To 7
                Mcom.CommandText = "DROP TABLE IF EXISTS temp_jadwal_pjr_estimasi_h" & i & ""
                Mcom.ExecuteNonQuery()

                Mcom.CommandText = "CREATE TABLE `pos`.`temp_jadwal_pjr_estimasi_h" & i & "` ( 
                                    `nik` VARCHAR(12), 
                                    `nama` VARCHAR(99), 
                                    `hari` VARCHAR(30), 
                                    `kode_modis` VARCHAR(99), 
                                    `h" & i & "` DOUBLE )"
                Mcom.ExecuteNonQuery()
            Next


            For j As Integer = 0 To Rtn.Rows.Count - 1

                For i As Integer = 1 To 7
                    Mcom.CommandText = "DROP TABLE IF EXISTS temp_jadwal_pjr_estimasi_" & j & "_" & i & ""
                    Mcom.ExecuteNonQuery()

                    Mcom.CommandText = "CREATE TABLE temp_jadwal_pjr_estimasi_" & j & "_" & i & " 
                                    SELECT IF(nik is NULL,'" & Rtn.Rows(j)("nik") & "','" & Rtn.Rows(j)("nik") & "') as nik,
                                            IF(nama is NULL,'" & Rtn.Rows(j)("nama") & "','" & Rtn.Rows(j)("nama") & "') as nama,
                                            IF(hari is NOT NULL,hari, 
                                            IF(" & i & "= 1,'Senin', 
                                            IF(" & i & "= 2,'Selasa',
                                            IF(" & i & "= 3,'Rabu', 
                                            IF(" & i & "= 4,'Kamis', 
                                            IF(" & i & "= 5,'Jumat', 
                                            IF(" & i & "= 6,'Sabtu',
                                            IF(" & i & "= 7,'Minggu',0))))))))

                                            as hari,


                                            IF(kode_modis is NULL,'" & Rtn.Rows(j)("kode_modis") & "','" & Rtn.Rows(j)("kode_modis") & "') as kode_modis,
                                            IF(SUM(totalestimasi) IS NULL,0,SUM(totalestimasi)) as h" & i & " FROM jadwal_penanggungjawabrak "
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
                    Mcom.CommandText &= " AND NIK = " & Rtn.Rows(j)("nik")
                    'AND 
                    '                                     (TANGGAL 
                    '                                     >= (SELECT DATE_ADD('" & tanggalawal & "', INTERVAL(1-DAYOFWEEK('" & tanggalawal & "')) DAY)) 
                    '                                      AND
                    '                                    TANGGAL 
                    '                                     <= (SELECT DATE_ADD('" & tanggalawal & "', INTERVAL(7-DAYOFWEEK('" & tanggalawal & "')) DAY))) 
                    '                                      "
                    '                    'If i = 7 Then
                    '    Mcom.CommandText &= " AND NIK = " & Rtn.Rows(j)("nik") & " AND 
                    '                (CAST(CONCAT(SUBSTRING(TANGGAL,7,4),'-',SUBSTRING(TANGGAL,4,2),'-',SUBSTRING(TANGGAL,1,2)) AS DATE) 
                    '                 >= (SELECT DATE_ADD('" & tanggalawal & "', INTERVAL(1-DAYOFWEEK('" & tanggalawal & "')-7) DAY)) 
                    '                  AND
                    '                CAST(CONCAT(SUBSTRING(TANGGAL,7,4),'-',SUBSTRING(TANGGAL,4,2),'-',SUBSTRING(TANGGAL,1,2)) AS DATE) 
                    '                 <= (SELECT DATE_ADD('" & tanggalawal & "', INTERVAL(7-DAYOFWEEK('" & tanggalawal & "')-7) DAY))) 
                    '                  "
                    '    ' (CAST(CONCAT(SUBSTRING(TANGGAL,7,4),'-',SUBSTRING(TANGGAL,4,2),'-',SUBSTRING(TANGGAL,1,2)) AS DATE) 
                    '    ' >= (SELECT DATE_ADD('" & tanggalawal & "', INTERVAL(2-DAYOFWEEK('" & tanggalawal & "')-7) DAY)) 
                    '    '  AND
                    '    'CAST(CONCAT(SUBSTRING(TANGGAL,7,4),'-',SUBSTRING(TANGGAL,4,2),'-',SUBSTRING(TANGGAL,1,2)) AS DATE) 
                    '    ' <= (SELECT DATE_ADD('" & tanggalawal & "', INTERVAL(8-DAYOFWEEK('" & tanggalawal & "')-7) DAY))) 
                    '    '  "

                    'Else
                    'Mcom.CommandText &= " AND NIK = " & Rtn.Rows(j)("nik") & " AND 
                    '                 (CAST(CONCAT(SUBSTRING(TANGGAL,7,4),'-',SUBSTRING(TANGGAL,4,2),'-',SUBSTRING(TANGGAL,1,2)) AS DATE) 
                    '                 >= (SELECT DATE_ADD('" & tanggalawal & "', INTERVAL(1-DAYOFWEEK('" & tanggalawal & "')) DAY)) 
                    '                  AND
                    '                CAST(CONCAT(SUBSTRING(TANGGAL,7,4),'-',SUBSTRING(TANGGAL,4,2),'-',SUBSTRING(TANGGAL,1,2)) AS DATE) 
                    '                 <= (SELECT DATE_ADD('" & tanggalawal & "', INTERVAL(7-DAYOFWEEK('" & tanggalawal & "')) DAY))) 
                    '                  "
                    'End If

                    'Console.WriteLine(Mcom.CommandText)
                    TraceLog("Kueri : " & Mcom.CommandText)
                    Mcom.ExecuteNonQuery()

                    'Mcom.CommandText = "ALTER TABLE `pos`.temp_jadwal_pjr_estimasi_" & j & "_" & i & " CHANGE `h1` `h1` VARCHAR(10) NULL;"
                    'Mcom.ExecuteNonQuery()
                    Mcom.CommandText = "SELECT COUNT(*) FROM temp_jadwal_pjr_estimasi_" & j & "_" & i & " "
                    'Console.WriteLine(Mcom.CommandText)

                    If Mcom.ExecuteScalar = 0 Then
                        Mcom.CommandText = "INSERT INTO temp_jadwal_pjr_estimasi_" & j & "_" & i & " 
                                        SELECT nik,nama,"
                        If i = 1 Then
                            Mcom.CommandText &= "'Senin' "
                        ElseIf i = 2 Then
                            Mcom.CommandText &= "'Selasa' "
                        ElseIf i = 3 Then
                            Mcom.CommandText &= "'Rabu' "
                        ElseIf i = 4 Then
                            Mcom.CommandText &= "'Kamis' "
                        ElseIf i = 5 Then
                            Mcom.CommandText &= "'Jumat' "
                        ElseIf i = 6 Then
                            Mcom.CommandText &= "'Sabtu' "
                        ElseIf i = 7 Then
                            Mcom.CommandText &= "'Minggu' "
                        End If

                        Mcom.CommandText &= ",'',0 as h" & i & " FROM jadwal_penanggungjawabrak WHERE NIK = '" & Rtn.Rows(j)("nik") & "' GROUP BY NIK"
                        'Console.WriteLine(Mcom.CommandText)
                        Mcom.ExecuteNonQuery()
                    End If
                    Mcom.CommandText = "INSERT INTO temp_jadwal_pjr_estimasi_h" & i & " SELECT * FROM temp_jadwal_pjr_estimasi_" & j & "_" & i & ""
                    Mcom.ExecuteNonQuery()
                    'Console.WriteLine(Mcom.CommandText)

                    'sqltampung &= "LEFT JOIN temp_jadwal_pjr_estimasi_" & j & "_" & i & " a" & j & "_" & i & " ON a.nik = a" & j & "_" & i & ".nik "

                Next


            Next
            'Console.WriteLine(sqltampung)
            Madp.SelectCommand.CommandText = "SELECT DISTINCT a.NIK,a.nama,a.jabatan,h1.h1,h2.h2,h3.h3,h4.h4,h5.h5,h6.h6,h7.h7 FROM jadwal_penanggungjawabrak a "
            'Madp.SelectCommand.CommandText &= sqltampung & "
            For i As Integer = 1 To 7
                Madp.SelectCommand.CommandText &= " LEFT JOIN temp_jadwal_pjr_estimasi_h" & i & " h" & i & " ON a.nik = h" & i & ".nik "

            Next

            Madp.SelectCommand.CommandText &= " GROUP BY NIK"
            'Console.WriteLine(Madp.SelectCommand.CommandText)
            Rtn.Clear()
            Madp.Fill(Rtn)






        Catch ex As Exception
            TraceLog(ex.Message & ex.StackTrace)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function cariJadwalHari(ByVal hari As String) As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim DtCP As New DataTable
        Dim Rtn As New DataTable
        Dim sqltampung As String = ""
        Dim counter As Integer = 0
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If


            Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`temp_jadwal_pjr_estimasi_Detail` ( 
                                `NIK` VARCHAR(12), 
                                `NAMA` VARCHAR(99), 
                                `JABATAN` VARCHAR(50),
                                `HARI` VARCHAR(10),
                                `TANGGAL` DATE,
                                `Kode_Modis` VARCHAR(20),
                                `MODIS` VARCHAR(99),
                                `Shelfing` VARCHAR(10),
                                `Addtime` DATE,
                                `Cat_cod` Varchar(8),
                                `Kemasan` Varchar(5),
                                `Totalitem` Varchar(5),
                                `Timesecond` Varchar(5)
                                ) ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci;"
            Mcom.ExecuteNonQuery()

            Madp.SelectCommand.CommandText = "SELECT  NIK,nama,hari,kode_modis FROM jadwal_penanggungjawabrak 
                                            where hari = '" & hari & "' AND STATUSAPPROVAL <> 'Y'
                                            group by nik"
            Rtn.Clear()
            Madp.Fill(Rtn)

            Mcom.CommandText = "DROP TABLE IF EXISTS temp_jadwal_pjr_estimasi_h_" & hari & ""
            Mcom.ExecuteNonQuery()

            Mcom.CommandText = "CREATE TABLE `pos`.`temp_jadwal_pjr_estimasi_h_" & hari & "` ( 
                                    `nik` VARCHAR(12), 
                                    `nama` VARCHAR(99), 
                                    `hari` VARCHAR(30), 
                                    `kode_modis` VARCHAR(99), 
                                    `h_" & hari & "` DOUBLE )"
            Mcom.ExecuteNonQuery()


            For j As Integer = 0 To Rtn.Rows.Count - 1

                Mcom.CommandText = "DROP TABLE IF EXISTS temp_jadwal_pjr_estimasi_" & j & "_" & hari & ";"
                Mcom.ExecuteNonQuery()

                Mcom.CommandText = "CREATE TABLE temp_jadwal_pjr_estimasi_" & j & "_" & hari & " 
                                    SELECT IF(nik is NULL,'" & Rtn.Rows(j)("nik") & "','" & Rtn.Rows(j)("nik") & "') as nik,
                                    IF(nama is NULL,'" & Rtn.Rows(j)("nama") & "','" & Rtn.Rows(j)("nama") & "') as nama,
                                    IF(hari is NULL,'" & hari & "',hari) as hari,
                                    IF(kode_modis is NULL,'" & Rtn.Rows(j)("kode_modis") & "','" & Rtn.Rows(j)("kode_modis") & "') as kode_modis,
                                    IF(SUM(totalestimasi) IS NULL,0,SUM(totalestimasi)) as h_" & hari & " 
                                    FROM jadwal_penanggungjawabrak 
                                    WHERE hari = '" & hari & "' 
                                    And NIK = '" & Rtn.Rows(j)("nik") & "'
                                    AND STATUSAPPROVAL <> 'Y';"

                Mcom.ExecuteNonQuery()

                Mcom.CommandText = "INSERT INTO temp_jadwal_pjr_estimasi_h_" & hari & " SELECT * FROM temp_jadwal_pjr_estimasi_" & j & "_" & hari & ";"
                Mcom.ExecuteNonQuery()

                'counter += 1
                'If counter = 10 Or j = Rtn.Rows.Count - 1 Then
                '    counter = 0
                '    TraceLog("Kueri : " & Mcom.CommandText)
                '    Mcom.ExecuteNonQuery()

                'End If
            Next
            'Console.WriteLine(sqltampung)
            Madp.SelectCommand.CommandText = "SELECT DISTINCT a.NIK,a.nama,a.jabatan,h.h_" & hari & " as h1 FROM jadwal_penanggungjawabrak a "
            'Madp.SelectCommand.CommandText &= sqltampung & "
            Madp.SelectCommand.CommandText &= " LEFT JOIN temp_jadwal_pjr_estimasi_h_" & hari & " h   ON a.nik = h.nik "
            Madp.SelectCommand.CommandText &= " WHERE a.statusapproval <> 'Y' AND h.h_" & hari & " IS NOT NULL "


            Madp.SelectCommand.CommandText &= " GROUP BY NIK"
            Console.WriteLine(Madp.SelectCommand.CommandText)
            Rtn.Clear()
            Madp.Fill(Rtn)

        Catch ex As Exception
            TraceLog(ex.Message & ex.StackTrace)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function


    Public Function GetHari(ByVal nik As String) As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable
        Dim Mcom As New MySqlCommand("", Conn)
        Dim dt As New DataTable
        Dim jabatan As String = ""
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            jabatan = getJabatanVirbacaprod()

            If FormMain.isPengganti = False Then

                Mcom.CommandText = "SET @@lc_time_names = 'id_ID';"
                Mcom.ExecuteNonQuery()
                'Revisi UAT 2, jagaan nik, hari dilepas
                '1 nik dalam 1 hari bisa > 1 modis
                'Madp.SelectCommand.CommandText = "SELECT tanggal FROM `soppagent`.`abssettingshift` 
                '                                WHERE nik = '" & nik & "'  AND DAYNAME(TANGGAL) <> 'Minggu' 
                '                                AND CONCAT(nik,'-',CAST(DAYNAME(tanggal) AS CHAR) ) NOT IN (SELECT CONCAT(NIK,'-',hari) FROM temp_jadwal_penanggungjawabrak)
                '                                group by DAYNAME(tanggal)"
                'Tambahan terkait hasil simulasi 9/15/22
                If FormMain.cbHariBukaToko.Contains(5) Then
                    Madp.SelectCommand.CommandText = "SELECT tanggal FROM `soppagent`.`abssettingshift` 
                                                WHERE nik = '" & nik & "'  AND DAYNAME(TANGGAL) <> 'Minggu' AND DAYNAME(TANGGAL) <> 'Sabtu'   
                                                group by DAYNAME(tanggal) ORDER BY FIELD(DAYNAME(tanggal),'Senin','Selasa','Rabu','Kamis','Jumat')"
                Else
                    Madp.SelectCommand.CommandText = "SELECT tanggal FROM `soppagent`.`abssettingshift` 
                                                WHERE nik = '" & nik & "'  AND DAYNAME(TANGGAL) <> 'Minggu' 
                                                group by DAYNAME(tanggal) ORDER BY FIELD(DAYNAME(tanggal),'Senin','Selasa','Rabu','Kamis','Jumat','Sabtu')"
                End If
                'Madp.SelectCommand.CommandText = "SELECT tanggal FROM `soppagent`.`abssettingshift` 
                '                                WHERE nik = '" & nik & "'  AND DAYNAME(TANGGAL) <> 'Minggu' 
                '                                group by DAYNAME(tanggal) ORDER BY FIELD(DAYNAME(tanggal),'Senin','Selasa','Rabu','Kamis','Jumat','Sabtu')"
                Console.WriteLine(Madp.SelectCommand.CommandText)
                Madp.Fill(Rtn)
                Mcom.CommandText = "SET @@lc_time_names = 'en_US';"
                Mcom.ExecuteNonQuery()

            Else

                'AMBIL DATA SHIFT DALAM MINGGU TERKAIT

                Madp.SelectCommand.CommandText = "SELECT a.NIK,a.TANGGAL FROM SOPPAGENT.ABSSETTINGSHIFT a  LEFT JOIN `soppagent`.`abspegawaimst` b ON a.nik = b.menoin
					                            WHERE b.jabatan IN (" & jabatan & ") AND 
                                               TANGGAL >= DATE_ADD(CAST(CURDATE() AS DATE), INTERVAL(2-DAYOFWEEK(CAST(CURDATE() AS DATE ))) DAY) 
                                               AND
                                               TANGGAL <= DATE_ADD(CAST(CURDATE() AS DATE), INTERVAL(8-DAYOFWEEK(CAST(CURDATE() AS DATE ))) DAY) AND a.NIK = '" & nik & "'"
                Console.WriteLine(Madp.SelectCommand.CommandText)
                dt.Clear()
                Madp.Fill(dt)
                Dim menoin As String = ""
                Dim tanggal As Date
                Dim tanggal2 As String = ""
                Dim concat As String = ""
                Dim hari As String = ""
                'looping utk cek absen
                For i As Integer = 0 To dt.Rows.Count - 1
                    menoin = dt.Rows(i)("nik").ToString
                    Console.WriteLine(menoin)
                    tanggal = dt.Rows(i)("tanggal")
                    tanggal2 = tanggal.ToString("yyyy-MM-dd")
                    hari = tanggal.ToString("dddd", New Globalization.CultureInfo("id-ID"))
                    Console.WriteLine(hari)
                    Console.WriteLine(tanggal2)
                    If tanggal2 <= Date.Now.ToString("yyyy-MM-dd") Then

                        'cek absen dari soppagent
                        'jika tidak ada = tidak absen
                        Mcom.CommandText = "SELECT count(*) from SOPPAGENT.absabsensitrnoffline where nik = '" & menoin & "' AND DECODE(tanggal,'EuVxq6hKnqe7pNtsP9c2dePyez6ABuD5') = '" & tanggal2 & "'"
                        Console.WriteLine(Mcom.CommandText)
                        'jika tidak absen di tanggal shift i
                        If Mcom.ExecuteScalar = "0" Then
                            Mcom.CommandText = "SELECT count(1) FROM TEMP_JADWAL_PJR WHERE   NIK = '" & menoin & "' AND TANGGAL = '" & tanggal2 & "'
                                                AND RECID = ''"
                            If Mcom.ExecuteScalar > 0 Then

                                'dt.Rows.Add(menoin)
                                'mau cek apakah sudah melaksanakan pjr di tanggal shift i
                                Mcom.CommandText = "SELECT HARI FROM TEMP_JADWAL_PJR WHERE  NIK = '" & menoin & "' AND TANGGAL = '" & tanggal2 & "'
                                                AND CONCAT(NIK,'-',hari)  IN (SELECT CONCAT(NIK,'-',hari) FROM temp_jadwal_penanggungjawabrak_pengganti) AND RECID = '1'"
                                Console.WriteLine(Mcom.CommandText)

                                'jika tidak ada / tidak ketemu
                                If Mcom.ExecuteScalar = "" Then

                                    concat &= "'" & hari & "',"

                                End If
                            End If

                            Console.WriteLine(concat)
                        Else
                            Mcom.CommandText = "SELECT count(1) FROM TEMP_JADWAL_PJR WHERE   NIK = '" & menoin & "' AND TANGGAL = '" & tanggal2 & "'
                                                AND RECID = ''"
                            If Mcom.ExecuteScalar > 0 Then

                                'jika absen tetapi tidak melaksakan pjr juga dicek
                                Mcom.CommandText = "SELECT HARI FROM TEMP_JADWAL_PJR WHERE   NIK = '" & menoin & "' AND TANGGAL = '" & tanggal2 & "'
                                                AND CONCAT(NIK,'-',hari)  IN (SELECT CONCAT(NIK,'-',hari) FROM temp_jadwal_penanggungjawabrak_pengganti) AND RECID = '1'"

                                Console.WriteLine(Mcom.CommandText)

                                'jika tidak ada / tidak ketemu
                                Console.WriteLine(Mcom.CommandText)
                                If Mcom.ExecuteScalar = "" Then
                                    concat &= "'" & hari & "',"
                                End If

                            End If
                        End If
                    End If
                Next
                If concat.Length <> 0 Then
                    concat = concat.Substring(0, concat.Length - 1)
                    Console.WriteLine(concat)
                Else
                    concat = "''"
                End If

                Madp.SelectCommand.CommandText = "SELECT distinct hari FROM jadwal_penanggungjawabrak WHERE NIK = '" & menoin & "' AND HARI IN(" & concat & ") AND
                                                (nik,hari,kode_modis,norak) NOT IN 
                                                (SELECT nik,hari,kode_modis,norak FROM temp_jadwal_penanggungjawabrak_pengganti WHERE NIK = '" & menoin & "' AND HARI IN( " & concat & "))
                                                AND (nik,hari,kode_modis,norak)  IN
                                                (SELECT nik,hari,kode_modis,norak FROM temp_jadwal_pjr WHERE NIK = '" & menoin & "' 
                                                AND TANGGAL >= DATE_ADD(CAST(CURDATE() AS DATE), INTERVAL(2-DAYOFWEEK(CAST(CURDATE() AS DATE ))) DAY) 
                                                AND TANGGAL <= DATE_ADD(CAST(CURDATE() AS DATE), INTERVAL(8-DAYOFWEEK(CAST(CURDATE() AS DATE ))) DAY)                                             
                                                AND TANGGAL <= CURDATE() AND RECID = '')"
                TraceLog(Madp.SelectCommand.CommandText)
                Madp.Fill(Rtn)
                'Console.WriteLine(Rtn.Rows(0)("HARI").ToString)
            End If


        Catch ex As Exception
            Rtn = New DataTable
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function GetHariPengganti(ByVal nik As String, ByVal hari As String) As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable
        Dim Mcom As New MySqlCommand("", Conn)
        Dim dt As New DataTable
        Dim str As String = ""
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Dim jabatan As String = ""
            jabatan = getJabatanVirbacaprod()

            Madp.SelectCommand.CommandText = "SELECT a.NIK,a.TANGGAL FROM SOPPAGENT.ABSSETTINGSHIFT a  LEFT JOIN `soppagent`.`abspegawaimst` b ON a.nik = b.menoin
                             WHERE b.jabatan IN (" & jabatan & ") AND 
                                           TANGGAL >= DATE_ADD(CAST(CURDATE() AS DATE), INTERVAL(2-DAYOFWEEK(CAST(CURDATE() AS DATE ))) DAY) 
                                           AND
                                           TANGGAL <= DATE_ADD(CAST(CURDATE() AS DATE), INTERVAL(8-DAYOFWEEK(CAST(CURDATE() AS DATE ))) DAY) 
                                           AND a.NIK = '" & nik & "'
                                           "
            Console.WriteLine(Madp.SelectCommand.CommandText)
            dt.Clear()
            Madp.Fill(dt)
            Dim menoin As String = ""
            Dim tanggal As Date
            Dim tanggal2 As String = ""
            Dim concat As String = ""
            For i As Integer = 0 To dt.Rows.Count - 1
                menoin = dt.Rows(i)("nik").ToString
                Console.WriteLine(menoin)
                tanggal = dt.Rows(i)("tanggal")
                tanggal2 = tanggal.ToString("yyyy-MM-dd")
                Console.WriteLine(tanggal2)

                concat &= "'" & tanggal2 & "',"
            Next
            If concat.Length <> 0 Then
                concat = concat.Substring(0, concat.Length - 1)
                Console.WriteLine(concat)
            Else
                concat = "''"

            End If
            Madp.SelectCommand.CommandText = "SELECT TANGGAL FROM SOPPAGENT.ABSSETTINGSHIFT WHERE nik = '" & nik & "' AND TANGGAL IN (" & concat & ") AND TANGGAL >= CURDATE()"
            Console.WriteLine(Madp.SelectCommand.CommandText)

            Rtn.Clear()
            Madp.Fill(Rtn)



        Catch ex As Exception
            Rtn = New DataTable
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function


    'Function untuk menampung perhitungan timesecond rak per catcod ambil dari tabel PENANGGUNGJAWAB_RAK (WRHO -> WRC -> CREATE PENANGGUNGJAWAB_RAK)
    Public Function getJadwal_menit() As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim DtCP As New DataTable
        Dim Rtn As New DataTable
        Dim dt1 As New DataTable
        Dim dt2 As New DataTable
        Dim dt3 As New DataTable
        Dim dt4 As New DataTable
        Dim dt5 As New DataTable
        Dim sqltampung As String = ""
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            Mcom.CommandText = "DROP TABLE IF EXISTS `pos`.`temp_jadwal_pjr_waktu` "
            Mcom.ExecuteNonQuery()
            'FormMain.PnlLoading.Visible = True

            Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`temp_jadwal_pjr_waktu` ( 
                                `NAMARAK` VARCHAR(30), 
                                `CATCOD` VARCHAR(8),
                                `KEMASAN` VARCHAR(4), 
                                `ESTIMASI` VARCHAR(5) )"
            Mcom.ExecuteNonQuery()

            Mcom.CommandText = "INSERT INTO TEMP_JADWAL_PJR_WAKTU SELECT kodemodis,cat_cod,b.kemasan,timesecond FROM prodmast a 
                                INNER JOIN penanggungjawab_rak b ON a.cat_cod = CONCAT(0,b.ctgr) 
                                INNER JOIN rak c ON a.prdcd = c.prdcd
                                INNER JOIN jadwal_penanggungjawabrak d ON c.kodemodis = d.kode_modis
                                WHERE d.STATUSAPPROVAL <> ''
                                GROUP BY a.prdcd,cat_cod,b.kemasan,kodemodis"
            Mcom.ExecuteNonQuery()

            'FormMain.PnlLoading.Visible = False

        Catch ex As Exception
            TraceLog("Error WDCP " & ex.Message & ex.StackTrace)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function getPJR() As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim DtCP As New DataTable
        Dim Rtn As New DataTable
        Dim dt1 As New DataTable
        Dim dt2 As New DataTable
        Dim dt3 As New DataTable
        Dim dt4 As New DataTable
        Dim dt5 As New DataTable

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Mcom.CommandText = "SELECT NIK FROM TEMP_JADWAL_PJR"


        Catch ex As Exception

        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function


    Public Function getLaporanPJR() As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)

        Dim Rtn As New DataTable
        Dim dt As New DataTable

        Dim hari As String


        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If


            For i As Integer = 1 To 7
                Dim sqltampung As String = ""

                Mcom.CommandText = "DROP TABLE IF EXISTS `temp_jadwal_pjr_" & i & "`"
                Mcom.ExecuteNonQuery()

                If i = 1 Then
                    hari = "Senin"
                ElseIf i = 2 Then
                    hari = "Selasa"
                ElseIf i = 3 Then
                    hari = "Rabu"
                ElseIf i = 4 Then
                    hari = "Kamis"
                ElseIf i = 5 Then
                    hari = "Jumat"
                ElseIf i = 6 Then
                    hari = "Sabtu"
                Else
                    hari = "Minggu"
                End If


                Madp.SelectCommand.CommandText = "SELECT DISTINCT NIK FROM TEMP_JADWAL_PJR"
                dt.Clear()
                Madp.Fill(dt)

                For j As Integer = 0 To dt.Rows.Count - 1
                    If j < dt.Rows.Count - 1 Then
                        sqltampung &= "SELECT nik,nama,jabatan,hari,kode_modis,shelfing FROM temp_jadwal_pjr WHERE nik = '" & dt.Rows(j)("nik") & "' AND hari = '" & hari & "' UNION "
                    Else
                        sqltampung &= "SELECT nik,nama,jabatan,hari,kode_modis,shelfing FROM temp_jadwal_pjr WHERE nik = '" & dt.Rows(j)("nik") & "' AND hari = '" & hari & "'"
                    End If
                Next
                ''Console.Writeline(sqltampung)

                Mcom.CommandText = "CREATE TABLE temp_jadwal_pjr_" & i & "
                                    " & sqltampung & ""
                ''Console.Writeline(Mcom.CommandText)
                Mcom.ExecuteNonQuery()


            Next
        Catch ex As Exception

        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function


    'Public Function approvePJR(ByVal approval As String, ByVal tanggal As String) As Boolean
    Public Function approvePJR(ByVal approval As String, ByVal hari As String) As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)

        Dim Rtn As Boolean
        Dim dt As New DataTable

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            Mcom.CommandText = "UPDATE jadwal_penanggungjawabrak SET STATUSAPPROVAL = '" & approval & "' where hari = '" & hari & "' "
            Mcom.ExecuteNonQuery()
            Rtn = True
        Catch ex As Exception
            TraceLog(ex.Message & ex.StackTrace)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function scanFinger(ByVal jenis_otorisasioperasional As String) As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)

        Dim jenis As String = ""
        Dim station As String = ""

        Dim TbName As String = ""
        TbName = "OtorisasiOperasional"

        Dim DoubleScan As Boolean = False

        If jenis_otorisasioperasional.ToUpper = "WDCP" Then
            jenis = "WDCP"
            'ElseIf jenis_otorisasioperasional.ToUpper = "BA AS/AM" Then
            '    jenis = "BA AS/AM"
            'ElseIf jenis_otorisasioperasional.ToUpper = "SO PRODUK KHUSUS" Then
            '    jenis = "SO PRODUK KHUSUS"
        ElseIf jenis_otorisasioperasional.ToUpper = "WDCP_PJR" Then
            jenis = "WDCP_PJR"
        ElseIf jenis_otorisasioperasional.ToUpper = "WDCP_PJR 2" Then
            jenis = "WDCP_PJR 2"

        End If


        IDM.Fungsi.TraceLog("Validasi " & jenis & " Finger Scan")

        Try

            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            If IDM.Fungsi.Get_Station.StartsWith("0") Then
                station = IDM.Fungsi.Get_Station().Substring(1)
            Else
                station = Get_Station()
            End If
            Mcom.CommandText = "SELECT JENIS FROM const WHERE rkey='ABS' and recid = '" & station & "'"

            IDM.Fungsi.TraceLog("Query: " & Mcom.CommandText)

            If Mcom.ExecuteScalar = "Y" Then
                If CekVersiProgram(Application.StartupPath & "\" & "ScanFinger.dll", "1.0.0.3") Then
                    Try
                        If Conn.State = ConnectionState.Closed Then
                            Conn.Open()
                        End If
                        Mcom.CommandText = "SHOW TABLES LIKE '" & TbName & "'"
                        IDM.Fungsi.TraceLog("Query: " & Mcom.CommandText)
                        If Mcom.ExecuteScalar & "" <> "" Then
                            Dim SQLQuery As String = ""

                            SQLQuery = "SELECT COUNT(*) FROM `" & TbName & "`"
                            SQLQuery &= " WHERE jenis = '" & jenis & "'"
                            SQLQuery &= " AND isAktif = '1';"
                            Mcom.CommandText = SQLQuery
                            TraceLog(SQLQuery)
                            If Mcom.ExecuteScalar = 0 Then
                                SQLQuery = "INSERT INTO otorisasioperasional(isAktif, Jenis, isDoubleApproval, Jabatan1, Jabatan2)"
                                SQLQuery &= " VALUES("
                                SQLQuery &= " '1',"
                                SQLQuery &= " '" & jenis & "',"
                                SQLQuery &= " 'N',"

                                If jenis = "WDCP" Then
                                    SQLQuery &= " 'Junior Supervisor,Area Jr. Manager',"
                                    SQLQuery &= " 'Junior Supervisor,Area Jr. Manager'"
                                    'ElseIf jenis = "BA AS/AM" Then
                                    '    SQLQuery &= " 'JUNIOR SUPERVISOR,AREA JR. MANAGER',"
                                    '    SQLQuery &= " 'JUNIOR SUPERVISOR,AREA JR. MANAGER'"
                                ElseIf jenis = "WDCP_PJR" Then
                                    SQLQuery &= " 'KASIR,KASIR (SS),PRAMUNIAGA,PRAMUNIAGA (SS),MERCHANDISER,MERCHANDISER (SS),ASISTEN KEPALA TOKO,ASISTEN KEPALA TOKO (SS),KEPALA TOKO,KEPALA TOKO (SS),STORE CREW,STORE CREW (SS),STORE JR. LEADER,STORE JR. LEADER (SS),STORE SR. LEADER,STORE SR. LEADER (SS),CHIEF OF STORE,CHIEF OF STORE (SS),STORE CREW BOY,STORE CREW GIRL,STORE CREW BOY (SS),STORE CREW GIRL (SS)',"
                                    SQLQuery &= " 'KASIR,KASIR (SS),PRAMUNIAGA,PRAMUNIAGA (SS),MERCHANDISER,MERCHANDISER (SS),ASISTEN KEPALA TOKO,ASISTEN KEPALA TOKO (SS),KEPALA TOKO,KEPALA TOKO (SS),STORE CREW,STORE CREW (SS),STORE JR. LEADER,STORE JR. LEADER (SS),STORE SR. LEADER,STORE SR. LEADER (SS),CHIEF OF STORE,CHIEF OF STORE (SS),STORE CREW BOY,STORE CREW GIRL,STORE CREW BOY (SS),STORE CREW GIRL (SS)'"

                                ElseIf jenis = "WDCP_PJR 2" Then
                                    SQLQuery &= " 'KEPALA TOKO,KEPALA TOKO (SS),ASISTEN KEPALA TOKO,ASISTEN KEPALA TOKO (SS),MERCHANDISER,MERCHANDISER (SS),STORE JR. LEADER,STORE JR. LEADER (SS),STORE SR. LEADER,STORE SR. LEADER (SS),CHIEF OF STORE,CHIEF OF STORE (SS)',"
                                    SQLQuery &= " 'KEPALA TOKO,KEPALA TOKO (SS),ASISTEN KEPALA TOKO,ASISTEN KEPALA TOKO (SS),MERCHANDISER,MERCHANDISER (SS),STORE JR. LEADER,STORE JR. LEADER (SS),STORE SR. LEADER,STORE SR. LEADER (SS),CHIEF OF STORE,CHIEF OF STORE (SS)'"
                                End If

                                SQLQuery &= " );"
                                Mcom.CommandText = SQLQuery
                                TraceLog(Mcom.CommandText)
                                Mcom.ExecuteNonQuery()
                            End If

                            SQLQuery = "select isDoubleApproval from `" & TbName & "`"
                            SQLQuery &= " where jenis = '" & jenis & "'"
                            SQLQuery &= " and isAktif = '1'"

                            Mcom.CommandText = SQLQuery
                            IDM.Fungsi.TraceLog("Query: " & Mcom.CommandText)
                            If Mcom.ExecuteScalar = "Y" Then
                                DoubleScan = True
                                TraceLog("Validasi " & jenis & " Finger Scan : Double Approval")
                            End If
                        Else
                            MsgBox("Table " & TbName & " Tidak Ada!")
                            Exit Function
                        End If

                    Catch ex As Exception
                        TraceLog("err: " & ex.Message & vbCrLf & ex.StackTrace)
                        Return False
                    Finally
                        If Conn.State <> ConnectionState.Closed Then
                            Conn.Close()
                        End If
                        Conn.Dispose()
                    End Try


                    Dim objScanFinger As New ScanFinger.ClsScan
                    Dim HasilScanFinger As String()
                    Dim HasilJabatan As String = ""
                    Dim HasilNIK As String = ""
                    Dim ScanFinger_1 As String() = objScanFinger.Otorisasi(jenis, "Scan Finger Otorisasi ke-1")
                    Dim ScanFinger_2 As String()


                    If ScanFinger_1(0) = "1" Then
                        HasilScanFinger = ScanFinger_1(2).Split("|")
                        HasilJabatan = HasilScanFinger(3)
                        HasilNIK = HasilScanFinger(0)
                        ''Console.Writeline(HasilNIK)
                        ConstNIKPJR(HasilNIK)
                        'FormMain.NikToko = HasilNIK
                        'NamaToko = HasilScanFinger(1)

                        TraceLog("HasilScanFinger Single : " & ScanFinger_1(2))

                        If DoubleScan = True Then
                            ScanFinger_2 = objScanFinger.Otorisasi(jenis, "Scan Finger Otorisasi ke-2", HasilNIK & "," & HasilJabatan)
                            If ScanFinger_2(0) = "1" Then
                                HasilScanFinger = ScanFinger_2(2).Split("|")
                                HasilJabatan &= "," & HasilScanFinger(3)
                                'NikIC = HasilScanFinger(0)
                                'NamaIC = HasilScanFinger(1)
                                TraceLog("HasilScanFinger Double : " & ScanFinger_2(2))
                            Else
                                TraceLog("Validasi " & jenis & " Finger Scan Double : " & ScanFinger_2(2))
                                MsgBox("Gagal scan finger kedua : " & ScanFinger_2(2))
                                Return False
                            End If
                        End If

                    Else
                        TraceLog("Validasi " & jenis & " Finger Scan Single : " & ScanFinger_1(2))
                        MsgBox("Gagal scan finger pertama : " & ScanFinger_1(2))
                        Return False
                    End If

                    If DoubleScan = True Then
                        If HasilJabatan.ToUpper = "JABATAN1,JABATAN2" Or HasilJabatan.ToUpper = "JABATAN2,JABATAN1" Then
                            Return True
                        Else
                            TraceLog("Validasi " & jenis & " Finger Scan Double : " & ScanFinger_2(2))
                            MsgBox("Validasi Finger Scan, tidak sesuai !")
                            Return False
                        End If

                    Else

                        If HasilJabatan.ToUpper = "JABATAN1" Or HasilJabatan.ToUpper = "JABATAN2" Then
                            Return True
                        Else
                            TraceLog("Validasi " & jenis & " Finger Scan Single : " & ScanFinger_1(2))
                            MsgBox("Validasi Finger Scan, tidak sesuai !")
                            Return False
                        End If
                    End If
                Else
                    TraceLog("Validasi " & jenis & " Pass Toko")
                    MsgBox("Validasi Finger Scan, VERSI PROGRAM tidak sesuai !")
                End If
            Else
                TraceLog("Validasi " & jenis & " Pass Toko")
                MsgBox("Validasi Finger Scan, CONST ABS tidak sesuai !")
            End If

        Catch ex As Exception
            IDM.Fungsi.ShowError("ERR: ", ex.Message & vbCrLf & ex.StackTrace)
        End Try

    End Function
    Public Function CekVersiProgram(ByVal ProgName As String, ByVal Version As String, Optional ByVal CekVersion As Boolean = False) As Boolean
        Try
            Dim myBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(ProgName)

            Dim VersiToko() As String = myBuildInfo.FileVersion.ToString.Split(".")
            Dim VersiServer() As String = Version.Split(".")

            Dim VERSI_MAJOR As Integer = 0
            Dim VERSI_MINOR As Integer = 0
            Dim VERSI_BUILD As Integer = 0
            Dim VERSI_REVISION As Integer = 0

            Dim VERSI_MAJOR_SERVER As Integer = 0
            Dim VERSI_MINOR_SERVER As Integer = 0
            Dim VERSI_BUILD_SERVER As Integer = 0
            Dim VERSI_REVISION_SERVER As Integer = 0

            For k As Short = 0 To VersiToko.Length - 1
                If k = 0 Then
                    VERSI_MAJOR = 0 & VersiToko(k)
                ElseIf k = 1 Then
                    VERSI_MINOR = 0 & VersiToko(k)
                ElseIf k = 2 Then
                    VERSI_BUILD = 0 & VersiToko(k)
                ElseIf k = 3 Then
                    VERSI_REVISION = 0 & VersiToko(k)
                End If
            Next

            For k As Short = 0 To VersiServer.Length - 1
                If k = 0 Then
                    VERSI_MAJOR_SERVER = 0 & VersiServer(k)
                ElseIf k = 1 Then
                    VERSI_MINOR_SERVER = 0 & VersiServer(k)
                ElseIf k = 2 Then
                    VERSI_BUILD_SERVER = 0 & VersiServer(k)
                ElseIf k = 3 Then
                    VERSI_REVISION_SERVER = 0 & VersiServer(k)
                End If
            Next

            If CekVersion Then
                TraceLog(ProgName & " " & VERSI_MAJOR & "." & VERSI_MINOR & "." & VERSI_BUILD & "." & VERSI_REVISION & " " & Version)
            End If

            If (VERSI_MAJOR * 1000000 + VERSI_MINOR * 10000 + VERSI_BUILD * 100 + VERSI_REVISION) >= (VERSI_MAJOR_SERVER * 1000000 + VERSI_MINOR_SERVER * 10000 + VERSI_BUILD_SERVER * 100 + VERSI_REVISION_SERVER) Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Return False
        End Try
    End Function
    'hari, NIK, MODIS, norak, noshelf
    Public Function ambilPersonilPJR(ByVal hari As String, ByVal nik As String, ByVal modis As String, ByVal norak As String, ByVal shelf As String) As ClsPJR
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Scon)
        Dim Madp As New MySqlDataAdapter("", Scon)
        Dim DtCP As New DataTable
        Dim Rtn As New Boolean
        Dim clsPJR As New ClsPJR
        Dim tmpDt As New DataTable


        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Madp.SelectCommand.CommandText = " SELECT nik,nama,  HARI, KODE_MODIS,MODIS,norak, SHELFING FROM TEMP_JADWAL_penanggungjawabrak
                                WHERE NIK = '" & nik & "'  AND HARI = '" & hari & "' 
                                  AND MODIS = '" & modis & "' AND SHELFING = '" & shelf & "' AND norak = '" & norak & "' "
            ''Console.Writeline(Madp.SelectCommand.CommandText)
            Madp.Fill(tmpDt)

            If tmpDt.Rows.Count > 0 Then
                clsPJR.NIK = tmpDt.Rows.Item(0)("nik").ToString
                clsPJR.NIK = tmpDt.Rows.Item(0)("nama").ToString
                clsPJR.HARI = tmpDt.Rows.Item(0)("HARI").ToString
                clsPJR.MODIS = tmpDt.Rows.Item(0)("KODE_MODIS").ToString
                clsPJR.NAMAMODIS = tmpDt.Rows.Item(0)("MODIS").ToString
                clsPJR.SHELFFROM = tmpDt.Rows.Item(0)("SHELFING").ToString.Split("-")(0)
                clsPJR.SHELFTO = tmpDt.Rows.Item(0)("SHELFING").ToString.Split("-")(1)
                clsPJR.NORAK = tmpDt.Rows.Item(0)("norak").ToString


                'Console.Writeline(clsPJR.HARI)
                'Console.Writeline(clsPJR.MODIS)
                'Console.Writeline(clsPJR.NIK)
                'Console.Writeline(clsPJR.SHELFFROM)
                'Console.Writeline(clsPJR.SHELFTO)
                'Console.Writeline(clsPJR.NORAK)

            End If



        Catch ex As Exception
            Rtn = False
            MsgBox(ex.Message & ex.StackTrace)
        Finally
            Scon.Close()
        End Try
        Return clsPJR
    End Function

    Public Function GetDeskripsiPJR(ByVal tabel_name As String, ByVal barcode_plu As String,
                                          ByVal NamaRak As String, ByVal ListShelf As String, ByVal User As ClsUser) As ClsPJRProduk
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Result As New ClsPJRProduk
        Dim batas_Retur As String = ""
        Dim temp_plu As String = ""

        If Conn Is Nothing Then
            'MsgBox("Gagal koneksi")
            Return Result
            Exit Function
        End If

        SyncLock Conn
            Try
                If Conn.State = ConnectionState.Closed Then
                    Conn.Open()
                End If

                Dim DtPlano As New DataTable
                Dim MaxBatasRetur As String = ""
                Dim MaxBatasReturS As String = ""
                Dim BatasPajang As String = ""
                Dim Price As String = ""
                'Console.Writeline(barcode_plu)
                'uat memo 092/CPS/22
                'where input plu dinonaktifkan
                'memo 1285/CPS/22
                'tambah baca barcode_cbr untuk item yg tidak ada barcode
                Mcom.CommandText = "SELECT PLU FROM BARCODE WHERE PLU = '" & barcode_plu & "'"
                IDM.Fungsi.TraceLog("Query: " & Mcom.CommandText)
                If Mcom.ExecuteScalar <> "" Then

                    Mcom.CommandText = "SELECT PLU FROM barcode_cbr WHERE PLU = '" & barcode_plu & "' "
                    IDM.Fungsi.TraceLog("Kueri " & Mcom.CommandText)

                    If Mcom.ExecuteScalar <> "" Then
                        IDM.Fungsi.TraceLog("Kueri " & Mcom.ExecuteScalar)

                        Result.Desc = "tolak"
                        GoTo skip

                    End If
                End If
                If ListShelf = "" Then
                    ListShelf = "'1'"
                End If
                batas_Retur = getBTRVirbacaprod()
                If tabel_name = "TINDAKLBTD" Then
                    Mcom.CommandText = "SELECT PLU FROM barcode WHERE BARCD = '" & barcode_plu & "'"
                    temp_plu = Mcom.ExecuteScalar
                End If

                Mcom.CommandText = " SELECT r.PRDCD, p.DESC2, bt.NORAK as NORAK, bt.kode_modis AS NAMA_RAK, r.NOSHELF, 
                                    IFNULL(t.PRDCD, If(o.tgl_akh >= CURDATE(), o.prdcd, NULL)) As PLU_PTAG, 
                                    t.TGL_AKH, IFNULL(br.MAX_RET_TOKO2DCI,0) AS MAX_RET_TOKO2DCI, br.MAX_RET_TOKO2DCI_S, p.PRICE, 
                                    IFNULL(t.PRICE, If(o.tgl_akh >= CURDATE(), o.promosi, NULL)) As PRICE_PTAG, 
                                    t.PROMOSI, s.QTY, p.DEPART, r.KIRIKANAN,
                                    CAST(DATE_FORMAT(DATE_ADD(CURDATE(), INTERVAL IF(max_ret_toko2dci_s='B',
                                    (max_ret_toko2dci*30),max_ret_toko2dci) DAY),'%d-%m-%Y') AS CHAR) AS Tanggal_Batas_Aman 
                                    From prodmast p 
                                    LEFT Join rak r ON p.PRDCD = r.PRDCD 
                                    Left Join stmast s ON p.PRDCD = s.PRDCD 
                                    Left Join barcode b ON p.PRDCD = b.PLU 
                                    LEFT JOIN barcode_cbr cbr ON p.PRDCD = cbr.PLU 
                                    Left Join ptag t ON p.PRDCD = t.PRDCD 
                                    Left Join ptag_old o ON p.PRDCD = o.PRDCD 
                                    Left Join batas_retur br ON p.PRDCD = br.FMKODE 
                                    LEFT JOIN (SELECT KODE_MODIS,NORAK FROM JADWAL_PENANGGUNGJAWABRAK WHERE kode_modis = '" & NamaRak & "' AND norak = '" & FormMain.norak_pjr & "' ) bt ON r.kodemodis = bt.kode_modis
                                    WHERE  (b.BARCD = '" & barcode_plu & "' OR cbr.BARCD = '" & barcode_plu & "')
                                    And (br.max_ret_toko2dci_s IN (" & batas_Retur & ") OR br.max_ret_toko2dci_s IS NULL OR br.max_ret_toko2dci_s ='0' OR br.max_ret_toko2dci_s = '')
                                    AND kode_modis = '" & NamaRak & "' AND  noshelf IN (" & ListShelf & ")  AND bt.norak = '" & FormMain.norak_pjr & "' "
                If tabel_name = "TINDAKLBTD" Then
                    'Mcom.CommandText &= " AND R.PRDCD NOT IN (SELECT PLU FROM CEKPJR WHERE PLU = '" & temp_plu & "' AND `STATUS` = 'B')"
                    Mcom.CommandText &= " AND R.PRDCD IN (SELECT PLU FROM CEKPJR WHERE PLU = '" & temp_plu & "' AND `JENISBARANG` = 'TT' AND TGLSCAN = CURDATE())"

                End If

                IDM.Fungsi.TraceLog("Kueri " & Mcom.CommandText)

                Dim sDap As New MySqlDataAdapter(Mcom)
                sDap.Fill(DtPlano)

                If DtPlano.Rows.Count > 15 Then
                    Result.Prdcd = barcode_plu
                    Result.Desc = "Rak Melebihi Batas"
                ElseIf DtPlano.Rows.Count > 0 Then
                    If Not IsNothing(DtPlano.Rows(0).Item("PLU_PTAG")) And IsDBNull(DtPlano.Rows(0).Item("PLU_PTAG")) Then
                        If IsDBNull(DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI")) Then
                            MaxBatasRetur = "" '0
                        Else
                            MaxBatasRetur = DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI")
                        End If

                        If IsDBNull(DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI_S")) Then
                            MaxBatasReturS = ""
                        Else
                            MaxBatasReturS = DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI_S")
                        End If

                        If IsDBNull(DtPlano.Rows(0).Item("Tanggal_Batas_Aman")) Then
                            BatasPajang = "-"
                        Else
                            BatasPajang = DtPlano.Rows(0).Item("Tanggal_Batas_Aman")
                        End If

                        If IsDBNull(DtPlano.Rows(0).Item("PRICE")) Then
                            Price = ""
                        Else
                            Price = "RP " & DtPlano.Rows(0).Item("PRICE").ToString.Split(".")(0)
                        End If
                    Else
                        If IsDBNull(DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI")) Then
                            MaxBatasRetur = "" '0
                        Else
                            MaxBatasRetur = DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI")
                        End If

                        If IsDBNull(DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI_S")) Then
                            MaxBatasReturS = ""
                        Else
                            MaxBatasReturS = DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI_S")
                        End If

                        If IsDBNull(DtPlano.Rows(0).Item("Tanggal_Batas_Aman")) Then
                            BatasPajang = "-"
                        Else
                            BatasPajang = DtPlano.Rows(0).Item("Tanggal_Batas_Aman")
                        End If

                        If IsDBNull(DtPlano.Rows(0).Item("PRICE_PTAG")) Then
                            Price = ""
                        Else
                            Price = "RP " & DtPlano.Rows(0).Item("PRICE_PTAG").ToString.Split(".")(0)
                        End If

                        If Not IsNothing(DtPlano.Rows(0).Item("TGL_AKH")) And Not IsDBNull(DtPlano.Rows(0).Item("TGL_AKH")) Then
                            Dim TglAkhTemp As Date = CDate(DtPlano.Rows(0).Item("TGL_AKH"))
                            If TglAkhTemp < Date.Now Then ' tgl_akh < CURDATE()
                                If IsDBNull(DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI")) Then
                                    MaxBatasRetur = "" '0
                                Else
                                    MaxBatasRetur = DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI")
                                End If

                                If IsDBNull(DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI_S")) Then
                                    MaxBatasReturS = ""
                                Else
                                    MaxBatasReturS = DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI_S")
                                End If

                                If IsDBNull(DtPlano.Rows(0).Item("Tanggal_Batas_Aman")) Then
                                    BatasPajang = "-"
                                Else
                                    BatasPajang = DtPlano.Rows(0).Item("Tanggal_Batas_Aman")
                                End If

                                If IsDBNull(DtPlano.Rows(0).Item("Price")) Then
                                    Price = ""
                                Else
                                    Price = "RP " & DtPlano.Rows(0).Item("Price").ToString.Split(".")(0)
                                End If
                            ElseIf TglAkhTemp >= Date.Now Then ' tgl_akh >= CURDATE()
                                If IsDBNull(DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI")) Then
                                    MaxBatasRetur = "" '0
                                Else
                                    MaxBatasRetur = DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI")
                                End If

                                If IsDBNull(DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI_S")) Then
                                    MaxBatasReturS = ""
                                Else
                                    MaxBatasReturS = DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI_S")
                                End If

                                If IsDBNull(DtPlano.Rows(0).Item("Tanggal_Batas_Aman")) Then
                                    BatasPajang = "-"
                                Else
                                    BatasPajang = DtPlano.Rows(0).Item("Tanggal_Batas_Aman")
                                End If

                                If IsDBNull(DtPlano.Rows(0).Item("PROMOSI")) Then
                                    Price = ""
                                Else
                                    Price = "RP " & DtPlano.Rows(0).Item("PROMOSI").ToString.Split(".")(0)
                                End If
                            End If
                        End If
                    End If
                    If IsDBNull(DtPlano.Rows(0).Item("PRDCD")) Then
                        Result.Prdcd = "Tidak Terdaftar"
                    Else
                        Result.Prdcd = DtPlano.Rows(0).Item("PRDCD")

                    End If
                    'Console.Writeline(Result.Prdcd)
                    If DtPlano.Rows(0).Item("DESC2").Length > 40 Then
                        Result.Desc = DtPlano.Rows(0).Item("DESC2").Substring(0, 40)
                    Else
                        Result.Desc = DtPlano.Rows(0).Item("DESC2").ToString.PadRight(40, " ")
                    End If
                    'If MaxBatasRetur = 0 Then
                    '    Result.MaxRet = MaxBatasRetur
                    'Else
                    '    Result.MaxRet = MaxBatasReturS & "-" & MaxBatasRetur
                    'End If
                    Result.MaxRet = BatasPajang

                    Result.Price = Price
                Else
                    Result.Desc = "Tidak Ditemukan"
                End If

                'Insert ke Table CekPlanogram
                Dim DtCekPlano As New DataTable
                Dim TempNamaRak As String = ""
                Dim TempNoShelf As String = ""
                Dim Status As String = ""
                Dim JenisBarang As String = ""

                If DtPlano.Rows.Count > 0 Then
                    If IsDBNull(DtPlano.Rows(0).Item("NAMA_RAK")) Then
                        TempNamaRak = ""
                    Else
                        TempNamaRak = DtPlano.Rows(0).Item("NAMA_RAK").ToString.ToUpper
                    End If
                    If IsDBNull(DtPlano.Rows(0).Item("NOSHELF")) Then
                        TempNoShelf = ""
                    Else
                        TempNoShelf = DtPlano.Rows(0).Item("NOSHELF").ToString.ToUpper
                    End If

                    'MsgBox("nama rak " & TempNamaRak)
                    'MsgBox("no shelf " & TempNoShelf)


                    Mcom.CommandText = "SELECT PLU FROM " & tabel_name
                    Mcom.CommandText &= " WHERE PLU = '" & DtPlano.Rows(0).Item("PRDCD") & "'"
                    Mcom.CommandText &= " AND DATE(TGLSCAN) = CURDATE()"
                    Mcom.CommandText &= " AND `STATUS` <> 'I' AND NORAKINPUT = '" & DtPlano.Rows(0).Item("NORAK") & "'"
                    Mcom.CommandText &= " AND NOSHELFINPUT = '" & DtPlano.Rows(0).Item("NOSHELF") & "';"
                    IDM.Fungsi.TraceLog("Kueri " & Mcom.CommandText)

                    Dim sDap2 As New MySqlDataAdapter(Mcom)
                    'Console.Writeline(Mcom.CommandText)
                    sDap2.Fill(DtCekPlano)

                    If DtCekPlano.Rows.Count = 0 Then
                        If NamaRak.ToUpper = TempNamaRak And ListShelf.ToUpper.Contains(TempNoShelf) Then
                            If DtPlano.Rows(0).Item("QTY") = "0" Then
                                Status = "S"
                                JenisBarang = "SO"
                            Else
                                Status = "B"
                                JenisBarang = ""
                            End If
                        Else
                            Status = "S"
                            JenisBarang = "SD"
                        End If

                        'Revisi 2020-04-30
                        'TGLSCAN yang seblumnya diset menggunakan SYSDATE(), diubah menggunakan DateTiem.Now dari VB
                        Dim TGLSCAN As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

                        Mcom.CommandText = "INSERT INTO " & tabel_name & " (PLU,TGLSCAN,NAMA,NOSHELFINPUT,NORAKINPUT,NAMARAK,NAMARAKINPUT,"
                        Mcom.CommandText &= "`STATUS`,NIK,NAMAUSER,KIRIKANAN,DIVISI,STOCK,JENISBARANG,MAXBATASRETUR,MAXBATASRETUR_S)"
                        Mcom.CommandText &= " VALUES ('" & DtPlano.Rows(0).Item("PRDCD") & "','" & TGLSCAN & "','','" & TempNoShelf & "',"
                        Mcom.CommandText &= " '" & DtPlano.Rows(0).Item("NORAK") & "','','" & TempNamaRak & "','" & Status & "',"
                        Mcom.CommandText &= " '" & User.ID & "', '" & User.Nama & "', '" & DtPlano.Rows(0).Item("KIRIKANAN") & "',"
                        Mcom.CommandText &= " '" & DtPlano.Rows(0).Item("DEPART") & "', '" & DtPlano.Rows(0).Item("QTY") & "',"
                        Mcom.CommandText &= " '" & JenisBarang & "', '" & MaxBatasRetur & "', '" & MaxBatasReturS & "');"
                        TraceLog("Kueri " & Mcom.CommandText)

                        'Console.Writeline(Mcom.CommandText)
                        Mcom.ExecuteNonQuery()
                    End If
                End If
skip:
                If Result.Desc.ToLower = "tolak" Then
                    Result.Desc = "tolak"
                ElseIf Result.Desc.ToLower = "Tidak Ditemukan" Then
                    Result.Desc = "Tidak Ditemukan"

                End If
            Catch ex As Exception
                Result.Desc = "Tidak Ditemukan"
                MsgBox(ex.Message & ex.StackTrace)

                IDM.Fungsi.TraceLog("Error" & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                Conn.Close()
            End Try
        End SyncLock

        Return Result
    End Function

    Public Function GetDeskripsiPJR_LBTD_BA_PJR(ByVal tabel_name As String, ByVal barcode_plu As String,
                                                ByVal User As ClsUser) As ClsPJRProduk
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Result As New ClsPJRProduk
        Dim batas_Retur As String = ""
        Dim temp_plu As String = ""

        If Conn Is Nothing Then
            'MsgBox("Gagal koneksi")
            Return Result
            Exit Function
        End If

        SyncLock Conn
            Try
                If Conn.State = ConnectionState.Closed Then
                    Conn.Open()
                End If

                Dim DtPlano As New DataTable
                Dim MaxBatasRetur As String = ""
                Dim MaxBatasReturS As String = ""
                Dim BatasPajang As String = ""
                Dim Price As String = ""
                'Console.Writeline(barcode_plu)
                'uat memo 092/CPS/22
                'where input plu dinonaktifkan
                'memo 1285/CPS/22
                'tambah baca barcode_cbr untuk item yg tidak ada barcode
                Mcom.CommandText = "SELECT PLU FROM BARCODE WHERE PLU = '" & barcode_plu & "'"
                IDM.Fungsi.TraceLog("Query: " & Mcom.CommandText)
                If Mcom.ExecuteScalar <> "" Then

                    Mcom.CommandText = "SELECT PLU FROM barcode_cbr WHERE PLU = '" & barcode_plu & "' "
                    IDM.Fungsi.TraceLog("Kueri " & Mcom.CommandText)

                    If Mcom.ExecuteScalar <> "" Then
                        IDM.Fungsi.TraceLog("Kueri " & Mcom.ExecuteScalar)

                        Result.Desc = "tolak"
                        GoTo skip

                    End If
                End If

                batas_Retur = getBTRVirbacaprod()
                If tabel_name = "TINDAKLBTD_BAPJR" Then
                    Mcom.CommandText = "SELECT PLU FROM barcode WHERE BARCD = '" & barcode_plu & "'"
                    temp_plu = Mcom.ExecuteScalar
                End If

                Mcom.CommandText = " SELECT r.PRDCD, p.DESC2, bt.NORAK as NORAK, bt.kode_modis AS NAMA_RAK, r.NOSHELF, 
                                    IFNULL(t.PRDCD, If(o.tgl_akh >= CURDATE(), o.prdcd, NULL)) As PLU_PTAG, 
                                    t.TGL_AKH, IFNULL(br.MAX_RET_TOKO2DCI,0) AS MAX_RET_TOKO2DCI, br.MAX_RET_TOKO2DCI_S, p.PRICE, 
                                    IFNULL(t.PRICE, If(o.tgl_akh >= CURDATE(), o.promosi, NULL)) As PRICE_PTAG, 
                                    t.PROMOSI, s.QTY, p.DEPART, r.KIRIKANAN,
                                    CAST(DATE_FORMAT(DATE_ADD(CURDATE(), INTERVAL IF(max_ret_toko2dci_s='B',
                                    (max_ret_toko2dci*30),max_ret_toko2dci) DAY),'%d-%m-%Y') AS CHAR) AS Tanggal_Batas_Aman 
                                    From prodmast p 
                                    LEFT Join rak r ON p.PRDCD = r.PRDCD 
                                    Left Join stmast s ON p.PRDCD = s.PRDCD 
                                    Left Join barcode b ON p.PRDCD = b.PLU 
                                    LEFT JOIN barcode_cbr cbr ON p.PRDCD = cbr.PLU 
                                    Left Join ptag t ON p.PRDCD = t.PRDCD 
                                    Left Join ptag_old o ON p.PRDCD = o.PRDCD 
                                    Left Join batas_retur br ON p.PRDCD = br.FMKODE 
                                    INNER JOIN (SELECT PRDCD,KODE_MODIS,NORAK FROM ITEMSO_PJR_BA_AS WHERE RECID='' AND (PRDCD = '" & barcode_plu & "' OR PRDCD = '" & temp_plu & "') ) bt ON r.PRDCD = bt.PRDCD
                                    WHERE  (b.BARCD = '" & barcode_plu & "' OR cbr.BARCD = '" & barcode_plu & "')
                                    And (br.max_ret_toko2dci_s IN (" & batas_Retur & ") OR br.max_ret_toko2dci_s IS NULL OR br.max_ret_toko2dci_s ='0' OR br.max_ret_toko2dci_s = '')"

                IDM.Fungsi.TraceLog("Kueri GetDeskripsiPJR_LBTD_BA_PJR : " & Mcom.CommandText)

                Dim sDap As New MySqlDataAdapter(Mcom)
                sDap.Fill(DtPlano)

                If DtPlano.Rows.Count > 15 Then
                    Result.Prdcd = barcode_plu
                    Result.Desc = "Rak Melebihi Batas"
                ElseIf DtPlano.Rows.Count > 0 Then
                    If Not IsNothing(DtPlano.Rows(0).Item("PLU_PTAG")) And IsDBNull(DtPlano.Rows(0).Item("PLU_PTAG")) Then
                        If IsDBNull(DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI")) Then
                            MaxBatasRetur = "" '0
                        Else
                            MaxBatasRetur = DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI")
                        End If

                        If IsDBNull(DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI_S")) Then
                            MaxBatasReturS = ""
                        Else
                            MaxBatasReturS = DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI_S")
                        End If

                        If IsDBNull(DtPlano.Rows(0).Item("Tanggal_Batas_Aman")) Then
                            BatasPajang = "-"
                        Else
                            BatasPajang = DtPlano.Rows(0).Item("Tanggal_Batas_Aman")
                        End If

                        If IsDBNull(DtPlano.Rows(0).Item("PRICE")) Then
                            Price = ""
                        Else
                            Price = "RP " & DtPlano.Rows(0).Item("PRICE").ToString.Split(".")(0)
                        End If
                    Else
                        If IsDBNull(DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI")) Then
                            MaxBatasRetur = "" '0
                        Else
                            MaxBatasRetur = DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI")
                        End If

                        If IsDBNull(DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI_S")) Then
                            MaxBatasReturS = ""
                        Else
                            MaxBatasReturS = DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI_S")
                        End If

                        If IsDBNull(DtPlano.Rows(0).Item("Tanggal_Batas_Aman")) Then
                            BatasPajang = "-"
                        Else
                            BatasPajang = DtPlano.Rows(0).Item("Tanggal_Batas_Aman")
                        End If

                        If IsDBNull(DtPlano.Rows(0).Item("PRICE_PTAG")) Then
                            Price = ""
                        Else
                            Price = "RP " & DtPlano.Rows(0).Item("PRICE_PTAG").ToString.Split(".")(0)
                        End If

                        If Not IsNothing(DtPlano.Rows(0).Item("TGL_AKH")) And Not IsDBNull(DtPlano.Rows(0).Item("TGL_AKH")) Then
                            Dim TglAkhTemp As Date = CDate(DtPlano.Rows(0).Item("TGL_AKH"))
                            If TglAkhTemp < Date.Now Then ' tgl_akh < CURDATE()
                                If IsDBNull(DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI")) Then
                                    MaxBatasRetur = "" '0
                                Else
                                    MaxBatasRetur = DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI")
                                End If

                                If IsDBNull(DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI_S")) Then
                                    MaxBatasReturS = ""
                                Else
                                    MaxBatasReturS = DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI_S")
                                End If

                                If IsDBNull(DtPlano.Rows(0).Item("Tanggal_Batas_Aman")) Then
                                    BatasPajang = "-"
                                Else
                                    BatasPajang = DtPlano.Rows(0).Item("Tanggal_Batas_Aman")
                                End If

                                If IsDBNull(DtPlano.Rows(0).Item("Price")) Then
                                    Price = ""
                                Else
                                    Price = "RP " & DtPlano.Rows(0).Item("Price").ToString.Split(".")(0)
                                End If
                            ElseIf TglAkhTemp >= Date.Now Then ' tgl_akh >= CURDATE()
                                If IsDBNull(DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI")) Then
                                    MaxBatasRetur = "" '0
                                Else
                                    MaxBatasRetur = DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI")
                                End If

                                If IsDBNull(DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI_S")) Then
                                    MaxBatasReturS = ""
                                Else
                                    MaxBatasReturS = DtPlano.Rows(0).Item("MAX_RET_TOKO2DCI_S")
                                End If

                                If IsDBNull(DtPlano.Rows(0).Item("Tanggal_Batas_Aman")) Then
                                    BatasPajang = "-"
                                Else
                                    BatasPajang = DtPlano.Rows(0).Item("Tanggal_Batas_Aman")
                                End If

                                If IsDBNull(DtPlano.Rows(0).Item("PROMOSI")) Then
                                    Price = ""
                                Else
                                    Price = "RP " & DtPlano.Rows(0).Item("PROMOSI").ToString.Split(".")(0)
                                End If
                            End If
                        End If
                    End If
                    If IsDBNull(DtPlano.Rows(0).Item("PRDCD")) Then
                        Result.Prdcd = "Tidak Terdaftar"
                    Else
                        Result.Prdcd = DtPlano.Rows(0).Item("PRDCD")

                    End If
                    'Console.Writeline(Result.Prdcd)
                    If DtPlano.Rows(0).Item("DESC2").Length > 40 Then
                        Result.Desc = DtPlano.Rows(0).Item("DESC2").Substring(0, 40)
                    Else
                        Result.Desc = DtPlano.Rows(0).Item("DESC2").ToString.PadRight(40, " ")
                    End If
                    'If MaxBatasRetur = 0 Then
                    '    Result.MaxRet = MaxBatasRetur
                    'Else
                    '    Result.MaxRet = MaxBatasReturS & "-" & MaxBatasRetur
                    'End If
                    Result.MaxRet = BatasPajang

                    Result.Price = Price
                Else
                    Result.Desc = "Tidak Ditemukan"
                End If

                'Insert ke Table CekPlanogram
                Dim DtCekPlano As New DataTable
                Dim TempNamaRak As String = ""
                Dim TempNoShelf As String = ""
                Dim Status As String = ""
                Dim JenisBarang As String = ""

                If DtPlano.Rows.Count > 0 Then
                    If IsDBNull(DtPlano.Rows(0).Item("NAMA_RAK")) Then
                        TempNamaRak = ""
                    Else
                        TempNamaRak = DtPlano.Rows(0).Item("NAMA_RAK").ToString.ToUpper
                    End If
                    If IsDBNull(DtPlano.Rows(0).Item("NOSHELF")) Then
                        TempNoShelf = ""
                    Else
                        TempNoShelf = DtPlano.Rows(0).Item("NOSHELF").ToString.ToUpper
                    End If

                    'MsgBox("nama rak " & TempNamaRak)
                    'MsgBox("no shelf " & TempNoShelf)


                    Mcom.CommandText = "SELECT PLU FROM " & tabel_name
                    Mcom.CommandText &= " WHERE PLU = '" & DtPlano.Rows(0).Item("PRDCD") & "'"
                    Mcom.CommandText &= " AND DATE(TGLSCAN) = CURDATE()"
                    Mcom.CommandText &= " AND `STATUS` <> 'I' AND NORAKINPUT = '" & DtPlano.Rows(0).Item("NORAK") & "'"
                    Mcom.CommandText &= " AND NOSHELFINPUT = '" & DtPlano.Rows(0).Item("NOSHELF") & "';"
                    IDM.Fungsi.TraceLog("Kueri " & Mcom.CommandText)

                    Dim sDap2 As New MySqlDataAdapter(Mcom)
                    'Console.Writeline(Mcom.CommandText)
                    sDap2.Fill(DtCekPlano)

                    If DtCekPlano.Rows.Count = 0 Then
                        If DtPlano.Rows(0).Item("QTY") = "0" Then
                            Status = "S"
                            JenisBarang = "SO"
                        Else
                            Status = "B"
                            JenisBarang = ""
                        End If


                        'Revisi 2020-04-30
                        'TGLSCAN yang seblumnya diset menggunakan SYSDATE(), diubah menggunakan DateTiem.Now dari VB
                        Dim TGLSCAN As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

                        Mcom.CommandText = "INSERT INTO " & tabel_name & " (PLU,TGLSCAN,NAMA,NOSHELFINPUT,NORAKINPUT,NAMARAK,NAMARAKINPUT,"
                        Mcom.CommandText &= "`STATUS`,NIK,NAMAUSER,KIRIKANAN,DIVISI,STOCK,JENISBARANG,MAXBATASRETUR,MAXBATASRETUR_S)"
                        Mcom.CommandText &= " VALUES ('" & DtPlano.Rows(0).Item("PRDCD") & "','" & TGLSCAN & "','','" & TempNoShelf & "',"
                        Mcom.CommandText &= " '" & DtPlano.Rows(0).Item("NORAK") & "','','" & TempNamaRak & "','" & Status & "',"
                        Mcom.CommandText &= " '" & User.ID & "', '" & User.Nama & "', '" & DtPlano.Rows(0).Item("KIRIKANAN") & "',"
                        Mcom.CommandText &= " '" & DtPlano.Rows(0).Item("DEPART") & "', '" & DtPlano.Rows(0).Item("QTY") & "',"
                        Mcom.CommandText &= " '" & JenisBarang & "', '" & MaxBatasRetur & "', '" & MaxBatasReturS & "');"
                        TraceLog("Kueri " & Mcom.CommandText)

                        'Console.Writeline(Mcom.CommandText)
                        Mcom.ExecuteNonQuery()
                    End If
                End If
skip:
                If Result.Desc.ToLower = "tolak" Then
                    Result.Desc = "tolak"
                ElseIf Result.Desc.ToLower = "Tidak Ditemukan" Then
                    Result.Desc = "Tidak Ditemukan"

                End If
            Catch ex As Exception
                Result.Desc = "Tidak Ditemukan"
                MsgBox(ex.Message & ex.StackTrace)

                IDM.Fungsi.TraceLog("Error" & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                Conn.Close()
            End Try
        End SyncLock

        Return Result
    End Function

    Public Function CekTablePJR() As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Mda As New MySqlDataAdapter("", Conn)
        Dim Rtn As New Boolean

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Try


                Mcom.CommandText = "Show tables like 'CekPJR'"
                If IsNothing(Mcom.ExecuteScalar) Then
                    Mcom.CommandText = "  Create table CekPJR ("
                    Mcom.CommandText &= " PLU varchar(8) not null, "
                    Mcom.CommandText &= " TglScan Date not null, "
                    Mcom.CommandText &= " Nama Varchar(50), "
                    Mcom.CommandText &= " NoShelf Integer, "
                    Mcom.CommandText &= " NoRak Integer, "
                    Mcom.CommandText &= " NamaRak Varchar(20), "
                    Mcom.CommandText &= " Status char(2), "
                    Mcom.CommandText &= " NoShelfInput Integer, "
                    Mcom.CommandText &= " NoRakInput Integer, "
                    Mcom.CommandText &= " KiriKanan Int(3), "
                    Mcom.CommandText &= " Divisi char(2), "
                    Mcom.CommandText &= " Stock decimal(12,0), "
                    Mcom.CommandText &= " JenisBarang Char(2), "
                    Mcom.CommandText &= " NIK Varchar(10) not null, "
                    Mcom.CommandText &= " NamaUser Varchar(50) not null, "
                    Mcom.CommandText &= " MaxBatasRetur VARCHAR(11) DEFAULT 0, "
                    'Mcom.CommandText &= " MaxBatasRetur int(11) DEFAULT NULL, "
                    Mcom.CommandText &= " MaxBatasRetur_S Varchar(4) DEFAULT 0, "
                    Mcom.CommandText &= " NamaRakInput Varchar(20), "
                    Mcom.CommandText &= " Primary Key(PLU,TglScan, NoRak, NoShelf)"
                    Mcom.CommandText &= " )"
                    Mcom.ExecuteNonQuery()
                Else
                    Mcom.CommandText = "Select column_type From Information_schema.Columns "
                    Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('CekPJR') "
                    Mcom.CommandText &= "And Column_Name='nama' "
                    If Mcom.ExecuteScalar & "" <> "varchar(50)" Then
                        Mcom.CommandText = "ALTER TABLE CekPJR modify COLUMN NAMA varchar(50) NOT NULL "
                        Mcom.ExecuteNonQuery()
                    End If

                    Mcom.CommandText = " Select count(*) From Information_schema.Columns "
                    Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('CekPJR') "
                    Mcom.CommandText &= " And Column_Name='MaxBatasRetur'"
                    If Mcom.ExecuteScalar = 0 Then
                        Mcom.CommandText = "Alter table CekPJR "
                        Mcom.CommandText &= "ADD COLUMN `MaxBatasRetur` VARCHAR(11) DEFAULT 0"
                        'Mcom.CommandText &= "ADD COLUMN `MaxBatasRetur` int(11) DEFAULT NULL"
                        Mcom.ExecuteNonQuery()
                    End If

                    Mcom.CommandText = " Select count(*) From Information_schema.Columns "
                    Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('CekPJR') "
                    Mcom.CommandText &= " And Column_Name='MaxBatasRetur_S'"
                    If Mcom.ExecuteScalar = 0 Then
                        Mcom.CommandText = "Alter table CekPJR "
                        Mcom.CommandText &= "ADD COLUMN `MaxBatasRetur_S` Varchar(4) DEFAULT 0"
                        Mcom.ExecuteNonQuery()
                    End If

                    Mcom.CommandText = " Select count(*) From Information_schema.Columns "
                    Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('CekPJR') "
                    Mcom.CommandText &= " And Column_Name='NamaRakInput'"
                    If Mcom.ExecuteScalar = 0 Then
                        Mcom.CommandText = "Alter table CekPJR "
                        Mcom.CommandText &= "ADD COLUMN `NamaRakInput` varchar(20)"
                        Mcom.ExecuteNonQuery()
                    End If

                    Mcom.CommandText = "Select column_type From Information_schema.Columns "
                    Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('CekPJR') "
                    Mcom.CommandText &= "And Column_Name='MaxBatasRetur' "
                    If Mcom.ExecuteScalar & "" <> "VARCHAR(11) DEFAULT 0" Then
                        'Mcom.CommandText = "ALTER TABLE CekPJR modify COLUMN MaxBatasRetur int(11) DEFAULT NULL "
                        Mcom.CommandText = "ALTER TABLE CekPJR modify COLUMN MaxBatasRetur VARCHAR(11) DEFAULT 0 "
                        Mcom.ExecuteNonQuery()
                    End If

                End If
            Catch ex As Exception
                TraceLog("Error " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)

            End Try
            Try
                'tabel tampungan untuk pjr
                Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`temp_jadwal_pjr` (
                                `RECID` VARCHAR(2) DEFAULT '', 
                                `NIK` VARCHAR(12), 
                                `NAMA` VARCHAR(99), 
                                `JABATAN` VARCHAR(50), 
                                `HARI` VARCHAR(30),
                                `TANGGAL` DATE,
                                `Kode_Modis` VARCHAR(20),
                                `MODIS` VARCHAR(99),
                                `Shelfing` VARCHAR(10),
                                `norak` VARCHAR(10),

                                `Addtime` DATE,
                                `Totalitem` Varchar(10),
                                `TotalEstimasi` Varchar(10),
                                `ITT` Varchar(10),
                                `FisikAda` Varchar(10),
                                `FisikTidakAda` Varchar(10),
                                `KetMinggu` Varchar(10),
                                `StatusApproval` Varchar(2),
                                `ITT_Adjust` Varchar(10)
                                ) ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci;"
                Mcom.ExecuteNonQuery()
            Catch ex As Exception
                TraceLog("Error " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)

            End Try
            Try
                Mcom.CommandText = "ALTER TABLE TEMP_JADWAL_PJR "
                Mcom.CommandText &= " ADD COLUMN `norak` VARCHAR(10) DEFAULT '' AFTER `Shelfing`"
                Mcom.ExecuteNonQuery()
            Catch ex As Exception

            End Try
            Try
                Mcom.CommandText = "Alter table temp_jadwal_pjr "
                Mcom.CommandText &= "ADD COLUMN `ITT_Adjust` varchar(10) DEFAULT ''"
                Mcom.ExecuteNonQuery()
            Catch ex As Exception

            End Try
            Try
                Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `barcode_cbr` (
                                  `RECID` char(3) DEFAULT NULL,
                                  `PLU` varchar(8) DEFAULT '',
                                  `BARCD` varchar(45) DEFAULT NULL,
                                  `KEMASAN` varchar(12) DEFAULT NULL,
                                  `QTY` decimal(13,0) DEFAULT NULL,
                                  `ADDID` varchar(135) DEFAULT NULL,
                                  `ADDTIME` datetime DEFAULT NULL
                                ) ENGINE=InnoDB DEFAULT CHARSET=latin1;"
                Mcom.ExecuteNonQuery()
            Catch ex As Exception
                TraceLog("Error " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)

            End Try
            'Try
            '    Mcom.CommandText = "ALTER TABLE `jadwal_penanggungjawabrak` DROP PRIMARY KEY"
            '    Mcom.ExecuteNonQuery()
            'Catch ex As Exception

            'End Try
            Rtn = True
        Catch ex As Exception
            Rtn = False
            TraceLog("Error " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        Finally
            Conn.Close()
        End Try

        Return Rtn
    End Function
    Public Function CekTableTINDAKLBTD() As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Mda As New MySqlDataAdapter("", Conn)
        Dim Rtn As New Boolean

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Mcom.CommandText = "Show tables like 'TINDAKLBTD'"
            If IsNothing(Mcom.ExecuteScalar) Then
                Mcom.CommandText = "Create table TINDAKLBTD ("
                Mcom.CommandText &= " PLU varchar(8) not null, "
                Mcom.CommandText &= " TglScan Date not null, "
                Mcom.CommandText &= " Nama Varchar(50), "
                Mcom.CommandText &= " NoShelf Integer, "
                Mcom.CommandText &= " NoRak Integer, "
                Mcom.CommandText &= " NamaRak Varchar(20), "
                Mcom.CommandText &= " Status char(2), "
                Mcom.CommandText &= " NoShelfInput Integer, "
                Mcom.CommandText &= " NoRakInput Integer, "
                Mcom.CommandText &= " KiriKanan Int(3), "
                Mcom.CommandText &= " Divisi char(2), "
                Mcom.CommandText &= " Stock decimal(12,0), "
                Mcom.CommandText &= " JenisBarang Char(2), "
                Mcom.CommandText &= " NIK Varchar(10) not null, "
                Mcom.CommandText &= " NamaUser Varchar(50) not null, "
                Mcom.CommandText &= " MaxBatasRetur varchar(11) DEFAULT 0, "
                Mcom.CommandText &= " MaxBatasRetur_S Varchar(4) DEFAULT NULL, "
                Mcom.CommandText &= " NamaRakInput Varchar(20), "
                Mcom.CommandText &= " Primary Key(PLU,TglScan, NoRak, NoShelf)"
                Mcom.CommandText &= " )"
                Mcom.ExecuteNonQuery()
            Else
                Mcom.CommandText = "Select column_type From Information_schema.Columns "
                Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('TINDAKLBTD') "
                Mcom.CommandText &= "And Column_Name='nama' "
                If Mcom.ExecuteScalar & "" <> "varchar(50)" Then
                    Mcom.CommandText = "ALTER TABLE TINDAKLBTD modify COLUMN NAMA varchar(50) NOT NULL "
                    Mcom.ExecuteNonQuery()
                End If

                Mcom.CommandText = " Select count(*) From Information_schema.Columns "
                Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('TINDAKLBTD') "
                Mcom.CommandText &= " And Column_Name='MaxBatasRetur'"
                If Mcom.ExecuteScalar = 0 Then
                    Mcom.CommandText = "Alter table TINDAKLBTD "
                    Mcom.CommandText &= "ADD COLUMN `MaxBatasRetur` varchar(11) DEFAULT 0"
                    Mcom.ExecuteNonQuery()
                End If

                Mcom.CommandText = " Select count(*) From Information_schema.Columns "
                Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('TINDAKLBTD') "
                Mcom.CommandText &= " And Column_Name='MaxBatasRetur_S'"
                If Mcom.ExecuteScalar = 0 Then
                    Mcom.CommandText = "Alter table TINDAKLBTD "
                    Mcom.CommandText &= "ADD COLUMN `MaxBatasRetur_S` Varchar(4) DEFAULT NULL"
                    Mcom.ExecuteNonQuery()
                End If

                Mcom.CommandText = " Select count(*) From Information_schema.Columns "
                Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('TINDAKLBTD') "
                Mcom.CommandText &= " And Column_Name='NamaRakInput'"
                If Mcom.ExecuteScalar = 0 Then
                    Mcom.CommandText = "Alter table TINDAKLBTD "
                    Mcom.CommandText &= "ADD COLUMN `NamaRakInput` varchar(20)"
                    Mcom.ExecuteNonQuery()
                End If

                Mcom.CommandText = "Select column_type From Information_schema.Columns "
                Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('TINDAKLBTD') "
                Mcom.CommandText &= "And Column_Name='MaxBatasRetur' "
                If Mcom.ExecuteScalar & "" <> "int(11) DEFAULT NULL" Then
                    Mcom.CommandText = "ALTER TABLE TINDAKLBTD modify COLUMN MaxBatasRetur varchar(11) DEFAULT 0 "
                    Mcom.ExecuteNonQuery()
                End If

            End If
            Rtn = True
        Catch ex As Exception
            Rtn = False
            TraceLog("Error " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        Finally
            Conn.Close()
        End Try

        Return Rtn
    End Function

    Public Function CekTableTINDAKLBTD_BAPJR() As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Mda As New MySqlDataAdapter("", Conn)
        Dim Rtn As New Boolean

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Mcom.CommandText = "Show tables like 'TINDAKLBTD_BAPJR'"
            If IsNothing(Mcom.ExecuteScalar) Then
                Mcom.CommandText = "Create table TINDAKLBTD_BAPJR ("
                Mcom.CommandText &= " PLU varchar(8) not null, "
                Mcom.CommandText &= " TglScan Date not null, "
                Mcom.CommandText &= " Nama Varchar(50), "
                Mcom.CommandText &= " NoShelf Integer, "
                Mcom.CommandText &= " NoRak Varchar(20), "
                Mcom.CommandText &= " NamaRak Varchar(20), "
                Mcom.CommandText &= " Status char(2), "
                Mcom.CommandText &= " NoShelfInput Integer, "
                Mcom.CommandText &= " NoRakInput Varchar(20), "
                Mcom.CommandText &= " KiriKanan Int(3), "
                Mcom.CommandText &= " Divisi char(2), "
                Mcom.CommandText &= " Stock decimal(12,0), "
                Mcom.CommandText &= " JenisBarang Char(2), "
                Mcom.CommandText &= " NIK Varchar(10) not null, "
                Mcom.CommandText &= " NamaUser Varchar(50) not null, "
                Mcom.CommandText &= " MaxBatasRetur varchar(11) DEFAULT 0, "
                Mcom.CommandText &= " MaxBatasRetur_S Varchar(4) DEFAULT NULL, "
                Mcom.CommandText &= " NamaRakInput Varchar(20), "
                Mcom.CommandText &= " Primary Key(PLU,TglScan, NoRak, NoShelf)"
                Mcom.CommandText &= " )"
                Mcom.ExecuteNonQuery()
            Else
                Mcom.CommandText = "Select column_type From Information_schema.Columns "
                Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('TINDAKLBTD_BAPJR') "
                Mcom.CommandText &= "And Column_Name='nama' "
                If Mcom.ExecuteScalar & "" <> "varchar(50)" Then
                    Mcom.CommandText = "ALTER TABLE TINDAKLBTD_BAPJR modify COLUMN NAMA varchar(50) NOT NULL "
                    Mcom.ExecuteNonQuery()
                End If

                Mcom.CommandText = " Select count(*) From Information_schema.Columns "
                Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('TINDAKLBTD_BAPJR') "
                Mcom.CommandText &= " And Column_Name='MaxBatasRetur'"
                If Mcom.ExecuteScalar = 0 Then
                    Mcom.CommandText = "Alter table TINDAKLBTD_BAPJR "
                    Mcom.CommandText &= "ADD COLUMN `MaxBatasRetur` varchar(11) DEFAULT 0"
                    Mcom.ExecuteNonQuery()
                End If

                Mcom.CommandText = " Select count(*) From Information_schema.Columns "
                Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('TINDAKLBTD_BAPJR') "
                Mcom.CommandText &= " And Column_Name='MaxBatasRetur_S'"
                If Mcom.ExecuteScalar = 0 Then
                    Mcom.CommandText = "Alter table TINDAKLBTD_BAPJR "
                    Mcom.CommandText &= "ADD COLUMN `MaxBatasRetur_S` Varchar(4) DEFAULT NULL"
                    Mcom.ExecuteNonQuery()
                End If

                Mcom.CommandText = " Select count(*) From Information_schema.Columns "
                Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('TINDAKLBTD_BAPJR') "
                Mcom.CommandText &= " And Column_Name='NamaRakInput'"
                If Mcom.ExecuteScalar = 0 Then
                    Mcom.CommandText = "Alter table TINDAKLBTD_BAPJR "
                    Mcom.CommandText &= "ADD COLUMN `NamaRakInput` varchar(20)"
                    Mcom.ExecuteNonQuery()
                End If

                Mcom.CommandText = "Select column_type From Information_schema.Columns "
                Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('TINDAKLBTD_BAPJR') "
                Mcom.CommandText &= "And Column_Name='MaxBatasRetur' "
                If Mcom.ExecuteScalar & "" <> "int(11) DEFAULT NULL" Then
                    Mcom.CommandText = "ALTER TABLE TINDAKLBTD_BAPJR modify COLUMN MaxBatasRetur varchar(11) DEFAULT 0 "
                    Mcom.ExecuteNonQuery()
                End If

            End If
            Rtn = True
        Catch ex As Exception
            Rtn = False
            TraceLog("Error " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        Finally
            Conn.Close()
        End Try

        Return Rtn
    End Function

    Public Sub bersih2JadwalPJR()
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Mda As New MySqlDataAdapter("", Conn)
        Dim Rtn As New Boolean

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Try
                Mcom.CommandText = "DELETE FROM JADWAL_PENANGGUNGJAWABRAK WHERE NIK = '' OR HARI = '' OR SHELFING = ''"
                Mcom.ExecuteNonQuery()
            Catch ex As Exception
            End Try
            Try
                Mcom.CommandText = "DELETE a FROM jadwal_penanggungjawabrak a JOIN (SELECT kode_modis, norak,shelfing FROM jadwal_penanggungjawabrak  GROUP BY kode_modis, norak,shelfing HAVING COUNT(*) > 1) b 
                                    ON a.kode_modis = b.kode_modis AND a.norak = b.norak AND a.shelfing = b.shelfing"
                Mcom.ExecuteNonQuery()

            Catch ex As Exception
            End Try
            Try
                'temp_ubah_jadwal_pjr
                Mcom.CommandText = "DELETE FROM JADWAL_PENANGGUNGJAWABRAK WHERE (KODE_MODIS,NORAK) IN (SELECT KODE_MODIS,NORAK FROM temp_ubah_jadwal_pjr)"
                Mcom.ExecuteNonQuery()

            Catch ex As Exception

            End Try

            Try
                Mcom.CommandText = "DELETE FROM JADWAL_PENANGGUNGJAWABRAK WHERE KODE_MODIS IN (SELECT KODEMODIS FROM RAK WHERE FLAGPROD NOT LIKE '%FJP=Y%')"
                Mcom.ExecuteNonQuery()
            Catch ex As Exception
            End Try
        Catch ex As Exception
            TraceLog("Error " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        Finally
            Conn.Close()
        End Try
    End Sub
    Public Sub cekUlangJadwalPJR()
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Mda As New MySqlDataAdapter("", Conn)
        Dim Rtn As New Boolean
        Dim sql As String = ""
        Dim dt As New DataTable
        Dim shelfing_awal As String = ""
        Dim shelfing_akhir As String = ""
        Dim norak As String = ""
        Dim modis As String = ""
        Dim hasil_1 As String = ""
        Dim hasil_2 As String = ""
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Try
                Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`temp_ubah_jadwal_pjr` ( 
                                `Kode_Modis` VARCHAR(99),
                                `NORAK` VARCHAR(20)
                                ) ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci;"
                Mcom.ExecuteNonQuery()
                Mcom.CommandText = "DELETE FROM temp_ubah_jadwal_pjr"
                Mcom.ExecuteNonQuery()
            Catch ex As Exception

            End Try


            Try
                'hapus jadwal yang double karna norak sama tapi beda shelfing
                Mcom.CommandText = "DELETE a FROM jadwal_penanggungjawabrak a JOIN 
                                    (SELECT kode_modis, norak FROM jadwal_penanggungjawabrak  GROUP BY kode_modis, norak HAVING COUNT(*) > 1) b 
                                    ON a.kode_modis = b.kode_modis AND a.norak = b.norak;"
                Mcom.ExecuteNonQuery()

                Mda.SelectCommand.CommandText = "SELECT KODE_MODIS,NORAK, SHELFING FROM JADWAL_PENANGGUNGJAWABRAK"
                dt.Clear()
                Mda.Fill(dt)

                For i As Integer = 0 To dt.Rows.Count - 1
                    'MODIS A
                    norak = dt.Rows(i)("NORAK").ToString
                    modis = dt.Rows(i)("KODE_MODIS").ToString
                    shelfing_awal = dt.Rows(i)("SHELFING").ToString.Split("-")(0) '1
                    shelfing_akhir = dt.Rows(i)("SHELFING").ToString.Split("-")(1) '6
                    'cek di bracket untuk modis norak dan shelfing
                    Mcom.CommandText = "SELECT COUNT(1) FROM BRACKET WHERE MODISP = '" & modis & "' 
                                        AND NO_RAK = '" & norak & "' 
                                        AND SHELFING_AWAL = '" & shelfing_awal & "' AND SHELFING_AKHIR = '" & shelfing_akhir & "'"
                    'jika tidak ketemu
                    If Mcom.ExecuteScalar = 0 Then
                        'cek di rak untuk modis dan shelfing

                        'Mcom.CommandText = "SELECT MIN(NOSHELF)  FROM RAK WHERE KODEMODIS = '" & modis & "'"
                        'TraceLog(Mcom.CommandText)

                        'If IsDBNull(Mcom.ExecuteScalar) Then
                        '    hasil_1 = ""
                        'Else
                        '    hasil_1 = Mcom.ExecuteScalar

                        'End If

                        'Mcom.CommandText = "SELECT MAX(NOSHELF) FROM RAK WHERE KODEMODIS = '" & modis & "'"
                        'TraceLog(Mcom.CommandText)

                        'If IsDBNull(Mcom.ExecuteScalar) Then

                        '    hasil_2 = ""
                        'Else
                        '    hasil_2 = Mcom.ExecuteScalar

                        'End If
                        'If hasil_1 = shelfing_awal And hasil_2 = shelfing_akhir Then
                        'Else
                        If norak = "1" And shelfing_awal = "1" And shelfing_akhir = "1" Then
                        Else
                            Mcom.CommandText = "SELECT  SHELFING_AWAL  FROM BRACKET WHERE MODISP = '" & modis & "'
                                                    AND NO_RAK = '" & norak & "'"

                            If IsDBNull(Mcom.ExecuteScalar) Then

                                hasil_1 = ""
                            Else
                                hasil_1 = Mcom.ExecuteScalar

                            End If
                            Mcom.CommandText = "SELECT  SHELFING_AKHIR   FROM BRACKET WHERE MODISP = '" & modis & "'
                                                    AND NO_RAK = '" & norak & "'"
                            If IsDBNull(Mcom.ExecuteScalar) Then

                                hasil_2 = ""
                            Else
                                hasil_2 = Mcom.ExecuteScalar

                            End If

                            If hasil_1 <> "" And hasil_2 <> "" Then
                                Mcom.CommandText = "UPDATE JADWAL_PENANGGUNGJAWABRAK SET shelfing = '" & hasil_1 & "-" & hasil_2 & "'
                                                        WHERE KODE_MODIS = '" & modis & "' AND NORAK = '" & norak & "'"
                                'Mcom.CommandText = "INSERT IGNORE INTO temp_ubah_jadwal_pjr VALUES ('" & modis & "', '" & norak & "')"
                                Mcom.ExecuteNonQuery()
                            Else
                                Mcom.CommandText = "SELECT MIN(NOSHELF) FROM RAK WHERE KODEMODIS = '" & modis & "'"
                                If IsDBNull(Mcom.ExecuteScalar) Then

                                    hasil_1 = "1"
                                Else
                                    hasil_1 = Mcom.ExecuteScalar

                                End If

                                Mcom.CommandText = "SELECT MAX(NOSHELF) FROM RAK WHERE KODEMODIS = '" & modis & "'"
                                If IsDBNull(Mcom.ExecuteScalar) Then

                                    hasil_2 = "1"
                                Else
                                    hasil_2 = Mcom.ExecuteScalar

                                End If

                                If hasil_1 = "" Then
                                    hasil_1 = "1"
                                End If
                                If hasil_2 = "" Then
                                    hasil_2 = "1"
                                End If
                                Mcom.CommandText = "SELECT COUNT(1) FROM JADWAL_PENANGGUNGJAWABRAK WHERE KODE_MODIS = '" & modis & "' AND NORAK = '1'
                                                    AND SHELFING = '" & hasil_1 & "-" & hasil_2 & "'"
                                TraceLog(Mcom.CommandText)
                                If Mcom.ExecuteScalar() = 0 Then
                                    '    Mcom.CommandText = "DELETE FROM JADWAL_PENANGGUNGJAWABRAK 
                                    '                        WHERE KODE_MODIS = '" & modis & "' AND NORAK = '" & norak & "'"

                                    '    Mcom.ExecuteNonQuery()
                                    'Else
                                    'Mcom.CommandText = "UPDATE JADWAL_PENANGGUNGJAWABRAK SET shelfing = '" & hasil_1 & "-" & hasil_2 & "', norak = '1'
                                    '                    WHERE KODE_MODIS = '" & modis & "' AND NORAK = '" & norak & "'"

                                    'Mcom.ExecuteNonQuery()
                                    Mcom.CommandText = "INSERT IGNORE INTO temp_ubah_jadwal_pjr VALUES ('" & modis & "', '" & norak & "')"
                                    Mcom.ExecuteNonQuery()
                                End If


                            End If
                        End If
                    End If
                    'End If

                Next

            Catch ex As Exception
                TraceLog("Error " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)

            End Try

        Catch ex As Exception
            TraceLog("Error " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        Finally
            Conn.Close()
        End Try
    End Sub

    Public Sub buatTabelItemSO()
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Mda As New MySqlDataAdapter("", Conn)
        Dim Rtn As New Boolean
        Dim sql As String = ""
        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
        End If

        Try
            Mcom.CommandText = "SHOW TABLES LIKE 'ITEMSO_PJR'"
            If IsNothing(Mcom.ExecuteScalar) Then
                sql = " CREATE TABLE  `itemso_PJR` (
                       `RECID` varchar(1) default NULL,
                       `DIV` varchar(2) default NULL,
                       `PRDCD` varchar(8) NOT NULL default '',

                       `DESC` VARCHAR(50),
                        SINGKAT VARCHAR(30),
                        BARCODE VARCHAR(15),
                        BARCODE2 VARCHAR(15),
        
                        FRAC LONG,
                        UNIT VARCHAR(4)

                       ,PTAG varchar(1)
                       ,CAT_COD varchar(6)
                       ,BKP varchar(1)
                       ,SUB_BKP varchar(1)
                       ,CTGR varchar(2)
                       ,KEMASAN varchar(3)
                       ,ACOST double
                       ,LCOST double
                       ,RCOST double
                       ,PRICE double,

                       `TIPERAK` varchar(1) default NULL,
                       `NORAK` int(11) default NULL,
                       `NOSHELF` int(11) default NULL,
                       `KIRIKANAN` int(11) default NULL,
                       `SOTYPE` varchar(1) default NULL,
                       `TANGGAL` datetime default NULL,
                       `NIK` VARCHAR(20) default NULL,
                       `NAMA_RAK` VARCHAR(30) default NULL,

                        PRIMARY KEY  (`PRDCD`)
                        ) ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='Data Buat SO Item'"
                Mcom.CommandText = sql
                Mcom.ExecuteNonQuery()
            Else
                sql = "DELETE from ItemSO_PJR where TANGGAL<>CURDATE()"
                Mcom.CommandText = sql
                Mcom.ExecuteNonQuery()
                Mcom.CommandText = "Select Count(*) From ITEMSO_PJR WHERE Tanggal=Curdate()"
                If Mcom.ExecuteScalar < 1 Then
                    Mcom.CommandText = "DROP TABLE IF EXISTS `itemso_PJR` "
                    Mcom.ExecuteNonQuery()

                    sql = " CREATE TABLE  `itemso_PJR` ("
                    sql &= "  `RECID` varchar(1) default NULL,"
                    sql &= "  `DIV` varchar(2) default NULL,"
                    sql &= "  `PRDCD` varchar(8) NOT NULL default '',"

                    sql &= "`DESC` VARCHAR(50),"
                    sql &= "SINGKAT VARCHAR(30),"
                    sql &= "BARCODE VARCHAR(15),"
                    sql &= "BARCODE2 VARCHAR(15),"

                    sql &= "FRAC LONG,"
                    sql &= "UNIT VARCHAR(4)"

                    sql &= ",PTAG varchar(1)"
                    sql &= ",CAT_COD varchar(6)"
                    sql &= ",BKP varchar(1)"
                    sql &= ",SUB_BKP varchar(1)"
                    sql &= ",CTGR varchar(2)"
                    sql &= ",KEMASAN varchar(3)"
                    sql &= ",ACOST double"
                    sql &= ",LCOST double"
                    sql &= ",RCOST double"
                    sql &= ",PRICE double,"

                    sql &= "  `TIPERAK` varchar(1) default NULL,"
                    sql &= "  `NORAK` int(11) default NULL,"
                    sql &= "  `NOSHELF` int(11) default NULL,"
                    sql &= "  `KIRIKANAN` int(11) default NULL,"
                    sql &= "  `SOTYPE` varchar(1) default NULL,"
                    sql &= "  `TANGGAL` datetime default NULL,"
                    sql &= "  `NIK` VARCHAR(20) default NULL,"
                    sql &= "  `NAMA_RAK` VARCHAR(30) default NULL,"
                    'sql &= "  `TANGGAL_JADWAL` datetime default NULL,"
                    sql &= "  PRIMARY KEY  (`PRDCD`)"
                    sql &= ") ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='Data Buat SO Item'"
                    Mcom.CommandText = sql
                    Mcom.ExecuteNonQuery()
                End If
            End If



        Catch ex As Exception
            ShowError("Error buatTabelItemSO", ex)
        Finally
            Conn.Close()
        End Try
    End Sub

    'Memo 405/CPS/23
    'Tabel tampungan utk akumulasi BA AS, dari ITT
    Public Sub buatTabelItemSO_BA_AS()
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Mda As New MySqlDataAdapter("", Conn)
        Dim Rtn As New Boolean
        Dim sql As String = ""
        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
        End If

        Try
            sql = " CREATE TABLE IF NOT EXISTS  `ITEMSO_PJR_BA_AS` (
                       `RECID` varchar(1) default NULL,
                       `PRDCD` varchar(8) NOT NULL default '',
                       `DESC` VARCHAR(50),
                        SINGKAT VARCHAR(30),
                        BARCODE VARCHAR(15),
                        BARCODE2 VARCHAR(15),
                       KEMASAN varchar(3),
                       `TIPERAK` varchar(1) default NULL,
                       `NORAK` int(11) default NULL,
                       `NOSHELF` int(11) default NULL,
                       `KIRIKANAN` int(11) default NULL,
                       `TANGGAL` date NOT NULL default '0000-00-00',
                       `HARI` VARCHAR(20) default NULL,
                       `NIK` VARCHAR(20) default NULL,
                       `KODE_MODIS` VARCHAR(30) NOT NULL default '',
                        PRIMARY KEY  (`PRDCD`, `TANGGAL`, `KODE_MODIS`)
                        ) ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='Data Buat SO Item BA AS PJR'"
            Mcom.CommandText = sql
            Mcom.ExecuteNonQuery()
            Try
                Mcom.CommandText = "ALTER TABLE `itemso_pjr_ba_as` CHANGE `TANGGAL` `TANGGAL` DATE NOT NULL, 
                            CHANGE `KODE_MODIS` `KODE_MODIS` VARCHAR(30) CHARSET latin1 COLLATE latin1_swedish_ci NOT NULL,
                            DROP PRIMARY KEY, ADD PRIMARY KEY (`PRDCD`, `TANGGAL`, `KODE_MODIS`)"
                Mcom.ExecuteNonQuery()
            Catch ex As Exception
            End Try


        Catch ex As Exception
            TraceLog("Error buatTabelItemSO " & ex.Message & ex.StackTrace)
            ShowError("Error buatTabelItemSO", ex)
        Finally
            Conn.Close()
        End Try
    End Sub

    Public Sub buatTabelBA()
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Mda As New MySqlDataAdapter("", Conn)
        Dim Rtn As New Boolean
        Dim sql As String = ""
        Dim cFileSO As String = "BS_PJR_" & Format(Now, "yyMMdd") & FormMain.Toko.Kode.Substring(0, 1) & ""
        'Console.Writeline(cFileSO)
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            sql = "CREATE TABLE IF NOT EXISTS " & cFileSO & "("
            sql &= "RECID VARCHAR(1),"
            sql &= "TIPERAK VARCHAR(1),"
            sql &= "NORAK LONG,"
            sql &= "NOSHELF LONG,"
            sql &= "KIRIKANAN LONG,"
            sql &= "`DIV` VARCHAR(2),"
            sql &= "PRDCD VARCHAR(8),"
            sql &= "`DESC` VARCHAR(50),"
            sql &= "SINGKAT VARCHAR(30),"
            sql &= "BARCODE VARCHAR(15),"
            sql &= "BARCODE2 VARCHAR(15),"

            sql &= "FRAC LONG,"
            sql &= "UNIT VARCHAR(4),"

            sql &= "PTAG varchar(1)"
            sql &= ",CAT_COD varchar(6)"
            sql &= ",BKP LOGICAL"
            sql &= ",SUB_BKP varchar(1)"
            sql &= ",CTGR varchar(2)"
            sql &= ",KEMASAN varchar(3)"
            sql &= ",ACOST double"
            sql &= ",LCOST double"
            sql &= ",RCOST double"
            sql &= ",PRICE double,"

            sql &= "TTL LONG,"
            sql &= "TTL1 LONG,"
            sql &= "TTL2 LONG,"
            sql &= "COM LONG,"
            sql &= "HPP LONG,"

            sql &= "SOID VARCHAR(1),"
            sql &= "EDIT VARCHAR(1),"
            sql &= "SOTYPE VARCHAR(1),"
            sql &= "SOTGL DateTime,"
            sql &= "SOTIME VARCHAR(8),"
            sql &= "ADJDT DateTime,"
            sql &= "ADJTIME VARCHAR(8),"
            sql &= "NIK VARCHAR(20),"
            sql &= "NAMA_RAK VARCHAR(30),"
            'sql &= "TANGGAL_JADWAL DATE,"

            'sql &= "RAKSEWA VARCHAR(1),"
            sql &= "DRAFT LOGICAL,"
            sql &= "DCP LOGICAL,"
            sql &= "CTK LOGICAL"

            sql = sql.Replace("LONG", "decimal(12,0) default 0 ")
            sql = sql.Replace("LOGICAL", "char(1) default '' ")
            sql = sql.Replace("VARCHAR(1)", "VARCHAR(1) default '' ")
            Mcom.CommandText = sql
            Mcom.CommandText &= ",PRIMARY KEY  (SOTYPE,`PRDCD`,TIPERAK,`NORAK`,`NOSHELF`,`KIRIKANAN`) )"
            TraceLog("buatTabelBA_01 : " & Mcom.CommandText)


            Mcom.ExecuteNonQuery()


        Catch ex As Exception
            TraceLog("Error " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        Finally
            Conn.Close()
        End Try

    End Sub

    Public Sub loadDataBAPJR(ByVal nik As String, ByVal namarak As String, ByVal norak As String)
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Mda As New MySqlDataAdapter("", Conn)
        Dim Rtn As New Boolean
        Dim sql As String = ""
        Dim dt As New DataTable

        Dim cFileSO As String = "BS_PJR_" & Format(Now, "yyMMdd") & FormMain.Toko.Kode.Substring(0, 1) & ""
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            'INSERT ITEMSO
            sql = "DELETE from ItemSO_PJR "
            Mcom.CommandText = sql
            Mcom.ExecuteNonQuery()

            sql = "INSERT INTO ITEMSO_PJR 
                   Select '' AS RECID,P.DEPART AS `DIV`,P.PRDCD

                   ,REPLACE(P.DESC2, '""', '''') AS `DESC`
                   ,REPLACE(P.SINGKATAN, '""', '''') AS SINGKAT
                   ,IFNULL(B.BARCODE,'') AS Barcode,IFNULL(B.BARCODE2,'') as BARCODE2

                   ,P.FRAC
                   ,P.UNIT

                   ,P.PTAG
                   ,P.CAT_COD
                   ,P.BKP
                   ,P.SUB_BKP
                   ,P.CTGR
                   ,P.KEMASAN
                   ,P.ACOST
                   ,P.LCOST
                   ,P.RCOST
                   ,P.PRICE
                   ,IF(COUNT(*)<=1,R.TIPERAK,null) AS TIPERAK,
                   c.NORAK AS NORAK,
                   If (COUNT(*) <= 1, R.NOSHELF, NULL)  As NOSHELF,
                   IF(COUNT(*)<=1,R.KIRIKANAN,NULL) AS KIRIKANAN,
                   '4' AS SOTYPE,CURDATE() AS TANGGAL ,
                   '" & nik & "' AS NIK,'" & namarak & "' AS NAMA_RAK 
                   FROM PRODMAST P LEFT JOIN RAK R ON P.PRDCD=R.PRDCD 
                   LEFT JOIN (SELECT PLU,MIN(BARCD) AS BARCODE,IF(MAX(BARCD)<>MIN(BARCD),MAX(BARCD),'') AS BARCODE2 FROM BARCODE GROUP BY PLU ) B ON B.PLU=R.PRDCD 
                   LEFT JOIN (SELECT KODE_MODIS,NORAK FROM jadwal_penanggungjawabrak WHERE NIK = '" & nik & "' AND kode_modis = '" & namarak & "' AND norak = '" & norak & "') c ON c.kode_modis = r.kodemodis
                   WHERE P.PRDCD IN (SELECT PLU FROM TINDAKLBTD WHERE `JENISBARANG` = 'TT' AND NIK = '" & nik & "' AND NamaRakInput = '" & namarak & "' and norak = '" & norak & "' AND TGLSCAN = CURDATE()) 
                   GROUP BY R.PRDCD"
            Mcom.CommandText = sql
            TraceLog("LoadDataBAPJR_01 : " & sql)
            Mcom.ExecuteScalar()



            sql = "Select '1' AS SOTYPE,P.PRDCD,P.DESC2,P.SINGKATAN,P.DEPART AS `DIV`,B.BARCODE,B.BARCODE2"
            sql &= ",P.FRAC"
            sql &= ",P.UNIT"
            sql &= ",P.PTAG"
            sql &= ",P.CAT_COD"
            sql &= ",if(ifnull(P.BKP,'N')='Y',1,0) as BKP"
            sql &= ",P.SUB_BKP"
            sql &= ",P.CTGR"
            sql &= ",P.KEMASAN"
            sql &= ",P.ACOST"
            sql &= ",P.LCOST"
            sql &= ",P.RCOST"
            sql &= ",P.PRICE"
            sql &= ",IF(P.ACOST>0,P.ACOST,IF(P.LCOST>0,P.LCOST,IF(P.RCOST>0,P.RCOST,IF(S.LCOST>0,S.LCOST,1)))) AS HPP1"
            'sql &= ",P.Price AS HPP1"
            sql &= ",IFNULL(S.BEGBAL,0)+(IFNULL(S.TRFIN,0)-IFNULL(S.TRFOUT,0))+(IFNULL(S.RETUR,0)-IFNULL(S.SALES,0))+IFNULL(S.ADJ,0)+IFNULL(S.BS,0)+IFNULL(S.BA,0) AS COM1"
            sql &= ",R.TIPERAK,jp.NORAK,R.NOSHELF,R.KIRIKANAN FROM ITEMSO_PJR SO "
            sql &= "LEFT JOIN PRODMAST P ON SO.PRDCD=P.PRDCD "
            sql &= "LEFT JOIN (SELECT PLU,MIN(BARCD) AS BARCODE,IF(MAX(BARCD)<>MIN(BARCD),MAX(BARCD),'') AS BARCODE2 FROM BARCODE GROUP BY PLU ) B ON B.PLU=SO.PRDCD "
            sql &= "LEFT JOIN (SELECT PRDCD,TIPERAK,NORAK,NOSHELF,KIRIKANAN,kodemodis FROM RAK WHERE KODETOKO='" & IDM.InfoToko.Get_TipeToko & "' AND KODEMODIS = '" & namarak & "' GROUP BY PRDCD) R ON R.PRDCD=P.PRDCD "
            sql &= "LEFT JOIN STMAST S ON S.PRDCD=P.PRDCD "
            sql &= "LEFT JOIN PROTECT PR ON PR.PRDCD=P.PRDCD "
            sql &= "LEFT JOIN PTAG_NR NR ON NR.PRDCD=P.PRDCD "
            sql &= "LEFT JOIN (SELECT norak,kode_modis FROM jadwal_penanggungjawabrak WHERE NIK = '" & nik & "' AND kode_modis = '" & namarak & "' AND norak = '" & norak & "')  jp ON r.kodemodis=jp.kode_modis "
            sql &= "Where 1=1 "
            sql &= "AND (PR.BPB NOT IN('X','Y','Z') OR PR.BPB IS NULL) "
            sql &= "AND (P.RECID<>'1' OR P.RECID IS NULL) "
            sql &= "AND (P.NONSO<>'Y' or P.NONSO is null) "
            sql &= "AND (NR.PRDCD IS NULL)"
            sql &= "AND (SO.PRDCD IN (SELECT PLU FROM TINDAKLBTD WHERE `JENISBARANG` = 'TT' AND NIK = '" & nik & "' AND NamaRakInput = '" & namarak & "' and norak = '" & norak & "'))"
            TraceLog("LoadDataBAPJR_02 : " & sql)

            Mda.SelectCommand.CommandText = sql
            dt.Clear()
            Mda.Fill(dt)

            For i As Integer = 0 To dt.Rows.Count - 1
                Application.DoEvents()

                sql = "SELECT COUNT(*) FROM " & cFileSO
                sql &= " Where SOTYPE='" & dt.Rows(i)("SOTYPE") & "' "
                sql &= "AND PRDCD='" & dt.Rows(i)("PRDCD") & "'"
                Mcom.CommandText = sql
                If Mcom.ExecuteScalar = 0 Then
                    sql = "INSERT INTO " & cFileSO & "("
                    sql &= "SOTYPE"
                    sql &= ",TIPERAK"
                    sql &= ",NORAK"
                    sql &= ",NOSHELF"
                    sql &= ",KIRIKANAN"
                    sql &= ",`DIV`"
                    sql &= ",PRDCD"
                    sql &= ",`DESC`"
                    sql &= ",SINGKAT"
                    sql &= ",BARCODE"
                    sql &= ",BARCODE2"

                    sql &= ",FRAC"
                    sql &= ",UNIT"

                    sql &= ",PTAG"
                    sql &= ",CAT_COD"
                    sql &= ",BKP"
                    sql &= ",SUB_BKP"
                    sql &= ",CTGR"
                    sql &= ",KEMASAN"
                    sql &= ",ACOST"
                    sql &= ",LCOST"
                    sql &= ",RCOST"
                    sql &= ",PRICE"

                    sql &= ",TTL"
                    sql &= ",TTL1"
                    sql &= ",TTL2"

                    sql &= ",COM"
                    sql &= ",HPP"
                    sql &= ",SOTGL"
                    sql &= ",SOTIME"
                    sql &= ",NIK"
                    sql &= ",NAMA_RAK"
                    'sql &= ",TANGGAL_JADWAL"
                    sql &= ")VALUES("
                    sql &= " '" & dt.Rows(i)("SOTYPE") & "'"
                    sql &= ",'" & dt.Rows(i)("TIPERAK") & "'"
                    sql &= ",0" & dt.Rows(i)("NORAK") & ""
                    sql &= ",0" & dt.Rows(i)("NOSHELF") & ""
                    sql &= ",0" & dt.Rows(i)("KIRIKANAN") & ""
                    sql &= ",'" & dt.Rows(i)("DIV") & "'"
                    sql &= ",'" & dt.Rows(i)("PRDCD") & "'"
                    sql &= ",'" & CType(dt.Rows(i)("DESC2") & "", String).Replace("'", "''") & "'"
                    sql &= ",'" & CType(dt.Rows(i)("SINGKATAN") & "", String).Replace("'", "''") & "'"
                    sql &= ",'" & dt.Rows(i)("BARCODE") & "'"
                    sql &= ",'" & dt.Rows(i)("BARCODE2") & "'"

                    sql &= ",0" & dt.Rows(i)("FRAC") & ""
                    sql &= ",'" & dt.Rows(i)("UNIT") & "'"

                    sql &= ",'" & dt.Rows(i)("PTAG") & "'"
                    sql &= ",'" & dt.Rows(i)("CAT_COD") & "'"
                    sql &= "," & dt.Rows(i)("BKP") & ""
                    sql &= ",'" & dt.Rows(i)("SUB_BKP") & "'"
                    sql &= ",'" & dt.Rows(i)("CTGR") & "'"
                    sql &= ",'" & dt.Rows(i)("KEMASAN") & "'"
                    sql &= ",0" & dt.Rows(i)("ACOST") & ""
                    sql &= ",0" & dt.Rows(i)("LCOST") & ""
                    sql &= ",0" & dt.Rows(i)("RCOST") & ""
                    sql &= ",0" & dt.Rows(i)("PRICE") & ""
                    sql &= ",0,0,0"
                    sql &= ", " & dt.Rows(i)("COM1") & ""
                    sql &= ", " & dt.Rows(i)("HPP1") & ""
                    sql &= ",'" & Format(Now, "yyyy-MM-dd") & "'"
                    sql &= ",'" & Format(Now, "HH:mm:ss") & "'"
                    sql &= ",'" & nik & "'"
                    sql &= ",'" & namarak & "'"
                    'sql &= ",'" & Format(Now, "HH:mm:ss") & "'"
                    sql &= ")"
                    Mcom.CommandText = sql
                    Mcom.ExecuteNonQuery()
                Else
                    sql = "UPDATE " & cFileSO & " SET "
                    sql &= "TTL=0"
                    sql &= ",TTL1=0"
                    sql &= ",SOTGL='" & Format(Now, "yyyy-MM-dd") & "'"
                    sql &= ",SOTIME='" & Format(Now, "HH:mm:ss") & "'"
                    sql &= ",HPP=" & dt.Rows(i)("HPP1") & ""
                    sql &= " Where SOTYPE='" & dt.Rows(i)("SOTYPE") & "' "
                    sql &= "AND TIPERAK='" & dt.Rows(i)("TIPERAK") & "' "
                    sql &= "AND NORAK=" & dt.Rows(i)("NORAK") & " "
                    sql &= "AND NOSHELF=" & dt.Rows(i)("NOSHELF") & " "
                    sql &= "AND KIRIKANAN=" & dt.Rows(i)("KIRIKANAN") & " "
                    sql &= "AND PRDCD='" & dt.Rows(i)("PRDCD") & "'"
                    sql &= "AND NIK='" & nik & "'"
                    sql &= "AND NAMA_RAK='" & namarak & "'"
                    Mcom.CommandText = sql

                    Mcom.ExecuteNonQuery()
                End If

Lewati:
            Next

        Catch ex As Exception
            TraceLog("Error " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        Finally
            Conn.Close()
        End Try

    End Sub


    Public Sub loadDataBAPJR_AS()
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Mda As New MySqlDataAdapter("", Conn)
        Dim Rtn As New Boolean
        Dim sql As String = ""
        Dim dt As New DataTable

        Dim cFileSO As String = "BS_PJR_" & Format(Now, "yyMMdd") & FormMain.Toko.Kode.Substring(0, 1) & ""
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            'INSERT ITEMSO
            'sql = "DELETE from ItemSO_PJR "
            'Mcom.CommandText = sql
            'Mcom.ExecuteNonQuery()

            'sql = "INSERT INTO ITEMSO_PJR 
            '       Select '' AS RECID,P.DEPART AS `DIV`,P.PRDCD

            '       ,REPLACE(P.DESC2, '""', '''') AS `DESC`
            '       ,REPLACE(P.SINGKATAN, '""', '''') AS SINGKAT
            '       ,IFNULL(B.BARCODE,'') AS Barcode,IFNULL(B.BARCODE2,'') as BARCODE2

            '       ,P.FRAC
            '       ,P.UNIT

            '       ,P.PTAG
            '       ,P.CAT_COD
            '       ,P.BKP
            '       ,P.SUB_BKP
            '       ,P.CTGR
            '       ,P.KEMASAN
            '       ,P.ACOST
            '       ,P.LCOST
            '       ,P.RCOST
            '       ,P.PRICE
            '       ,IF(COUNT(*)<=1,R.TIPERAK,null) AS TIPERAK,
            '       c.NORAK AS NORAK,
            '       If (COUNT(*) <= 1, R.NOSHELF, NULL)  As NOSHELF,
            '       IF(COUNT(*)<=1,R.KIRIKANAN,NULL) AS KIRIKANAN,
            '       '4' AS SOTYPE,CURDATE() AS TANGGAL ,
            '       '" & nik & "' AS NIK,'" & namarak & "' AS NAMA_RAK 
            '       FROM PRODMAST P LEFT JOIN RAK R ON P.PRDCD=R.PRDCD 
            '       LEFT JOIN (SELECT PLU,MIN(BARCD) AS BARCODE,IF(MAX(BARCD)<>MIN(BARCD),MAX(BARCD),'') AS BARCODE2 FROM BARCODE GROUP BY PLU ) B ON B.PLU=R.PRDCD 
            '       LEFT JOIN (SELECT KODE_MODIS,NORAK FROM jadwal_penanggungjawabrak WHERE NIK = '" & nik & "' AND kode_modis = '" & namarak & "' AND norak = '" & norak & "') c ON c.kode_modis = r.kodemodis
            '       WHERE P.PRDCD IN (SELECT PLU FROM TINDAKLBTD WHERE `JENISBARANG` = 'TT' AND NIK = '" & nik & "' AND NamaRakInput = '" & namarak & "' and norak = '" & norak & "' AND TGLSCAN = CURDATE()) 
            '       GROUP BY R.PRDCD"
            'Mcom.CommandText = sql
            'TraceLog("LoadDataBAPJR_01 : " & sql)
            'Mcom.ExecuteScalar()

            sql = "Select '1' AS SOTYPE,P.PRDCD,P.DESC2,P.SINGKATAN,P.DEPART AS `DIV`,B.BARCODE,B.BARCODE2"
            sql &= ",P.FRAC"
            sql &= ",P.UNIT"
            sql &= ",P.PTAG"
            sql &= ",P.CAT_COD"
            sql &= ",if(ifnull(P.BKP,'N')='Y',1,0) as BKP"
            sql &= ",P.SUB_BKP"
            sql &= ",P.CTGR"
            sql &= ",P.KEMASAN"
            sql &= ",P.ACOST"
            sql &= ",P.LCOST"
            sql &= ",P.RCOST"
            sql &= ",P.PRICE"
            sql &= ",IF(P.ACOST>0,P.ACOST,IF(P.LCOST>0,P.LCOST,IF(P.RCOST>0,P.RCOST,IF(S.LCOST>0,S.LCOST,1)))) AS HPP1"
            'sql &= ",P.Price AS HPP1"
            sql &= ",IFNULL(S.BEGBAL,0)+(IFNULL(S.TRFIN,0)-IFNULL(S.TRFOUT,0))+(IFNULL(S.RETUR,0)-IFNULL(S.SALES,0))+IFNULL(S.ADJ,0)+IFNULL(S.BS,0)+IFNULL(S.BA,0) AS COM1"
            sql &= ",R.TIPERAK,SO.NORAK,R.NOSHELF,R.KIRIKANAN,SO.NIK,SO.KODE_MODIS FROM ITEMSO_PJR_BA_AS SO "
            sql &= "LEFT JOIN PRODMAST P ON SO.PRDCD=P.PRDCD "
            sql &= "LEFT JOIN (SELECT PLU,MIN(BARCD) AS BARCODE,IF(MAX(BARCD)<>MIN(BARCD),MAX(BARCD),'') AS BARCODE2 FROM BARCODE GROUP BY PLU ) B ON B.PLU=SO.PRDCD "
            sql &= "LEFT JOIN (SELECT PRDCD,TIPERAK,NORAK,NOSHELF,KIRIKANAN,kodemodis FROM RAK GROUP BY PRDCD) R ON R.PRDCD=P.PRDCD "
            sql &= "LEFT JOIN STMAST S ON S.PRDCD=P.PRDCD "
            sql &= "LEFT JOIN PROTECT PR ON PR.PRDCD=P.PRDCD "
            sql &= "LEFT JOIN PTAG_NR NR ON NR.PRDCD=P.PRDCD "
            'sql &= "LEFT JOIN (SELECT norak,kode_modis FROM jadwal_penanggungjawabrak WHERE NIK = '" & nik & "' AND kode_modis = '" & namarak & "' AND norak = '" & norak & "')  jp ON r.kodemodis=jp.kode_modis "
            sql &= "Where 1=1 "
            sql &= "AND (PR.BPB NOT IN('X','Y','Z') OR PR.BPB IS NULL) "
            sql &= "AND (P.RECID<>'1' OR P.RECID IS NULL) "
            sql &= "AND (P.NONSO<>'Y' or P.NONSO is null) "
            sql &= "AND (NR.PRDCD IS NULL)"
            sql &= "AND (SO.PRDCD IN (SELECT PLU FROM TINDAKLBTD_BAPJR WHERE `JENISBARANG` = 'TT' AND TGLSCAN = CURDATE() ))"
            TraceLog("LoadDataBAPJR_02 : " & sql)

            Mda.SelectCommand.CommandText = sql
            dt.Clear()
            Mda.Fill(dt)

            For i As Integer = 0 To dt.Rows.Count - 1
                Application.DoEvents()

                sql = "SELECT COUNT(*) FROM " & cFileSO
                sql &= " Where SOTYPE='" & dt.Rows(i)("SOTYPE") & "' "
                sql &= "AND PRDCD='" & dt.Rows(i)("PRDCD") & "'"
                Mcom.CommandText = sql
                If Mcom.ExecuteScalar = 0 Then
                    sql = "INSERT INTO " & cFileSO & "("
                    sql &= "SOTYPE"
                    sql &= ",TIPERAK"
                    sql &= ",NORAK"
                    sql &= ",NOSHELF"
                    sql &= ",KIRIKANAN"
                    sql &= ",`DIV`"
                    sql &= ",PRDCD"
                    sql &= ",`DESC`"
                    sql &= ",SINGKAT"
                    sql &= ",BARCODE"
                    sql &= ",BARCODE2"

                    sql &= ",FRAC"
                    sql &= ",UNIT"

                    sql &= ",PTAG"
                    sql &= ",CAT_COD"
                    sql &= ",BKP"
                    sql &= ",SUB_BKP"
                    sql &= ",CTGR"
                    sql &= ",KEMASAN"
                    sql &= ",ACOST"
                    sql &= ",LCOST"
                    sql &= ",RCOST"
                    sql &= ",PRICE"

                    sql &= ",TTL"
                    sql &= ",TTL1"
                    sql &= ",TTL2"

                    sql &= ",COM"
                    sql &= ",HPP"
                    sql &= ",SOTGL"
                    sql &= ",SOTIME"
                    sql &= ",NIK"
                    sql &= ",NAMA_RAK"
                    'sql &= ",TANGGAL_JADWAL"
                    sql &= ")VALUES("
                    sql &= " '" & dt.Rows(i)("SOTYPE") & "'"
                    sql &= ",'" & dt.Rows(i)("TIPERAK") & "'"
                    sql &= ",0" & dt.Rows(i)("NORAK") & ""
                    sql &= ",0" & dt.Rows(i)("NOSHELF") & ""
                    sql &= ",0" & dt.Rows(i)("KIRIKANAN") & ""
                    sql &= ",'" & dt.Rows(i)("DIV") & "'"
                    sql &= ",'" & dt.Rows(i)("PRDCD") & "'"
                    sql &= ",'" & CType(dt.Rows(i)("DESC2") & "", String).Replace("'", "''") & "'"
                    sql &= ",'" & CType(dt.Rows(i)("SINGKATAN") & "", String).Replace("'", "''") & "'"
                    sql &= ",'" & dt.Rows(i)("BARCODE") & "'"
                    sql &= ",'" & dt.Rows(i)("BARCODE2") & "'"

                    sql &= ",0" & dt.Rows(i)("FRAC") & ""
                    sql &= ",'" & dt.Rows(i)("UNIT") & "'"

                    sql &= ",'" & dt.Rows(i)("PTAG") & "'"
                    sql &= ",'" & dt.Rows(i)("CAT_COD") & "'"
                    sql &= "," & dt.Rows(i)("BKP") & ""
                    sql &= ",'" & dt.Rows(i)("SUB_BKP") & "'"
                    sql &= ",'" & dt.Rows(i)("CTGR") & "'"
                    sql &= ",'" & dt.Rows(i)("KEMASAN") & "'"
                    sql &= ",0" & dt.Rows(i)("ACOST") & ""
                    sql &= ",0" & dt.Rows(i)("LCOST") & ""
                    sql &= ",0" & dt.Rows(i)("RCOST") & ""
                    sql &= ",0" & dt.Rows(i)("PRICE") & ""
                    sql &= ",0,0,0"
                    sql &= ", " & dt.Rows(i)("COM1") & ""
                    sql &= ", " & dt.Rows(i)("HPP1") & ""
                    sql &= ",'" & Format(Now, "yyyy-MM-dd") & "'"
                    sql &= ",'" & Format(Now, "HH:mm:ss") & "'"
                    sql &= ",'" & dt.Rows(i)("NIK") & "'"
                    sql &= ",'" & dt.Rows(i)("KODE_MODIS") & "'"
                    'sql &= ",'" & Format(Now, "HH:mm:ss") & "'"
                    sql &= ")"
                    Mcom.CommandText = sql
                    Mcom.ExecuteNonQuery()
                Else
                    sql = "UPDATE " & cFileSO & " SET "
                    sql &= "TTL=0"
                    sql &= ",TTL1=0"
                    sql &= ",SOTGL='" & Format(Now, "yyyy-MM-dd") & "'"
                    sql &= ",SOTIME='" & Format(Now, "HH:mm:ss") & "'"
                    sql &= ",HPP=" & dt.Rows(i)("HPP1") & ""
                    sql &= " Where SOTYPE='" & dt.Rows(i)("SOTYPE") & "' "
                    sql &= "AND TIPERAK='" & dt.Rows(i)("TIPERAK") & "' "
                    sql &= "AND NORAK=" & dt.Rows(i)("NORAK") & " "
                    sql &= "AND NOSHELF=" & dt.Rows(i)("NOSHELF") & " "
                    sql &= "AND KIRIKANAN=" & dt.Rows(i)("KIRIKANAN") & " "
                    sql &= "AND PRDCD='" & dt.Rows(i)("PRDCD") & "'"
                    sql &= "AND NIK='" & dt.Rows(i)("NIK") & "'"
                    sql &= "AND NAMA_RAK='" & dt.Rows(i)("KODE_MODIS") & "'"
                    Mcom.CommandText = sql

                    Mcom.ExecuteNonQuery()
                End If

Lewati:
            Next

        Catch ex As Exception
            TraceLog("Error " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        Finally
            Conn.Close()
        End Try

    End Sub


    Public Sub insertDataBAPJR_AS(ByVal nik As String, ByVal namarak As String, ByVal norak As String)
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Mda As New MySqlDataAdapter("", Conn)
        Dim Rtn As New Boolean
        Dim sql As String = ""
        Dim dt As New DataTable

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            ''INSERT ITEMSO
            'sql = "DELETE from ItemSO_PJR "
            'Mcom.CommandText = sql
            'Mcom.ExecuteNonQuery()

            sql = "INSERT IGNORE INTO ITEMSO_PJR_BA_AS 
                    Select '' AS RECID,
                    P.PRDCD,
                    REPLACE(P.DESC2, '""', '''') AS `DESC`,
                    REPLACE(P.SINGKATAN, '""', '''') AS SINGKAT,
                    IFNULL(B.BARCODE,'') AS Barcode,
                    IFNULL(B.BARCODE2,'') as BARCODE2,
                    P.KEMASAN,
                    IF(COUNT(*)<=1,R.TIPERAK,null) AS TIPERAK,
                    c.NORAK AS NORAK,
                    If (COUNT(*) <= 1, R.NOSHELF, NULL)  As NOSHELF,
                    IF(COUNT(*)<=1,R.KIRIKANAN,NULL) AS KIRIKANAN,
                    CURDATE() AS TANGGAL , c.HARI,
                    '" & nik & "' AS NIK,'" & namarak & "' AS NAMA_RAK 
                    FROM PRODMAST P LEFT JOIN (SELECT * FROM RAK WHERE KODEMODIS = '" & namarak & "') R ON P.PRDCD=R.PRDCD 
                    LEFT JOIN (SELECT PLU,MIN(BARCD) AS BARCODE,IF(MAX(BARCD)<>MIN(BARCD),MAX(BARCD),'') AS BARCODE2 FROM BARCODE GROUP BY PLU ) B ON B.PLU=R.PRDCD 
                    LEFT JOIN (SELECT KODE_MODIS,NORAK,HARI FROM jadwal_penanggungjawabrak WHERE NIK = '" & nik & "' AND kode_modis = '" & namarak & "' AND norak = '" & norak & "') c ON c.kode_modis = r.kodemodis
                    WHERE P.PRDCD IN (SELECT PLU FROM TINDAKLBTD WHERE `JENISBARANG` = 'TT' AND NIK = '" & nik & "' AND NamaRakInput = '" & namarak & "' and norak = '" & norak & "' AND TGLSCAN = CURDATE()) 
                    GROUP BY R.PRDCD"
            Mcom.CommandText = sql
            TraceLog("insertDataBAPJR_01 : " & sql)
            Mcom.ExecuteScalar()
Lewati:
        Catch ex As Exception
            TraceLog("Error " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        Finally
            Conn.Close()
        End Try

    End Sub

    Public Function SelesaiCekPJR(ByVal tabel_name As String, ByVal NamaRak As String,
                                    ByVal ListShelf As String, ByVal User As ClsUser, ByVal tanggal As String, ByVal norak As String) As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim DtCP As New DataTable
        Dim Rtn As New Boolean

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Madp.SelectCommand.CommandText = "  SELECT r.PRDCD,p.DESC2,r.NOSHELF,br.NORAK NORAK,r.KODEMODIS NAMA_RAK,r.KIRIKANAN,"
            Madp.SelectCommand.CommandText &= " p.DEPART,s.QTY,IFNULL(b.MAX_RET_TOKO2DCI,0) AS MAX_RET_TOKO2DCI,b.MAX_RET_TOKO2DCI_S"
            Madp.SelectCommand.CommandText &= " FROM rak r JOIN prodmast p ON r.PRDCD = p.PRDCD"
            Madp.SelectCommand.CommandText &= " JOIN stmast s ON s.PRDCD = r.PRDCD"
            Madp.SelectCommand.CommandText &= " LEFT JOIN batas_retur b ON p.PRDCD = b.FMKODE"
            Madp.SelectCommand.CommandText &= " LEFT JOIN JADWAL_PENANGGUNGJAWABRAK br ON r.kodemodis = br.kode_modis "


            Madp.SelectCommand.CommandText &= " WHERE (r.PRDCD,br.NORAK,r.NOSHELF) NOT IN (SELECT PLU,NORAKINPUT,NOSHELFINPUT"
            Madp.SelectCommand.CommandText &= " FROM " & tabel_name & " WHERE DATE(TGLSCAN) = CURDATE())"
            Madp.SelectCommand.CommandText &= " AND r.KODEMODIS = '" & NamaRak & "'"
            Madp.SelectCommand.CommandText &= " AND r.NOSHELF IN (" & ListShelf & ") "
            Madp.SelectCommand.CommandText &= " AND br.norak ='" & norak & "' "
            If tabel_name = "TINDAKLBTD" Then
                Madp.SelectCommand.CommandText &= " AND R.PRDCD NOT IN (SELECT PLU FROM CEKPJR WHERE (`STATUS` = 'B' OR JENISBARANG = 'SO') AND DATE(TGLSCAN) = CURDATE())"
            End If
            Madp.SelectCommand.CommandText &= " GROUP BY r.PRDCD;"
            Madp.Fill(DtCP)
            TraceLog("Selesai " & tabel_name & " : " & Madp.SelectCommand.CommandText)
            If DtCP.Rows.Count > 0 Then
                Mcom.CommandText = "INSERT INTO " & tabel_name & " (PLU,TGLSCAN,NAMA,NOSHELF,NORAK,NAMARAK,`STATUS`,"
                Mcom.CommandText &= "NOSHELFINPUT,NORAKINPUT,KIRIKANAN,DIVISI,STOCK,JENISBARANG,"
                Mcom.CommandText &= "NIK,NAMAUSER,MAXBATASRETUR,MAXBATASRETUR_S,NAMARAKINPUT) VALUES"
                For Each Dr As DataRow In DtCP.Rows
                    If Dr("MAX_RET_TOKO2DCI").ToString = "" Then
                        Dr("MAX_RET_TOKO2DCI") = 0
                    End If
                    'Revisi 2020-04-30
                    'TGLSCAN yang seblumnya diset menggunakan SYSDATE(), diubah menggunakan DateTiem.Now dari VB
                    Dim TGLSCAN As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

                    Mcom.CommandText &= " ('" & Dr("PRDCD") & "', '" & TGLSCAN & "',"
                    Mcom.CommandText &= " '" & Dr("DESC2").Replace("'", "") & "', '" & Dr("NOSHELF") & "',"
                    Mcom.CommandText &= " '" & Dr("NORAK") & "', '" & Dr("NAMA_RAK") & "', 'S', NULL, NULL,"
                    Mcom.CommandText &= " '" & Dr("KIRIKANAN") & "', '" & Dr("DEPART") & "', '" & Dr("QTY") & "',"
                    If Convert.ToInt32(Dr("QTY").ToString) <= 0 Then
                        Mcom.CommandText &= " 'SO',"
                    Else
                        Mcom.CommandText &= " 'TT',"
                    End If
                    Mcom.CommandText &= " '" & User.ID & "', '" & User.Nama & "', '" & Dr("MAX_RET_TOKO2DCI") & "',"
                    Mcom.CommandText &= " '" & Dr("MAX_RET_TOKO2DCI_S") & "', '" & NamaRak & "'),"
                Next
                Mcom.CommandText = Mcom.CommandText.Remove(Mcom.CommandText.Length - 1, 1)
                Mcom.ExecuteNonQuery()
            End If

            DtCP = New DataTable
            Madp.SelectCommand.CommandText = "  SELECT PLU, jenisbarang FROM " & tabel_name
            Madp.SelectCommand.CommandText &= " WHERE DATE(tglscan) = CURDATE()"
            Madp.SelectCommand.CommandText &= " AND NIK = '" & User.ID & "';"
            Madp.Fill(DtCP)

            If DtCP.Rows.Count > 0 Then
                For Each Dr As DataRow In DtCP.Rows
                    Mcom.CommandText = "UPDATE " & tabel_name & " c, ("
                    Mcom.CommandText &= " SELECT r.prdcd AS plu, r.kodemodis AS rak, p.desc2 AS nama,"
                    Mcom.CommandText &= " r.noshelf AS noshelf, br.norak AS norak"
                    Mcom.CommandText &= " FROM rak r inner join prodmast p on r.prdcd = p.prdcd"
                    Mcom.CommandText &= " LEFT JOIN JADWAL_PENANGGUNGJAWABRAK br ON r.kodemodis = br.kode_modis "
                    Mcom.CommandText &= " WHERE(r.prdcd = p.prdcd) and br.norak = '" & norak & "'"
                    If Dr("jenisbarang").ToString.ToUpper = "SD" Then
                        Mcom.CommandText &= " AND r.prdcd = '" & Dr("PLU") & "'"
                    Else
                        Mcom.CommandText &= " AND r.prdcd = '" & Dr("PLU") & "' AND r.kodemodis = '" & NamaRak & "'"
                    End If
                    Mcom.CommandText &= "GROUP BY r.prdcd) t"
                    Mcom.CommandText &= " SET c.nama = t.nama, c.namarak = t.rak,"
                    Mcom.CommandText &= " c.noshelf = t.noshelf, c.norak = t.norak"
                    Mcom.CommandText &= " WHERE(C.plu = t.plu) AND c.NamaRakInput = '" & NamaRak & "';"
                    Mcom.ExecuteNonQuery()
                Next
            End If

            Mcom.CommandText = "SELECT COUNT(*) FROM " & tabel_name & " WHERE JENISBARANG = 'TT'"
            Mcom.CommandText &= " AND DATE(tglscan) = CURDATE() AND NIK = '" & User.ID & "' AND namarakinput = '" & NamaRak & "' AND NORAK = '" & norak & "';"
            TraceLog("Check Jumlah data ITT: " & Mcom.CommandText)
            If Mcom.ExecuteScalar > 0 Then
                Rtn = True
            End If

            Mcom.CommandText = "UPDATE " & tabel_name & " SET `Status` = 'I'"
            Mcom.CommandText &= " WHERE DATE(tglscan) = CURDATE() AND `Status` = 'S' AND NIK = '" & User.ID & "'  AND namarakinput = '" & NamaRak & "' AND NORAK = '" & norak & "';"
            Mcom.ExecuteNonQuery()

            'Tambah RECID = P untuk flag selesai PJR
            'Tambah RECID = T untuk flag selesai LBTD
            'Tambah RECID = 1 untuk flag selesai

            TraceLog("Nama Tabel: " & tabel_name.ToLower)

            If tabel_name.ToLower = "cekpjr" Then
                TraceLog("RTN: " & Rtn)
                Mcom.CommandText = "UPDATE TEMP_JADWAL_PJR SET `RECID` = 'P' "
                Mcom.CommandText &= "WHERE NIK = '" & User.ID & "' AND kode_modis = '" & NamaRak & "' AND TANGGAL = '" & tanggal.Replace("/", "-") & "' AND NORAK = '" & norak & "';"
                Mcom.ExecuteNonQuery()
                If Rtn = False Then
                    Mcom.CommandText = "UPDATE TEMP_JADWAL_PJR SET `RECID` = '1' "
                    Mcom.CommandText &= "WHERE NIK = '" & User.ID & "' AND kode_modis = '" & NamaRak & "' AND TANGGAL = '" & tanggal.Replace("/", "-") & "'  AND NORAK = '" & norak & "';"
                    Mcom.ExecuteNonQuery()

                    Mcom.CommandText = "SELECT COUNT(*) FROM cekpjr cpjr WHERE DATE(cpjr.tglscan) = '" & tanggal.Replace("/", "-") & "' AND cpjr.NIK = '" & User.ID & "' AND cpjr.NAMARAKINPUT = '" & NamaRak & "' AND cpjr.norak = '" & norak & "' AND cpjr.STATUS = 'B';"
                    TraceLog("SelesaiCekPJR - Hitung Fisik: " & Mcom.CommandText)
                    Dim jumlahFisik As Int16 = Convert.ToInt16(Mcom.ExecuteScalar)
                    Mcom.CommandText = "UPDATE TEMP_JADWAL_PJR SET ITT = 0, FISIKTIDAKADA = 0, FISIKADA = '" & jumlahFisik & "' WHERE nik = '" & User.ID & "' AND TANGGAL = '" & tanggal.Replace("/", "-") & "' AND kode_modis = '" & NamaRak & "' and norak = '" & norak & "';"
                    TraceLog("SelesaiCekPJR - Update Fisik Ada: " & Mcom.CommandText)
                    Mcom.ExecuteNonQuery()
                End If

            ElseIf tabel_name.ToLower = "tindaklbtd" Then
                Mcom.CommandText = "UPDATE TEMP_JADWAL_PJR SET `RECID` = 'T' "
                Mcom.CommandText &= "WHERE NIK = '" & User.ID & "' AND kode_modis = '" & NamaRak & "' AND TANGGAL = '" & tanggal.Replace("/", "-") & "'  AND NORAK = '" & norak & "';"
                Mcom.ExecuteNonQuery()
            End If
        Catch ex As Exception
            Rtn = False
            TraceLog("Error " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function SelesaiCekLBTD_BAPJR(ByVal tabel_name As String, ByVal User As ClsUser) As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim DtCP As New DataTable
        Dim Rtn As New Boolean

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Madp.SelectCommand.CommandText = "  SELECT r.PRDCD,p.DESC2,r.NOSHELF,br.NORAK NORAK,r.KODEMODIS NAMA_RAK,r.KIRIKANAN,"
            Madp.SelectCommand.CommandText &= " p.DEPART,s.QTY,IFNULL(b.MAX_RET_TOKO2DCI,0) AS MAX_RET_TOKO2DCI,b.MAX_RET_TOKO2DCI_S"
            Madp.SelectCommand.CommandText &= " FROM rak r JOIN prodmast p ON r.PRDCD = p.PRDCD"
            Madp.SelectCommand.CommandText &= " JOIN stmast s ON s.PRDCD = r.PRDCD"
            Madp.SelectCommand.CommandText &= " LEFT JOIN batas_retur b ON p.PRDCD = b.FMKODE"
            Madp.SelectCommand.CommandText &= " INNER JOIN ITEMSO_PJR_BA_AS br ON r.PRDCD = br.PRDCD AND r.kodemodis = br.kode_modis"
            Madp.SelectCommand.CommandText &= " WHERE (r.PRDCD) NOT IN (SELECT PLU"
            Madp.SelectCommand.CommandText &= " FROM " & tabel_name & " WHERE DATE(TGLSCAN) = CURDATE()) AND br.recid = ''"

            Madp.SelectCommand.CommandText &= " GROUP BY r.PRDCD,r.kodemodis;"
            TraceLog("Kueri Selesai : " & Madp.SelectCommand.CommandText)
            Madp.Fill(DtCP)
            If DtCP.Rows.Count > 0 Then
                Mcom.CommandText = "INSERT INTO " & tabel_name & " (PLU,TGLSCAN,NAMA,NOSHELF,NORAK,NAMARAK,`STATUS`,"
                Mcom.CommandText &= "NOSHELFINPUT,NORAKINPUT,KIRIKANAN,DIVISI,STOCK,JENISBARANG,"
                Mcom.CommandText &= "NIK,NAMAUSER,MAXBATASRETUR,MAXBATASRETUR_S,NAMARAKINPUT) VALUES"
                For Each Dr As DataRow In DtCP.Rows
                    If Dr("MAX_RET_TOKO2DCI").ToString = "" Then
                        Dr("MAX_RET_TOKO2DCI") = 0
                    End If
                    'Revisi 2020-04-30
                    'TGLSCAN yang seblumnya diset menggunakan SYSDATE(), diubah menggunakan DateTiem.Now dari VB
                    Dim TGLSCAN As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

                    Mcom.CommandText &= " ('" & Dr("PRDCD") & "', '" & TGLSCAN & "',"
                    Mcom.CommandText &= " '" & Dr("DESC2").Replace("'", "") & "', '" & Dr("NOSHELF") & "',"
                    Mcom.CommandText &= " '" & Dr("NORAK") & "', '" & Dr("NAMA_RAK") & "', 'S', NULL, NULL,"
                    Mcom.CommandText &= " '" & Dr("KIRIKANAN") & "', '" & Dr("DEPART") & "', '" & Dr("QTY") & "',"
                    If Convert.ToInt32(Dr("QTY").ToString) <= 0 Then
                        Mcom.CommandText &= " 'SO',"
                    Else
                        Mcom.CommandText &= " 'TT',"
                    End If
                    Mcom.CommandText &= " '" & User.ID & "', '" & User.Nama & "', '" & Dr("MAX_RET_TOKO2DCI") & "',"
                    Mcom.CommandText &= " '" & Dr("MAX_RET_TOKO2DCI_S") & "', '" & Dr("NAMA_RAK") & "'),"
                Next
                Mcom.CommandText = Mcom.CommandText.Remove(Mcom.CommandText.Length - 1, 1)
                Mcom.ExecuteNonQuery()
            End If

            DtCP = New DataTable
            Madp.SelectCommand.CommandText = "  SELECT PLU, jenisbarang FROM " & tabel_name
            Madp.SelectCommand.CommandText &= " WHERE DATE(tglscan) = CURDATE()"
            'Madp.SelectCommand.CommandText &= " AND NIK = '" & User.ID & "';"
            Madp.Fill(DtCP)

            If DtCP.Rows.Count > 0 Then
                For Each Dr As DataRow In DtCP.Rows
                    Mcom.CommandText = "UPDATE " & tabel_name & " c, ("
                    Mcom.CommandText &= " SELECT r.prdcd AS plu, r.kodemodis AS rak, p.desc2 AS nama,"
                    Mcom.CommandText &= " r.noshelf AS noshelf, br.norak AS norak"
                    Mcom.CommandText &= " FROM rak r inner join prodmast p on r.prdcd = p.prdcd"
                    Mcom.CommandText &= " LEFT JOIN ITEMSO_PJR_BA_AS br ON r.kodemodis = br.kode_modis "
                    Mcom.CommandText &= " WHERE(r.prdcd = p.prdcd) "
                    If Dr("jenisbarang").ToString.ToUpper = "SD" Then
                        Mcom.CommandText &= " AND r.prdcd = '" & Dr("PLU") & "'"
                    Else
                        Mcom.CommandText &= " AND r.prdcd = '" & Dr("PLU") & "'"
                    End If

                    Mcom.CommandText &= " AND r.kodemodis IN(SELECT kode_modis FROM itemso_pjr_ba_as WHERE RECID='') "

                    Mcom.CommandText &= " GROUP BY r.prdcd) t"
                    Mcom.CommandText &= " SET c.nama = t.nama, c.namarak = t.rak,"
                    Mcom.CommandText &= " c.noshelf = t.noshelf, c.norak = t.norak"
                    Mcom.CommandText &= " WHERE(C.plu = t.plu);"

                    Mcom.ExecuteNonQuery()
                Next
            End If

            Mcom.CommandText = "SELECT COUNT(*) FROM " & tabel_name & " WHERE STATUS = 'S'"
            Mcom.CommandText &= " AND DATE(tglscan) = CURDATE()  ;"
            If Mcom.ExecuteScalar > 0 Then
                Rtn = True
            End If

            Mcom.CommandText = "UPDATE " & tabel_name & " SET `Status` = 'I'"
            Mcom.CommandText &= " WHERE DATE(tglscan) = CURDATE() AND `Status` = 'S'  ;"
            Mcom.ExecuteNonQuery()

            'Tambah RECID = P untuk flag selesai PJR
            'Tambah RECID = T untuk flag selesai LBTD
            'Tambah RECID = 1 untuk flag selesai
            Mcom.CommandText = "UPDATE ITEMSO_PJR_BA_AS SET `RECID` = '1' "
            Mcom.CommandText &= "WHERE RECID = '' AND PRDCD IN (SELECT PLU FROM TINDAKLBTD_BAPJR WHERE TGLSCAN = CURDATE());"
            Mcom.ExecuteNonQuery()

        Catch ex As Exception
            Rtn = False
            TraceLog("Error " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function HASILBTD(ByVal tglscan As String, ByVal nik As String, ByVal rak As String, ByVal norak As String) As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable
        Dim Mcom As New MySqlCommand("", Conn)

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Mcom.CommandText = "DROP TABLE IF EXISTS `temp_lbtd`"
            Mcom.ExecuteNonQuery()

            Mcom.CommandText = "CREATE TABLE temp_lbtd SELECT * FROM(SELECT COUNT(*) AS LBTD, nik,tglscan,namarakinput FROM cekpjr WHERE DATE(tglscan) = '" & tglscan & "' AND NIK = '" & nik & "' AND NAMARAKINPUT = '" & rak & "' and norak = '" & norak & "' AND  JENISBARANG = 'TT') a;"

            Mcom.ExecuteNonQuery()

            Mcom.CommandText = "DROP TABLE IF EXISTS `temp_fisikada`"
            Mcom.ExecuteNonQuery()
            Mcom.CommandText = "CREATE TABLE temp_fisikada SELECT * FROM(SELECT COUNT(*) AS LBTD, '" & nik & "' as NIK,'" & tglscan & "' as tglscan,'" & rak & "' AS NAMARAKINPUT FROM tindaklbtd WHERE DATE(tglscan) = '" & tglscan & "' AND NIK = '" & nik & "' AND NAMARAKINPUT = '" & rak & "' and norak = '" & norak & "' AND `STATUS` = 'B' ) b;"
            Mcom.ExecuteNonQuery()

            Mcom.CommandText = "DROP TABLE IF EXISTS `temp_tidakada`"
            Mcom.ExecuteNonQuery()
            Mcom.CommandText = "CREATE TABLE temp_tidakada SELECT * FROM(SELECT COUNT(*) AS LBTD, '" & nik & "' as NIK,'" & tglscan & "' as tglscan,'" & rak & "' AS NAMARAKINPUT FROM tindaklbtd WHERE DATE(tglscan) = '" & tglscan & "' AND NIK = '" & nik & "' AND NAMARAKINPUT = '" & rak & "' and norak = '" & norak & "' AND `STATUS` = 'I' AND JENISBARANG = 'TT') c;"
            Mcom.ExecuteNonQuery()

            Mcom.CommandText = "DROP TABLE IF EXISTS `temp_lbtd_hpp`"
            Mcom.ExecuteNonQuery()
            Mcom.CommandText = "CREATE TABLE temp_lbtd_hpp SELECT * FROM (SELECT SUM(IF(P.ACOST <=1 OR P.ACOST IS NULL,
                                  IF(P.RCOST <=1 OR P.RCOST IS NULL,
                                  IF(P.LCOST <=1 OR P.LCOST IS NULL,
                                  ROUND(IF(BEGBAL=0,1,IF(S.RP_SLD_AKH/S.BEGBAL < 1, 1, S.RP_SLD_AKH/S.BEGBAL)))
                                 ,ROUND(P.LCOST))
                                 ,ROUND(P.RCOST))
                                 ,ROUND(P.ACOST))) AS HPP, '" & nik & "' as NIK,'" & tglscan & "' as tglscan,'" & rak & "' AS NAMARAKINPUT FROM tindaklbtd a LEFT JOIN stmast s ON a.plu = s.prdcd LEFT JOIN prodmast p ON a.plu = p.prdcd WHERE DATE(tglscan) = '" & tglscan & "' AND NIK = '" & nik & "' AND NAMARAKINPUT = '" & rak & "' and norak = '" & norak & "') c;"
            Mcom.ExecuteNonQuery()

            Madp.SelectCommand.CommandText = "SELECT a.nik,a.namarakinput AS KODE_MODIS,FORMAT(a.LBTD,0) AS ITT,FORMAT(b.LBTD,0) AS FISIK_ADA,FORMAT(c.LBTD,0) AS FISIK_TIDAK_ADA,FORMAT(d.HPP,0) AS TOTAL_HPP  FROM temp_lbtd a LEFT JOIN 
                                                temp_fisikada b ON a.nik=b.nik LEFT JOIN temp_tidakada c ON a.nik = c.nik LEFT JOIN temp_lbtd_hpp d ON a.nik = d.nik"
            Console.WriteLine(Madp.SelectCommand.CommandText)
            Madp.Fill(Rtn)

        Catch ex As Exception
            TraceLog("Error WDCP_HASILLBTD : " & ex.Message & ex.StackTrace)

            Rtn = New DataTable
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function getTanggal(ByVal tanggal As String, ByVal keterangan As String) As String
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As String = ""
        Dim Mcom As New MySqlCommand("", Conn)
        Dim hari As String = ""
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            If keterangan = "awal" Then
                keterangan = "1"
            ElseIf keterangan = "akhir" Then
                keterangan = "7"
            End If
            Mcom.CommandText = "SELECT DAYNAME('" & tanggal & "')"
            If Mcom.ExecuteScalar <> "Sunday" Or Mcom.ExecuteScalar <> "Minggu" Then
                Mcom.CommandText = "SELECT DATE_ADD(CAST('" & tanggal & "' AS DATE), INTERVAL(" & keterangan & "-DAYOFWEEK(CAST('" & tanggal & "' AS DATE ))) DAY) AS tanggal"

            Else
                Mcom.CommandText = "SELECT DATE_ADD(CAST('" & tanggal & "' AS DATE), INTERVAL(" & keterangan & "-DAYOFWEEK(CAST('" & tanggal & "' AS DATE ))) DAY) AS tanggal"
                'Mcom.CommandText = "SELECT DATE_ADD(CAST('" & tanggal & "' AS DATE), INTERVAL(" & keterangan & "-DAYOFWEEK(CAST('" & tanggal & "' AS DATE))-7) DAY) AS tanggal"

            End If
            'Console.Writeline(Mcom.CommandText)
            Rtn = Mcom.ExecuteScalar.ToString

        Catch ex As Exception
            TraceLog("Error WDCP_GETTANGGAL : " & ex.Message & ex.StackTrace)

        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function ConstRakPJR(ByVal tabel_name As String) As String
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim DtCP As New DataTable
        Dim Rtn As String = ""

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            'untuk update const

            Mcom.CommandText = "SELECT `DESC` FROM CONST WHERE RKEY = 'PJA'"
            If Mcom.ExecuteScalar <> "" Or IsDBNull(Mcom.ExecuteScalar) Then
                Mcom.CommandText = "UPDATE CONST SET `DESC` = '" & tabel_name & "' WHERE RKEY = 'PJA'"
                Mcom.ExecuteNonQuery()
            Else
                Mcom.CommandText = "INSERT IGNORE INTO CONST(`RKEY`,`DESC`) VALUES('PJA','" & tabel_name & "')"
                Mcom.ExecuteNonQuery()

            End If
            'untuk ambil nilai

        Catch ex As Exception
            TraceLog("Error " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function getConstRakPJR() As String
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim DtCP As New DataTable
        Dim Rtn As String = ""

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            'untuk ambil nilai
            Mcom.CommandText = "SELECT `DESC` FROM CONST WHERE RKEY = 'PJA'"
            Rtn = Mcom.ExecuteScalar
        Catch ex As Exception
            TraceLog("Error " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Shared Function ConstNIKPJR(ByVal NIK As String) As String
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim DtCP As New DataTable
        Dim Rtn As String = ""

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            'untuk update const

            Mcom.CommandText = "SELECT `DESC` FROM CONST WHERE RKEY = 'PJN'"
            If Mcom.ExecuteScalar <> "" Or IsDBNull(Mcom.ExecuteScalar) Then
                Mcom.CommandText = "UPDATE CONST SET `DESC` = '" & NIK & "' WHERE RKEY = 'PJN'"
                Mcom.ExecuteNonQuery()
            Else
                Mcom.CommandText = "INSERT IGNORE INTO CONST(`RKEY`,`DESC`) VALUES('PJN','" & NIK & "')"
                Mcom.ExecuteNonQuery()
            End If
            'Console.Writeline(Mcom.CommandText)
        Catch ex As Exception
            TraceLog("Error " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function getConstNIKPJR() As String
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim DtCP As New DataTable
        Dim Rtn As String = ""

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            'untuk ambil nilai
            Mcom.CommandText = "SELECT `DESC` FROM CONST WHERE RKEY = 'PJN'"
            Rtn = Mcom.ExecuteScalar
        Catch ex As Exception
            TraceLog("Error " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function approveLBTD(ByVal approval As String, ByVal tanggal As String, ByVal nik As String, ByVal kode_modis As String, ByVal norak As String) As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim approval1 As String = ""
        Dim oknok As String = ""
        Dim Rtn As Boolean
        Dim dt As New DataTable
        Dim itt As String = ""
        Dim fisikada As String = ""
        Dim fisiktidakada As String = ""
        Dim itt_adjust As String = ""
        Dim cFileSO As String = "BS_PJR_" & Format(Now, "yyMMdd") & FormMain.Toko.Kode.Substring(0, 1) & ""

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            If approval = "Y" Then
                approval1 = "1"
                oknok = "OK"
            Else
                approval1 = """''"""
                oknok = "NOK"

            End If
            Mcom.CommandText = "UPDATE TEMP_JADWAL_PJR SET RECID = '" & approval1 & "' where nik = '" & nik & "' AND TANGGAL = '" & tanggal & "' AND kode_modis = '" & kode_modis & "' and norak = '" & norak & "'"
            TraceLog("Approve LBTD Update RECID: " & Mcom.CommandText)
            Mcom.ExecuteNonQuery()

            Mcom.CommandText = "SELECT LBTD FROM temp_lbtd"
            TraceLog("Approve LBTD Get Value ITT: " & Mcom.CommandText)
            itt = Mcom.ExecuteScalar

            Mcom.CommandText = "SELECT LBTD FROM temp_fisikada"
            TraceLog("Approve LBTD Get Value Fisik Ada: " & Mcom.CommandText)
            fisikada = Mcom.ExecuteScalar

            Mcom.CommandText = "SELECT LBTD FROM temp_tidakada"
            TraceLog("Approve LBTD Get Value Fisik Tidak Ada: " & Mcom.CommandText)
            fisiktidakada = Mcom.ExecuteScalar

            Mcom.CommandText = "UPDATE TEMP_JADWAL_PJR SET ITT = '" & itt & "', FisikAda = '" & fisikada & "', FisikTidakAda ='" & fisiktidakada & "' WHERE nik = '" & nik & "' AND TANGGAL = '" & tanggal & "' AND kode_modis = '" & kode_modis & "' and norak = '" & norak & "'"
            TraceLog("Approve LBTD Update ITT, FisikAda, FisikTidakAda: " & Mcom.CommandText)
            Mcom.ExecuteNonQuery()

            Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`temp_tindaklbtd_detail` (
                                `Modis` VARCHAR(20), 
                                `Shelf` VARCHAR(12), 
                                `Tanggal` DATE, 
                                `NIK` VARCHAR(15), 
                                `Nama` VARCHAR(50),
                                `ITT` VARCHAR(5),
                                `Ada` VARCHAR(5),
                                `Tidak` VARCHAR(5),
                                `PJR` VARCHAR(4),
                                `Minggu` VARCHAR(2)
                                ) ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci;"
            Mcom.ExecuteNonQuery()

            Mcom.CommandText = "INSERT IGNORE INTO temp_tindaklbtd_detail 
                                SELECT namarak, CONCAT(MIN(noshelf),'-',MAX(noshelf)) AS Shelf,DATE_FORMAT(CURDATE(),'%Y-%m-%d') AS tanggal,nik,namauser,
                                COUNT(plu) AS ITT,
                                (SELECT COUNT(*) AS LBTD FROM tindaklbtd WHERE `STATUS` = 'B' AND namarak = '" & kode_modis & "' AND tglscan = CURDATE() AND NIK = '" & nik & "' and norak = '" & norak & "') AS ADA,
                                (SELECT COUNT(*) AS LBTD FROM tindaklbtd WHERE `STATUS` = 'I' AND JENISBARANG = 'TT' AND namarak = '" & kode_modis & "' AND tglscan = CURDATE() AND NIK = '" & nik & "' and norak = '" & norak & "') AS TIDAK,
                                '" & oknok & "',
                                (SELECT (WEEK(CURDATE()) - WEEK(DATE_FORMAT(CURDATE(),'%Y-%m-01'))) +1) AS Minggu
                                FROM tindaklbtd WHERE namarak = '" & kode_modis & "' AND NIK = '" & nik & "' AND tglscan = CURDATE() AND (`STATUS` = 'B' OR `JENISBARANG` = 'TT') and norak = '" & norak & "'"
            Console.WriteLine(Mcom.CommandText)
            Mcom.ExecuteNonQuery()

            Rtn = True
        Catch ex As Exception
            TraceLog(ex.Message & ex.StackTrace)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function updateITT_ADJUST(ByVal tanggal As String, ByVal nik As String, ByVal kode_modis As String, ByVal norak As String) As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim cFileSO As String = "BS_PJR_" & Format(Now, "yyMMdd") & FormMain.Toko.Kode.Substring(0, 1) & ""
        Dim itt_adjust As String = ""

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            Mcom.CommandText = "SELECT count(1) FROM " & cFileSO & " WHERE nik = '" & nik & "' AND DATE(SOTGL) = CURDATE() AND  NAMA_RAK = '" & kode_modis & "'  and norak = '" & norak & "'"
            TraceLog("updateITT_ADJUST - Q1: " & Mcom.CommandText)
            itt_adjust = Mcom.ExecuteScalar

            Mcom.CommandText = "UPDATE TEMP_JADWAL_PJR SET ITT_ADJUST = '" & itt_adjust & "' WHERE nik = '" & nik & "' AND TANGGAL = '" & tanggal & "' AND kode_modis = '" & kode_modis & "' and norak = '" & norak & "'"
            TraceLog("updateITT_ADJUST - Q2: " & Mcom.CommandText)
            Mcom.ExecuteNonQuery()
        Catch ex As Exception
            TraceLog(ex.Message & ex.StackTrace)
            Return False
        Finally
            Conn.Close()
        End Try
        Return True
    End Function

    Public Function CekPJR(ByVal tabel_name As String, ByVal NamaRak As String,
                                    ByVal ListShelf As String) As String
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim DtCP As New DataTable
        Dim Rtn As String

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Madp.SelectCommand.CommandText = " SELECT PLU,NORAKINPUT,NOSHELFINPUT FROM " & tabel_name & " WHERE DATE(TGLSCAN) = CURDATE() AND NAMARAKINPUT = '" & NamaRak & "'"
            Madp.SelectCommand.CommandText &= " GROUP BY PLU;"
            Madp.Fill(DtCP)
            ''Console.Writeline(Madp.SelectCommand.CommandText)
            If DtCP.Rows.Count = 0 Then
                Rtn = ""
            Else
                Rtn = "Ada"
            End If

        Catch ex As Exception
            Rtn = "Err"
            'Utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "SelesaiCekPlano", Conn)
            TraceLog(ex.Message & ex.StackTrace)

        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function cekApproveLBTD(ByVal tanggal As String, ByVal nik As String, ByVal kode_modis As String) As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim approval1 As String = ""
        Dim oknok As String = ""
        Dim Rtn As Boolean
        Dim dt As New DataTable

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            Mcom.CommandText = "SELECT RECID FROM TEMP_JADWAL_PJR where nik = '" & nik & "' AND TANGGAL = '" & tanggal & "' AND kode_modis = '" & kode_modis & "'"
            If Mcom.ExecuteScalar = "1" Then
                Rtn = False
            Else
                Rtn = True
            End If
        Catch ex As Exception
            TraceLog(ex.Message & ex.StackTrace)
            MsgBox(ex.Message & ex.StackTrace)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function cekJadwalPersonil() As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As Boolean = False
        Dim Mcom As New MySqlCommand("", Conn)

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Mcom.CommandText = "SELECT COUNT(nik) FROM `soppagent`.`abssettingshift` where nik NOT IN (SELECT DISTINCT NIK FROM TEMP_PJR)"
            If Mcom.ExecuteScalar = 0 Then
                Rtn = True
            Else
                Rtn = False
            End If

        Catch ex As Exception
            IDM.Fungsi.TraceLog("Error WDCP_cekJadwalPersonil " & ex.Message & ex.StackTrace)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function SimpanJadwalPJR() As Boolean
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
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Dim a As New FrmRegistPJR
            a.ProgressBar1.Value = 100
            Application.DoEvents()

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
            '
            Mcom.CommandText = "SELECT DISTINCT nama_rak FROM rak WHERE nama_Rak NOT IN (SELECT KODE_MODIS FROM temp_jadwal_penanggungjawabrak WHERE NIK <> '')  AND ket_rak <> ''; "
            If Mcom.ExecuteScalar <> 0 Then
                MessageBox.Show("Maaf, Proses Simpan belum dapat dilakukan jika seluruh MODIS belum didaftarkan !", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Rtn = False
                Exit Try
            Else


                Madp.SelectCommand.CommandText = "SELECT nik,nama,jabatan,hari,kode_modis,modis,shelfing,norak FROM temp_jadwal_penanggungjawabrak GROUP BY kode_modis"
                dtrak.Clear()
                Madp.Fill(dtrak)
                'progres = (80) / (dtrak.Rows.Count + progres)
                For j As Integer = 0 To dtrak.Rows.Count - 1

                    'FrmRegistPJR.ProgressBar1.Value = FrmRegistPJR.ProgressBar1.Value + progres

                    'estimasi waktu
                    Madp.SelectCommand.CommandText = "SELECT a.cat_cod, kemasan From prodmast a 
                                LEFT Join rak b ON a.prdcd = b.prdcd  Where NAMA_RAK = '" & dtrak.Rows(j)("Kode_Modis").ToString & "'  GROUP BY CAT_COD,KEMASAn ORDER BY CAT_COD "
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
                        ''Console.Writeline(timesecond)
                        'ambil jumlah item per rak, cat_cod dan kemasan
                        Mcom.CommandText = "SELECT COUNT(a.cat_cod) FROM prodmast a
                                    LEFT JOIN rak b ON a.prdcd = b.prdcd 
                                    WHERE NAMA_RAK = '" & dtrak.Rows(j)("Kode_Modis").ToString & "' AND CAT_COD LIKE '%" & dt1.Rows(i)("cat_cod") & "%' 
                                    AND KEMASAN LIKE '%" & dt1.Rows(i)("kemasan") & "%' ORDER BY CAT_COD "
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
                                WHERE NIK = '" & dtrak.Rows(j)("nik").ToString & "'  AND KODE_MODIS = '" & dtrak.Rows(j)("Kode_Modis").ToString & "';"
                        TraceLog("Kueri  : " & Mcom.CommandText)
                        totalestimasi = Mcom.ExecuteScalar
                        Mcom.CommandText = "UPDATE temp_jadwal_penanggungjawabrak SET `Totalitem` = '" & totalitem & "' , TotalEstimasi = '" & totalestimasi & "'
                                WHERE NIK = '" & dtrak.Rows(j)("nik").ToString & "' AND KODE_MODIS = '" & dtrak.Rows(j)("Kode_Modis").ToString & "'; "
                        TraceLog("Kueri  : " & Mcom.CommandText)

                        Mcom.ExecuteNonQuery()
                    End If
                Next

                Mcom.CommandText = "INSERT IGNORE INTO jadwal_penanggungjawabrak SELECT *,'' FROM temp_jadwal_penanggungjawabrak"
                Mcom.ExecuteNonQuery()
                MessageBox.Show("Berhasil Proses Simpan !", "Berhasil", MessageBoxButtons.OK)

                Rtn = True
            End If

        Catch ex As Exception
            IDM.Fungsi.TraceLog("Error WDCP_SimpanJadwalPJR " & ex.Message & ex.StackTrace)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function
    Public Function getJabatanVirbacaprod() As String
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable
        Dim listjabatan As String = ""
        Dim jabatan() As String
        Dim Mcom As New MySqlCommand("", Conn)

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            Mcom.CommandText = "SELECT jenis FROM VIR_BACAPROD WHERE JENIS = 'jabatan_pjr';"
            TraceLog("getJabatanVirbacaprod-Q1: " & Mcom.CommandText)
            If Mcom.ExecuteScalar = "" Then
                Mcom.CommandText = "INSERT IGNORE INTO VIR_BACAPROD(JENIS,FILTER,KET) VALUES('JABATAN_PJR','Chief Of Store (Ss),Store Sr. Leader (Ss),Store Jr. Leader (Ss),"
                Mcom.CommandText &= "Store Crew Boy (Ss),Store Crew Girl (Ss),Chief Of Store,Store Sr. Leader,Store Jr. Leader,"
                Mcom.CommandText &= "Store Crew Boy,Store Crew Girl,Store Junior Leader,Store Junior Leader (Ss),Store Senior Leader,Store Senior Leader (Ss)', 'JABATAN UNTUK PJR');"
                TraceLog("getJabatanVirbacaprod-Q2: " & Mcom.CommandText)
                Mcom.ExecuteNonQuery()
            Else
                'UPDATE JABATAN PJR TERBARU
                Mcom.CommandText = "UPDATE VIR_BACAPROD SET FILTER = 'Chief Of Store (Ss),Store Sr. Leader (Ss),Store Jr. Leader (Ss),"
                Mcom.CommandText &= "Store Crew Boy (Ss),Store Crew Girl (Ss),Chief Of Store,Store Sr. Leader,Store Jr. Leader,"
                Mcom.CommandText &= "Store Crew Boy,Store Crew Girl,Store Junior Leader,Store Junior Leader (Ss),Store Senior Leader,Store Senior Leader (Ss)' "
                Mcom.CommandText &= "WHERE JENIS = 'jabatan_pjr';"
                TraceLog("getJabatanVirbacaprod-Q2: " & Mcom.CommandText)
                Mcom.ExecuteNonQuery()
            End If

            Mcom.CommandText = "SELECT FILTER FROM VIR_BACAPROD WHERE JENIS = 'jabatan_pjr';"
            listjabatan = Mcom.ExecuteScalar

            jabatan = listjabatan.Split(",")
            listjabatan = ""
            For i As Integer = 0 To jabatan.Length - 1
                listjabatan &= "'" & jabatan(i) & "',"
                Console.WriteLine(listjabatan)

            Next
            listjabatan = listjabatan.Substring(0, listjabatan.Length - 1)
            Console.WriteLine(listjabatan)

        Catch ex As Exception
            TraceLog("Error WDCP_JABATANVIRBACAPROD : " & ex.Message & ex.StackTrace)
        Finally
            Conn.Close()
        End Try
        Return listjabatan
    End Function

    Public Function getBTRVirbacaprod() As String
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)

        Dim batasretur As String = ""
        Dim listbatasretur() As String
        Dim rtn As String = ""
        Dim Mcom As New MySqlCommand("", Conn)

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            Mcom.CommandText = "SELECT jenis FROM VIR_BACAPROD WHERE JENIS='BTR' "
            If Mcom.ExecuteScalar = "" Then
                Mcom.CommandText = "INSERT IGNORE INTO VIR_BACAPROD(JENIS,FILTER,KET) VALUES('BTR','B,H,J', 'MAX BATAS RETUR PJR')"
                Mcom.ExecuteNonQuery()
            Else

            End If
            Mcom.CommandText = "SELECT FILTER FROM VIR_BACAPROD WHERE JENIS='BTR' "
            batasretur = Mcom.ExecuteScalar
            Console.WriteLine(batasretur)

            listbatasretur = batasretur.Split(",")
            For i As Integer = 0 To listbatasretur.Length - 1
                rtn &= "'" & listbatasretur(i) & "',"
                Console.WriteLine(rtn)

            Next
            rtn = rtn.Substring(0, rtn.Length - 1)
            Console.WriteLine(rtn)

        Catch ex As Exception
            TraceLog("Error WDCP_JABATANVIRBACAPROD : " & ex.Message & ex.StackTrace)
        Finally
            Conn.Close()
        End Try
        Return rtn
    End Function


    Public Function notif_cekJadwal_personil() As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable
        Dim listtanggal As String = ""
        Dim Mcom As New MySqlCommand("", Conn)
        Dim jabatan As String = ""
        Dim cPJR As New ClsPJRController
        Dim cek1 As String = ""
        Dim cek2 As String = ""

        Dim result As Boolean = True
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            jabatan = cPJR.getJabatanVirbacaprod

            Mcom.CommandText = "SELECT COUNT(1) FROM SOPPAGENT.ABSPEGAWAIMST WHERE JABATAN IN (" & jabatan & ") 
                                                    AND MENOIN NOT IN (SELECT DISTINCT nik FROM JADWAL_PENANGGUNGJAWABRAK  ) AND pinjaman = 0"
            TraceLog("notif_cekJadwal_personil_1 : " & Mcom.CommandText)
            cek1 = Mcom.ExecuteScalar

            'Mcom.CommandText = "SELECT COUNT(distinct nik) FROM JADWAL_PENANGGUNGJAWABRAK  "
            Mcom.CommandText = "SELECT COUNT(1) FROM JADWAL_PENANGGUNGJAWABRAK WHERE NIK NOT IN 
                                                (SELECT MENOIN FROM SOPPAGENT.ABSPEGAWAIMST WHERE JABATAN IN 
                                                (" & jabatan & ") AND pinjaman = 0) ;"
            TraceLog("notif_cekJadwal_personil_2 : " & Mcom.CommandText)

            cek2 = Mcom.ExecuteScalar

            Console.WriteLine(cek1)
            Console.WriteLine(cek2)
            If cek1 <> 0 Or cek2 <> 0 Then
                result = False
            End If



        Catch ex As Exception
            TraceLog("Error notif_cekJadwal_personil : " & ex.Message & ex.StackTrace)
        Finally
            Conn.Close()
        End Try
        Return result
    End Function

    Public Function notif_cekJadwal_Modis() As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable
        Dim listtanggal As String = ""
        Dim Mcom As New MySqlCommand("", Conn)
        Dim jabatan As String = ""
        Dim cPJR As New ClsPJRController
        Dim cek1 As String = ""
        Dim cek2 As String = ""

        Dim result As Boolean = True
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            'Memo 447/cps/23
            'PJR hnya FJP=Y

            Mcom.CommandText = "SELECT COUNT(DISTINCT modisp,no_rak) FROM bracket a JOIN rak b ON a.modisp = b.kodemodis WHERE flagprod LIKE '%FJP=Y%';"

            TraceLog("notif_cekJadwal_modis_1 : " & Mcom.CommandText)

            cek1 = Mcom.ExecuteScalar
            Mcom.CommandText = "SELECT COUNT(DISTINCT kode_modis,norak) FROM JADWAL_PENANGGUNGJAWABRAK;"
            TraceLog("notif_cekJadwal_modis_2 : " & Mcom.CommandText)

            cek2 = Mcom.ExecuteScalar
            Console.WriteLine(cek1)
            Console.WriteLine(cek2)
            If cek1 <> cek2 Then
                'Memo 447/cps/23
                'PJR hnya FJP=Y

                Mcom.CommandText = "SELECT COUNT(1)  FROM JADWAL_PENANGGUNGJAWABRAK WHERE kode_modis NOT IN (SELECT  DISTINCT kodemodis FROM RAK WHERE flagprod LIKE '%FJP=Y%')"

                TraceLog("notif_cekJadwal_modis_3 : " & Mcom.CommandText)

                cek1 = Mcom.ExecuteScalar
                'Memo 447/cps/23
                'PJR hnya FJP=Y

                Mcom.CommandText = "SELECT COUNT(1) FROM RAK WHERE flagprod LIKE '%FJP=Y%' AND KODEMODIS  
                                                    NOT IN (SELECT DISTINCT kode_modis FROM JADWAL_PENANGGUNGJAWABRAK)"
                TraceLog("notif_cekJadwal_modis_4: " & Mcom.CommandText)

                cek2 = Mcom.ExecuteScalar

                If cek1 <> 0 Or cek2 <> 0 Then
                    result = False
                End If

            End If
        Catch ex As Exception
            TraceLog("Error notif_cekJadwal_Modis : " & ex.Message & ex.StackTrace)
        Finally
            Conn.Close()
        End Try
        Return result
    End Function

    Public Function ReloadJadwal(Optional ByVal isPelaksanaan As Boolean = False) As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable
        Dim listtanggal As String = ""
        Dim Mcom As New MySqlCommand("", Conn)
        Dim jabatan As String = ""
        Dim cPJR As New ClsPJRController
        Dim cek1 As String = ""
        Dim cek2 As String = ""

        Dim result As Boolean = True
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            Mcom.CommandText = "DELETE FROM JADWAL_PENANGGUNGJAWABRAK WHERE NIK IN (SELECT MENOIN FROM SOPPAGENT.ABSPEGAWAIMST WHERE pinjaman IN('1'))"
            Mcom.ExecuteNonQuery()

            Mcom.CommandText = "DELETE FROM TEMP_JADWAL_PJR WHERE NIK IN (SELECT MENOIN FROM SOPPAGENT.ABSPEGAWAIMST WHERE pinjaman IN('1'))"
            Mcom.ExecuteNonQuery()

            If isPelaksanaan = True Then
                Mcom.CommandText = "DELETE FROM TEMP_JADWAL_PJR WHERE recid = '' and tanggal = CURDATE()"
                Mcom.ExecuteNonQuery()
            End If

        Catch ex As Exception
            TraceLog("Error notif_cekJadwal_Modis : " & ex.Message & ex.StackTrace)
        Finally
            Conn.Close()
        End Try
        Return result
    End Function

    Public Shared Sub ReloadCekPJR(ByVal namarakinput As String)
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable
        Dim listtanggal As String = ""
        Dim Mcom As New MySqlCommand("", Conn)
        Dim jabatan As String = ""
        Dim cPJR As New ClsPJRController
        Dim cek1 As String = ""
        Dim cek2 As String = ""

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            Mcom.CommandText = "DELETE FROM CEKPJR WHERE DATE(TGLSCAN) = CURDATE() AND NAMA = '' AND NAMARAK = '' AND NAMARAKINPUT <> '" & namarakinput & "'"
            TraceLog("ReloadCekPJR : " & Mcom.CommandText)
            Mcom.ExecuteNonQuery()


        Catch ex As Exception
            TraceLog("Error ReloadCekPJR : " & ex.Message & ex.StackTrace)
        Finally
            Conn.Close()
        End Try
    End Sub

    Public Shared Function loadDataBA_AS_PJR(ByVal cFileSO As String) As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable
        Dim listtanggal As String = ""
        Dim Mcom As New MySqlCommand("", Conn)
        Dim dt As New DataTable

        Dim SQL As String

        Try
            Conn.Open()
            Mcom.CommandText = "Show Tables Like 'ITEMSO_PJR_BA_AS'"
            TraceLog("loadData_BAPJR_01 : " & Mcom.CommandText)

            If Mcom.ExecuteScalar & "" <> "" Then
                Mcom.CommandText = "Show Tables Like '" & cFileSO & "'"
                TraceLog("loadData_BAPJR_02 : " & Mcom.CommandText)

                If Mcom.ExecuteScalar & "" <> "" Then
                    'SQL = "Select I.PRDCD AS PLU,I.DESC AS NAMA,B.COM AS QTY From ITEMSO_PJR_BA_AS I JOIN " & cFileSO & " B ON I.PRDCD=B.PRDCD"
                    SQL = "Select PRDCD AS PLU,`DESC` AS NAMA,COM AS QTY From  " & cFileSO & " "

                    'Console.WriteLine(SQL)
                    TraceLog("loadData_BAPJR_03 : " & SQL)
                    Madp.SelectCommand.CommandText = SQL
                    dt.Clear()
                    Madp.Fill(dt)
                End If
            End If
        Catch ex As Exception
            ShowError("err", ex)
        Finally
            Conn.Close()
        End Try
        Return dt

    End Function

    Public Function cekPerbandinganPersonilVSModis(ByVal jumlahHari As String) As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)

        Dim Mcom As New MySqlCommand("", Conn)
        Dim rtn As Boolean = False
        Dim jabatan As String = ""
        Dim jumlah_personil As Integer = 0
        Dim jumlah_modis1 As Integer = 0
        Dim jumlah_modis2 As Integer = 0

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            jabatan = getJabatanVirbacaprod()
            If jumlahHari.Contains("5") Then
                jumlahHari = 5
            ElseIf jumlahHari.Contains("6") Then
                jumlahHari = 6
            End If

            Mcom.CommandText = "SELECT COUNT(1) * " & jumlahHari & " FROM SOPPAGENT.ABSPEGAWAIMST WHERE JABATAN IN (" & jabatan & ")
                                AND pinjaman = 0 "
            jumlah_personil = Mcom.ExecuteScalar
            Mcom.CommandText = "SELECT count(DISTINCT MODISP,NO_RAK) FROM BRACKET a 
                                inner join (select kodemodis,flagprod from rak) b 
                                on a.modisp = b.kodemodis where b.flagprod like '%FJP=Y%';"
            jumlah_modis1 = Mcom.ExecuteScalar

            Mcom.CommandText = "select count(distinct kodemodis) from rak 
                                where kodemodis not in (select modisp from bracket) and flagprod like '%FJP=Y%'"
            jumlah_modis2 = Mcom.ExecuteScalar
            jumlah_modis2 = jumlah_modis1 + jumlah_modis2

            If jumlah_personil > jumlah_modis2 Then
                'personil lbih banyak
                rtn = True
            Else
                rtn = False
            End If

        Catch ex As Exception
            rtn = False
            TraceLog("Error WDCP_GetNamaRak : " & ex.Message & ex.StackTrace)
        Finally
            Conn.Close()
        End Try
        Return rtn
    End Function


    Public Function selisihWaktu(ByVal jamAwal As Date, ByVal jamAkhir As Date) As String
        Dim hasil As String = ""

        Try
            Dim temp As Long = DateDiff(DateInterval.Second, jamAwal, jamAkhir)
            'Dim temp As Integer = DateDiff(DateInterval.Second, jamAwal, jamAkhir)
            Dim ttlSec As Long = temp
            Dim jam As Integer = Math.Floor(temp / 3600)

            temp = temp Mod 3600
            Dim mnt As Integer = Math.Floor(temp / 60)
            temp = temp Mod 60
            Dim det As Integer = temp

            hasil = jam.ToString.PadLeft(2, "0") & ":" & mnt.ToString.PadLeft(2, "0") & ":" & det.ToString.PadLeft(2, "0") & " (" & ttlSec & "s)"
        Catch ex As Exception
        End Try

        Return hasil
    End Function

End Class
