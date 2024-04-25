Imports MySql.Data.MySqlClient
Imports IDM.Fungsi
Public Class ClsKesegaranController

    Private utility As New Utility

    ''' <summary>
    ''' untuk cek table proses Kesegaran
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CekTableBatasKesegaran() As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Mda As New MySqlDataAdapter("", Conn)
        Dim Rtn As New Boolean

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Mcom.CommandText = "SHOW TABLES LIKE 'CekKesegaran'"
            If Mcom.ExecuteScalar = "" Then
                Mcom.CommandText = "Create table if not exists CekKesegaran ("
                Mcom.CommandText &= " PLU varchar(8) not null, "
                Mcom.CommandText &= " TglScan Datetime not null, "
                Mcom.CommandText &= " Nama Varchar(50), "
                Mcom.CommandText &= " NoShelf Integer, "
                Mcom.CommandText &= " NoRak Integer, "
                Mcom.CommandText &= " NamaRak Varchar(50), "
                Mcom.CommandText &= " NIK Varchar(10) not null, "
                Mcom.CommandText &= " NamaUser Varchar(50) not null, "
                Mcom.CommandText &= " BatasRetur Varchar(5) DEFAULT NULL, "
                Mcom.CommandText &= " TanggalTurunPajang Date, "
                Mcom.CommandText &= " TanggalExpTerakhir Date, "
                Mcom.CommandText &= " Qty Varchar(10), "
                Mcom.CommandText &= " Status_retur Varchar(5), "
                Mcom.CommandText &= " Status char(2), "
                Mcom.CommandText &= " Primary Key(PLU,TglScan, NoRak, NoShelf,TanggalExpTerakhir)"
                Mcom.CommandText &= " )"
                Mcom.ExecuteNonQuery()
            End If
            'Memo 1221/CPS/21 - tambah bersih2 tabel cekkesegaran
            Mcom.CommandText = "Delete from CEKKESEGARAN where TglScan <= DATE_ADD(CURDATE(), INTERVAL -3 MONTH)"
            Mcom.ExecuteNonQuery()

            Mcom.CommandText = "Show tables like 'draft_batas_kesegaran_plu'"
            If Mcom.ExecuteScalar = "" Then
                MsgBox("Table draft_batas_kesegaran_plu tidak ada")
                Rtn = False
            Else
                Rtn = True
            End If

            utility.Tracelog("Kueri: ", Mcom.CommandText, "CekTableBatasKesegaran", Conn)


        Catch ex As Exception
            Rtn = False
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "CekTablePlano", Conn)
        Finally
            Conn.Close()
        End Try

        Return Rtn
    End Function


    ''' <summary>
    ''' untuk ambil data nama rak Planogram
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNamaRak_byKesegaran() As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            'Madp.SelectCommand.CommandText = "SELECT DISTINCT nama_rak FROM rak r WHERE prdcd IN (SELECT prdcd FROM draft_batas_kesegaran_plu);"
            Madp.SelectCommand.CommandText = "SELECT DISTINCT nama_RAk FROM (SELECT a.*, nama_rak FROM draft_batas_kesegaran_plu a LEFT JOIN rak b 
                ON a.prdcd = b.prdcd WHERE nama_Rak IS NOT NULL ORDER BY nama_Rak) c LEFT JOIN cekkesegaran d ON c.prdcd = d.plu AND C.NAMA_RAK = D.NAMARAK
                AND c.tanggal_exp_terakhir = d.tanggalexpterakhir LEFT JOIN prodmast e ON c.prdcd = e.prdcd WHERE PLU IS NULL AND recid IS NOT NULL"
            Madp.Fill(Rtn)
            utility.Tracelog("Kueri: ", Madp.SelectCommand.CommandText, "GetNamaRak_byKesegaran", Conn)
        Catch ex As Exception
            Rtn = New DataTable
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetNamaRak", Conn)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    ''' <summary>
    ''' untuk cek modis proses Planogram
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CekModis(ByVal Modis As String, ByRef JumlahItem As String, ByRef NamaModis As String) As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As Boolean = True

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            'Mcom.CommandText = "SELECT COUNT(DISTINCT p.prdcd) FROM draft_batas_kesegaran_plu d LEFT JOIN rak r ON d.prdcd = r.prdcd LEFT JOIN PRODMAST p ON d.prdcd = p.prdcd WHERE nama_rak = '" & Modis & "'"
            'Mcom.CommandText = " SELECT COUNT(*) FROM (SELECT a.*, nama_rak FROM draft_batas_kesegaran_plu a LEFT JOIN rak b 
            '                     ON a.prdcd = b.prdcd WHERE nama_Rak IS NOT NULL AND nama_Rak ='" & Modis & "' GROUP BY a.prdcd ORDER BY nama_Rak) c LEFT JOIN cekkesegaran d ON c.prdcd = d.plu AND C.NAMA_RAK = D.NAMARAK
            '                     AND c.tanggal_exp_terakhir = d.tanggalexpterakhir WHERE PLU IS NULL AND nama_Rak ='" & Modis & "'"
            Mcom.CommandText = " SELECT COUNT(DISTINCT PRDCD) FROM temp_draft_kesegaran WHERE nama_rak = '" & Modis & "'"
            JumlahItem = Mcom.ExecuteScalar
            TraceLog("WDCP CekKesegaran_CekModis : " & Mcom.CommandText)

            Mcom.CommandText = "select KET_RAK from rak where nama_rak = '" & Modis & "';"
            NamaModis = Mcom.ExecuteScalar


        Catch ex As Exception
            JumlahItem = 0
            Rtn = False
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "CekModis", Conn)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    ''' <summary>
    ''' Get deskripsi produk
    ''' </summary>
    ''' <param name="tabel_name"></param>
    ''' <param name="barcode_plu"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDeskripsiKesegaran(ByVal tabel_name As String, ByVal barcode_plu As String,
                                          ByVal NamaRak As String, ByVal User As ClsUser, ByVal qty As String) As ClsKesegaran
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Result As New ClsKesegaran

        If Conn Is Nothing Then
            utility.TraceLogTxt("Error - GetDeskripsiPlanogram (connection Nothing) " & vbCrLf & "PLU:" & barcode_plu)
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
                Dim temp_tgl As Date



                Mcom.CommandText = "SELECT r.PRDCD, p.DESC2, r.NORAK, r.NAMA_RAK, r.NOSHELF, d.batas_retur,d.TANGGAL_TURUN_PAJANG, p.status_retur,  
                                     tanggal_Exp_terakhir as `tgl_exp`, st.qty 
                                      FROM prodmast p LEFT JOIN 
                                       rak r ON p.PRDCD = r.PRDCD 
                                      LEFT JOIN draft_batas_kesegaran_plu d ON p.prdcd = d.prdcd
                                      Left Join barcode b ON p.PRDCD = b.PLU 
                                      LEFT JOIN stmast st ON p.PRDCD = st.prdcd 
                                       LEFT JOIN temp_cekkesegaran t ON p.prdcd = t.plu

                                      WHERE(b.PLU = '" & barcode_plu & "' OR b.BARCD = '" & barcode_plu & "')
                                      And nama_Rak = '" & NamaRak & "' AND r.prdcd IN (t.plu) GROUP BY TGL_EXP 
                                      ORDER BY TGL_EXP DESC "

                'ORDER BY TANGGAL_TURUN_PAJANG "

                IDM.Fungsi.TraceLog("WDCP_GetDeskripsiKesegaran : " & Mcom.CommandText)
                Dim sDap As New MySqlDataAdapter(Mcom)
                sDap.Fill(DtPlano)

                If DtPlano.Rows.Count = 0 Then
                    Result.Prdcd = ""
                    Result.Desc = "Tidak Ditemukan"
                ElseIf DtPlano.Rows.Count > 0 Then

                    Result.Prdcd = DtPlano.Rows(0).Item("PRDCD")
                    If DtPlano.Rows(0).Item("DESC2").Length > 40 Then
                        Result.Desc = DtPlano.Rows(0).Item("DESC2").Substring(0, 40)
                    Else
                        Result.Desc = DtPlano.Rows(0).Item("DESC2").ToString.PadRight(40, " ")
                    End If
                    For i As Integer = 0 To DtPlano.Rows.Count - 1
                        temp_tgl = DtPlano.Rows(i).Item("tgl_exp")
                        If i = DtPlano.Rows.Count - 1 Then
                            Result.MaxRet &= temp_tgl.ToString("dd/MM/yy") & ""
                        Else
                            Result.MaxRet &= temp_tgl.ToString("dd/MM/yy") & ","
                        End If
                    Next





                Else
                    Result.Desc = "Tidak Ditemukan"
                End If

                If qty <> "" Then
                    'Insert ke Table Cekkesegaran
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

                        If qty > DtPlano.Rows(0).Item("qty") Then
                            Result.Prdcd = ""
                            Result.Desc = "QTY melebihi LPP!"
                            Result.MaxRet = ""

                        Else

                            'Revisi 25/06/2021
                            'Data yang discan masuk ke tabel temporary terlebih dahulu

                            Dim TGLSCAN As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

                            Dim tgl_exp As String = Date.Parse(DtPlano.Rows(0).Item("tgl_exp")).ToString("yyyy-MM-dd HH:mm:ss")
                            Mcom.CommandText = "SELECT COUNT(*) FROM TEMP_CEKKESEGARAN WHERE PLU = '" & DtPlano.Rows(0).Item("PRDCD") & "' 
                                            AND DATE(TGLSCAN) = CURDATE() AND NAMARAK = '" & TempNamaRak & "'"
                            IDM.Fungsi.TraceLog("WDCP_GetDeskripsiKesegaran_1(" & 0 & ") : " & Mcom.CommandText)
                            If Mcom.ExecuteScalar > 0 Then
                                Mcom.CommandText = "UPDATE `TEMP_CEKKESEGARAN` SET QTY = '" & qty & "', TGLSCAN = NOW(), NIK = '" & User.ID & "', 
                                                    NAMAUSER = '" & User.Nama & "', STATUS = '1' 
                                                    where PLU = '" & DtPlano.Rows(0).Item("PRDCD") & "' 
                                                    AND DATE(TGLSCAN) = CURDATE() AND NAMARAK = '" & TempNamaRak & "' AND tanggalEXPTERAKHIR = '" & tgl_exp & "'"
                                IDM.Fungsi.TraceLog("WDCP_GetDeskripsiKesegaran_2 : " & Mcom.CommandText)
                                Mcom.ExecuteNonQuery()
                            Else
                                Mcom.CommandText = "INSERT ignore INTO `TEMP_CEKKESEGARAN` (`PLU`,`TGLSCAN`,`NAMA`,`NOSHELF`,`NORAK`,`NAMARAK`,
                                                    `NIK`,`NAMAUSER`,`BATASRETUR`,`TANGGALTURUNPAJANG`,`TANGGALEXPTERAKHIR`,`QTY`,`STATUS_RETUR`,`Status`)
                                                    VALUES ('" & DtPlano.Rows(0).Item("PRDCD") & "','" & TGLSCAN & "','" & Result.Desc & "','" & TempNoShelf & "',
                                                    '" & DtPlano.Rows(0).Item("NORAK") & "','" & TempNamaRak & "',
                                                    '" & User.ID & "', '" & User.Nama & "', 
                                                    '" & DtPlano.Rows(0).Item("batas_retur") & "','" & Date.Parse(DtPlano.Rows(0).Item("TANGGAL_TURUN_PAJANG")).ToString("yyyy-MM-dd HH:mm:ss") & "', '" & tgl_exp & "', '" & qty & "','" & DtPlano.Rows(0).Item("STATUS_RETUR") & "', '1')"
                                IDM.Fungsi.TraceLog("WDCP_GetDeskripsiKesegaran_3 : " & Mcom.CommandText)
                            End If



                            Mcom.ExecuteNonQuery()


                        End If
                    End If
                End If

            Catch ex As Exception
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetDeskripsiKesegaran", Conn)
                utility.TraceLogTxt("Error - GetDeskripsiKesegaran " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                Conn.Close()
            End Try
        End SyncLock

        Return Result
    End Function

    Public Function SimpanQTYKesegaran(ByVal tabel_name As String, ByVal barcode_plu As String,
                                          ByVal NamaRak As String, ByVal User As ClsUser) As ClsKesegaran
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Result As New ClsKesegaran

        If Conn Is Nothing Then
            utility.TraceLogTxt("Error - SimpanQTYKesegaran (connection Nothing) " & vbCrLf & "PLU:" & barcode_plu)
            Return Result
            Exit Function
        End If

        SyncLock Conn
            Try
                If Conn.State = ConnectionState.Closed Then
                    Conn.Open()
                End If



            Catch ex As Exception
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "SimpanQTYKesegaran", Conn)
                utility.TraceLogTxt("Error - SimpanQTYKesegaran " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                Conn.Close()
            End Try
        End SyncLock

        Return Result
    End Function


    'Public Function GetDeskripsiKesegaran_byIndex(ByVal tabel_name As String, ByVal barcode_plu As String,
    '                                      ByVal NamaRak As String, ByVal User As ClsUser, ByVal index As String) As ClsPlanogram
    '    Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
    '    Dim tmpDt As New DataTable
    '    Dim result As New ClsSo
    '    Dim mcom As New MySqlCommand("", conn)
    '    Dim da As New MySqlDataAdapter
    '    Dim dt As New DataTable
    '    Dim no_rak As String = ""
    '    Dim no_shelf As String = ""
    '    Console.WriteLine(indeks)
    '    If conn Is Nothing Then
    '        utility.TraceLogTxt("Error - GetDeskripsiProduk_byIndex (connection Nothing) " & vbCrLf & "index:" & indeks)
    '        Return result
    '        Exit Function
    '    End If

    '    SyncLock conn
    '        Try
    '            If conn.State = ConnectionState.Closed Then
    '                conn.Open()
    '            End If


    '            mcom.CommandText = "SELECT * FROM " & tabel_name & ""
    '            da.SelectCommand = mcom
    '            da.Fill(dt)
    '            Console.WriteLine(dt.Rows.Count - 1)
    '            If indeks > dt.Rows.Count - 1 And dt.Rows.Count - 1 <> "0" Then
    '                result.PRDCD = ""
    '                result.Unit = ""
    '                result.Deskripsi = "Sudah semua"
    '                result.Rak = ""
    '                result.QTYToko = ""
    '                result.QTYGudang = ""
    '                result.QTYTotal = ""
    '                result.QTYCom = ""

    '            Else
    '                mcom.CommandText = "SELECT distinct T.TIPERAK,T.NORAK,T.NOSHELF,T.PRDCD,T.SINGKAT,T.TTL,T.TTL1,T.TTL2,"
    '                mcom.CommandText &= "T.SOID,T.SOTIME,T.DCP,T.KIRIKANAN,T.Unit,T.COM+T.BPB-T.RETUR_K-T.SALES+T.RETUR+T.BPB_2+T.ADJ-T.TTL2 AS COM"
    '                mcom.CommandText &= " FROM " & tabel_name & " T left join BARCODE B "
    '                mcom.CommandText &= " on T.PRDCD = B.PLU "
    '                mcom.CommandText &= " WHERE B.BARCD = '" & dt.Rows(indeks)("prdcd") & "' or T.PRDCD ='" & dt.Rows(indeks)("prdcd") & "'"
    '                mcom.CommandText &= " ORDER BY T.NORAK,T.NOSHELF,T.TIPERAK,T.KIRIKANAN"
    '                Console.WriteLine(mcom.CommandText)
    '                Dim sDap As New MySqlDataAdapter(mcom)
    '                sDap.Fill(tmpDt)

    '                result.BarcodePlu = dt.Rows(indeks)("prdcd")
    '                If tmpDt.Rows.Count > 0 Then
    '                    no_rak = CInt(tmpDt.Rows.Item(0)("NORAK"))
    '                    no_rak = no_rak.PadLeft(3, "0")
    '                    no_shelf = CInt(tmpDt.Rows.Item(0)("NOSHELF"))
    '                    no_shelf = no_shelf.PadLeft(3, "0")

    '                    result.PRDCD = tmpDt.Rows(0)("PRDCD")
    '                    result.Unit = tmpDt.Rows(0)("Unit")

    '                    result.Deskripsi = tmpDt.Rows(0)("SINGKAT")
    '                    If result.Deskripsi.Length > 20 Then
    '                        result.Deskripsi = result.Deskripsi.Substring(0, 20)
    '                    End If

    '                    result.Rak = no_rak & "/" & no_shelf
    '                    result.QTYToko = tmpDt.Rows(0)("TTL1")
    '                    result.QTYGudang = tmpDt.Rows(0)("TTL2")
    '                    result.QTYTotal = tmpDt.Rows(0)("TTL")
    '                    result.QTYCom = tmpDt.Rows(0)("COM")

    '                    'Revisi 20 November 2019 (Memo 1081/CPS/19)
    '                    'Hitung lokasi RAK untuk item, Jika ada lebih dari 1 lokasi aktifkan fitur NEXT WDCP
    '                    mcom.CommandText = "SELECT COUNT(NORAK) FROM RAK R"
    '                    mcom.CommandText &= " LEFT JOIN BARCODE B  ON R.PRDCD = B.PLU"
    '                    mcom.CommandText &= " WHERE B.BARCD = '" & dt.Rows(indeks)("prdcd") & "' OR R.PRDCD ='" & dt.Rows(indeks)("prdcd") & "';"
    '                    Dim CountRak = mcom.ExecuteScalar
    '                    If Not IsDBNull(CountRak) Then
    '                        result.TotalRak = CountRak
    '                    Else
    '                        result.TotalRak = 0
    '                    End If
    '                Else
    '                    result.PRDCD = ""
    '                    result.Unit = ""
    '                    result.Deskripsi = "Tidak Ditemukan"
    '                    result.Rak = ""
    '                    result.QTYToko = ""
    '                    result.QTYGudang = ""
    '                    result.QTYTotal = ""
    '                    result.QTYCom = ""
    '                End If
    '            End If
    '        Catch ex As Exception
    '            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetDeskripsiProduk", conn)
    '            utility.TraceLogTxt("Error - GetDeskripsiProduk " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
    '        Finally
    '            conn.Close()
    '        End Try

    '    End SyncLock

    '    Return result
    'End Function


    ''' <summary>
    ''' proses selesai cek Planogram
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SelesaiCekKesegaran(ByVal tabel_name As String, ByVal NamaRak As String
                                     ) As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim DtCP As New DataTable
        Dim DtCP2 As New DataTable

        Dim Rtn As New Boolean
        Dim Rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Rpt = New rptListingTurunPajang
        Dim sqltampung As String = ""
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Mcom.CommandText = "UPDATE TEMP_CEKKESEGARAN SET STATUS ='1'"
            Mcom.ExecuteNonQuery()
            Mcom.CommandText = "INSERT IGNORE INTO CEKKESEGARAN SELECT * FROM TEMP_CEKKESEGARAN"
            Mcom.ExecuteNonQuery()

            Madp.SelectCommand.CommandText = "SELECT r.PRDCD,p.DESC2 as `DESC`,r.NOSHELF,r.NORAK,r.NAMA_RAK AS `MODIS`,r.KIRIKANAN,
                                                p.DEPART, s.QTY As QTYTP, BATASRETUR, TANGGALTURUNPAJANG, TANGGALEXPTERAKHIR, p.STATUS_RETUR AS `STATUSPTRT` 
                                                 FROM rak r JOIN prodmast p ON r.PRDCD = p.PRDCD 
                                                Join CEKKESEGARAN s ON s.PLU = r.PRDCD 
                                                 WHERE (r.PRDCD) IN (SELECT PLU 
                                                 From cekkesegaran Where Date(TGLSCAN) = CURDATE()) 
                                                 AND r.NAMA_RAK = '" & NamaRak & "' 
                                                 GROUP BY r.PRDCD;"

            Madp.Fill(DtCP)

            'If DtCP.Rows.Count > 0 Then
            Mcom.CommandText = "INSERT IGNORE INTO CEKKESEGARAN Select t.prdcd, Now(), t.singkatan, r.noshelf, r.norak, t.nama_rak,'','',d.batas_retur,d.tanggal_turun_pajang,d.tanggal_Exp_terakhir,0,t.status_retur,''
                         From TEMP_DRAFT_KESEGARAN t LEFT Join rak r ON t.prdcd = r.prdcd
                         LEFT JOIN draft_Batas_kesegaran_plu d ON t.prdcd = d.prdcd WHERE  t.nama_Rak = '" & NamaRak & "' AND r.nama_Rak = '" & NamaRak & "' 
                         AND t.prdcd NOT IN (SELECT plu FROM cekkesegaran WHERE namarak = '" & NamaRak & "') AND singkatan IS NOT NULL"
            Mcom.ExecuteNonQuery()
            FormMain.cmbmodisText = NamaRak

            'Dim f As New FrmLITP
            'f.Show()
            Rtn = True
            'Else
            '    Rtn = False
            'End If


        Catch ex As Exception
            Rtn = False
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "SelesaiCekPlano", Conn)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function getRCKB() As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Mda As New MySqlDataAdapter("", Conn)
        Dim Rtn As New Boolean
        Dim DtCP As New DataTable
        Dim sqltampung As String = ""
        Dim temp_tanggal As Date
        Dim temp_tanggal2 As String = ""
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            Mcom.CommandText = "SELECT KASIR_NAME FROM INITIAL WHERE STATION = '" & IDM.Fungsi.Get_Station() & "' AND TANGGAL = CURDATE()"
            If Mcom.ExecuteScalar <> "" Then


                Mcom.CommandText = "DROP TABLE IF EXISTS `temp_draft_kesegaran`"
                Mcom.ExecuteNonQuery()

                Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `temp_draft_kesegaran` ( 
                                `prdcd` VARCHAR(8), 
                                `singkatan` VARCHAR(99), 
                                `status_retur` VARCHAR(3), 
                                `tanggal_exp_terakhir` VARCHAR(30) , 
                                `nama_rak` VARCHAR(30) ) ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci"
                Mcom.ExecuteNonQuery()

                Mcom.CommandText = "SELECT DISTINCT PRDCD FROM DRAFT_BATAS_KESEGARAN_PLU group by prdcd"
                Mda.SelectCommand = Mcom
                Mda.Fill(DtCP)
                'Console.WriteLine(DtCP.Rows.Count)
                For i As Integer = 0 To DtCP.Rows.Count - 1
                    Mcom.CommandText = "SELECT TANGGAL_EXP_TERAKHIR FROM DRAFT_BATAS_KESEGARAN_PLU WHERE prdcd = '" & DtCP.Rows(i)("prdcd").ToString & "'"
                    Dim dttgl As New DataTable
                    Dim datgl As New MySqlDataAdapter
                    dttgl.Clear()
                    datgl.SelectCommand = Mcom
                    datgl.Fill(dttgl)
                    'Console.WriteLine(dttgl.Rows.Count - 1)
                    For j As Integer = 0 To dttgl.Rows.Count - 1
                        temp_tanggal = dttgl.Rows(j)("tanggal_Exp_terakhir")

                        temp_tanggal2 = temp_tanggal.ToString("dd/MM/yyyy")
                        'Console.WriteLine(temp_tanggal2)
                        Mcom.CommandText = "INSERT IGNORE INTO temp_draft_kesegaran "
                        Mcom.CommandText &= "SELECT a.prdcd,b.singkatan,b.status_retur,'" & temp_tanggal2 & "',c.nama_rak FROM ( 
                               SELECT DISTINCT prdcd FROM draft_batas_kesegaran_plu 
                               WHERE prdcd = '" & DtCP.Rows(i)("prdcd").ToString & "' ORDER BY tanggal_exp_terakhir asc 
                                ) a
                               LEFT JOIN prodmast b ON a.prdcd = b.prdcd LEFT JOIN rak c ON a.prdcd = c.prdcd where b.recid is not null;

"
                        Mcom.ExecuteNonQuery()
                        sqltampung &= Mcom.CommandText
                    Next


                    'Console.WriteLine(Mcom.CommandText)
                Next
                IDM.Fungsi.TraceLog("WDCP_getRCKB : " & sqltampung)
                Rtn = True

                Dim a As New FrmRKCB

                a.Show()
                a.Activate()

            Else
                Rtn = False
                MessageBox.Show("Tidak Ada initial di komputer ini", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If


        Catch ex As Exception

            Rtn = False
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "CekTablePlano", Conn)
        Finally
            Conn.Close()
        End Try

        Return Rtn
    End Function

    Public Function MulaiCekKesegaran(ByVal modis As String) As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Result As Boolean

        If Conn Is Nothing Then
            Return False
            Exit Function
        End If

        SyncLock Conn
            Try
                If Conn.State = ConnectionState.Closed Then
                    Conn.Open()
                End If
                'Mcom.CommandText = "SELECT COUNT(*) FROM CEKKESEGARAN WHERE NAMARAK = '" & modis & "'"
                Mcom.CommandText = " SELECT COUNT(*) FROM (SELECT a.*, nama_rak FROM draft_batas_kesegaran_plu a LEFT JOIN rak b 
                ON a.prdcd = b.prdcd WHERE nama_Rak IS NOT NULL ORDER BY nama_Rak) c LEFT JOIN cekkesegaran d ON c.prdcd = d.plu AND C.NAMA_RAK = D.NAMARAK
                AND c.tanggal_exp_terakhir = d.tanggalexpterakhir left join prodmast e on c.prdcd = e.prdcd WHERE PLU IS NULL AND nama_Rak ='" & modis & "'"
                'Console.WriteLine(Mcom.CommandText)
                If Mcom.ExecuteScalar > 0 Then
                    Mcom.CommandText = "DROP TABLE IF EXISTS `temp_CekKesegaran`"
                    Mcom.ExecuteNonQuery()

                    Mcom.CommandText = "Create table if not exists temp_CekKesegaran (
                                       PLU varchar(8) Not null, 
                                        TglScan DateTime Not null, 
                                        Nama Varchar(50),
                                        NoShelf Integer, 
                                        NoRak Integer,
                                        NamaRak Varchar(50), 
                                        NIK Varchar(10) Not null, 
                                        NamaUser Varchar(50) Not null, 
                                        BatasRetur Varchar(5) DEFAULT NULL, 
                                        TanggalTurunPajang Date, 
                                        TanggalExpTerakhir Date,
                                        Qty Varchar(10), 
                                        Status_retur Varchar(5), 
                                        Status Char(2) DEFAULT NULL,
                                        Primary Key(PLU, TglScan, NoRak, NoShelf,TanggalExpTerakhir)
                                        )"
                    Mcom.ExecuteNonQuery()
                    Mcom.CommandText = "INSERT IGNORE INTO temp_CekKesegaran SELECT c.prdcd,CURDATE(),e.singkatan,c.noshelf,c.norak,c.nama_rak,'','',batas_retur,tanggal_turun_pajang,tanggal_exp_terakhir,0,e.status_Retur,NULL FROM (SELECT a.*,nama_rak,noshelf,norak FROM draft_batas_kesegaran_plu a LEFT JOIN rak b 
                ON a.prdcd = b.prdcd WHERE nama_Rak IS NOT NULL ORDER BY nama_Rak) c LEFT JOIN cekkesegaran d ON c.prdcd = d.plu AND C.NAMA_RAK = D.NAMARAK
                AND c.tanggal_exp_terakhir = d.tanggalexpterakhir LEFT JOIN prodmast e ON c.prdcd = e.prdcd WHERE plu is null and nama_Rak ='" & modis & "' AND e.prdcd IS NOT NULL ORDER BY tanggal_exp_terakhir"
                    Mcom.ExecuteNonQuery()
                    Result = True
                Else
                    Result = False
                End If




                'Mcom.CommandText = "INSERT ignore INTO cekkesegaran 
                '                    SELECT r.PRDCD,CURDATE(), p.DESC2,r.NOSHELF, r.NORAK, r.NAMA_RAK,'','',
                '                    d.batas_retur,d.TANGGAL_TURUN_PAJANG,tanggal_Exp_terakhir AS `tgl_exp`,0,
                '                    p.status_retur,NULL FROM prodmast p LEFT JOIN  
                '                    rak r ON p.PRDCD = r.PRDCD  LEFT JOIN draft_batas_kesegaran_plu d ON p.prdcd = d.prdcd 
                '                    LEFT JOIN barcode b ON p.PRDCD = b.PLU  LEFT JOIN stmast st ON p.PRDCD = st.prdcd 
                '                     WHERE r.prdcd IN (SELECT prdcd FROM draft_batas_kesegaran_plu) AND nama_Rak = '" & NamaRak & "' GROUP BY prdcd ORDER BY TANGGAL_TURUN_PAJANG "
                'Mcom.ExecuteNonQuery()
                'IDM.Fungsi.TraceLog("WDCP_MulaiCekKesegaran : " & Mcom.CommandText)
                'Result = True


            Catch ex As Exception
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetDeskripsiKesegaran", Conn)
                utility.TraceLogTxt("Error - GetDeskripsiKesegaran " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                Conn.Close()
            End Try
        End SyncLock

        Return Result
    End Function
    Public Function CekKesegaran(ByVal NamaRak As String,
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
            Madp.SelectCommand.CommandText = " SELECT PLU FROM temp_CekKesegaran WHERE DATE(TGLSCAN) = CURDATE() AND NAMARAK = '" & NamaRak & "'"
            Madp.SelectCommand.CommandText &= " AND NIK <>'' AND NAMAUSER<>''"
            Madp.SelectCommand.CommandText &= " GROUP BY PLU;"
            Madp.Fill(DtCP)
            'Console.WriteLine(Madp.SelectCommand.CommandText)
            If DtCP.Rows.Count = 0 Then
                Rtn = ""
                Mcom.CommandText = "DELETE FROM TEMP_CEKKESEGARAN WHERE DATE(TGLSCAN) = CURDATE() AND NAMARAK = '" & NamaRak & "'"
                Mcom.ExecuteNonQuery()
            Else
                Rtn = "Ada"
            End If

        Catch ex As Exception
            Rtn = "Err"
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "SelesaiCekKesegaran", Conn)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function
End Class
