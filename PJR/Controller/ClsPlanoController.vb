Imports MySql.Data.MySqlClient

Public Class ClsPlanoController

    Private utility As New Utility

    ''' <summary>
    ''' untuk cek table proses Planogram
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CekTablePlano() As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Mda As New MySqlDataAdapter("", Conn)
        Dim Rtn As New Boolean

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Mcom.CommandText = "Show tables like 'CekPlanogram'"
            If IsNothing(Mcom.ExecuteScalar) Then
                Mcom.CommandText = "Create table CekPlanogram ("
                Mcom.CommandText &= " PLU varchar(8) not null, "
                Mcom.CommandText &= " TglScan Datetime not null, "
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
                Mcom.CommandText &= " MaxBatasRetur int(11) DEFAULT NULL, "
                Mcom.CommandText &= " MaxBatasRetur_S Varchar(4) DEFAULT NULL, "
                Mcom.CommandText &= " NamaRakInput Varchar(20), "
                Mcom.CommandText &= " Primary Key(PLU,TglScan, NoRak, NoShelf)"
                Mcom.CommandText &= " )"
                Mcom.ExecuteNonQuery()
            Else
                Mcom.CommandText = "Select column_type From Information_schema.Columns "
                Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('cekplanogram') "
                Mcom.CommandText &= "And Column_Name='nama' "
                If Mcom.ExecuteScalar & "" <> "varchar(50)" Then
                    Mcom.CommandText = "ALTER TABLE cekplanogram modify COLUMN NAMA varchar(50) NOT NULL "
                    Mcom.ExecuteNonQuery()
                End If

                Mcom.CommandText = " Select count(*) From Information_schema.Columns "
                Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('cekplanogram') "
                Mcom.CommandText &= " And Column_Name='MaxBatasRetur'"
                If Mcom.ExecuteScalar = 0 Then
                    Mcom.CommandText = "Alter table cekplanogram "
                    Mcom.CommandText &= "ADD COLUMN `MaxBatasRetur` int(11) DEFAULT NULL"
                    Mcom.ExecuteNonQuery()
                End If

                Mcom.CommandText = " Select count(*) From Information_schema.Columns "
                Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('cekplanogram') "
                Mcom.CommandText &= " And Column_Name='MaxBatasRetur_S'"
                If Mcom.ExecuteScalar = 0 Then
                    Mcom.CommandText = "Alter table cekplanogram "
                    Mcom.CommandText &= "ADD COLUMN `MaxBatasRetur_S` Varchar(4) DEFAULT NULL"
                    Mcom.ExecuteNonQuery()
                End If

                Mcom.CommandText = " Select count(*) From Information_schema.Columns "
                Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('cekplanogram') "
                Mcom.CommandText &= " And Column_Name='NamaRakInput'"
                If Mcom.ExecuteScalar = 0 Then
                    Mcom.CommandText = "Alter table cekplanogram "
                    Mcom.CommandText &= "ADD COLUMN `NamaRakInput` varchar(20)"
                    Mcom.ExecuteNonQuery()
                End If

                Mcom.CommandText = "Select column_type From Information_schema.Columns "
                Mcom.CommandText &= " Where TABLE_SCHEMA='pos' AND Table_Name In('cekplanogram') "
                Mcom.CommandText &= "And Column_Name='MaxBatasRetur' "
                If Mcom.ExecuteScalar & "" <> "int(11) DEFAULT NULL" Then
                    Mcom.CommandText = "ALTER TABLE cekplanogram modify COLUMN MaxBatasRetur int(11) DEFAULT NULL "
                    Mcom.ExecuteNonQuery()
                End If
                
            End If
            Rtn = True
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
    Public Function GetNamaRak() As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Madp.SelectCommand.CommandText = "select distinct nama_rak from rak r;"
            Madp.Fill(Rtn)
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
    Public Function CekModis(ByVal Modis As String, ByRef CountModis As Integer, ByRef NamaModis As String) As DataTable
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Mcom.CommandText = "select count(*) from rak r "
            Mcom.CommandText &= " where nama_rak = '" & Modis & "' "
            Mcom.CommandText &= " and noshelf not in (select noshelf from cekplanogram where namarak = '" & Modis & "' "
            Mcom.CommandText &= "and date(tglscan) = date(now()) and jenisbarang <> 'SD')"
            Mcom.CommandText &= " order by noshelf asc;"
            CountModis = Mcom.ExecuteScalar

            If CountModis = 0 Then
                Return Rtn
                Exit Function
            End If

            Mcom.CommandText = "select KET_RAK from rak where nama_rak = '" & Modis & "';"
            NamaModis = Mcom.ExecuteScalar

            'Revisi (15 April 2019)
            'Email: RE: Permasalahan  ITT Item ISMOD (Andry)
            'Khusus modis / rak Ecommerce tidak perlu dimasukan kedalam list yang harus dicek
            Madp.SelectCommand.CommandText = "select distinct r.noshelf from rak r, prodmast p"
            Madp.SelectCommand.CommandText &= " where r.prdcd = p.prdcd"
            Madp.SelectCommand.CommandText &= " and r.nama_rak = '" & Modis & "'"
            Madp.SelectCommand.CommandText &= " and r.noshelf not in (select noshelf from cekplanogram where namarak = '" & Modis & "'"
            Madp.SelectCommand.CommandText &= " and date(tglscan) = date(now()) and jenisbarang <> 'SD')"
            Madp.SelectCommand.CommandText &= " and p.flagprod not like '%eco = Y%'"
            Madp.SelectCommand.CommandText &= " order by r.noshelf asc;"
            Madp.Fill(Rtn)
        Catch ex As Exception
            CountModis = 0
            Rtn = Nothing
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
    Public Function GetDeskripsiPlanogram(ByVal tabel_name As String, ByVal barcode_plu As String, _
                                          ByVal NamaRak As String, ByVal ListShelf As String, ByVal User As ClsUser) As ClsPlanogram
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Result As New ClsPlanogram

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

                'Ambil data planogram
                'Revisi: 20 Juni 2019
                'MaxBatasRetur diganti dengan BatasPajang (Format Tanggal)
                Dim DtPlano As New DataTable
                Dim MaxBatasRetur As String = ""
                Dim MaxBatasReturS As String = ""
                Dim BatasPajang As String = ""
                Dim Price As String = ""

                'Mcom.CommandText = "SELECT r.PRDCD, p.DESC2, r.NORAK, r.NAMA_RAK, r.NOSHELF, "
                'Mcom.CommandText &= " IFNULL(t.PRDCD,IF(o.tgl_akh>=CURDATE(),o.prdcd,NULL)) AS PLU_PTAG, "
                'Mcom.CommandText &= " t.TGL_AKH, IFNULL(br.MAX_RET_TOKO2DCI,0) AS MAX_RET_TOKO2DCI, br.MAX_RET_TOKO2DCI_S, p.PRICE, "
                'Mcom.CommandText &= " IFNULL(t.PRICE,IF(o.tgl_akh>=CURDATE(),o.promosi,NULL)) AS PRICE_PTAG, "
                'Mcom.CommandText &= " t.PROMOSI, s.QTY, p.DEPART, r.KIRIKANAN "
                'Mcom.CommandText &= " FROM prodmast p LEFT JOIN rak r ON p.PRDCD = r.PRDCD "
                'Mcom.CommandText &= " LEFT JOIN stmast s ON p.PRDCD = s.PRDCD "
                'Mcom.CommandText &= " LEFT JOIN barcode b ON p.PRDCD = b.PLU "
                'Mcom.CommandText &= " LEFT JOIN ptag t ON p.PRDCD = t.PRDCD "
                'Mcom.CommandText &= " LEFT JOIN ptag_old o ON p.PRDCD = o.PRDCD "
                'Mcom.CommandText &= " LEFT JOIN batas_retur br ON p.PRDCD = br.FMKODE "
                'Mcom.CommandText &= " WHERE b.PLU = '" & barcode_plu & "' OR b.BARCD = '" & barcode_plu & "';"

                'Revisi terkait BatasPajang, perubahan query dari Andry
                'Mcom.CommandText = "SELECT r.PRDCD, p.DESC2, r.NORAK, r.NAMA_RAK, r.NOSHELF, "
                'Mcom.CommandText &= " IFNULL(t.PRDCD,IF(o.tgl_akh>=CURDATE(),o.prdcd,NULL)) AS PLU_PTAG, "
                'Mcom.CommandText &= " t.TGL_AKH, IFNULL(br.MAX_RET_TOKO2DCI,0) AS MAX_RET_TOKO2DCI, br.MAX_RET_TOKO2DCI_S, p.PRICE, "
                'Mcom.CommandText &= " IFNULL(t.PRICE,IF(o.tgl_akh>=CURDATE(),o.promosi,NULL)) AS PRICE_PTAG, "
                'Mcom.CommandText &= " t.PROMOSI, s.QTY, p.DEPART, r.KIRIKANAN, "
                'Mcom.CommandText &= " CAST(DATE_FORMAT(DATE_ADD(CURDATE(),INTERVAL IF(max_ret_toko2dci_s='B'," & _
                '                    "(max_ret_toko2dci*30),max_ret_toko2dci) DAY),'%d-%m-%Y') AS CHAR) AS Tanggal_Batas_Aman "
                'Mcom.CommandText &= " FROM prodmast p LEFT JOIN rak r ON p.PRDCD = r.PRDCD "
                'Mcom.CommandText &= " LEFT JOIN stmast s ON p.PRDCD = s.PRDCD "
                'Mcom.CommandText &= " LEFT JOIN barcode b ON p.PRDCD = b.PLU "
                'Mcom.CommandText &= " LEFT JOIN ptag t ON p.PRDCD = t.PRDCD "
                'Mcom.CommandText &= " LEFT JOIN ptag_old o ON p.PRDCD = o.PRDCD "
                'Mcom.CommandText &= " LEFT JOIN batas_retur br ON p.PRDCD = br.FMKODE "
                'Mcom.CommandText &= " WHERE b.PLU = '" & barcode_plu & "' OR b.BARCD = '" & barcode_plu & "'"
                ''Mcom.CommandText &= " AND br.max_ret_toko2dci_s IN ('B','H');"
                'Mcom.CommandText &= " AND (br.max_ret_toko2dci_s IN ('B','H') OR br.max_ret_toko2dci_s IS NULL);"

                'Revisi Scan Planogram tambah filter NAMA_RAK (Request Pak Beny & Pak YYN)
                Mcom.CommandText = "SELECT r.PRDCD, p.DESC2, r.NORAK, r.NAMA_RAK, r.NOSHELF, "
                Mcom.CommandText &= " IFNULL(t.PRDCD,IF(o.tgl_akh>=CURDATE(),o.prdcd,NULL)) AS PLU_PTAG, "
                Mcom.CommandText &= " t.TGL_AKH, IFNULL(br.MAX_RET_TOKO2DCI,0) AS MAX_RET_TOKO2DCI, br.MAX_RET_TOKO2DCI_S, p.PRICE, "
                Mcom.CommandText &= " IFNULL(t.PRICE,IF(o.tgl_akh>=CURDATE(),o.promosi,NULL)) AS PRICE_PTAG, "
                Mcom.CommandText &= " t.PROMOSI, s.QTY, p.DEPART, r.KIRIKANAN, "
                Mcom.CommandText &= " CAST(DATE_FORMAT(DATE_ADD(CURDATE(),INTERVAL IF(max_ret_toko2dci_s='B'," & _
                                    "(max_ret_toko2dci*30),max_ret_toko2dci) DAY),'%d-%m-%Y') AS CHAR) AS Tanggal_Batas_Aman "
                Mcom.CommandText &= " FROM prodmast p LEFT JOIN "
                Mcom.CommandText &= " (SELECT * FROM rak WHERE NAMA_RAK = '" & NamaRak & "') r ON p.PRDCD = r.PRDCD "
                Mcom.CommandText &= " LEFT JOIN stmast s ON p.PRDCD = s.PRDCD "
                Mcom.CommandText &= " LEFT JOIN barcode b ON p.PRDCD = b.PLU "
                Mcom.CommandText &= " LEFT JOIN ptag t ON p.PRDCD = t.PRDCD "
                Mcom.CommandText &= " LEFT JOIN ptag_old o ON p.PRDCD = o.PRDCD "
                Mcom.CommandText &= " LEFT JOIN batas_retur br ON p.PRDCD = br.FMKODE "
                Mcom.CommandText &= " WHERE b.PLU = '" & barcode_plu & "' OR b.BARCD = '" & barcode_plu & "'"
                Mcom.CommandText &= " AND (br.max_ret_toko2dci_s IN ('B','H') OR br.max_ret_toko2dci_s IS NULL);"
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
                    Result.Prdcd = DtPlano.Rows(0).Item("PRDCD")
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

                    Mcom.CommandText = "SELECT PLU FROM " & tabel_name
                    Mcom.CommandText &= " WHERE PLU = '" & DtPlano.Rows(0).Item("PRDCD") & "'"
                    Mcom.CommandText &= " AND DATE(TGLSCAN) = CURDATE()"
                    Mcom.CommandText &= " AND `STATUS` <> 'I' AND NORAKINPUT = '" & DtPlano.Rows(0).Item("NORAK") & "'"
                    Mcom.CommandText &= " AND NOSHELFINPUT = '" & DtPlano.Rows(0).Item("NOSHELF") & "';"
                    Dim sDap2 As New MySqlDataAdapter(Mcom)
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
                        Mcom.ExecuteNonQuery()
                    End If
                End If

            Catch ex As Exception
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetDeskripsiPlanogram", Conn)
                utility.TraceLogTxt("Error - GetDeskripsiPlanogram " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                Conn.Close()
            End Try
        End SyncLock

        Return Result
    End Function

    ''' <summary>
    ''' proses selesai cek Planogram
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SelesaiCekPlano(ByVal tabel_name As String, ByVal NamaRak As String, _
                                    ByVal ListShelf As String, ByVal User As ClsUser) As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim DtCP As New DataTable
        Dim Rtn As New Boolean

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Madp.SelectCommand.CommandText = "SELECT r.PRDCD,p.DESC2,r.NOSHELF,r.NORAK,r.NAMA_RAK,r.KIRIKANAN,"
            Madp.SelectCommand.CommandText &= " p.DEPART,s.QTY,s.QTY,IFNULL(b.MAX_RET_TOKO2DCI,0) AS MAX_RET_TOKO2DCI,b.MAX_RET_TOKO2DCI_S"
            'Madp.SelectCommand.CommandText &= " FROM rak r JOIN prodmast p ON r.PRDCD = p.PRDCD JOIN cekplanogram c"
            Madp.SelectCommand.CommandText &= " FROM rak r JOIN prodmast p ON r.PRDCD = p.PRDCD"
            Madp.SelectCommand.CommandText &= " JOIN stmast s ON s.PRDCD = r.PRDCD"
            Madp.SelectCommand.CommandText &= " LEFT JOIN batas_retur b ON p.PRDCD = b.FMKODE"
            Madp.SelectCommand.CommandText &= " WHERE (r.PRDCD,r.NORAK,r.NOSHELF) NOT IN (SELECT PLU,NORAKINPUT,NOSHELFINPUT"
            Madp.SelectCommand.CommandText &= " FROM cekplanogram WHERE DATE(TGLSCAN) = CURDATE())"
            Madp.SelectCommand.CommandText &= " AND r.NAMA_RAK = '" & NamaRak & "'"
            Madp.SelectCommand.CommandText &= " AND r.NOSHELF IN (" & ListShelf & ")"
            Madp.SelectCommand.CommandText &= " GROUP BY r.PRDCD;"
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
                    Mcom.CommandText &= " '" & Dr("MAX_RET_TOKO2DCI_S") & "', '" & NamaRak & "'),"
                Next
                Mcom.CommandText = Mcom.CommandText.Remove(Mcom.CommandText.Length - 1, 1)
                Mcom.ExecuteNonQuery()
            End If

            DtCP = New DataTable
            Madp.SelectCommand.CommandText = "SELECT PLU, jenisbarang FROM " & tabel_name
            Madp.SelectCommand.CommandText &= " WHERE DATE(tglscan) = CURDATE()"
            Madp.SelectCommand.CommandText &= " AND NIK = '" & User.ID & "';"
            Madp.Fill(DtCP)
            
            If DtCP.Rows.Count > 0 Then
                For Each Dr As DataRow In DtCP.Rows
                    Mcom.CommandText = "UPDATE " & tabel_name & " c, ("
                    Mcom.CommandText &= " SELECT r.prdcd AS plu, r.nama_rak AS rak, p.desc2 AS nama,"
                    Mcom.CommandText &= " r.noshelf AS noshelf, r.norak AS norak"
                    Mcom.CommandText &= " FROM rak r, prodmast p"
                    Mcom.CommandText &= " WHERE(r.prdcd = p.prdcd)"
                    If Dr("jenisbarang").ToString.ToUpper = "SD" Then
                        Mcom.CommandText &= " AND r.prdcd = '" & Dr("PLU") & "'"
                    Else
                        Mcom.CommandText &= " AND r.prdcd = '" & Dr("PLU") & "' AND r.nama_rak = '" & NamaRak & "'"
                    End If
                    Mcom.CommandText &= ") t"
                    Mcom.CommandText &= " SET c.nama = t.nama, c.namarak = t.rak,"
                    Mcom.CommandText &= " c.noshelf = t.noshelf, c.norak = t.norak"
                    Mcom.CommandText &= " WHERE(C.plu = t.plu);"
                    Mcom.ExecuteNonQuery()
                Next
            End If

            Mcom.CommandText = "SELECT COUNT(*) FROM " & tabel_name & " WHERE STATUS = 'S'"
            Mcom.CommandText &= " AND DATE(tglscan) = CURDATE() AND NIK = '" & User.ID & "';"
            If Mcom.ExecuteScalar > 0 Then
                Rtn = True
            End If

            Mcom.CommandText = "UPDATE " & tabel_name & " SET `Status` = 'I'"
            Mcom.CommandText &= " WHERE DATE(tglscan) = CURDATE() AND `Status` = 'S' AND NIK = '" & User.ID & "';"
            Mcom.ExecuteNonQuery()
        Catch ex As Exception
            Rtn = False
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "SelesaiCekPlano", Conn)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function CekPlano(ByVal tabel_name As String, ByVal NamaRak As String,
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
            Madp.SelectCommand.CommandText = " SELECT PLU,NORAKINPUT,NOSHELFINPUT FROM cekplanogram WHERE DATE(TGLSCAN) = CURDATE() AND NAMARAKINPUT = '" & NamaRak & "'"
            Madp.SelectCommand.CommandText &= " GROUP BY PLU;"
            Madp.Fill(DtCP)
            Console.WriteLine(Madp.SelectCommand.CommandText)
            If DtCP.Rows.Count = 0 Then
                Rtn = ""
            Else
                Rtn = "Ada"
            End If

        Catch ex As Exception
            Rtn = "Err"
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "SelesaiCekPlano", Conn)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function
End Class
