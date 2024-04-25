Imports MySql.Data.MySqlClient
Imports PJR.FormMain
Imports IDM.Fungsi
Imports IDM.InfoToko

Public Class ClsProdukController
    Private utility As New Utility

    Public Function GetDeskripsiProdukSO(ByVal tabel_name As String, ByVal barcode_plu As String,
                                         ByVal ketInput As String, Optional isSOIC As Boolean = False,
                                         Optional mainCBR As Boolean = False, Optional flagTTL3 As Boolean = False,
                                         Optional lokasi_so As String = "") As ClsSo

        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim tmpDt As New DataTable
        Dim result As New ClsSo
        Dim Scom As New MySqlCommand("", Scon)

        Dim no_rak As String = ""
        Dim no_shelf As String = ""

        Dim cSOIC As New ClsSOICController

        If Scon Is Nothing Then
            utility.TraceLogTxt("Error - GetDeskripsiProduk (connection Nothing) " & vbCrLf & "PLU:" & barcode_plu)
            Return result
            Exit Function
        End If

        SyncLock Scon
            Try
                If Scon.State = ConnectionState.Closed Then
                    Scon.Open()
                End If

                Scom.CommandText = "SELECT distinct T.TIPERAK,T.NORAK,T.NOSHELF,T.PRDCD,T.SINGKAT,T.TTL,T.TTL1,T.TTL2,"

                If isSOIC And flagTTL3 Then
                    Scom.CommandText &= "TTL3,"
                End If

                Scom.CommandText &= "T.SOID,T.SOTIME,T.DCP,T.KIRIKANAN,T.Unit,T.COM+T.BPB-T.RETUR_K-T.SALES+T.RETUR+T.BPB_2+T.ADJ-T.TTL2 AS COM"
                Scom.CommandText &= " FROM " & tabel_name & " T left join BARCODE B "
                Scom.CommandText &= " on T.PRDCD = B.PLU "
                Scom.CommandText &= " WHERE B.BARCD = '" & barcode_plu & "' or T.PRDCD ='" & barcode_plu & "'"
                Scom.CommandText &= " ORDER BY T.NORAK,T.NOSHELF,T.TIPERAK,T.KIRIKANAN;"
                TraceLog("GetDeskripsiProdukSO-Q1: " & Scom.CommandText)
                Dim sDap As New MySqlDataAdapter(Scom)
                sDap.Fill(tmpDt)

                result.BarcodePlu = barcode_plu
                result.statusBarcode = ""

                If tmpDt.Rows.Count > 0 Then
                    'Revisi 20 November 2019 (Memo 1081/CPS/19)
                    'Hitung lokasi RAK untuk item, Jika ada lebih dari 1 lokasi aktifkan fitur NEXT WDCP
                    Scom.CommandText = "SELECT COUNT(NORAK) FROM RAK R"
                    Scom.CommandText &= " LEFT JOIN BARCODE B  ON R.PRDCD = B.PLU"
                    Scom.CommandText &= " WHERE (B.BARCD = '" & barcode_plu & "' OR R.PRDCD ='" & barcode_plu & "') AND B.QTY=1;"
                    TraceLog("GetDeskripsiProdukSO-Q2: " & Scom.CommandText)

                    Dim CountRak = Scom.ExecuteScalar
                    If Not IsDBNull(CountRak) Then
                        result.TotalRak = CountRak
                    Else
                        result.TotalRak = 0
                    End If

                    If isSOIC And mainCBR Then
                        Scom.CommandText = "SELECT COUNT(*) FROM PRODMAST P"
                        Scom.CommandText &= " LEFT JOIN BARCODE B  ON P.PRDCD = B.PLU"
                        Scom.CommandText &= " WHERE (B.BARCD = '" & barcode_plu & "' OR P.PRDCD ='" & barcode_plu & "')"
                        Scom.CommandText &= " AND FLAGPROD LIKE '%CBR=Y%' AND B.QTY=1;"
                        TraceLog("GetDeskripsiProdukSO-Q2: " & Scom.CommandText)
                        If Scom.ExecuteScalar <> "0" Then
                            result.statusBarcode = "CBRY"
                        Else
                            result.statusBarcode = "CBRN"
                        End If
                    End If

                    result.PRDCD = tmpDt.Rows(0)("PRDCD")

                    If isSOIC And flagTTL3 And lokasi_so = "Barang Rusak" Then
                        If cSOIC.isItemActive(result.PRDCD) = False Then
                            result.Deskripsi = "Bukan Item Aktif"
                            GoTo GAGAL
                        ElseIf cSOIC.cekFisikStmast(result.PRDCD) = 0 Then
                            result.Deskripsi = "Tidak Ada Fisik"
                            GoTo GAGAL
                        ElseIf cSOIC.cekAcostItem(result.PRDCD) = 0 Then
                            result.Deskripsi = "Nilai HPP Kosong"
                            GoTo GAGAL
                        Else
                            If cSOIC.isItemBKL(result.PRDCD) Then
                                If cSOIC.isBarangPutus(result.PRDCD) Then
                                    If cSOIC.IsItemBAP(result.PRDCD) Then
                                        result.isBADraft = True
                                        result.isWtran = False
                                    Else
                                        result.isBADraft = False
                                        result.isWtran = False
                                        result.Deskripsi = "Item Tidak Valid"
                                        GoTo GAGAL
                                    End If
                                Else
                                    result.isBADraft = False
                                    result.isWtran = False
                                    result.Deskripsi = "Item Tidak Valid"
                                    GoTo GAGAL
                                End If
                            Else
                                If cSOIC.IsItemBAP(result.PRDCD) Then
                                    result.isBADraft = True
                                    result.isWtran = False
                                Else
                                    If cSOIC.IsValidNonBAP(result.PRDCD) Then
                                        result.isBADraft = False
                                        result.isWtran = True
                                    Else
                                        result.isBADraft = False
                                        result.isWtran = False
                                        result.Deskripsi = "Item Tidak Valid"
                                        GoTo GAGAL
                                    End If
                                End If
                            End If

                        End If
                    End If

                    no_rak = CInt(tmpDt.Rows.Item(0)("NORAK"))
                    no_rak = no_rak.PadLeft(3, "0")
                    no_shelf = CInt(tmpDt.Rows.Item(0)("NOSHELF"))
                    no_shelf = no_shelf.PadLeft(3, "0")

                    result.Unit = tmpDt.Rows(0)("Unit")

                    result.Deskripsi = tmpDt.Rows(0)("SINGKAT")
                    If result.Deskripsi.Length > 20 Then
                        result.Deskripsi = result.Deskripsi.Substring(0, 20)
                    End If

                    result.Rak = no_rak & "/" & no_shelf
                    result.QTYToko = tmpDt.Rows(0)("TTL1")
                    result.QTYGudang = tmpDt.Rows(0)("TTL2")
                    result.QTYTotal = tmpDt.Rows(0)("TTL")
                    result.QTYCom = tmpDt.Rows(0)("COM")
                Else
                    result.Deskripsi = "Tidak Ditemukan"
GAGAL:
                    result.PRDCD = ""
                    result.Unit = ""
                    result.Rak = ""
                    result.QTYToko = ""
                    result.QTYGudang = ""
                    result.QTYTotal = ""
                    result.QTYCom = ""
                    If isSOIC And mainCBR Then
                        result.statusBarcode = "CBRY"
                    End If
                End If

            Catch ex As Exception
                TraceLog("Last Query: " & Scom.CommandText)
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetDeskripsiProduk", Scon)
                utility.TraceLogTxt("Error - GetDeskripsiProduk " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                Scon.Close()
            End Try

        End SyncLock

        Return result
    End Function

    Public Function GetDeskripsiProdukBPB(ByVal tabel_name As String, ByVal container_no As String, ByVal barcode_plu As String) As ClsBPB
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim tmpDt As New DataTable
        Dim result As New ClsBPB
        Dim Scom As New MySqlCommand("", Scon)

        Dim no_rak As String = ""
        Dim no_shelf As String = ""

        If Scon Is Nothing Then
            utility.TraceLogTxt("Error - GetDeskripsiProduk (connection Nothing) " & vbCrLf & "PLU:" & barcode_plu)
            Return result
            Exit Function
        End If

        SyncLock Scon
            Try
                If Scon.State = ConnectionState.Closed Then
                    Scon.Open()
                End If

                Scom.CommandText = "SELECT * FROM " & tabel_name & " WHERE dus_no = '" & container_no & "'"
                Scom.CommandText &= " AND prdcd = '" & barcode_plu & "';"

                Dim sDap As New MySqlDataAdapter(Scom)
                sDap.Fill(tmpDt)

                result.Prdcd = barcode_plu
                If tmpDt.Rows.Count > 0 Then
                    result.Prdcd = tmpDt.Rows(0)("PRDCD")
                    result.Desc = tmpDt.Rows(0)("NAMA")
                    result.Qty = tmpDt.Rows(0)("qty")
                Else
                    result.Prdcd = ""
                    result.Desc = "Tidak Ditemukan"
                    result.Qty = ""
                End If

            Catch ex As Exception
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetDeskripsiProduk", Scon)
                utility.TraceLogTxt("Error - GetDeskripsiProduk " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                Scon.Close()
            End Try

        End SyncLock

        Return result
    End Function

    Public Function GetListRakProdukSO(ByVal barcode_plu As String) As DataTable
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim RtnDt As New DataTable
        Dim mcom As New MySqlCommand("", conn)

        If conn Is Nothing Then
            utility.TraceLogTxt("Error - GetListRakProdukSO (connection Nothing) " & vbCrLf & "PLU:" & barcode_plu)
            Return New DataTable
            Exit Function
        End If

        SyncLock conn
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If

                mcom.CommandText = "SET @row_number = 0; "
                mcom.ExecuteNonQuery()
                mcom.CommandText = "SELECT DISTINCT R.TIPERAK,R.NORAK,R.NOSHELF, (@row_number:=@row_number + 1) AS NO"
                mcom.CommandText &= " FROM RAK R"
                mcom.CommandText &= " LEFT JOIN BARCODE B ON R.PRDCD = B.PLU"
                mcom.CommandText &= " WHERE B.BARCD = '" & barcode_plu & "' OR R.PRDCD ='" & barcode_plu & "';"

                Dim sDap As New MySqlDataAdapter(mcom)
                sDap.Fill(RtnDt)

            Catch ex As Exception
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetListRakProdukSO", conn)
                utility.TraceLogTxt("Error - GetListRakProdukSO " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                conn.Close()
            End Try
        End SyncLock

        Return RtnDt
    End Function

    Public Function UpdateTotalProdukSO(ByVal tabel_name As String, ByVal barcode_plu As String, ByVal qty_total As String, ByVal lokasi As String, Optional ByVal jenis_so As String = "") As Boolean
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)
        Dim result As Boolean = False
        Dim shift As String = ""
        If Scon Is Nothing Then
            utility.TraceLogTxt("Error - UpdateTotalProduk_ (connection Nothing)" & vbCrLf & "PLU:" & barcode_plu)
            Return result
            Exit Function
        End If

        SyncLock Scon
            Try
                If Scon.State = ConnectionState.Closed Then
                    Scon.Open()
                End If

                If jenis_so = "Khusus" Then
                    'memo 209/cps/23
                    'tambah baca kolom shift
                    shift = GetShift()
                End If

                If qty_total <> 0 Then
                    If lokasi = "Toko" Then
                        If jenis_so = "Khusus" Then
                            'memo 209/cps/23
                            'tambah baca kolom shift
                            Scom.CommandText = "UPDATE " & tabel_name & " SET TTL1 = " & qty_total & ", SOID = 'T',SOTIME = CURTIME(), DCP = '3' "
                            Scom.CommandText &= " WHERE ((PRDCD IN (SELECT plu FROM barcode WHERE barcd='" & barcode_plu & "')) OR PRDCD = '" & barcode_plu & "') AND SHIFT = '" & shift & "';"
                        Else
                            Scom.CommandText = "UPDATE " & tabel_name & " SET TTL1 = TTL1 +" & qty_total & ", SOID = 'T',SOTIME = CURTIME(), DCP = '3' "
                            Scom.CommandText &= " WHERE (PRDCD IN (SELECT plu FROM barcode WHERE barcd='" & barcode_plu & "')) OR PRDCD = '" & barcode_plu & "';"
                        End If
                    ElseIf lokasi = "Gudang" Then
                        If FormMain.Toko.Kode.ToUpper.StartsWith("B") Or FormMain.Toko.Kode.ToUpper.StartsWith("P") Then
                            Scom.CommandText = "UPDATE " & tabel_name & " SET TTL2 = TTL2 + " & qty_total & ", SOID = 'G', DCP = '3'  "
                        Else
                            Scom.CommandText = "UPDATE " & tabel_name & " SET TTL2 = TTL2 + " & qty_total & ", SOID = 'G', DCP = '3',SOTIME = CURTIME() "
                        End If
                        'memo 209/cps/23
                        'tambah baca kolom shift
                        If jenis_so = "Khusus" Then
                            Scom.CommandText &= " WHERE ((PRDCD IN (SELECT plu FROM barcode WHERE barcd='" & barcode_plu & "')) OR PRDCD = '" & barcode_plu & "') AND SHIFT = '" & shift & "';"
                        Else
                            Scom.CommandText &= " WHERE (PRDCD IN (SELECT plu FROM barcode WHERE barcd='" & barcode_plu & "')) OR PRDCD = '" & barcode_plu & "';"
                        End If
                    Else
                        Scom.CommandText = "UPDATE " & tabel_name & " SET TTL3 = TTL3 + " & qty_total & ", SOID = 'R', DCP = '3',SOTIME = CURTIME() "
                        Scom.CommandText &= " WHERE (PRDCD IN (SELECT plu FROM barcode WHERE barcd='" & barcode_plu & "')) OR PRDCD = '" & barcode_plu & "';"
                    End If
                Else
                    If lokasi = "Toko" Then
                        'memo 209/cps/23
                        'tambah baca kolom shift
                        If jenis_so = "Khusus" Then
                            Scom.CommandText = "UPDATE " & tabel_name & " SET TTL1 = " & qty_total & ", SOID = 'T',SOTIME = CURTIME(), DCP = '3' "
                            Scom.CommandText &= " WHERE ((PRDCD IN (SELECT plu FROM barcode WHERE barcd='" & barcode_plu & "')) OR PRDCD = '" & barcode_plu & "')  AND SHIFT = '" & shift & "';"
                        Else
                            Scom.CommandText = "UPDATE " & tabel_name & " SET TTL1 = " & qty_total & ", SOID = 'T',SOTIME = CURTIME(), DCP = '3' "
                            Scom.CommandText &= " WHERE (PRDCD IN (SELECT plu FROM barcode WHERE barcd='" & barcode_plu & "')) OR PRDCD = '" & barcode_plu & "';"
                        End If
                    ElseIf lokasi = "Gudang" Then
                        If FormMain.Toko.Kode.ToUpper.StartsWith("B") Or FormMain.Toko.Kode.ToUpper.StartsWith("P") Then
                            Scom.CommandText = "UPDATE " & tabel_name & " SET TTL2 = " & qty_total & ", SOID = 'G' , DCP = '3' "
                        Else
                            Scom.CommandText = "UPDATE " & tabel_name & " SET TTL2 = " & qty_total & ", SOID = 'G',SOTIME = CURTIME(), DCP = '3' "
                        End If
                        'memo 209/cps/23
                        'tambah baca kolom shift
                        If jenis_so = "Khusus" Then
                            Scom.CommandText &= " WHERE ((PRDCD IN (SELECT plu FROM barcode WHERE barcd='" & barcode_plu & "')) OR PRDCD = '" & barcode_plu & "') AND SHIFT = '" & shift & "' ;"
                        Else
                            Scom.CommandText &= " WHERE (PRDCD IN (SELECT plu FROM barcode WHERE barcd='" & barcode_plu & "')) OR PRDCD = '" & barcode_plu & "';"
                        End If
                    Else
                        Scom.CommandText = "UPDATE " & tabel_name & " SET TTL3 = " & qty_total & ", SOID = 'R',SOTIME = CURTIME(), DCP = '3' "
                        Scom.CommandText &= " WHERE (PRDCD IN (SELECT plu FROM barcode WHERE barcd='" & barcode_plu & "')) OR PRDCD = '" & barcode_plu & "';"
                    End If
                End If

                Scom.ExecuteNonQuery()
                result = True
                utility.Tracelog("Debug", "QTY_TOTAL : " & qty_total, "UpdateTotalProduk", Scon)

            Catch ex As Exception
                result = False
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "UpdateTotalProduk", Scon)
                utility.TraceLogTxt("Error - UpdateTotalProduk " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                Scon.Close()
            End Try
        End SyncLock

        Return result
    End Function

    Public Function SkipProduk(ByVal tabel_name As String, ByVal barcode_plu As String, ByVal lokasi As String) As Boolean
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim result As Boolean = False
        Dim Scom As New MySqlCommand("", Scon)
        Dim shift As String
        If Scon Is Nothing Then
            utility.TraceLogTxt("Error - SkipProduk (connection Nothing)" & vbCrLf & "PLU:" & barcode_plu)
            Return result
            Exit Function
        End If

        SyncLock Scon
            Try
                If Scon.State = ConnectionState.Closed Then
                    Scon.Open()
                End If
                'memo 209/cps/23
                'tambah baca kolom shift
                If tabel_name.StartsWith("SP") Then
                    shift = GetShift()

                    If lokasi = "Toko" Then
                        Scom.CommandText = "UPDATE " & tabel_name & " SET SOID = 'T',SOTIME = CURTIME(), DCP = '3' "
                        Scom.CommandText &= " WHERE ((PRDCD IN (SELECT plu FROM barcode WHERE barcd='" & barcode_plu & "')) OR PRDCD = '" & barcode_plu & "') AND SHIFT = '" & shift & "';"
                    Else
                        Scom.CommandText = "UPDATE " & tabel_name & " SET SOID = 'G',SOTIME = CURTIME(), DCP = '3' "
                        Scom.CommandText &= " WHERE ((PRDCD IN (SELECT plu FROM barcode WHERE barcd='" & barcode_plu & "')) OR PRDCD = '" & barcode_plu & "') AND SHIFT = '" & shift & "';"
                    End If
                Else
                    If lokasi = "Toko" Then
                        Scom.CommandText = "UPDATE " & tabel_name & " SET SOID = 'T',SOTIME = CURTIME(), DCP = '3' "
                        Scom.CommandText &= " WHERE (PRDCD IN (SELECT plu FROM barcode WHERE barcd='" & barcode_plu & "')) OR PRDCD = '" & barcode_plu & "';"
                    ElseIf lokasi = "Gudang" Then
                        Scom.CommandText = "UPDATE " & tabel_name & " SET SOID = 'G',SOTIME = CURTIME(), DCP = '3' "
                        Scom.CommandText &= " WHERE (PRDCD IN (SELECT plu FROM barcode WHERE barcd='" & barcode_plu & "')) OR PRDCD = '" & barcode_plu & "';"
                    Else
                        Scom.CommandText = "UPDATE " & tabel_name & " SET SOID = 'R',SOTIME = CURTIME(), DCP = '3' "
                        Scom.CommandText &= " WHERE (PRDCD IN (SELECT plu FROM barcode WHERE barcd='" & barcode_plu & "')) OR PRDCD = '" & barcode_plu & "';"
                    End If
                End If

                Scom.ExecuteNonQuery()
                result = True

            Catch ex As Exception
                result = False
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "SkipProduk", Scon)
                utility.TraceLogTxt("Error - SkipProduk " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                Scon.Close()
            End Try
        End SyncLock

        Return result
    End Function

    Public Function LihatTabelSO(ByVal tabel_name As String, ByVal mode_run As String) As DataTable
        'Dim connection As New ClsConnection
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)
        Dim tmpDt As New DataTable
        Dim shift As String

        Dim cVirBacaprod As New ClsVirBacaprodController

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If
            If tabel_name.StartsWith("SP") Then
                'memo 209/cps/23
                'tambah baca kolom shift
                shift = GetShift()

                If mode_run.ToUpper = "E" Then
                    Scom.CommandText = "SELECT TIPERAK, NORAK, NOSHELF, KIRIKANAN, PRDCD, SINGKAT FROM " & tabel_name
                    Scom.CommandText &= " WHERE (SOID <> 'T' and SOID <> 'G') AND SHIFT = '" & shift & "' ORDER BY NORAK, NOSHELF, PRDCD;"
                Else
                    Scom.CommandText = "SELECT TIPERAK, NORAK, NOSHELF, KIRIKANAN, PRDCD, SINGKAT FROM " & tabel_name
                    Scom.CommandText &= " WHERE (NORAK,NOSHELF) NOT IN ("
                    Scom.CommandText &= " SELECT DISTINCT NORAK, NOSHELF FROM " & tabel_name
                    Scom.CommandText &= " WHERE SOID = 'T' AND SHIFT = '" & shift & "') AND COM <> 0 AND SHIFT = '" & shift & "';"
                End If
            Else
                If mode_run.ToUpper = "E" Then
                    Scom.CommandText = "SELECT TIPERAK, NORAK, NOSHELF, KIRIKANAN, PRDCD, SINGKAT FROM " & tabel_name
                    If cVirBacaprod.get1230TTL3_Virbacaprod = "ON" And FormMain.jenis_so = "BIC" Then
                        Scom.CommandText &= " WHERE (SOID <> 'T' AND SOID <> 'G' AND SOID <> 'R') "
                    Else
                        Scom.CommandText &= " WHERE (SOID <> 'T' and SOID <> 'G') "
                    End If
                    Scom.CommandText &= "ORDER BY NORAK, NOSHELF, PRDCD;"
                Else
                    Scom.CommandText = "SELECT TIPERAK, NORAK, NOSHELF, KIRIKANAN, PRDCD, SINGKAT FROM " & tabel_name
                    Scom.CommandText &= " WHERE (NORAK,NOSHELF) NOT IN ("
                    Scom.CommandText &= " SELECT DISTINCT NORAK, NOSHELF FROM " & tabel_name
                    Scom.CommandText &= " WHERE SOID = 'T') AND COM <> 0;"
                End If
            End If

            Dim sDap As New MySqlDataAdapter(Scom)
            sDap.Fill(tmpDt)

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "LihatTabelSO", Scon)
        Finally
            Scon.Close()
        End Try

        Return tmpDt
    End Function

    Public Function LihatTabelSO_Bazar(ByVal tabel_name As String, ByVal mode_run As String) As DataTable
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)
        Dim tmpDt As New DataTable

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If


            Scom.CommandText = "SELECT PRDCD, SINGKAT, BULAN_EXP FROM " & tabel_name & "
            WHERE (PRDCD) Not In (
             SELECT DISTINCT PRDCD FROM " & tabel_name & "
            WHERE RECID = 'P' OR RECID = 'B') "


            Dim sDap As New MySqlDataAdapter(Scom)
            sDap.Fill(tmpDt)

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "LihatTabelSO", Scon)
        Finally
            Scon.Close()
        End Try

        Return tmpDt
    End Function

    Public Function LihatTabelSO2(ByVal tabel_name As String, ByVal no_rak As String) As DataTable
        'Dim connection As New ClsConnection
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)
        Dim tmpDt As New DataTable

        Dim cVirBacaprod As New ClsVirBacaprodController

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SELECT TIPERAK, NORAK, NOSHELF, KIRIKANAN,PRDCD, SINGKAT FROM " & tabel_name
            Scom.CommandText &= " WHERE NORAK =  " & no_rak & " AND "
            If cVirBacaprod.get1230TTL3_Virbacaprod And FormMain.jenis_so = "BIC" Then
                Scom.CommandText &= "( SOID <> 'T' AND SOID <> 'G' AND SOID <> 'R') "
            Else
                Scom.CommandText &= "( SOID <> 'T' AND SOID <> 'G') "
            End If
            Scom.CommandText &= "ORDER BY NORAK, NOSHELF, PRDCD;"

            Dim sDap As New MySqlDataAdapter(Scom)
            sDap.Fill(tmpDt)

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "LihatTabelSO", Scon)
        Finally
            Scon.Close()
        End Try

        Return tmpDt
    End Function

    Public Function FindTableSO(ByVal tabel_name As String) As String
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim result As String = "0|0|0"
        Dim Scom As New MySqlCommand("", Scon)

        Dim total As Integer
        Dim total_soid As Integer
        Dim total_soid_toko As Integer

        Dim cVirBacaprod As New ClsVirBacaprodController

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If
            If tabel_name.ToUpper.Contains("SZ") Then
                'total SZ
                Try
                    Scom.CommandText = "SELECT COUNT(*) FROM " & tabel_name
                    total = Scom.ExecuteScalar
                Catch
                    total = 0
                End Try
                'tidak dipakai
                Try
                    Scom.CommandText = "SELECT COUNT(*) FROM " & tabel_name
                    Scom.CommandText &= " WHERE SOID = 'A' "
                    total_soid = Scom.ExecuteScalar
                Catch
                    total_soid = 0
                End Try
            ElseIf tabel_name.ToUpper.StartsWith("SE") Then
                'total SE
                Try
                    Scom.CommandText = "SELECT COUNT(*) FROM " & tabel_name
                    total = Scom.ExecuteScalar
                Catch
                    total = 0
                End Try
                'tidak dipakai
                Try
                    Scom.CommandText = "SELECT COUNT(*) FROM " & tabel_name
                    Scom.CommandText &= " WHERE STATUS='';"
                    total_soid = Scom.ExecuteScalar
                Catch
                    total_soid = 0
                End Try
            Else
                Try
                    Scom.CommandText = "SELECT COUNT(*) FROM " & tabel_name
                    total = Scom.ExecuteScalar
                Catch
                    total = 0
                End Try
                Try
                    Scom.CommandText = "SELECT COUNT(*) FROM " & tabel_name
                    Scom.CommandText &= " WHERE SOID like 'I' "
                    total_soid = Scom.ExecuteScalar
                Catch
                    total_soid = 0
                End Try
                Try
                    Scom.CommandText = "SELECT COUNT(*) FROM " & tabel_name
                    If cVirBacaprod.get1230TTL3_Virbacaprod And FormMain.jenis_so = "BIC" Then
                        Scom.CommandText &= " WHERE SOID='T' OR SOID='G' OR SOID='R'"
                    Else
                        Scom.CommandText &= " WHERE SOID = 'T' OR SOID = 'G' "
                    End If
                    total_soid_toko = Scom.ExecuteScalar
                Catch
                    total_soid_toko = 0
                End Try

                result = total & "|" & total_soid & "|" & total_soid_toko
            End If

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "FindTableSO", Scon)
        Finally
            Scon.Close()
        End Try

        Return result
    End Function

    Public Function ResetTableSO(ByVal table As String) As Boolean
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)
        Dim result As Boolean = False

        Dim mJam, mWaktu, mSOID As String

        Dim cVirBacaprod As New ClsVirBacaprodController

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            mWaktu = Date.Now.TimeOfDay.ToString
            mJam = Mid(mWaktu, 1, 8)
            mSOID = ""

            Scom.CommandText = "UPDATE " & table
            If jenis_so = "BIC" And cVirBacaprod.get1230TTL3_Virbacaprod Then
                Scom.CommandText &= " SET TTL = 0, TTL1 = 0, TTL2 = 0, TTL3=0, SOID = '" & mSOID & "', SOTIME = '" & mJam & "', DCP = 0"
            Else
                Scom.CommandText &= " SET TTL = 0, TTL1 = 0, TTL2 = 0, SOID = '" & mSOID & "', SOTIME = '" & mJam & "', DCP = 0"
            End If

            Scom.ExecuteScalar()

            result = True
        Catch ex As Exception
            result = False

            TraceLog("Last Query: " & Scom.CommandText)
            TraceLog("ResetTableSO Error:  " & ex.ToString)
            MsgBox("Error: " & ex.ToString, MsgBoxStyle.Exclamation, "Error ResetTableSO")
        Finally
            Scon.Close()
        End Try

        Return result
    End Function

    ''' <summary>
    ''' Untuk mengecet apakah tabel sbe dan sbde sudah ada, jika ada maka bisa menjalankan fungsi edit so
    ''' </summary>
    ''' <param name="kode_toko"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function cekTableSBEdit(ByVal kode_toko As String) As Boolean
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)
        Dim result As Boolean = False

        Dim mSbe As Boolean
        Dim mSbde As Boolean

        Dim cVirBacaprod As New ClsVirBacaprodController
        Dim cFileSO As String
        Dim cFileSOD As String

        SyncLock Scon
            Try
                If Scon.State = ConnectionState.Closed Then
                    Scon.Open()
                End If

                If cVirBacaprod.get1230Tabel_Virbacaprod() Then
                    cFileSO = "SBE" & Format(Date.Now, "yyMMdd") & Mid(kode_toko, 1, 1)
                Else
                    cFileSO = "SBE" & Format(Date.Now, "yyMM") & Mid(kode_toko, 1, 1)
                End If

                Scom.CommandText = "show tables like '" & cFileSO & "';"
                If IsNothing(Scom.ExecuteScalar) Then
                    mSbe = False
                Else
                    mSbe = True
                End If

                If cVirBacaprod.get1230Tabel_Virbacaprod() Then
                    cFileSOD = "SBDE" & Format(Date.Now, "yyMMdd") & Mid(kode_toko, 1, 1)
                Else
                    cFileSOD = "SBDE" & Format(Date.Now, "yyMM") & Mid(kode_toko, 1, 1)
                End If

                Scom.CommandText = "show tables like '" & cFileSOD & "';"
                If IsNothing(Scom.ExecuteScalar) Then
                    mSbde = False
                Else
                    mSbde = True
                End If

                If mSbe = True And mSbde = True Then
                    result = True
                Else
                    result = False
                End If
            Catch ex As Exception
                result = False
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "cekTableSBEdit", Scon)
            Finally
                Scon.Close()
            End Try
        End SyncLock

        Return result
    End Function

    Public Function cekTableSKEdit(ByVal kode_toko As String) As Boolean
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mcom As New MySqlCommand("", conn)
        Dim result As Boolean = False

        Dim mSKE As Boolean
        Dim mSKDE As Boolean

        SyncLock conn
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If

                mcom.CommandText = "show tables like '" & "SKE" & Format(Date.Now, "yyMMdd") & Mid(kode_toko, 1, 1) & "'"
                If IsNothing(mcom.ExecuteScalar) Then
                    mSKE = False
                Else
                    mSKE = True
                End If

                mcom.CommandText = "show tables like '" & "SKDE" & Format(Date.Now, "yyMMdd") & Mid(kode_toko, 1, 1) & "'"
                If IsNothing(mcom.ExecuteScalar) Then
                    mSKDE = False
                Else
                    mSKDE = True
                End If

                If mSKE = True And mSKDE = True Then
                    result = True
                Else
                    result = False
                End If
            Catch ex As Exception
                result = False
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "cekTableSKEdit", conn)
            Finally
                conn.Close()
            End Try
        End SyncLock


        Return result
    End Function

    ''' <summary>
    ''' Untuk mengecet apakah tabel sne dan snde sudah ada, jika ada maka bisa menjalankan fungsi edit sonas
    ''' </summary>
    ''' <param name="kode_toko"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' 
    Public Function cekTableSNEdit(ByVal kode_toko As String) As Boolean
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mcom As New MySqlCommand("", conn)
        Dim result As Boolean = False

        Dim mSbe As Boolean
        Dim mSbde As Boolean

        SyncLock conn
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If

                mcom.CommandText = "show tables like '" & "SNE" & Format(Date.Now, "yyMM") & Mid(kode_toko, 1, 1) & "'"
                If IsNothing(mcom.ExecuteScalar) Then
                    mSbe = False
                Else
                    mSbe = True
                End If

                mcom.CommandText = "show tables like '" & "SNDE" & Format(Date.Now, "yyMM") & Mid(kode_toko, 1, 1) & "'"
                If IsNothing(mcom.ExecuteScalar) Then
                    mSbde = False
                Else
                    mSbde = True
                End If

                If mSbe = True And mSbde = True Then
                    result = True
                Else
                    result = False
                End If
            Catch ex As Exception
                result = False
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "cekTableSNEdit", conn)
            Finally
                conn.Close()
            End Try
        End SyncLock


        Return result
    End Function

    Public Function GetDeskripsiProdukSOKhusus(ByVal tabel_name As String, ByVal barcode_plu As String) As ClsSo
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim tmpDt As New DataTable
        Dim result As New ClsSo
        Dim mcom As New MySqlCommand("", conn)

        Dim no_rak As String = ""
        Dim no_shelf As String = ""

        Dim shift As String = ""

        If conn Is Nothing Then
            utility.TraceLogTxt("Error - GetDeskripsiProduk (connection Nothing) " & vbCrLf & "PLU:" & barcode_plu)
            Return result
            Exit Function
        End If

        SyncLock conn
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If

                'memo 209/cps/23
                'tambah baca kolom shift
                shift = GetShift()

                mcom.CommandText = " SELECT distinct T.TIPERAK,T.NORAK,T.NOSHELF,T.PRDCD,T.SINGKAT,T.TTL,T.TTL1,T.TTL2,"
                mcom.CommandText &= "T.SOID,T.SOTIME,T.DCP,T.KIRIKANAN,T.Unit,T.COM+T.BPB-T.RETUR_K-T.SALES+T.RETUR+T.BPB_2+T.ADJ-T.TTL2 AS COM, TTL1_OLD,TTL2_OLD"
                mcom.CommandText &= " FROM " & tabel_name & " T left join BARCODE B "
                mcom.CommandText &= " on T.PRDCD = B.PLU "
                mcom.CommandText &= " WHERE (B.BARCD = '" & barcode_plu & "' or T.PRDCD ='" & barcode_plu & "') AND T.SHIFT = '" & shift & "'"
                mcom.CommandText &= " ORDER BY T.NORAK,T.NOSHELF,T.TIPERAK,T.KIRIKANAN"

                Dim sDap As New MySqlDataAdapter(mcom)
                sDap.Fill(tmpDt)

                result.BarcodePlu = barcode_plu
                If tmpDt.Rows.Count > 0 Then
                    no_rak = CInt(tmpDt.Rows.Item(0)("NORAK"))
                    no_rak = no_rak.PadLeft(3, "0")
                    no_shelf = CInt(tmpDt.Rows.Item(0)("NOSHELF"))
                    no_shelf = no_shelf.PadLeft(3, "0")

                    result.PRDCD = tmpDt.Rows(0)("PRDCD")
                    result.Unit = tmpDt.Rows(0)("Unit")

                    result.Deskripsi = tmpDt.Rows(0)("SINGKAT")
                    If result.Deskripsi.Length > 20 Then
                        result.Deskripsi = result.Deskripsi.Substring(0, 20)
                    End If

                    result.Rak = no_rak & "/" & no_shelf
                    result.QTYToko = tmpDt.Rows(0)("TTL1")
                    result.QTYGudang = tmpDt.Rows(0)("TTL2")
                    result.QTYTotal = tmpDt.Rows(0)("TTL")
                    result.QTYCom = tmpDt.Rows(0)("COM")
                    result.qtyTTL1_OLD = tmpDt.Rows(0)("TTL1_OLD")

                    'Revisi 20 November 2019 (Memo 1081/CPS/19)
                    'Hitung lokasi RAK untuk item, Jika ada lebih dari 1 lokasi aktifkan fitur NEXT WDCP
                    mcom.CommandText = "SELECT COUNT(NORAK) FROM RAK R"
                    mcom.CommandText &= " LEFT JOIN BARCODE B  ON R.PRDCD = B.PLU"
                    mcom.CommandText &= " WHERE B.BARCD = '" & barcode_plu & "' OR R.PRDCD ='" & barcode_plu & "';"
                    Dim CountRak = mcom.ExecuteScalar
                    If Not IsDBNull(CountRak) Then
                        result.TotalRak = CountRak
                    Else
                        result.TotalRak = 0
                    End If

                    'revisi lepas flag CBR SO PRODUK KHUSUS
                    '09/09/2021
                    If FormMain.isFlagCBR = True Then
                        mcom.CommandText = "SELECT COUNT(*) FROM PRODMAST P"
                        mcom.CommandText &= " LEFT JOIN BARCODE B  ON P.PRDCD = B.PLU"
                        mcom.CommandText &= " WHERE (B.BARCD = '" & barcode_plu & "' OR P.PRDCD ='" & barcode_plu & "')"
                        mcom.CommandText &= " AND FLAGPROD LIKE '%CBR=Y%'"
                        Console.WriteLine(mcom.CommandText)
                        If mcom.ExecuteScalar <> "0" Then
                            result.statusBarcode = "CBRY"
                        Else
                            result.statusBarcode = "CBRN"

                        End If
                    End If
                Else
                    result.PRDCD = ""
                    result.Unit = ""
                    result.Deskripsi = "Tidak Ditemukan"
                    result.Rak = ""
                    result.QTYToko = ""
                    result.QTYGudang = ""
                    result.QTYTotal = ""
                    result.QTYCom = ""
                    result.statusBarcode = ""
                    result.qtyTTL1_OLD = ""

                    If FormMain.isFlagCBR = True Then
                        result.statusBarcode = "CBRY"
                    End If


                End If
                'MsgBox(result.statusBarcode)
            Catch ex As Exception
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetDeskripsiProduk", conn)
                utility.TraceLogTxt("Error - GetDeskripsiProduk " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                conn.Close()
            End Try

        End SyncLock

        Return result
    End Function

    Public Function cekTableSPEdit(ByVal kode_toko As String) As Boolean
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mcom As New MySqlCommand("", conn)
        Dim result As Boolean = False
        Dim tabelEdit As String = ""
        Dim mSpe As Boolean
        Dim mSpde As Boolean

        SyncLock conn
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                mcom.CommandText = "SELECT TABELNAME_EDIT FROM SO_PRODUK_KHUSUS_HIST where recid = 'A'"
                tabelEdit = mcom.ExecuteScalar()
                'Console.WriteLine(tabelEdit)

                'mcom.CommandText = "show tables like '" & tabelEdit & "'"
                'Console.WriteLine(mcom.CommandText)
                If tabelEdit = "" Then
                    mSpe = False
                Else
                    mSpe = True
                End If

                mcom.CommandText = "SELECT TABELNAME_EDITDETAIL FROM SO_PRODUK_KHUSUS_HIST where recid = 'A'"
                tabelEdit = mcom.ExecuteScalar()
                'mcom.CommandText = "show tables like '" & tabelEdit & "'"
                If tabelEdit = "" Then
                    mSpde = False
                Else
                    mSpde = True
                End If

                If mSpe = True And mSpde = True Then
                    result = True
                Else
                    result = False
                End If
            Catch ex As Exception
                result = False
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "cekTableSPEdit", conn)
            Finally
                conn.Close()
            End Try
        End SyncLock


        Return result
    End Function

    Public Function GetTableName(ByVal tabel As String, ByVal moderun As String) As String
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim tmpDt As New DataTable
        Dim result As String = ""
        Dim mcom As New MySqlCommand("", conn)
        Dim da As New MySqlDataAdapter
        Dim dt As New DataTable
        Dim no_rak As String = ""
        Dim no_shelf As String = ""

        If conn Is Nothing Then
            utility.TraceLogTxt("Error - GetTableName (connection Nothing) " & vbCrLf)
            Return result
            Exit Function
        End If

        SyncLock conn
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                mcom.CommandText = "SELECT TABELNAME FROM " & tabel & " where recid = 'A'"
                result = mcom.ExecuteScalar()
                'Console.WriteLine(result)
                If moderun = "E" Then
                    mcom.CommandText = "SELECT TABELNAME_EDIT FROM " & tabel & " where recid = 'A'"
                    Console.WriteLine(mcom.CommandText)

                    result = mcom.ExecuteScalar
                    Console.WriteLine(result)
                End If


            Catch ex As Exception
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetDeskripsiProduk", conn)
                utility.TraceLogTxt("Error - GetDeskripsiProduk " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                conn.Close()
            End Try

        End SyncLock

        Return result
    End Function

    Public Function GetDeskripsiProdukBazar(ByVal tabel_name As String, ByVal barcode_plu As String) As ClsSo
        'Dim connection As New ClsConnection
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim tmpDt As New DataTable
        Dim result As New ClsSo
        Dim Scom As New MySqlCommand("", Scon)

        Dim no_rak As String = ""
        Dim no_shelf As String = ""

        If Scon Is Nothing Then
            utility.TraceLogTxt("Error - GetDeskripsiProdukBazar (connection Nothing) " & vbCrLf & "PLU:" & barcode_plu)
            Return result
            Exit Function
        End If

        SyncLock Scon
            Try
                If Scon.State = ConnectionState.Closed Then
                    Scon.Open()
                End If

                Scom.CommandText = "SELECT T.Recid,T.TIPERAK,T.NORAK,T.NOSHELF,T.KIRIKANAN,T.PRDCD,"
                Scom.CommandText &= "T.SINGKAT,T.TTL,B.BARCD,T.SOID,T.COM"
                Scom.CommandText &= " FROM " & tabel_name & " T left join BARCODE B "
                Scom.CommandText &= " on T.PRDCD = B.PLU "
                Scom.CommandText &= " WHERE B.BARCD = '" & barcode_plu & "' or T.PRDCD ='" & barcode_plu & "'"
                Scom.CommandText &= " ORDER BY T.NORAK,T.NOSHELF,T.TIPERAK,T.KIRIKANAN"
                Console.WriteLine(Scom.CommandText)
                Dim sDap As New MySqlDataAdapter(Scom)
                sDap.Fill(tmpDt)

                result.BarcodePlu = barcode_plu
                If tmpDt.Rows.Count > 0 Then
                    no_rak = CInt(tmpDt.Rows.Item(0)("NORAK"))
                    no_rak = no_rak.PadLeft(3, "0")
                    no_shelf = CInt(tmpDt.Rows.Item(0)("NOSHELF"))
                    no_shelf = no_shelf.PadLeft(3, "0")

                    result.PRDCD = tmpDt.Rows(0)("PRDCD")

                    result.Deskripsi = tmpDt.Rows(0)("SINGKAT")
                    If result.Deskripsi.Length > 20 Then
                        result.Deskripsi = result.Deskripsi.Substring(0, 20)
                    End If

                    result.Rak = no_rak & "/" & no_shelf
                    result.QTYTotal = tmpDt.Rows(0)("TTL")
                    result.QTYCom = tmpDt.Rows(0)("COM")

                    'Revisi 20 November 2019 (Memo 1081/CPS/19)
                    'Hitung lokasi RAK untuk item, Jika ada lebih dari 1 lokasi aktifkan fitur NEXT WDCP
                    'mcom.CommandText = "SELECT COUNT(NORAK) FROM RAK R"
                    'mcom.CommandText &= " LEFT JOIN BARCODE B  ON R.PRDCD = B.PLU"
                    'mcom.CommandText &= " WHERE B.BARCD = '" & barcode_plu & "' OR R.PRDCD ='" & barcode_plu & "';"
                    'Dim CountRak = mcom.ExecuteScalar
                    'If Not IsDBNull(CountRak) Then
                    '    result.TotalRak = CountRak
                    'Else
                    '    result.TotalRak = 0
                    'End If
                Else
                    result.PRDCD = ""
                    result.Deskripsi = "Tidak Ditemukan"
                    result.QTYTotal = ""
                    result.QTYCom = ""
                    result.Tgl_exp = ""
                    result.Rak = ""
                End If

            Catch ex As Exception
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetDeskripsiProduk", Scon)
                utility.TraceLogTxt("Error - GetDeskripsiProduk " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                Scon.Close()
            End Try

        End SyncLock

        Return result
    End Function

    Public Function GetShift() As String
        'memo 209/cps/23
        'tambah baca kolom shift
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim tmpDt As New DataTable
        Dim result As String = ""
        Dim mcom As New MySqlCommand("", conn)


        If conn Is Nothing Then
            utility.TraceLogTxt("Error - GetShift (connection Nothing) " & vbCrLf)
            Return result
            Exit Function
        End If

        SyncLock conn
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                mcom.CommandText = "SELECT Shift FROM Initial WHERE Station='" & IDM.Fungsi.Get_Station & "' " &
                "Order By Tanggal Desc"

                result = mcom.ExecuteScalar()

                utility.Tracelog("", result, "GetShift", conn)

            Catch ex As Exception
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetDeskripsiProduk", conn)
                utility.TraceLogTxt("Error - GetDeskripsiProduk " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                conn.Close()
            End Try

        End SyncLock

        Return result
    End Function

    Public Function InsertTabelSPD(ByVal tabel_name As String, ByVal mode_run As String, ByVal kode_toko As String, ByVal prdcd As String, ByVal qty As Integer, ByVal lokasi As String, ByVal nik As String, ByVal nama As String) As Boolean
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim result As Boolean = False
        Dim mcom As New MySqlCommand("", conn)
        Dim mTable As String = ""
        Dim shift As String = ""
        SyncLock conn
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If

                If mode_run = "B" Then
                    mTable = tabel_name.Substring(0, 2) & "D" & tabel_name.Substring(2, tabel_name.Length - 2)
                Else
                    mTable = tabel_name.Substring(0, 2) & "D" & tabel_name.Substring(2, tabel_name.Length - 2)
                End If

                shift = GetShift()

                'set lokasi(toko=1,gudang=2)
                If lokasi = "Toko" Then
                    lokasi = "1"
                Else
                    lokasi = "2"
                End If

                If qty = 0 Then
                    Dim m_qty As Integer
                    mcom.CommandText = "SELECT count(*) FROM " & mTable & " WHERE NIK = '" & nik & "' And PRDCD = '" & prdcd & "' And Lokasi = '" & lokasi & "' and shift = '" & shift & "'"
                    If mcom.ExecuteScalar > 0 Then
                        mcom.CommandText = "SELECT sum(TTL) FROM " & mTable & " WHERE NIK = '" & nik & "' And PRDCD = '" & prdcd & "' And Lokasi = '" & lokasi & "' and shift = '" & shift & "'"
                        m_qty = CInt(mcom.ExecuteScalar)
                        If m_qty > 0 Then
                            mcom.CommandText = "Insert Ignore Into " & mTable
                            mcom.CommandText &= " values('" & prdcd & "', " & m_qty * -1 & ", '" & lokasi & "','" & Format(Date.Now, "yyyy-MM-dd") & "',"
                            mcom.CommandText &= "'" & Format(Date.Now, "HH:mm:ss") & "', '" & nik & "','" & nama & "','" & shift & "')"
                            mcom.ExecuteNonQuery()
                        End If
                    End If
                Else
                    mcom.CommandText = "Insert Ignore Into " & mTable
                    mcom.CommandText &= " values('" & prdcd & "', " & qty & ", '" & lokasi & "','" & Format(Date.Now, "yyyy-MM-dd") & "',"
                    mcom.CommandText &= "'" & Format(Date.Now, "HH:mm:ss") & "', '" & nik & "','" & nama & "','" & shift & "')"
                    mcom.ExecuteNonQuery()
                End If

            Catch ex As Exception
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "InsertTabel" & mTable, conn)
                utility.TraceLogTxt("Error - InsertTabel " & mTable & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                conn.Close()
            End Try
        End SyncLock


        Return result
    End Function

    Public Function cekKolomSBDE(ByVal kode_toko As String) As Boolean
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)
        Dim result As Boolean = False

        Dim cVirBacaprod As New ClsVirBacaprodController
        Dim cFileSODE As String

        SyncLock Scon
            Try
                If Scon.State = ConnectionState.Closed Then
                    Scon.Open()
                End If

                If cVirBacaprod.get1230Tabel_Virbacaprod() Then
                    cFileSODE = "SBDE" & Format(Date.Now, "yyMMdd") & Mid(kode_toko, 1, 1)
                Else
                    cFileSODE = "SBDE" & Format(Date.Now, "yyMM") & Mid(kode_toko, 1, 1)
                End If

                Scom.CommandText = "SHOW COLUMNS FROM `" & cFileSODE & "` LIKE 'NIK_PEMEGANG_SHIFT'"
                If Scom.ExecuteScalar = "" Then
                    Scom.CommandText = "ALTER TABLE " & cFileSODE & " ADD COLUMN `NIK_PEMEGANG_SHIFT` VARCHAR(20) NULL"
                    Scom.ExecuteNonQuery()
                End If

                Scom.CommandText = "SHOW COLUMNS FROM `" & cFileSODE & "` LIKE 'NAMA_PEMEGANG_SHIFT'"
                If Scom.ExecuteScalar = "" Then
                    Scom.CommandText = "ALTER TABLE " & cFileSODE & " ADD COLUMN `NAMA_PEMEGANG_SHIFT` VARCHAR(20) NULL"
                    Scom.ExecuteNonQuery()
                End If


            Catch ex As Exception
                result = False
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "cekTableSBEdit", Scon)
            Finally
                Scon.Close()
            End Try
        End SyncLock

        Return result
    End Function

    Public Function cekKolomSKDE(ByVal kode_toko As String) As Boolean
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mcom As New MySqlCommand("", conn)
        Dim result As Boolean = False

        SyncLock conn
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If

                mcom.CommandText = "SHOW COLUMNS FROM `SKDE" & Format(Date.Now, "yyMMdd") & Mid(kode_toko, 1, 1) & "` LIKE 'NIK_PEMEGANG_SHIFT'"
                If mcom.ExecuteScalar = "" Then
                    mcom.CommandText = "ALTER TABLE SKDE" & Format(Date.Now, "yyMMdd") & Mid(kode_toko, 1, 1) & " ADD COLUMN `NIK_PEMEGANG_SHIFT` VARCHAR(20) NULL"
                    mcom.ExecuteNonQuery()
                End If

                mcom.CommandText = "SHOW COLUMNS FROM `SKDE" & Format(Date.Now, "yyMMdd") & Mid(kode_toko, 1, 1) & "` LIKE 'NAMA_PEMEGANG_SHIFT'"
                If mcom.ExecuteScalar = "" Then
                    mcom.CommandText = "ALTER TABLE SKDE" & Format(Date.Now, "yyMMdd") & Mid(kode_toko, 1, 1) & " ADD COLUMN `NAMA_PEMEGANG_SHIFT` VARCHAR(20) NULL"
                    mcom.ExecuteNonQuery()
                End If

            Catch ex As Exception
                result = False
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "cekKolomSKDEEdit", conn)
            Finally
                conn.Close()
            End Try
        End SyncLock

        Return result

    End Function

    Public Function InsertTabelSO(ByVal tabel_name As String, ByVal mode_run As String, ByVal kode_toko As String, ByVal prdcd As String, ByVal qty As Integer, ByVal lokasi As String, ByVal nik As String, ByVal nama As String) As Boolean
        Dim sCon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim sCom As New MySqlCommand("", sCon)
        Dim result As Boolean = False
        Dim mTable As String = ""
        Dim alasanSO As String = ""

        Dim cVirBacaprod As New ClsVirBacaprodController
        Dim isFormatBaru As Boolean

        SyncLock sCon
            Try
                If sCon.State = ConnectionState.Closed Then
                    sCon.Open()
                End If

                If tabel_name.ToUpper.Contains("SN") Then
                    If mode_run = "B" Then
                        mTable = "SND" & Format(Now, "yyMM") & kode_toko.Substring(0, 1)
                    Else
                        mTable = "SNDE" & Format(Now, "yyMM") & kode_toko.Substring(0, 1)
                    End If
                ElseIf tabel_name.ToUpper.Contains("SP") Then
                    If mode_run = "B" Then
                        mTable = tabel_name.Substring(0, 2) & "D" & tabel_name.Substring(2, tabel_name.Length - 2)
                    Else
                        mTable = tabel_name.Substring(0, 2) & "D" & tabel_name.Substring(2, tabel_name.Length - 2)
                    End If
                    'Revisi MEMO 1074/CPS/22 'Kukuh
                ElseIf tabel_name.ToUpper.Contains("SK") Then
                    If mode_run = "B" Then
                        mTable = "SKD" & Format(Now, "yyMMdd") & kode_toko.Substring(0, 1)
                    Else
                        mTable = "SKDE" & Format(Now, "yyMMdd") & kode_toko.Substring(0, 1)
                    End If

                    'Baca Alasan SO Kasus
                    sCom.CommandText = "SELECT DISTINCT alasan from SK" & Format(Now, "yyMMdd") & kode_toko.Substring(0, 1) & ";"
                    TraceLog("InsertTabelSO: " & sCom.CommandText)
                    Dim sReader As MySqlDataReader = sCom.ExecuteReader()
                    If sReader.Read Then
                        alasanSO = sReader.GetString(0)
                    End If

                    sReader.Close()
                    TraceLog("Alasan SO: " & alasanSO)
                Else
                    isFormatBaru = cVirBacaprod.get1230Tabel_Virbacaprod

                    If mode_run = "B" Then
                        If isFormatBaru Then
                            mTable = "SBD" & Format(Now, "yyMMdd") & kode_toko.Substring(0, 1)
                        Else
                            mTable = "SBD" & Format(Now, "yyMM") & kode_toko.Substring(0, 1)
                        End If
                    Else
                        If isFormatBaru Then
                            mTable = "SBDE" & Format(Now, "yyMMdd") & kode_toko.Substring(0, 1)
                        Else
                            mTable = "SBDE" & Format(Now, "yyMM") & kode_toko.Substring(0, 1)
                        End If
                    End If
                End If

                'set lokasi(toko=1,gudang=2)
                If lokasi = "Toko" Then
                    lokasi = "1"
                ElseIf lokasi = "Gudang" Then
                    lokasi = "2"
                Else
                    lokasi = "3"
                End If

                TraceLog("InsertTabelSO-QTY: " & qty)

                If qty = 0 Then
                    Dim m_qty As Integer
                    sCom.CommandText = "SELECT count(*) FROM " & mTable & " WHERE NIK = '" & nik & "' And PRDCD = '" & prdcd & "' And Lokasi = '" & lokasi & "'"
                    If sCom.ExecuteScalar > 0 Then
                        sCom.CommandText = "SELECT sum(TTL) FROM " & mTable & " WHERE NIK = '" & nik & "' And PRDCD = '" & prdcd & "' And Lokasi = '" & lokasi & "'"
                        m_qty = CInt(sCom.ExecuteScalar)

                        'Memo 1457/cps/22
                        'Tambah simpan nik dan nama pemegang shift (edit SO IC)

                        If mTable.StartsWith("SBDE") Then
                            If m_qty > 0 Then
                                sCom.CommandText = "Insert Ignore Into " & mTable
                                sCom.CommandText &= " values('" & prdcd & "', " & m_qty * -1 & ", '" & lokasi & "','" & Format(Date.Now, "yyyy-MM-dd") & "',"
                                sCom.CommandText &= "'" & Format(Date.Now, "HH:mm:ss") & "', '" & nik & "','" & nama & "', '" & edit_so_nik_pemegang_shift & "','" & edit_so_nama_pemegang_shift & "')"
                                TraceLog("InsertTabelSO-T1-SBDE: " & sCom.CommandText)
                                sCom.ExecuteNonQuery()
                            End If
                        ElseIf mTable.StartsWith("SKDE") Then
                            If m_qty > 0 Then
                                sCom.CommandText = "Insert Ignore Into " & mTable
                                sCom.CommandText &= " values('" & prdcd & "', " & m_qty * -1 & ", '" & lokasi & "','" & Format(Date.Now, "yyyy-MM-dd") & "',"
                                sCom.CommandText &= "'" & Format(Date.Now, "HH:mm:ss") & "', '" & alasanSO & "', '" & nik & "','" & nama & "', '" & edit_so_nik_pemegang_shift & "','" & edit_so_nama_pemegang_shift & "')"
                                TraceLog("InsertTabelSO-T2-SKDE: " & sCom.CommandText)
                                sCom.ExecuteNonQuery()
                            End If
                        Else
                            If m_qty > 0 Then
                                sCom.CommandText = "Insert Ignore Into " & mTable
                                sCom.CommandText &= " values('" & prdcd & "', " & m_qty * -1 & ", '" & lokasi & "','" & Format(Date.Now, "yyyy-MM-dd") & "',"
                                sCom.CommandText &= "'" & Format(Date.Now, "HH:mm:ss") & "', '" & nik & "','" & nama & "')"
                                TraceLog("InsertTabelSO-T3-" & mTable & ": " & sCom.CommandText)
                                sCom.ExecuteNonQuery()
                            End If
                        End If
                    End If
                Else
                    'Revisi MEMO 1074/CPS/22 'Kukuh
                    If mTable.StartsWith("SBDE") Then
                        sCom.CommandText = "Insert Ignore Into " & mTable
                        sCom.CommandText &= " values('" & prdcd & "', " & qty & ", '" & lokasi & "','" & Format(Date.Now, "yyyy-MM-dd") & "',"
                        sCom.CommandText &= "'" & Format(Date.Now, "HH:mm:ss") & "', '" & nik & "','" & nama & "', '" & edit_so_nik_pemegang_shift & "','" & edit_so_nama_pemegang_shift & "')"
                        TraceLog("InsertTabelSO-T1-SBDE: " & sCom.CommandText)
                        sCom.ExecuteNonQuery()
                    ElseIf mTable.StartsWith("SKDE") Then
                        sCom.CommandText = "Insert Ignore Into " & mTable
                        sCom.CommandText &= " values('" & prdcd & "', " & qty & ", '" & lokasi & "','" & Format(Date.Now, "yyyy-MM-dd") & "',"
                        sCom.CommandText &= "'" & Format(Date.Now, "HH:mm:ss") & "', '" & alasanSO & "', '" & nik & "','" & nama & "', '" & edit_so_nik_pemegang_shift & "','" & edit_so_nama_pemegang_shift & "');"
                        TraceLog("InsertTabelSO-T2-SKDE: " & sCom.CommandText)
                        sCom.ExecuteNonQuery()
                    ElseIf mTable.StartsWith("SKD") Then
                        sCom.CommandText = "Insert Ignore Into " & mTable
                        sCom.CommandText &= " values('" & prdcd & "', " & qty & ", '" & lokasi & "','" & Format(Date.Now, "yyyy-MM-dd") & "',"
                        sCom.CommandText &= "'" & Format(Date.Now, "HH:mm:ss") & "', '" & nik & "','" & nama & "','" & alasanSO & "')"
                        TraceLog("InsertTabelSO-T3-SKD: " & sCom.CommandText)
                        sCom.ExecuteNonQuery()
                    Else
                        sCom.CommandText = "Insert Ignore Into " & mTable
                        sCom.CommandText &= " values('" & prdcd & "', " & qty & ", '" & lokasi & "','" & Format(Date.Now, "yyyy-MM-dd") & "',"
                        sCom.CommandText &= "'" & Format(Date.Now, "HH:mm:ss") & "', '" & nik & "','" & nama & "')"
                        TraceLog("InsertTabelSO-T4-" & mTable & ": " & sCom.CommandText)
                        sCom.ExecuteNonQuery()
                    End If
                End If
            Catch ex As Exception
                TraceLog("Error :" & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "InsertTabel" & mTable, sCon)
                utility.TraceLogTxt("Error - InsertTabel " & mTable & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                sCon.Close()
            End Try
        End SyncLock

        Return result

    End Function

End Class
