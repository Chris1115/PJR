Imports IDM.Fungsi
Imports MySql.Data.MySqlClient

Public Class ClsSOEDController
    Private utility As New Utility

    Public Function getDeskripsiExpiredDate(ByVal tabel_name As String, ByVal barcode_plu As String,
                                            Optional expDate As String = "", Optional mainCBR As Boolean = False) As ClsSOED
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)

        Dim no_rak As String
        Dim no_shelf As String
        Dim no_kirikanan As String
        Dim tmpDt As New DataTable
        Dim dtSOED As New DataTable

        Dim result As New ClsSOED

        If Scon Is Nothing Then
            utility.TraceLogTxt("Error - getDeskripsiExpiredDate (connection Nothing) " & vbCrLf & "PLU:" & barcode_plu)
            Return result
            Exit Function
        End If

        SyncLock Scon
            Try
                If Scon.State = ConnectionState.Closed Then
                    Scon.Open()
                End If

                Scom.CommandText = "SELECT ED.PRDCD,R.NORAK,R.NOSHELF,R.KIRIKANAN,"
                Scom.CommandText &= "ED.DESKRIPSI AS SINGKAT,ED.BULAN_EXP,ED.TTL,ED.BARCODE"
                Scom.CommandText &= " FROM " & tabel_name & " ED LEFT JOIN (SELECT PRDCD,TIPERAK,NORAK,NOSHELF,KIRIKANAN FROM RAK WHERE KODETOKO='" & IDM.InfoToko.Get_TipeToko & "' GROUP BY PRDCD) R "
                Scom.CommandText &= " ON ED.PRDCD = R.PRDCD "
                Scom.CommandText &= " WHERE ED.BARCODE = '" & barcode_plu & "' OR ED.PRDCD ='" & barcode_plu & "'"
                If expDate <> "" Then
                    Scom.CommandText &= " AND ED.BULAN_EXP ='" & expDate & "'"
                End If
                Scom.CommandText &= " ORDER BY ED.BULAN_EXP;"
                TraceLog("getDeskripsiExpiredDate-Q1: " & Scom.CommandText)

                Dim Sdap As New MySqlDataAdapter(Scom)
                Sdap.Fill(tmpDt)

                result.PRDCD = barcode_plu

                If tmpDt.Rows.Count > 0 Then
                    If mainCBR Then
                        Scom.CommandText = "SELECT COUNT(*) FROM PRODMAST P"
                        Scom.CommandText &= " LEFT JOIN BARCODE B  ON P.PRDCD = B.PLU"
                        Scom.CommandText &= " WHERE (B.BARCD = '" & barcode_plu & "' OR P.PRDCD ='" & barcode_plu & "')"
                        Scom.CommandText &= " AND FLAGPROD LIKE '%CBR=Y%' AND B.QTY=1;"
                        TraceLog("getDeskripsiExpiredDate-Q2: " & Scom.CommandText)

                        If Scom.ExecuteScalar <> "0" Then
                            result.StatusBarcode = "CBRY"
                        Else
                            result.StatusBarcode = "CBRN"
                        End If
                    Else
                        result.StatusBarcode = "CBRY"
                    End If

                    no_rak = CInt(tmpDt.Rows.Item(0)("NORAK"))
                    no_rak = no_rak.PadLeft(3, "0")
                    no_shelf = CInt(tmpDt.Rows.Item(0)("NOSHELF"))
                    no_shelf = no_shelf.PadLeft(3, "0")
                    no_kirikanan = CInt(tmpDt.Rows.Item(0)("KIRIKANAN"))
                    no_kirikanan = no_kirikanan.PadLeft(3, "0")

                    result.PRDCD = tmpDt.Rows(0)("PRDCD")
                    result.Deskripsi = tmpDt.Rows(0)("SINGKAT")
                    result.Lokasi = no_rak & "-" & no_shelf & "-" & no_kirikanan

                    If result.Deskripsi.Length > 20 Then
                        result.Deskripsi = result.Deskripsi.Substring(0, 20)
                    End If

                    If expDate <> "" Then
                        result.ExpDate = tmpDt.Rows(0)("BULAN_EXP")
                    End If

                    result.Feedback = "1"
                Else
                    result.PRDCD = ""
                    result.Deskripsi = "Tidak Ditemukan"
                    result.Lokasi = ""
                    result.Feedback = "0"
                End If
            Catch ex As Exception
                TraceLog("Last Query: " & Scom.CommandText)
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "getDeskripsiExpiredDate", Scon)
                utility.TraceLogTxt("Error - getDeskripsiExpiredDate " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                Scon.Close()
            End Try
        End SyncLock

        Return result
    End Function

    Public Function inputTglExp(ByVal table_name As String, ByVal prdcd As String, ByVal expDateInput As String,
                                Optional mainCBR As Boolean = False) As ClsSOED
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)

        Dim tempDtSOED As New DataTable
        Dim Results As New ClsSOED
        Dim statusItem As String = ""
        Dim noPropED As String = ""

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            If expDateInput.Length = "6" Then
                Try
                    Scom.CommandText = "SELECT COUNT(*) FROM master_so_expire_date"
                    Scom.CommandText &= " WHERE PRDCD='" & prdcd & "' AND DATE_FORMAT(BULAN_EXP, '%m%Y')='" & expDateInput & "';"
                    TraceLog("inputTglExp-Q1: " & Scom.CommandText)

                    If Scom.ExecuteScalar > 0 Then
                        statusItem = "Mendekati Expired"
                    Else
                        statusItem = "Baik"
                    End If

                    Scom.CommandText = "SELECT COUNT(1) FROM " & table_name & " WHERE PRDCD='" & prdcd & "' AND BULAN_EXP='" & expDateInput & "';"
                    TraceLog("inputTglExp-Q2: " & Scom.CommandText)

                    If Scom.ExecuteScalar = 0 Then
                        Scom.CommandText = "SELECT PRDCD, DESKRIPSI, DATE_FORMAT(TGL_AWAL, '%Y-%m-%d') as TGL_AWAL, DATE_FORMAT(TGL_AKHIR, '%Y-%m-%d') as TGL_AKHIR, HPP, PRICE, BARCODE FROM " & table_name & " WHERE PRDCD='" & prdcd & "' GROUP BY PRDCD;"
                        TraceLog("inputTglExp-Q3: " & Scom.CommandText)
                        Dim Sdap As New MySqlDataAdapter(Scom)
                        Sdap.Fill(tempDtSOED)

                        Scom.CommandText = "INSERT IGNORE INTO " & table_name & "("
                        Scom.CommandText &= "`COUNTER`,`PRDCD`,`DESKRIPSI`,`TGL_AWAL`,`TGL_AKHIR`,`BULAN_EXP`,`TTL`,`HPP`,`PRICE`,`BARCODE`,`STATUS`,`ADDTIME`,`UPDTIME_WDCP`) "
                        Scom.CommandText &= "VALUES("
                        Scom.CommandText &= "'0','" & tempDtSOED.Rows.Item(0)("PRDCD") & "','" & tempDtSOED.Rows.Item(0)("DESKRIPSI") & "','"
                        Scom.CommandText &= tempDtSOED.Rows.Item(0)("TGL_AWAL") & "','" & tempDtSOED.Rows.Item(0)("TGL_AKHIR") & "',"
                        Scom.CommandText &= "'" & expDateInput & "','0','" & tempDtSOED.Rows.Item(0)("HPP") & "','"
                        Scom.CommandText &= tempDtSOED.Rows.Item(0)("PRICE") & "','" & tempDtSOED.Rows.Item(0)("BARCODE") & "','" & statusItem & "',"
                        Scom.CommandText &= "NOW(),NOW());"
                        TraceLog("inputTglExp-Q4: " & Scom.CommandText)
                        Scom.ExecuteNonQuery()
                    Else

                        Scom.CommandText = "SELECT ID_SO_EXPIRED FROM master_so_expire_date"
                        Scom.CommandText &= " WHERE PRDCD='" & prdcd & "' AND DATE_FORMAT(BULAN_EXP, '%m%Y')='" & expDateInput & "';"
                        TraceLog("inputTglExp-Q3: " & Scom.CommandText)
                        noPropED = Scom.ExecuteScalar

                        Scom.CommandText = "UPDATE " & table_name & " SET STATUS='" & statusItem & "'"
                        Scom.CommandText &= " WHERE PRDCD='" & prdcd & "' AND BULAN_EXP='" & expDateInput & "' AND NO_PROP_ED='" & noPropED & "';"
                        TraceLog("inputTglExp-Q4: " & Scom.CommandText)
                        Scom.ExecuteNonQuery()

                    End If

                    Results = getDeskripsiExpiredDate(table_name, prdcd, expDateInput, mainCBR)
                    Results.noPropED = noPropED
                    Results.Feedback = "1"
                Catch ex As Exception
                    TraceLog("Last Query: " & Scom.CommandText)
                    IDM.Fungsi.TraceLog("Gagal inputTglExp SE " & ex.Message & ex.StackTrace)
                    Results.Feedback = "3"
                End Try
            Else
                Results.Feedback = "2"
            End If
        Catch ex As Exception
            TraceLog("Last Query: " & Scom.CommandText)
            IDM.Fungsi.TraceLog("Gagal inputTglExp SE " & ex.Message & ex.StackTrace)
            MsgBox("Gagal inputTglExp SE " & ex.Message & ex.StackTrace)
        Finally
            Scon.Close()
        End Try

        Return Results
    End Function

    Public Function SimpanQty_ED(ByVal table As String, ByVal prdcd As String, ByVal expDate As String,
                                 ByVal qty As String, ByVal noPropED As String) As ClsSOED
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)
        Dim dt As New DataTable
        Dim descSOED As New ClsSOED
        Dim result As New ClsSOED

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "UPDATE " & table & " SET counter=counter+1, TTL=" & qty & ",UPDTIME_WDCP=NOW() "
            Scom.CommandText &= "WHERE PRDCD='" & prdcd & "' AND BULAN_EXP='" & expDate & "' AND NO_PROP_ED='" & noPropED & "';"
            TraceLog("SimpanQty_ED-Q1: " & Scom.CommandText)
            Scom.ExecuteNonQuery()

            result.Feedback = "4"
        Catch ex As Exception
            result.Feedback = "1"
            TraceLog("Last Query: " & Scom.CommandText)
            IDM.Fungsi.TraceLog("Gagal update table SE " & ex.Message & ex.StackTrace)
            MsgBox("Gagal update table SE " & ex.Message & ex.StackTrace)
        Finally
            Scon.Close()
        End Try

        Return result
    End Function

    Public Function lihatTabel_SOED(ByVal tabel_name As String)
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)
        Dim tmpDt As New DataTable

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SELECT PRDCD, DESKRIPSI, BULAN_EXP FROM " & tabel_name & " "
            Scom.CommandText &= "WHERE STATUS='' "
            Scom.CommandText &= "AND TGL_AKHIR >= NOW();"
            TraceLog("lihatTabel_SOED-Q1: " & Scom.CommandText)

            Dim sDap As New MySqlDataAdapter(Scom)
            sDap.Fill(tmpDt)

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "lihatTabel_SOED", Scon)
        Finally
            Scon.Close()
        End Try

        Return tmpDt
    End Function

    Public Function SOED_cekVirBacaprod() As Boolean
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)

        Dim resultFilter As Boolean = False

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SELECT COUNT(1) FROM vir_bacaprod WHERE jenis='1314_SOED';"
            TraceLog("SOED_cekVirBacaprod-Q1: " & Scom.CommandText)

            If Scom.ExecuteScalar = 1 Then
                Scom.CommandText = "SELECT filter FROM vir_bacaprod WHERE jenis='1314_SOED';"
                TraceLog("SOED_cekVirBacaprod-Q2: " & Scom.CommandText)

                If Scom.ExecuteScalar = "ON" Then
                    resultFilter = True
                Else
                    resultFilter = False
                End If
            End If
        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "SOED_cekVirBacaprod", Scon)
            utility.TraceLogTxt("Error - SOED_cekVirBacaprod " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        Finally
            Scon.Close()
        End Try

        Return resultFilter
    End Function
End Class
