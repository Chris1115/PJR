Imports MySql.Data.MySqlClient
Imports IDM.Fungsi
Public Class ClsAktivaController

    Private utility As New Utility

    ''' <summary>
    ''' Get deskripsi produk
    ''' </summary>
    ''' <param name="tabel_name"></param>
    ''' <param name="NSeri"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDeskripsiAktiva(ByVal tabel_name As String, ByVal NSeri As String, ByVal Toko As ClsToko, Optional ByVal qtyinput As String = "") As ClsAktiva
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Result As New ClsAktiva
        Dim IsLengthValid As Boolean

        If Conn Is Nothing Then
            utility.TraceLogTxt("Error - GetDeskripsiAktiva (connection Nothing) " & vbCrLf & "NSeri:" & NSeri)
            Return Result
            Exit Function
        End If

        SyncLock Conn
            Try
                'Cek format NSeri
                IsLengthValid = True
                Result.NSeri = NSeri
                Result.Deskripsi = ""
                If NSeri.Trim.Length < 9 Then
                    IsLengthValid = False
                    Result.Deskripsi2 = "Scan Nomor Seri dulu"
                    'Validasi pembeda aktiva Reguler dan Franchise ditiadakan
                    'Semua AT bisa diinput
                    '22/04/21
                Else
                    If NSeri.Trim.Length > 16 Then
                        IsLengthValid = False
                        Result.Deskripsi2 = "Scan Nomor Seri dulu"
                    End If
                    If Not Char.IsLetter(NSeri.Substring(0, 1)) Then
                        IsLengthValid = False
                        Result.Deskripsi2 = "Scan Nomor Seri dulu"
                    End If
                End If

                If IsLengthValid Then
                    'Ambil data planogram
                    Dim DtAktiva As New DataTable
                    Dim IsATBaru As Boolean = False
                    If Conn.State = ConnectionState.Closed Then
                        Conn.Open()
                    End If
                    Mcom.CommandText = "SELECT RECID3,NSERI,BARANG,MERK,QTY,QTYB,QTYH,QTYR,QTYSO,FRAK"
                    Mcom.CommandText &= " FROM " & tabel_name & " WHERE NSERI = '" & NSeri & "' LIMIT 1;"
                    Dim sDap As New MySqlDataAdapter(Mcom)
                    sDap.Fill(DtAktiva)

                    If DtAktiva.Rows.Count = 0 Then
                        IsATBaru = True
                    Else
                        If DtAktiva.Rows(0).Item("RECID3") & "" = "" Then
                            IsATBaru = True
                        End If
                    End If

                    Result.QtyMax = "-"
                    If IsATBaru = False Then
                        'revisi 18/11/2020
                        If qtyinput <> "" Then
                            If qtyinput = DtAktiva.Rows(0)("QTY").ToString Then
                                Result.statusQty = True
                            Else
                                Result.statusQty = False
                            End If

                        Else
                            Result.statusQty = False
                        End If

                        'Tampil AT
                        Result.NSeri = NSeri
                        Result.QtyMax = DtAktiva.Rows(0).Item("QTY").ToString
                        Result.Deskripsi = DtAktiva.Rows(0).Item("BARANG").ToString
                        Result.Deskripsi2 = ""
                    Else
                        'AT Baru
                        Result.NSeri = NSeri
                        Result.Deskripsi = ""
                        Result.Deskripsi2 = "AT baru !"
                        If DtAktiva.Rows.Count > 0 Then
                            Result.QtyMax = DtAktiva.Rows(0).Item("QTY").ToString
                        End If
                    End If
                End If

            Catch ex As Exception
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetDeskripsiAktiva", Conn)
                utility.TraceLogTxt("Error - GetDeskripsiAktiva " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                Conn.Close()
            End Try
        End SyncLock

        Return Result
    End Function

    ''' <summary>
    ''' Simpan
    ''' </summary>
    ''' <param name="tabel_name"></param>
    ''' <param name="barcode_plu"></param>
    ''' <param name="qtySO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UpdateQtyProduk(ByVal tabel_name As String, ByVal barcode_plu As String, ByVal deskripsi As String,
                                    ByVal qtyMax As String, ByVal qtySO As Integer, ByVal statusBarcode As String, ByVal Toko As ClsToko) As ClsAktiva
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Result As New ClsAktiva
        Dim Scom As New MySqlCommand("", Scon)
        Dim qtySO_total As Integer
        Dim qtyHilang As Integer
        Dim qty As Integer

        SyncLock Scon
            Try
                'revisi Memo No. 208 - CPS - 20
                '13/10/2020
                qtySO_total = qtyMax + qtySO 'Baik + rusak buat dibandingin dengan qty (db)
                If Scon.State = ConnectionState.Closed Then
                    Scon.Open()
                End If
                If qtySO_total >= 0 Then
                    Scom.CommandText = "SELECT QTY FROM `" & tabel_name & "` WHERE NSERI = '" & barcode_plu & "';" 'qty(db)
                    qty = Scom.ExecuteScalar
                    Scom.CommandText = "SELECT recid3 FROM `" & tabel_name & "` WHERE NSERI = '" & barcode_plu & "';" 'qty(db)
                    Dim sDap As New MySqlDataAdapter(Scom)
                    Dim dt As New DataTable
                    sDap.Fill(dt)
                    If qty > 1 Or deskripsi.Contains("AT baru") Then ' qty dat lebih dari 1
                        If qtySO_total <> qty Then
                            qtyHilang = qty - qtySO_total
                        End If
                        If qtyHilang < 0 Then
                            qtyHilang = 0
                        End If
                        '21/8/16
                        'case qtyso = 0, tidak bisa reset aktiva di SOTKIDM
                        If qtySO_total = 0 Then
                            qtySO_total = qtyHilang
                        End If
                        If deskripsi.Contains("AT baru") And dt.Rows.Count = 0 Then
                            Scom.CommandText = "INSERT INTO " & tabel_name & " (NSERI,KCAB,KLOK,QTY,QTYB,QTYR,QTYH,QTYSO,"
                            Scom.CommandText &= "RECID2,SFTJUAL,KETJUAL,HrgJual,TGLMUT,KCABMUT,KLOKMUT,TRANSMUT,"
                            Scom.CommandText &= "PEMAKAI,TGLTRANS,TROLD,TGLTROLD,NACCT,KPP,SATUAN,KETR,"
                            Scom.CommandText &= "JRNPERI,RCSUR,NOPP,NILPAT,TGLPAT,JNSP,TPDOKT,NODOKT,SQNO,TGDOKT,UMURAT,"
                            Scom.CommandText &= "TGLMS1,NSBLN1,JBLNS1,NASTB1,NASTL1,TAKS1,TGLTB1,MPERI,YPERI,TGLTB,"
                            Scom.CommandText &= "MISC,D_MISC,FVEH,VNAMA,VJENIS,VTAHUN,VRANGKA,VMEREK,VMESIN,VSPEC,VSTNK,"
                            Scom.CommandText &= "VBPKB,VPOLI,NEWAT,JNSAT,TGLOAT,TGLKKOAT,KDCETAK,TGLUP,JAMUP,USRUP,STARTOAT)"
                            Scom.CommandText &= " VALUES ('" & barcode_plu & "','" & Toko.KCabATK & "','" & Toko.Kode & "',0," & qtyMax & "," & qtySO & ",0," & qtySO_total & ", "
                            Scom.CommandText &= "'','','',0,NOW(),'','','',"
                            Scom.CommandText &= "'',NOW(),'',NOW(),'','','','',"
                            Scom.CommandText &= "'','','',0,NOW(),'','','',0,NOW(),0,"
                            Scom.CommandText &= "NOW(),0,0,0,0,0,NOW(),0,0,NOW(),"
                            Scom.CommandText &= "'',NOW(),'','','',0,'','','','','',"
                            Scom.CommandText &= "'','','','',NOW(),NOW(),'',NOW(),NOW(),'',CURTIME());"
                            'revisi 18/11/20
                        ElseIf deskripsi.Contains("AT baru") And dt.Rows.Count <> 0 Then
                            Scom.CommandText = "UPDATE " & tabel_name & " SET QTYB = " & qtyMax & " , QTYR = " & qtySO & ",QTYSO = " & qtySO_total & ", STARTOAT = CURTIME() "
                            Scom.CommandText &= " WHERE NSERI = '" & barcode_plu & "';"
                        Else
                            'revisi Memo No. 208 - CPS - 20
                            '13/10/2020
                            Scom.CommandText = "UPDATE " & tabel_name & " SET QTYB = " & qtyMax & " , QTYR = " & qtySO & ", QTYH = " & qtyHilang & ",QTYSO = " & qtySO_total & ",RECID2='" & statusBarcode & "', STARTOAT = CURTIME()"

                            Scom.CommandText &= " WHERE NSERI = '" & barcode_plu & "';"
                        End If
                        Scom.ExecuteNonQuery()

                        'Jika Barcode/PLU diinputkan maka dianggap Sticker Rusak
                        If statusBarcode = "I" Then
                            Scom.CommandText = "UPDATE " & tabel_name & " SET STIKER_RUSAK = 1 WHERE NSERI = '" & barcode_plu & "';"
                            TraceLog("Update stiker_rusak: " & Scom.CommandText)
                            Scom.ExecuteNonQuery()
                        End If

                        'Set output
                        Result.NSeri = ""
                        Result.Deskripsi = ""
                        Result.Deskripsi2 = "Data sudah direkam !"
                    Else ' qty dat sama dengan 1
                        If qtySO_total > qty Then
                            Result.NSeri = barcode_plu
                            Result.Deskripsi = deskripsi
                            Result.Deskripsi2 = "Qty lbh, isi ulang!"
                            Exit Try

                        End If
                        If qtySO_total <> qty Then
                            qtyHilang = qty - qtySO_total
                        End If
                        '21/8/16
                        'case qtyso = 0, tidak bisa reset aktiva di SOTKIDM
                        If qtySO_total = 0 Then
                            qtySO_total = qtyHilang
                        End If
                        If deskripsi.Contains("AT baru") And dt.Rows.Count = 0 Then
                            Scom.CommandText = "INSERT INTO " & tabel_name & " (NSERI,KCAB,KLOK,QTY,QTYB,QTYR,QTYH,QTYSO,"
                            Scom.CommandText &= "RECID2,SFTJUAL,KETJUAL,HrgJual,TGLMUT,KCABMUT,KLOKMUT,TRANSMUT,"
                            Scom.CommandText &= "PEMAKAI,TGLTRANS,TROLD,TGLTROLD,NACCT,KPP,SATUAN,KETR,"
                            Scom.CommandText &= "JRNPERI,RCSUR,NOPP,NILPAT,TGLPAT,JNSP,TPDOKT,NODOKT,SQNO,TGDOKT,UMURAT,"
                            Scom.CommandText &= "TGLMS1,NSBLN1,JBLNS1,NASTB1,NASTL1,TAKS1,TGLTB1,MPERI,YPERI,TGLTB,"
                            Scom.CommandText &= "MISC,D_MISC,FVEH,VNAMA,VJENIS,VTAHUN,VRANGKA,VMEREK,VMESIN,VSPEC,VSTNK,"
                            Scom.CommandText &= "VBPKB,VPOLI,NEWAT,JNSAT,TGLOAT,TGLKKOAT,KDCETAK,TGLUP,JAMUP,USRUP,STARTOAT)"
                            Scom.CommandText &= " VALUES ('" & barcode_plu & "','" & Toko.KCabATK & "','" & Toko.Kode & "',0," & qtyMax & "," & qtySO & ",0," & qtySO_total & ", "
                            Scom.CommandText &= "'','','',0,NOW(),'','','',"
                            Scom.CommandText &= "'',NOW(),'',NOW(),'','','','',"
                            Scom.CommandText &= "'','','',0,NOW(),'','','',0,NOW(),0,"
                            Scom.CommandText &= "NOW(),0,0,0,0,0,NOW(),0,0,NOW(),"
                            Scom.CommandText &= "'',NOW(),'','','',0,'','','','','',"
                            Scom.CommandText &= "'','','','',NOW(),NOW(),'',NOW(),NOW(),'',CURTIME());"
                        ElseIf deskripsi.Contains("AT baru") And dt.Rows.Count <> 0 Then
                            Scom.CommandText = "UPDATE " & tabel_name & " SET QTYB = " & qtyMax & " , QTYR = " & qtySO & ",QTYSO = " & qtySO_total & ", STARTOAT = CURTIME() "
                            Scom.CommandText &= " WHERE NSERI = '" & barcode_plu & "';"
                        Else
                            'revisi Memo No. 208 - CPS - 20
                            '13/10/2020
                            Scom.CommandText = "UPDATE " & tabel_name & " SET QTYB = " & qtyMax & " , QTYR = " & qtySO & ", QTYH = " & qtyHilang & ",QTYSO = " & qtySO_total & ",RECID2='" & statusBarcode & "', STARTOAT = CURTIME() "
                            Scom.CommandText &= " WHERE NSERI = '" & barcode_plu & "';"
                        End If
                        Scom.ExecuteNonQuery()

                        'Jika Barcode/PLU diinputkan maka dianggap Sticker Rusak
                        If statusBarcode = "I" Then
                            Scom.CommandText = "UPDATE " & tabel_name & " SET STIKER_RUSAK = 1 WHERE NSERI = '" & barcode_plu & "';"
                            TraceLog("Update stiker_rusak: " & Scom.CommandText)
                            Scom.ExecuteNonQuery()
                        End If

                        'Set output
                        Result.NSeri = ""
                        Result.Deskripsi = ""
                        Result.Deskripsi2 = "Data sudah direkam !"
                    End If
                Else
                    'Set output
                    Result.NSeri = barcode_plu
                    Result.Deskripsi = deskripsi
                    Result.Deskripsi2 = "QTY " & qtySO & ", isi ulang!"
                End If

            Catch ex As Exception
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "UpdateQtyProduk", Scon)
                utility.TraceLogTxt("Error - UpdateQtyProduk " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                Scon.Close()
            End Try
        End SyncLock

        Return Result
    End Function

    'revisi Memo No. 208 - CPS - 20
    '13/10/2020
    Public Function CekKolomAktiva(ByVal tabelname As String) As Boolean
        'Dim connection As New ClsConnection
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim result As Boolean = True
        Dim Scom As New MySqlCommand("", Scon)

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SHOW COLUMNS FROM " & tabelname & " LIKE 'STARTOAT';"
            If Scom.ExecuteScalar <> "" Then
                Scom.CommandText = "SHOW COLUMNS FROM " & tabelname & " LIKE 'ENDOAT';"
                If Scom.ExecuteScalar <> "" Then
                    result = True
                Else
                    result = False
                End If
            Else
                result = False
            End If

            Scom.CommandText = "SHOW COLUMNS FROM " & tabelname & " LIKE 'STIKER_RUSAK';"
            If ("" & Scom.ExecuteScalar) = "" Then
                Scom.CommandText = "INSERT IGNORE INTO VIR_BACAPROD(JENIS,FILTER,KET) VALUES ('main_sticker_rusak', '', 'Memo 208-CPS-20 SOTKIDM Kolom STIKER_RUSAK')"
                Scom.ExecuteNonQuery()

                Scom.CommandText = "ALTER TABLE " & tabelname & " ADD COLUMN STIKER_RUSAK INTEGER DEFAULT 0"
                TraceLog("Alter tabel stiker_rusak: " & Scom.CommandText)
                Scom.ExecuteNonQuery()
            End If
        Catch ex As Exception
            result = False
        Finally
            Scon.Close()
        End Try

        Return result

    End Function

    'revisi Memo No. 208 - CPS - 20
    '13/10/2020
    Public Function GetListAktiva(ByVal tabel_name As String, ByVal keyNSeri As String, ByVal Toko As ClsToko) As List(Of ClsAktiva)
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)
        Dim ResultList As New List(Of ClsAktiva)
        Dim IsLengthValid As Boolean = True

        If Scon Is Nothing Then
            utility.TraceLogTxt("Error - GetListAktiva (connection Nothing) " & vbCrLf & "keyNSeri:" & keyNSeri)
            Return ResultList
            Exit Function
        End If

        SyncLock Scon
            Try
                If IsLengthValid Then
                    'Ambil data planogram
                    Dim DtAktiva As New DataTable
                    Dim IsATBaru As Boolean = False
                    If Scon.State = ConnectionState.Closed Then
                        Scon.Open()
                    End If
                    Scom.CommandText = "SELECT NSERI,BARANG FROM " & tabel_name & " WHERE nseri LIKE '%" & keyNSeri & "%';"
                    TraceLog("GetListAktiva-Q1: " & Scom.CommandText)
                    Dim sDap As New MySqlDataAdapter(Scom)
                    sDap.Fill(DtAktiva)
                    If DtAktiva.Rows.Count > 0 Then
                        For i As Integer = 0 To DtAktiva.Rows.Count - 1
                            ResultList.Add(New ClsAktiva() With {.NSeri = DtAktiva.Rows(i)("NSERI").ToString, .Deskripsi = DtAktiva.Rows(i)("BARANG")})
                        Next
                        Return ResultList
                    End If
                End If

            Catch ex As Exception
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetListAktiva", Scon)
                utility.TraceLogTxt("Error - GetListAktiva " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                Scon.Close()
            End Try
        End SyncLock

        Return ResultList
    End Function

End Class
