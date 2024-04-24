Imports MySql.Data.MySqlClient
Imports IDM.Fungsi
Public Class ClsCekDisplayController

    Private utility As New Utility

    ''' <summary>
    ''' untuk cek table proses Planogram
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CekTableDisplay() As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Mda As New MySqlDataAdapter("", Conn)
        Dim Rtn As New Boolean

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            Mcom.CommandText = "Create table if not exists CekDisplay ("
            Mcom.CommandText &= " PLU varchar(8) not null, "
            Mcom.CommandText &= " TglScan Date not null, "
            Mcom.CommandText &= " Deskripsi Varchar(50), "
            Mcom.CommandText &= " NoShelf Integer, "
            Mcom.CommandText &= " NoRak Integer, "
            Mcom.CommandText &= " KodeModis Varchar(20), "
            Mcom.CommandText &= " NamaModis Varchar(99), "
            Mcom.CommandText &= " KiriKanan Int(3), "
            Mcom.CommandText &= " Kap_disp decimal(12,0), "
            Mcom.CommandText &= " Qty_disp decimal(12,0), "
            Mcom.CommandText &= " NIK VARCHAR(20), "
            Mcom.CommandText &= " NAMA VARCHAR(50), "
            Mcom.CommandText &= " JABATAN VARCHAR(50), "
            Mcom.CommandText &= " Primary Key(PLU,KodeModis,TglScan)"
            Mcom.CommandText &= " )"
            Mcom.ExecuteNonQuery()

            Try
                Mcom.CommandText = " Select count(*) From Information_schema.Columns  
                                 Where TABLE_SCHEMA='pos' 
                                 AND Table_Name In('CEKDISPLAY')  And Column_Name='NAMA'"
                If Mcom.ExecuteScalar = 0 Then
                    Mcom.CommandText = "Alter table CEKDISPLAY "
                    Mcom.CommandText &= "ADD COLUMN `NAMA` Varchar(50) DEFAULT ''"
                    Mcom.ExecuteNonQuery()
                End If
            Catch ex As Exception

            End Try
            Try
                Mcom.CommandText = " Select count(*) From Information_schema.Columns  
                                 Where TABLE_SCHEMA='pos' 
                                 AND Table_Name In('CEKDISPLAY')  And Column_Name='JABATAN'"
                If Mcom.ExecuteScalar = 0 Then
                    Mcom.CommandText = "Alter table CEKDISPLAY "
                    Mcom.CommandText &= "ADD COLUMN `JABATAN` Varchar(50) DEFAULT ''"
                    Mcom.ExecuteNonQuery()
                End If
            Catch ex As Exception

            End Try

            Mcom.CommandText = "Create table if not exists temp_CekDisplay ("
            Mcom.CommandText &= " PLU varchar(8) not null, Primary Key(PLU)"
            Mcom.CommandText &= " )"
            Mcom.ExecuteNonQuery()


            Rtn = True
        Catch ex As Exception
            Rtn = False
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "CekTableDisplay", Conn)
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
            Madp.SelectCommand.CommandText = "select distinct kodemodis from rak r;"
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
    Public Function CekModis(ByVal Modis As String) As String
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As String = ""

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            'hasil uat 9/6/23
            Mcom.CommandText = "select TIPERAK from rak where KODEMODIS = '" & Modis & "';"
            If Mcom.ExecuteScalar = "G" Then
                Rtn = "Rak Sewa"

            Else
                Mcom.CommandText = "select KET_RAK from rak where KODEMODIS = '" & Modis & "';"
                Rtn = Mcom.ExecuteScalar

            End If

        Catch ex As Exception
            Rtn = ""
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
    Public Function GetDeskripsiListingDisplay(ByVal tabel_name As String, ByVal barcode_plu As String,
                                          ByVal NamaRak As String, ByVal User As ClsUser) As ClsCekDisplay
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Result As New ClsCekDisplay

        If Conn Is Nothing Then
            utility.TraceLogTxt("Error - GetDeskripsiListingDisplay (connection Nothing) " & vbCrLf & "PLU:" & barcode_plu)
            Return Result
            Exit Function
        End If

        SyncLock Conn
            Try
                If Conn.State = ConnectionState.Closed Then
                    Conn.Open()
                End If

                Dim DtPlano As New DataTable
                Mcom.CommandText = "SELECT a.prdcd,a.DESC2 FROM prodmast a 
                                    LEFT JOIN (SELECT prdcd, kodemodis,ket_rak, noshelf, kirikanan FROM rak ) b ON a.prdcd = b.prdcd 
                                    LEFT JOIN (SELECT prdcd,KAP_DISP FROM stmast) c ON a.prdcd = c.prdcd 
                                    LEFT JOIN (SELECT plu,barcd FROM barcode) d ON a.prdcd = d.plu
                                    WHERE (a.prdcd = '" & barcode_plu & "' OR d.BARCD = '" & barcode_plu & "') AND b.kodemodis = '" & NamaRak & "'"

                Dim sDap As New MySqlDataAdapter(Mcom)
                sDap.Fill(DtPlano)
                utility.Tracelog("Query", Mcom.CommandText, "GetDeskripsiListingDisplay", Conn)
                'utility.Tracelog("Query", DtPlano.Rows.Count, "GetDeskripsiListingDisplay", Conn)

                If DtPlano.Rows.Count > 0 Then

                    Result.Prdcd = DtPlano.Rows(0)("prdcd").ToString

                    Result.Desc = DtPlano.Rows(0)("DESC2").ToString
                    If Result.Desc.Length > 20 Then
                        Result.Desc = Result.Desc.Substring(0, 20)
                    End If
                Else
                    Result.Prdcd = ""
                    Result.Desc = "Tidak Ditemukan"

                End If

            Catch ex As Exception
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetDeskripsiListingDisplay", Conn)
                utility.TraceLogTxt("Error - GetDeskripsiListingDisplay " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
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
    Public Function SelesaiCekPlano(ByVal tabel_name As String, ByVal NamaRak As String,
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
            Madp.SelectCommand.CommandText = " SELECT PLU FROM CekDisplay WHERE DATE(TGLSCAN) = CURDATE() AND KODEMODIS = '" & NamaRak & "'"
            Madp.SelectCommand.CommandText &= " GROUP BY PLU;"
            Madp.Fill(DtCP)
            If DtCP.Rows.Count = 0 Then
                Rtn = ""
            Else
                Rtn = "Ada"
            End If

        Catch ex As Exception
            Rtn = "Err"
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "SelesaiCekDisplay", Conn)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function DeleteTempCekDisplay(ByVal tabel_name As String) As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim DtCP As New DataTable
        Dim Rtn As Boolean

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Mcom.CommandText = "DELETE FROM TEMP_CEKDISPLAY"
            Mcom.ExecuteNonQuery()
            Rtn = True

        Catch ex As Exception
            Rtn = False
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "DeleteTempCekDisplay", Conn)
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function


    Public Function SimpanQTYCekDisplay(ByVal tabel_name As String, ByVal barcode_plu As String, ByVal qty As String, ByVal namarak As String, ByVal nik As String, ByVal nama As String, ByVal jabatan As String) As Boolean
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim result As Boolean = False
        Dim mcom As New MySqlCommand("", conn)

        SyncLock conn
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                mcom.CommandText = "SELECT COUNT(1) FROM " & tabel_name & " WHERE PLU = '" & barcode_plu & "' and kodemodis = '" & namarak & "' AND TGLSCAN = CURDATE()"
                If mcom.ExecuteScalar = 0 Then
                    mcom.CommandText = "INSERT IGNORE INTO " & tabel_name & "
                                   SELECT a.prdcd,curdate(),a.DESC2,B.noshelf,IF(e.no_rak IS NOT NULL, e.no_rak, '1'),b.kodemodis,b.ket_rak,b.kirikanan,IF(b.tiperak ='G', a.nplus,c.kap_disp),'" & qty & "', '" & nik & "', '" & nama & "','" & jabatan & "' FROM prodmast a 
                                    LEFT JOIN (SELECT prdcd,TIPERAK, kodemodis,ket_rak, noshelf,norak,kirikanan FROM rak ) b ON a.prdcd = b.prdcd 
                                    LEFT JOIN (SELECT prdcd,KAP_DISP FROM stmast) c ON a.prdcd = c.prdcd
                                    LEFT JOIN (SELECT plu,barcd FROM barcode) d ON a.prdcd = d.plu
                                    LEFT JOIN (SELECT no_rak, modisp FROM bracket) e ON b.kodemodis = e.modisp
                                    WHERE (a.prdcd = '" & barcode_plu & "' OR d.barcd = '" & barcode_plu & "') and b.kodemodis = '" & namarak & "'"
                Else
                    mcom.CommandText = "UPDATE " & tabel_name & " SET QTY_DISP = '" & qty & "'
                                    WHERE PLU = '" & barcode_plu & "' and kodemodis = '" & namarak & "' AND TGLSCAN = CURDATE()"
                End If
                mcom.ExecuteNonQuery()

                utility.Tracelog("Query", mcom.CommandText, "SimpanQTYCekDisplay", conn)

                mcom.CommandText = "INSERT IGNORE INTO TEMP_CEKDISPLAY VALUES('" & barcode_plu & "')"
                mcom.ExecuteNonQuery()

                result = True

            Catch ex As Exception
                result = False
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "SimpanQTYCekDisplay", conn)
                utility.TraceLogTxt("Error - SimpanQTYCekDisplay " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                conn.Close()
            End Try
        End SyncLock

        Return result
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


            Madp.SelectCommand.CommandText = "SELECT DISTINCT nik FROM SOPPAGENT.ABSSETTINGSHIFT a LEFT JOIN `soppagent`.`abspegawaimst` b ON a.nik = b.menoin"


            Madp.Fill(Rtn)


        Catch ex As Exception
            TraceLog("Error GetPersonilCetak: " & ex.Message & ex.StackTrace)

            Rtn = New DataTable
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function

    Public Function GetJabatan(ByVal nik As String) As String
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As String = ""
        Dim Mcom As New MySqlCommand("", Conn)
        Dim dt As New DataTable
        Dim jabatan As String = ""
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If


            Mcom.CommandText = "SELECT JABATAN FROM soppagent.abspegawaimst WHERE MENOIN = '" & nik & "'"
            Rtn = Mcom.ExecuteScalar


        Catch ex As Exception
            TraceLog("Error GetJabatan: " & ex.Message & ex.StackTrace)

            Rtn = ""
        Finally
            Conn.Close()
        End Try
        Return Rtn
    End Function



End Class
