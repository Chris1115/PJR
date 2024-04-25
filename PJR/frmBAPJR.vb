Imports MySql.Data.MySqlClient
Imports IDM.InfoToko
Imports PJR.ClsPJRController
Imports IDM.Fungsi
Imports IDM.Report
Public Class frmBAPJR
    Dim dt As New DataTable
    Dim cFileSO As String = "BS_PJR_" & Format(Now, "yyMMdd") & FormMain.Toko.Kode.Substring(0, 1) & ""

    Private Sub frmBAPJR_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dt.Clear()
        dt = loadDataBA_AS_PJR(cFileSO)
        'MsgBox(dt.Rows.Count)
        DataGridView1.DataSource = dt

        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).Width = 300
        DataGridView1.Columns(2).Width = 100

        DataGridView1.ReadOnly = True
        DataGridView1.Refresh()

    End Sub

    Private Sub btnAdjust_Click(sender As Object, e As EventArgs) Handles btnAdjust.Click
        If MessageBox.Show(("Lanjutkan Adjust Berita Acara ?"), Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then
            Exit Sub
        End If

        'Dim Scon As New MySqlConnection(Get_KoneksiSQL)
        'Dim Scom As New MySqlCommand("", Scon)
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Madp As New MySqlDataAdapter("", Conn)

        Dim Stran As MySqlTransaction

        Dim dt As New DataTable

        Try
            Conn.Open()

            If Not IsTableExists(cFileSO) Then
                MessageBox.Show(("File Berita Acara tidak ada."), Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If


            Stran = Conn.BeginTransaction
            Mcom.Transaction = Stran

            'Cek Cetak 
            Mcom.CommandText = "Select Count(*) From " & cFileSO & " Where SOID='A'  "
            If Mcom.ExecuteScalar > 0 Then
                MessageBox.Show(("BA AM/AS Sudah Adjust."), Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            Else
                Mcom.CommandText = "Select Count(*) From " & cFileSO & " Where (SOID<>'L' or SOID IS NULL)  "
                If Mcom.ExecuteScalar > 0 Then
                    'MessageBox.Show(frmSO.lgg.getmsg(696, "Opname BA belum di-Input/Cetak."), Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    'Exit Sub
                End If
            End If

            'Ambil NKL
            Dim nNKL As Integer
            Dim nSeqno As Integer = 1

            Mcom.CommandText = "Update " & cFileSO & " Set ttl=com Where ttl>com "
            TraceLog("PROSES BA AM/AS: " & Mcom.CommandText, TipeTracelog.Info)
            Mcom.ExecuteNonQuery()

            Mcom.CommandText = "Select max(cast(bukti_no as unsigned)) From mstran Where year(bukti_tgl)=year(now()) " &
            "And RTYPE='X' AND LOKASI='01'"
            nNKL = "0" & Nb(Mcom.ExecuteScalar)

            If nNKL = 0 Then
                Mcom.CommandText = "Select DOCNO From Const Where RKEY='NKL'"
                nNKL = "0" & Mcom.ExecuteScalar
                If nNKL = 0 Then
                    nNKL = 1
                    Mcom.CommandText = "Delete CONST Where RKEY='NKL';"
                    Mcom.CommandText = "INSERT INTO CONST(RKEY,`DESC`,DOCNO)VALUES('NKL','NOTA KURANG LEBIH',1);"
                    Mcom.ExecuteNonQuery()
                End If
            Else
                nNKL = nNKL + 1
                Mcom.CommandText = "Update Const Set DOCNO='" & nNKL & "' Where RKEY='NKL'"
                Mcom.ExecuteNonQuery()
            End If

            'Flag SO1 to A
            Mcom.CommandText = "UPDATE Const SET JENIS='A',PERIOD2=IF(PERIOD1<>CURDATE(),PERIOD1,NULL),PERIOD1=CURDATE() WHERE RKEY='SO2'"
            Mcom.ExecuteNonQuery()

            'Insert WTRAN
            Madp.SelectCommand.CommandText = "Select PRDCD,`DESC` as Deskripsi,HPP,TTL,COM,(IIF(TTL is NULL,0,TTL)-COM) As Selisih from " &
                cFileSO & " Where ( (IIF(TTL is NULL,0,TTL)-COM)<>0) AND SOTYPE IN ('1','2')  "
            dt.Clear()
            Madp.SelectCommand.CommandText = CType(Madp.SelectCommand.CommandText, String).Replace("IIF(", "IF(")
            Madp.Fill(dt)

            Dim SQL As String = ""
            Dim cToko As String = Get_KodeToko()
            Dim nRMPP, nRNBH, nSMPP, nSNBH As Double
            Dim sReport As String = ""

            Dim sw As New IO.StreamWriter(Get_PathIDM() & "\ADJ_" & cFileSO & ".TXT", False)

            sw.WriteLine("".PadLeft(1) & Chr(218) & Chr(196).ToString.PadRight(90, Chr(196)) & Chr(191))
            sw.Write(Chr(18) + (Chr(27) + "W" + "0"))
            sw.WriteLine("".PadLeft(1) & Chr(179) & "".PadRight(90, " ") & Chr(179))
            'sw.WriteLine("".PadLeft(1) & Chr(179) & Get_KodeToko() & " - " & Strings.Left(Get_LokasiToko(), 30) & "".PadRight(61, " ") & Chr(179))
            'sw.WriteLine(Chr(15) + Chr(14) + (Chr(27) + "W" + "1"))
            sw.WriteLine("".PadLeft(1) & Chr(179) & Chr(15) + Chr(14) + (Chr(27) + "W" + "1") & "BERITA ACARA AM/AS".PadLeft(45) & Chr(20) + Chr(18) + (Chr(27) + "W" + "0") & "".PadRight(36, " ") & Chr(179))
            'sw.WriteLine(Chr(20) + Chr(18) + (Chr(27) + "W" + "0"))
            sw.WriteLine("".PadLeft(1) & Chr(179) & "".PadRight(90, " ") & Chr(179))
            sw.WriteLine("".PadLeft(1) & Chr(179) & "Pada hari ini tgl " & Format(Now, "dd-MM-yyyy") & " Jam " & Format(Now, "HH:mm:ss") & " telah dilakukan koreksi data LPP" & "".PadRight(16, " ") & Chr(179))
            sw.WriteLine("".PadLeft(1) & Chr(179) & "di Toko Idm." & Get_KodeToko() & " - " & Strings.Left(Get_LokasiToko(), 30) & "".PadRight(49, " ") & Chr(179))
            sw.WriteLine("".PadLeft(1) & Chr(179) & "".PadRight(90, " ") & Chr(179))
            'sw.Write("".PadLeft(1) & Chr(179) & strCenter("REF : " & cFileSO, 77)) : sw.WriteLine(Chr(20) + Chr(18) + (Chr(27) + "W" + "0") & "".PadRight(13, " ") & Chr(179))
            sw.WriteLine("".PadLeft(1) & Chr(179) & "".PadRight(90, " ") & Chr(179))
            sw.WriteLine("".PadLeft(1) & Chr(179) & "Dengan rincian sbb." & "".PadRight(71, " ") & Chr(179))
            sw.WriteLine("".PadLeft(1) & Chr(179) & "".PadRight(90, " ") & Chr(179))
            sw.WriteLine("".PadLeft(1) & Chr(179) & "".PadRight(90, " ") & Chr(179))

            For i As Integer = 0 To dt.Rows.Count - 1
                Mcom.CommandText = "Select SINGKATAN FROM PRODMAST Where PRDCD='" & dt.Rows(i)("PRDCD") & "'"
                'lblStatus.Text = dt.Rows(i)("PRDCD") & "-" & Scom.ExecuteScalar
                Application.DoEvents()
                Dim nPrice_Prodmast As Integer = 0
                Mcom.CommandText = "Select PRICE FROM PRODMAST Where PRDCD='" & dt.Rows(i)("PRDCD") & "'"
                nPrice_Prodmast = Mcom.ExecuteScalar

                'ProgressBar1.Value += 1
                SQL = "Select Count(*) From MSTRAN WHERE RTYPE='X' AND LOKASI='01' "
                SQL &= "AND BUKTI_NO=" & nNKL & " AND PRDCD='" & dt.Rows(i)("PRDCD") & "' And year(bukti_tgl)=year(now()) "
                Mcom.CommandText = SQL
                If Mcom.ExecuteScalar = 0 Then
                    Mcom.CommandText = "Show Columns From Pos.Mstran Like 'Price_jual'"
                    If Mcom.ExecuteScalar & "" <> "" Then
                        Mcom.CommandText = "Show Columns From Pos.Mstran Like 'Gross_Jual'"
                        If Mcom.ExecuteScalar & "" <> "" Then
                            SQL = "INSERT IGNORE MSTRAN("
                            SQL &= "PRDCD"
                            SQL &= ",RTYPE"
                            SQL &= ",BUKTI_NO"
                            SQL &= ",BUKTI_TGL"
                            'SQL &= ",SEQNO"
                            SQL &= ",ISTYPE"
                            SQL &= ",LOKASI"
                            SQL &= ",KETER"
                            SQL &= ",QTY"
                            SQL &= ",PRICE"
                            SQL &= ",GROSS"
                            SQL &= ",PRICE_JUAL"
                            SQL &= ",GROSS_JUAL"
                            SQL &= ")VALUES("
                            SQL &= "'" & dt.Rows(i)("PRDCD") & "'"
                            SQL &= ",'X'"
                            SQL &= "," & nNKL & ""
                            SQL &= ",NOW()"
                            'SQL &= "," & nSeqno & "" : nSeqno += 1
                            SQL &= ",'BS'"
                            SQL &= ",'01'"
                            SQL &= ",'PJR'"
                            SQL &= "," & dt.Rows(i)("Selisih") & ""
                            SQL &= "," & dt.Rows(i)("HPP") & ""
                            SQL &= "," & dt.Rows(i)("Selisih") * dt.Rows(i)("HPP") & ""
                            SQL &= "," & nPrice_Prodmast & ""
                            SQL &= "," & dt.Rows(i)("Selisih") * nPrice_Prodmast & ""
                            SQL &= ")"
                            Mcom.CommandText = SQL
                            Mcom.ExecuteNonQuery()

                            If dt.Rows(i)("Selisih") > 0 Then
                                nSMPP += 1
                                nRMPP += dt.Rows(i)("Selisih") * dt.Rows(i)("HPP")
                            Else
                                nSNBH += 1
                                nRNBH += dt.Rows(i)("Selisih") * dt.Rows(i)("HPP")
                            End If

                            'Kirim Laporan ke Monitoring utk adjust plus dan ada transaksi di MTRAN
                            If dt.Rows(i)("Selisih") > 0 Then
                                Mcom.CommandText = "Select Sum(if(rtype='J',qty,qty*-1)) As Qty From Mtran Where Tanggal=Curdate() " &
                                "And PLU='" & dt.Rows(i)("PRDCD") & "' Group By PLU"
                                Dim nAdjustPlus As Integer = 0 & Mcom.ExecuteScalar
                                If nAdjustPlus > 0 Then
                                    Mcom.CommandText = "Select Cast(concat(Kasir_Name,'|',Tanggal,'|',Shift,'|',Station) As Char) As Kasir From Initial Where Station='" & IDM.Fungsi.Get_Station & "' " &
                                    "Order By tanggal Desc "
                                    Dim sLaporan As String = "" & Mcom.ExecuteScalar
                                    sLaporan &= "|" & dt.Rows(i)("PRDCD") & "|" & dt.Rows(i)("Selisih").ToString & "|" & (dt.Rows(i)("Selisih") * dt.Rows(i)("HPP")).ToString
                                    sReport &= sLaporan & ";"
                                End If
                            End If

                            sw.WriteLine("".PadLeft(1) & Chr(179) & "PLU - Deskrip: " & dt.Rows(i)("PRDCD") & " - " & Strings.Left(dt.Rows(i)("Deskripsi"), 30) & "".PadRight(34, " ") & Chr(179))
                            sw.WriteLine("".PadLeft(1) & Chr(179) & "Qty.Input     Qty.LPP     Selisih Qty     Selisih Rp" & "".PadRight(38, " ") & Chr(179))
                            sw.WriteLine("".PadLeft(1) & Chr(179) & dt.Rows(i)("TTL").ToString.PadLeft(9) & dt.Rows(i)("COM").ToString.PadLeft(12) & ("" & dt.Rows(i)("Selisih")).ToString.PadLeft(16) & "" & ("" & (dt.Rows(i)("Selisih") * nPrice_Prodmast)).ToString.PadLeft(15) & "" & "" & "".PadRight(38, " ") & Chr(179))
                            sw.WriteLine("".PadLeft(1) & Chr(179) & "".PadRight(90, " ") & Chr(179))

                            ''9-jun-08 beben: Adj dikumulatif dengan selisih hasil So-nya
                            'Scom.CommandText = "UPDATE STMAST SET ADJ=ADJ+0" & dt.Rows(i)("SELISIH") & _
                            '    " ,QTY=0" & dt.Rows(i)("TTL") & " where PRDCD='" & dt.Rows(i)("PRDCD") & "' "
                            'Scom.ExecuteNonQuery()

                            '20-aug-08 beben: QTY dihitung ulang
                            'begbal-sales+trfin-trfout+adj+retur
                            Mcom.CommandText = "Show Columns From Pos.Stmast Like 'rp_adj_k'"
                            If Mcom.ExecuteScalar & "" <> "" Then
                                Mcom.CommandText = "Show Columns From Pos.Stmast Like 'rp_adj_l'"
                                If Mcom.ExecuteScalar & "" <> "" Then
                                    Mcom.CommandText = "UPDATE STMAST SET ADJ=ADJ+" & dt.Rows(i)("SELISIH")
                                    Mcom.CommandText &= ",QTY=begbal-sales+trfin-trfout+adj+retur" ' & dt.Rows(i)("SELISIH")
                                    If dt.Rows(i)("Selisih") < 0 Then
                                        Mcom.CommandText &= ",rp_adj_k=rp_adj_k+" & (-1 * dt.Rows(i)("Selisih") * nPrice_Prodmast)
                                    End If
                                    If dt.Rows(i)("Selisih") >= 0 Then
                                        Mcom.CommandText &= ",rp_adj_l=rp_adj_l+" & (dt.Rows(i)("Selisih") * nPrice_Prodmast)
                                    End If
                                    Mcom.CommandText &= " where PRDCD='" & dt.Rows(i)("PRDCD") & "' "
                                    Mcom.ExecuteNonQuery()


                                    Mcom.CommandText = "UPDATE " & cFileSO & " SET SOID='A',ADJDT=CURDATE()," &
                                        "ADJTIME=curtime() " &
                                        "WHERE  PRDCD='" & dt.Rows(i)("PRDCD") & "' "

                                    Mcom.ExecuteNonQuery()
                                End If
                            Else
                                Mcom.CommandText = "UPDATE STMAST SET ADJ=ADJ+" & dt.Rows(i)("SELISIH")
                                Mcom.CommandText &= ",QTY=begbal-sales+trfin-trfout+adj+retur" ' & dt.Rows(i)("SELISIH")
                                Mcom.CommandText &= " where PRDCD='" & dt.Rows(i)("PRDCD") & "' "
                                Mcom.ExecuteNonQuery()


                                Mcom.CommandText = "UPDATE " & cFileSO & " SET SOID='A',ADJDT=CURDATE()," &
                                    "ADJTIME=curtime() " &
                                    "WHERE  PRDCD='" & dt.Rows(i)("PRDCD") & "'  "

                                Mcom.ExecuteNonQuery()
                            End If
                        End If
                    Else
                        SQL = "INSERT IGNORE MSTRAN("
                        SQL &= "PRDCD"
                        SQL &= ",RTYPE"
                        SQL &= ",BUKTI_NO"
                        SQL &= ",BUKTI_TGL"
                        'SQL &= ",SEQNO"
                        SQL &= ",ISTYPE"
                        SQL &= ",LOKASI"
                        SQL &= ",KETER"
                        SQL &= ",QTY"
                        SQL &= ",PRICE"
                        SQL &= ",GROSS"
                        SQL &= ")VALUES("
                        SQL &= "'" & dt.Rows(i)("PRDCD") & "'"
                        SQL &= ",'X'"
                        SQL &= "," & nNKL & ""
                        SQL &= ",NOW()"
                        'SQL &= "," & nSeqno & "" : nSeqno += 1
                        SQL &= ",'BS'"
                        SQL &= ",'01'"
                        SQL &= ",'PJR'"
                        SQL &= "," & dt.Rows(i)("Selisih") & ""
                        SQL &= "," & dt.Rows(i)("HPP") & ""
                        SQL &= "," & dt.Rows(i)("Selisih") * dt.Rows(i)("HPP") & ""
                        SQL &= ")"
                        Mcom.CommandText = SQL
                        Mcom.ExecuteNonQuery()

                        If dt.Rows(i)("Selisih") > 0 Then
                            nSMPP += 1
                            nRMPP += dt.Rows(i)("Selisih") * dt.Rows(i)("HPP")
                        Else
                            nSNBH += 1
                            nRNBH += dt.Rows(i)("Selisih") * dt.Rows(i)("HPP")
                        End If

                        'Kirim Laporan ke Monitoring utk adjust plus dan ada transaksi di MTRAN
                        If dt.Rows(i)("Selisih") > 0 Then
                            Mcom.CommandText = "Select Sum(if(rtype='J',qty,qty*-1)) As Qty From Mtran Where Tanggal=Curdate() " &
                            "And PLU='" & dt.Rows(i)("PRDCD") & "' Group By PLU"
                            Dim nAdjustPlus As Integer = 0 & Mcom.ExecuteScalar
                            If nAdjustPlus > 0 Then
                                Mcom.CommandText = "Select Cast(concat(Kasir_Name,'|',Tanggal,'|',Shift,'|',Station) As Char) As Kasir From Initial Where Station='" & IDM.Fungsi.Get_Station & "' " &
                                "Order By tanggal Desc "
                                Dim sLaporan As String = "" & Mcom.ExecuteScalar
                                sLaporan &= "|" & dt.Rows(i)("PRDCD") & "|" & dt.Rows(i)("Selisih").ToString & "|" & (dt.Rows(i)("Selisih") * dt.Rows(i)("HPP")).ToString
                                sReport &= sLaporan & ";"
                            End If
                        End If

                        sw.WriteLine("".PadLeft(1) & Chr(179) & "PLU - Deskrip: " & dt.Rows(i)("PRDCD") & " - " & Strings.Left(dt.Rows(i)("Deskripsi"), 30) & "".PadRight(34, " ") & Chr(179))
                        sw.WriteLine("".PadLeft(1) & Chr(179) & "Qty.Input     Qty.LPP     Selisih Qty     Selisih Rp" & "".PadRight(38, " ") & Chr(179))
                        sw.WriteLine("".PadLeft(1) & Chr(179) & dt.Rows(i)("TTL").ToString.PadLeft(9) & dt.Rows(i)("COM").ToString.PadLeft(12) & ("" & dt.Rows(i)("Selisih")).ToString.PadLeft(16) & "" & ("" & (dt.Rows(i)("Selisih") * nPrice_Prodmast)).ToString.PadLeft(15) & "" & "" & "".PadRight(38, " ") & Chr(179))
                        sw.WriteLine("".PadLeft(1) & Chr(179) & "".PadRight(90, " ") & Chr(179))


                        '20-aug-08 beben: QTY dihitung ulang
                        'begbal-sales+trfin-trfout+adj+retur
                        Mcom.CommandText = "UPDATE STMAST SET ADJ=ADJ+" & dt.Rows(i)("SELISIH")
                        Mcom.CommandText &= ",QTY=begbal-sales+trfin-trfout+adj+retur" ' & dt.Rows(i)("SELISIH")
                        Mcom.CommandText &= " where PRDCD='" & dt.Rows(i)("PRDCD") & "' "
                        Mcom.ExecuteNonQuery()


                        Mcom.CommandText = "UPDATE " & cFileSO & " SET SOID='A',ADJDT=CURDATE()," &
                                    "ADJTIME=curtime() " &
                                    "WHERE  PRDCD='" & dt.Rows(i)("PRDCD") & "'  "

                        Mcom.ExecuteNonQuery()
                    End If
                End If
            Next

            'sw.WriteLine("")
            sw.WriteLine("".PadLeft(1) & Chr(179) & "".PadRight(90, " ") & Chr(179))
            sw.WriteLine("".PadLeft(1) & Chr(179) & "Demikian Berita Acara ini dibuat dengan sebenar-benarnya." & "".PadRight(33, " ") & Chr(179))
            sw.WriteLine("".PadLeft(1) & Chr(179) & "".PadRight(90, " ") & Chr(179))
            sw.WriteLine("".PadLeft(1) & Chr(179) & "".PadLeft(10) & "     AM / AS             Chief of Store/Ast" & "".PadRight(37, " ") & Chr(179))
            sw.WriteLine("".PadLeft(1) & Chr(179) & "".PadRight(90, " ") & Chr(179))
            sw.WriteLine("".PadLeft(1) & Chr(179) & "".PadRight(90, " ") & Chr(179))
            sw.WriteLine("".PadLeft(1) & Chr(179) & "".PadRight(90, " ") & Chr(179))
            sw.WriteLine("".PadLeft(1) & Chr(179) & "".PadLeft(10) & "(              )          (               )" & "".PadRight(37, " ") & Chr(179))
            sw.WriteLine("".PadLeft(1) & Chr(179) & "+----------------------------------------------------------+" & "".PadRight(30, " ") & Chr(179))
            sw.WriteLine("".PadLeft(1) & Chr(179) & "SO.Net v." & Application.ProductVersion.PadLeft(10, "") & "".PadRight(73, " ") & Chr(179))
            sw.WriteLine("".PadLeft(1) & Chr(179) & "".PadRight(90, " ") & Chr(179))
            sw.WriteLine("".PadLeft(1) & Chr(192) & Chr(196).ToString.PadRight(90, Chr(196)) & Chr(217))
            sw.Write(Chr(18))
            sw.Flush()
            sw.Close()

            CetakText(Get_PathIDM() & "\ADJ_" & cFileSO & ".TXT")

            'Scom.CommandText = "UPDATE WTRAN W,PRODMAST P SET W.DIVISI=P.DIVISI,W.PTAG=P.PTAG,W.CAT_COD=P.CAT_COD,W.BKP=P.BKP,W.SUB_BKP=P.SUB_BKP WHERE (W.PRDCD=P.PRDCD) AND (W.DIVISI='' OR W.DIVISI IS NULL)"
            'Scom.ExecuteScalar()
            Dim testAlamat As String = ""
            SQL = "Select almt From Toko"
            Mcom.CommandText = SQL
            testAlamat = Nb(Mcom.ExecuteScalar()) ': MsgBox("Alamat Toko: " & testAlamat)
            Dim rpt As New IDM.Report.MonitorToko(IDM.InfoToko.Get_KodeToko, IDM.InfoToko.Get_LokasiToko, IDM.InfoToko.Get_Cabang, testAlamat)
            rpt.dt.AddStatusRow(20, "Status Adjust Plus : ", sReport)
            rpt.SendStatusReport() : rpt = Nothing

            If dt.Rows.Count > 0 Then
                Mcom.CommandText = "UPDATE CONST SET DOCNO=DOCNO+1 WHERE RKEY='NKL'"
                Mcom.ExecuteNonQuery()
            End If

            'Cetak Bukti ADJ
            'CetakBukti(nMPP, nNBH, nJENIS, nRMPP, nRNBH, cFILE, nNKL, nSMPP, nSNBH)

            'Commit
            Stran.Commit()


            'UPDATE kolom PLU_NAS di MSTRAN sesuai dgn PLU MD (Permintaaan Gideon email tgl 19-03-2012)
            Mcom.CommandText = "Update Mstran a, Prodmast b Set a.Plu_Nas=b.PluMD Where a.Prdcd=b.Prdcd " &
            "And (a.Plu_Nas is Null or a.Plu_Nas='') " &
            "And a.Istype='BS' " &
            "And a.Keter='PJR' " &
            "And Date(a.Bukti_Tgl)=Curdate() " &
            "And (b.Prdcd is Not Null or a.Plu_Nas='')"
            Mcom.ExecuteNonQuery()

            'MessageBox.Show(frmSO.lgg.getmsg(697, "Proses Adjust Berita Acara Selesai."), Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
            MessageBox.Show(("Proses Adjust Berita Acara Selesai." & vbCrLf & "Program akan memanggil form Bukti Peminjaman Tabung (Jika ada)"), Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)

            BuktiPeminjamanTabung()
        Catch ex As Exception
            Try
                Stran.Rollback()
            Catch
            End Try
            ShowError("Error Adjust Berita Acara.", ex)
            Exit Sub
        Finally
            'IO.File.Delete(Get_PathIDM() & "\" & cFileSO & ".DBF")
            Conn.Close()
        End Try
        Me.Close()
    End Sub

    Public Shared Sub BuktiPeminjamanTabung()
        Try
            Dim sProg As String = "POSUtil.dll"
            Dim versiToko As String = ""
            If IO.File.Exists(Application.StartupPath & "\" & sProg) Then
                Dim myBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Application.StartupPath & "\" & sProg)
                versiToko = myBuildInfo.FileVersion
            End If
            If Compare_Version(versiToko, "2.1.2", "POSUtil.dll") < 1 Then
                PanggilBPT()
            Else
                MsgBox("POSUtil harap diupdate ke versi 2.1.2")
            End If
        Catch ex As Exception
            TraceLog("Error Hide (BuktiPeminjamanTabung) : " & ex.ToString & vbCrLf & ex.StackTrace)
            ex = Nothing
        End Try
    End Sub


    Friend Shared Function Compare_Version(ByVal CurrentVersion As String, ByVal AnotherVersion As String, ByVal NamaProgram As String) As Integer
        Dim hasil As Integer
        Dim major_ftp As String = ""
        Dim minor_ftp As String = ""
        Dim build_ftp As String = ""
        Dim revision_ftp As String = ""
        Dim major_toko As String = ""
        Dim minor_toko As String = ""
        Dim build_toko As String = ""
        Dim revision_toko As String = ""
        Dim iMajor_ftp As Integer
        Dim iMinor_ftp As Integer
        Dim iBuild_ftp As Integer
        Dim iRevision_ftp As Integer
        Dim iMajor_toko As Integer
        Dim iMinor_toko As Integer
        Dim iBuild_toko As Integer
        Dim iRevision_toko As Integer
        Try
            Dim x1() As String
            Dim p1 As Integer
            If CurrentVersion.IndexOf(".") >= 0 Then
                If AnotherVersion.IndexOf(".") >= 0 Then
                    x1 = AnotherVersion.Split(".")
                    For p1 = 0 To x1.GetUpperBound(0)
                        If p1 = 0 Then major_ftp = x1(p1)
                        If p1 = 1 Then minor_ftp = x1(p1)
                        If p1 = 2 Then build_ftp = x1(p1)
                        If p1 = 3 Then revision_ftp = x1(p1)
                    Next
                    Dim x2() As String
                    Dim p2 As Integer
                    x2 = CurrentVersion.Split(".")
                    For p2 = 0 To x2.GetUpperBound(0)
                        If p2 = 0 Then major_toko = x2(p2)
                        If p2 = 1 Then minor_toko = x2(p2)
                        If p2 = 2 Then build_toko = x2(p2)
                        If p2 = 3 Then revision_toko = x2(p2)
                    Next
                    'Konversi String ke Integer
                    iMajor_ftp = "0" & major_ftp
                    iMinor_ftp = "0" & minor_ftp
                    iBuild_ftp = "0" & build_ftp
                    iRevision_ftp = "0" & revision_ftp
                    iMajor_toko = "0" & major_toko
                    iMinor_toko = "0" & minor_toko
                    iBuild_toko = "0" & build_toko
                    iRevision_toko = "0" & revision_toko
                    If iMajor_ftp > iMajor_toko Then
                        hasil = 1
                    ElseIf iMajor_ftp = iMajor_toko And iMinor_ftp > iMinor_toko Then
                        hasil = 1
                    ElseIf iMajor_ftp = iMajor_toko And iMinor_ftp = iMinor_toko And iBuild_ftp > iBuild_toko Then
                        hasil = 1
                    ElseIf iMajor_ftp = iMajor_toko And iMinor_ftp = iMinor_toko And iBuild_ftp = iBuild_toko Then
                        hasil = 0
                    Else
                        hasil = -1
                    End If
                Else
                    hasil = -1
                End If
            Else
                hasil = -1
            End If
        Catch ex As Exception
            ShowError("Error di Compare_Version, program='" & NamaProgram & "', version 1='" & CurrentVersion & "', version 2='" & AnotherVersion & "'", ex)
        End Try
        Return hasil
    End Function

    Public Shared Function PanggilBPT() As Boolean
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand
        Dim Sdap As New MySqlDataAdapter(Scom)
        Dim sql As String = ""
        Dim Filter As String = ""
        Dim dtPLU As New DataTable

        Try
            Scon.Open()
            Scom.Connection = Scon
            Scom.CommandText = ""

            'cek vir_bacaprod ada atau tidak
            Scom.CommandText = "show tables like 'vir_bacaprod'"
            If Scom.ExecuteNonQuery & "" <> "" Then
                sql = "select filter from vir_bacaprod where program = 'POSUTIL' and jenis = 'PINJAMAN' "
                Scom.CommandText = sql

                If Scom.ExecuteScalar & "" <> "" Then
                    Filter = Scom.ExecuteScalar & ""
                Else
                    TraceLog("Filter dari vir_bacaprod kosong")
                    Exit Function
                End If
            End If

            'jika ada tabel plu_pinjaman ada
            Scom.CommandText = "" & Filter & ""
            Dim Filter2 As String = IFNULL(Scom.ExecuteScalar, "")
            If Filter2 <> "" Then

                Scom.CommandText = "SELECT prdcd,singkatan "
                Scom.CommandText &= "FROM prodmast "
                Scom.CommandText &= "WHERE prdcd IN (" & Filter2 & ") "

                Sdap.SelectCommand = Scom
                Sdap.Fill(dtPLU)

                For i As Short = 0 To dtPLU.Rows.Count - 1
                    Dim BPTTampil As New PosUtil.FrmPinjamTabung
                    BPTTampil.BPT_Tampil(dtPLU.Rows(i)("prdcd"))
                Next

            Else
                TraceLog("Tidak ada PLU LPG Pinjaman yang dipilih")
            End If
        Catch ex As Exception
            TraceLog("Error PanggilBPT: " & ex.Message)
        End Try

    End Function


End Class