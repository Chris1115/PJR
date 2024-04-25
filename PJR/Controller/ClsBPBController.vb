Imports MySql.Data.MySqlClient
Imports PJR.FormMain
Imports IDM.Fungsi
Public Class ClsBPBController

    Private utility As New Utility

    ''' <summary>
    ''' untuk cek container proses BPB
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CekPluContainer(ByVal ContainerNo As String, ByVal KodeDC As String,
                                    ByRef Box1 As Integer, ByRef Box2 As Integer) As DataTable
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mcom As New MySqlCommand("", conn)
        Dim mda As New MySqlDataAdapter("", conn)
        Dim Rtn As New DataTable
        Dim DC As String = KodeDC.Split("-")(0).ToString
        Dim DOCNO As String = KodeDC.Split("-")(1).ToString
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            'Menghitung jumlah container
            '438/03-23/E/PMO
            mcom.CommandText = "SELECT COUNT(*) FROM dcp_boxplu " &
                               "WHERE Dus_No LIKE '" & ContainerNo & "%' " &
                               "AND KIRIM = '" & DC & "' AND DOCNO = '" & DOCNO & "';"
            TraceLog("CekPluContainer-Q1: " & mcom.CommandText)
            Box1 = Convert.ToInt32(mcom.ExecuteScalar)
            'MsgBox(mcom.CommandText)
            If Box1 > 0 Then
                '438/03-23/E/PMO
                mda.SelectCommand.CommandText = "SELECT  DUS_NO, Prdcd, Nama, Qty FROM dcp_boxplu " &
                                                "WHERE recid <> 1 " &
                                                "AND Dus_no LIKE '" & ContainerNo & "%' " &
                                                "AND (DPDID = '' OR DPDID IS NULL) " &
                                                "AND KIRIM = '" & DC & "' AND DOCNO = '" & DOCNO & "';"
                TraceLog("CekPluContainer-Q1: " & mda.SelectCommand.CommandText)
                mda.Fill(Rtn)
                Rtn.TableName = "boxplu"
                Box2 = Rtn.Rows.Count
            End If

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "CekPluContainer", conn)
        Finally
            conn.Close()
        End Try

        Return Rtn
    End Function

    ''' <summary>
    ''' untuk cek PLU barang
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CekPluBarang(ByVal Barcode_Plu As String, ByVal ContainerNo As String, ByVal KodeDC As String) As ClsBPB
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mcom As New MySqlCommand("", conn)
        Dim mda As New MySqlDataAdapter("", conn)
        Dim DtPlu As New DataTable
        Dim Result As New ClsBPB
        Dim KodeGudang As String = ""
        Dim docno As String = ""
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            KodeGudang = KodeDC.Split("-")(0)
            docno = KodeDC.Split("-")(1)
            'PROGRAM CEK BARANG BISA PILIH PER DOCNO (NO NPB) - 438/03-23/E/PMO
            mda.SelectCommand.CommandText = "SELECT DISTINCT d.* FROM pos.dcp_boxplu d " &
                                "LEFT JOIN pos.barcode b ON d.PRDCD = b.PLU " &
                                "WHERE KIRIM = '" & KodeGudang & "' " &
                                "AND d.RECID <> '1' " &
                                "AND d.DUS_NO LIKE '" & ContainerNo & "%' " &
                                "AND d.DOCNO = '" & docno & "' " &
                                "AND b.BARCD = '" & Barcode_Plu & "' OR d.PRDCD = '" & Barcode_Plu & "';"
            TraceLog("CekPluBarang-Q1: " & mda.SelectCommand.CommandText)
            mda.Fill(DtPlu)
            DtPlu.TableName = "Plu"

            If DtPlu.Rows.Count > 0 Then

                Barcode_Plu = DtPlu.Rows(0)("PRDCD")
                If IsNothing(DtPlu.Rows(0).Item("DPDID")) Or IsDBNull(DtPlu.Rows(0).Item("DPDID")) _
                   Or DtPlu.Rows(0).Item("DPDID").ToString.Trim = "" Then
                    'PROGRAM CEK BARANG BISA PILIH PER DOCNO (NO NPB) - 438/03-23/E/PMO
                    mcom.CommandText = "UPDATE dcp_boxplu SET DPDID='@',TGL_SCAN =CURDATE() " &
                                       "WHERE Dus_No LIKE '" & ContainerNo & "%' " &
                                       "AND Prdcd = '" & Barcode_Plu & "' " &
                                       "AND KIRIM = '" & KodeGudang & "' " &
                                       "AND DOCNO = '" & docno & "' ;"
                    TraceLog("CekPluBarang-Q2: " & mcom.CommandText)
                    mcom.ExecuteNonQuery()
                End If

                Result.Prdcd = DtPlu.Rows(0)("PRDCD")
                Result.Desc = DtPlu.Rows(0)("NAMA")
                Result.Qty = DtPlu.Rows(0)("qty")
                If Not IsNothing(DtPlu.Rows(0).Item("qtyqc")) And Not IsDBNull(DtPlu.Rows(0).Item("qtyqc")) Then
                    If DtPlu.Rows(0).Item("qtyqc").ToString.Trim <> "" Then
                        If Not IsNothing(DtPlu.Rows(0).Item("DPDID")) And Not IsDBNull(DtPlu.Rows(0).Item("DPDID")) Then
                            If DtPlu.Rows(0).Item("DPDID").ToString.Trim <> "" And DtPlu.Rows(0).Item("DPDID").ToString.Trim <> "@" Then
                                Result.Qty = DtPlu.Rows(0)("qtyqc")
                            End If
                        End If
                    End If
                End If
            Else
                Result.Prdcd = ""
                Result.Desc = "Tidak Ditemukan"
                Result.Qty = ""
            End If

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "CekPluBarang", conn)
        Finally
            conn.Close()
        End Try

        Return Result
    End Function

    Public Function CekPluBarangBKL(ByVal Barcode_Plu As String, ByVal supco As String) As ClsBPBBKL
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mcom As New MySqlCommand("", conn)
        Dim mda As New MySqlDataAdapter("", conn)
        Dim DtPluBKL As New DataTable
        Dim Result As New ClsBPBBKL
        'Dim minor As Integer
        'Dim fraction_pcs As Integer
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            mda.SelectCommand.CommandText = "SELECT DISTINCT d.*, NAMA FROM pos.bpbbkl_wdcp d " &
                                "LEFT JOIN pos.barcode b ON d.PRDCD = b.PLU " &
                                "LEFT JOIN pos.prodmast p ON p.prdcd = b.plu " &
                                "WHERE " &
                                "b.BARCD = '" & Barcode_Plu & "' OR d.PRDCD = '" & Barcode_Plu & "';"
            Console.WriteLine(mda.SelectCommand.CommandText)
            mda.Fill(DtPluBKL)
            'If DtPluBKL.Rows.Count > 0 Then
            '    minor = DtPluBKL.Rows(0)("minor")
            '    fraction_pcs = DtPluBKL.Rows(0)("fraction_pcs")
            'End If

            DtPluBKL.TableName = "Plu"
            Result.BKL = New ClsBKL
            'If (minor < fraction_pcs) And DtPluBKL.Rows.Count > 0 Then
            '    Result.BKL.Prdcd = ""
            '    Result.BKL.Desc = "Min.Or < Fraction"
            '    Result.BKL.Qty = ""
            '    Result.StatusDesc = "3"
            If DtPluBKL.Rows.Count > 0 Then

                'If Not IsNothing(DtPlu.Rows(0).Item("DPDID")) And Not IsDBNull(DtPlu.Rows(0).Item("DPDID")) Then
                '    If DtPlu.Rows(0).Item("DPDID") = "@" Then
                '        Result.Desc = "Barang Sudah Dicek"
                '        Exit Try
                '    End If
                'End If

                Barcode_Plu = DtPluBKL.Rows(0)("PRDCD")
                'If IsNothing(DtPluBKL.Rows(0).Item("FINISHW")) Or IsDBNull(DtPluBKL.Rows(0).Item("FINISHW")) _
                '   Or DtPluBKL.Rows(0).Item("FINISHW").ToString.Trim = "" Then
                'mcom.CommandText = "UPDATE BPBBKL_WDCP SET TGL_BPBW =CURDATE() " &
                '                   "WHERE Prdcd = '" & Barcode_Plu & "' ;"
                'mcom.ExecuteNonQuery()
                'End If

                Result.BKL.Prdcd = DtPluBKL.Rows(0)("PRDCD")
                Result.BKL.Desc = DtPluBKL.Rows(0)("NAMA")
                Result.BKL.Qty = DtPluBKL.Rows(0)("qty")
                Result.BKL.fraction_pcs = DtPluBKL.Rows(0)("fraction_pcs")
                Result.StatusDesc = "2"
                'If Not IsNothing(DtPluBKL.Rows(0).Item("QTY_BPBW")) And Not IsDBNull(DtPluBKL.Rows(0).Item("QTY_BPBW")) Then
                '    If DtPluBKL.Rows(0).Item("QTY_BPBW").ToString.Trim <> "" Then
                '        If Not IsNothing(DtPluBKL.Rows(0).Item("FINISHW")) And Not IsDBNull(DtPluBKL.Rows(0).Item("FINISHW")) Then
                '            If DtPluBKL.Rows(0).Item("FINISHW").ToString.Trim <> "" And DtPluBKL.Rows(0).Item("FINISHW").ToString.Trim <> "@" Then
                '                Result.BKL.Qty = DtPluBKL.Rows(0)("QTY_BPBW")
                '            End If
                '        End If
                '    End If
                'End If
            Else

                Result.BKL.Prdcd = ""
                Result.BKL.Desc = "Tidak Terdaftar"
                Result.BKL.Qty = ""
                Result.StatusDesc = "1"
                Result.BKL.fraction_pcs = ""

            End If

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "CekPluBarangBKL", conn)
        Finally
            conn.Close()
        End Try

        Return Result
    End Function

    Public Function CekPluBarangNPS(ByVal Barcode_Plu As String, ByVal nopo As String) As ClsBPBNPS
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mcom As New MySqlCommand("", conn)
        Dim mda As New MySqlDataAdapter("", conn)
        Dim DtPluNPS As New DataTable
        Dim Result As New ClsBPBNPS
        'Dim minor As Integer
        'Dim fraction_pcs As Integer
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            mda.SelectCommand.CommandText = "SELECT DISTINCT d.*, NAMA FROM pos.bpbnps_wdcp d " &
                                "LEFT JOIN pos.barcode b ON d.PRDCD = b.PLU " &
                                "LEFT JOIN pos.prodmast p ON p.prdcd = b.plu " &
                                "WHERE " &
                                "(b.BARCD = '" & Barcode_Plu & "' OR d.PRDCD = '" & Barcode_Plu & "')
                                AND d.nopo = '" & nopo & "';"
            Console.WriteLine(mda.SelectCommand.CommandText)
            mda.Fill(DtPluNPS)

            DtPluNPS.TableName = "Plu"
            Result.NPS = New CLSNPS

            If DtPluNPS.Rows.Count > 0 Then
                Barcode_Plu = DtPluNPS.Rows(0)("PRDCD")
                Result.NPS.Prdcd = DtPluNPS.Rows(0)("PRDCD")
                Result.NPS.Desc = DtPluNPS.Rows(0)("NAMA")
                Result.NPS.Qty = DtPluNPS.Rows(0)("qty")
                Result.StatusDesc = "2"

            Else

                Result.NPS.Prdcd = ""
                Result.NPS.Desc = "Tidak Terdaftar"
                Result.NPS.Qty = ""
                Result.StatusDesc = "1"
            End If

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "CekPluBarangBKL", conn)
        Finally
            conn.Close()
        End Try

        Return Result
    End Function


    ''' <summary>
    ''' Revisi Qty barang
    ''' </summary>
    ''' <param name="kodeDC"></param>
    ''' <returns>true or false</returns>
    ''' <remarks></remarks>
    Public Function RevisiQtyBarang(ByVal Barcode_Plu As String, ByVal ContainerNo As String, ByVal KodeDC As String,
                                    ByVal QtyRev As Integer) As ClsBPB
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mcom As New MySqlCommand("", conn)
        Dim mda As New MySqlDataAdapter("", conn)
        Dim DtPlu As New DataTable
        Dim Result As New ClsBPB
        Dim DpdID As String = ""
        Dim QtyAwal As Int32
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            mda.SelectCommand.CommandText = "SELECT DISTINCT d.* FROM pos.dcp_boxplu d " &
                               "LEFT JOIN pos.barcode b ON d.PRDCD = b.PLU " &
                               "WHERE KIRIM = '" & KodeDC & "' " &
                               "AND d.RECID <> '1' " &
                               "AND d.DUS_NO LIKE '" & ContainerNo & "%' " &
                               "AND b.BARCD = '" & Barcode_Plu & "' OR d.PRDCD = '" & Barcode_Plu & "';"
            mda.Fill(DtPlu)
            DtPlu.TableName = "Plu"
            If DtPlu.Rows.Count > 0 Then
                QtyAwal = 0
                If Not IsNothing(DtPlu.Rows(0).Item("Qty")) And Not IsDBNull(DtPlu.Rows(0).Item("Qty")) Then
                    If DtPlu.Rows(0).Item("Qty").ToString.Trim <> "" Then
                        QtyAwal = Convert.ToInt32(DtPlu.Rows(0).Item("Qty"))
                    End If
                End If

                If QtyRev < QtyAwal Then
                    DpdID = "-"
                ElseIf QtyRev > QtyAwal Then
                    DpdID = "+"
                Else
                    DpdID = "="
                End If

                Barcode_Plu = DtPlu.Rows(0)("PRDCD")
                mcom.CommandText = "UPDATE Dcp_Boxplu SET QTYQC = '" & QtyRev & "', DPDID = '" & DpdID & "', TGL_SCAN = NOW() " &
                                   "WHERE DUS_NO LIKE '" & ContainerNo & "%' " &
                                   "AND PRDCD = '" & Barcode_Plu & "' " &
                                   "AND KIRIM = '" & KodeDC & "';"
                mcom.ExecuteNonQuery()

                Result.Prdcd = DtPlu.Rows(0)("PRDCD")
                Result.Desc = DtPlu.Rows(0)("NAMA")
                Result.Qty = QtyRev
            Else
                Result.Prdcd = ""
                Result.Desc = "Tidak Ditemukan"
                Result.Qty = ""
            End If

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "RevisiQtyBarang", conn)
        Finally
            conn.Close()
        End Try
        Return Result
    End Function

    Public Function cekQTYBKL(ByVal docno As String, ByVal qtyinput As String, ByVal prdcd As String, ByVal supco As String, ByVal tgl As String, ByVal namauser As String) As ClsBPBBKL
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim tmpDt As New DataTable
        Dim result As New ClsBPBBKL
        Dim mcom As New MySqlCommand("", conn)
        Dim tempQtytotal As Integer
        Dim tgl_exp_convert As Date
        Dim temp_tgl As String
        Dim hasilMod As Integer
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            '30/1/23
            'Perubahan baca parameter dari DOCNO menjadi GROUP DOCNO
            mcom.CommandText = "SELECT DISTINCT d.*, NAMA FROM pos.bpbbkl_wdcp d  "
            mcom.CommandText &= "LEFT JOIN pos.barcode b ON d.PRDCD = b.PLU "
            mcom.CommandText &= "  LEFT JOIN pos.prodmast p ON p.prdcd = b.plu "
            mcom.CommandText &= "  WHERE d.supco = '" & supco & "' "
            If main_groupdocno = False Then
                mcom.CommandText &= "  AND d.DOCNO = '" & docno & "' "

            Else
                mcom.CommandText &= "  AND d.GROUP_DOCNO = '" & docno & "' "

            End If

            mcom.CommandText &= "  AND (b.BARCD ='" & prdcd & "'  OR d.PRDCD = '" & prdcd & "')"
            Dim sDap As New MySqlDataAdapter(mcom)
            sDap.Fill(tmpDt)

            If tmpDt.Rows.Count = 0 Then
                result.StatusQTY = "1"
                result.Feedback = "Tidak ditemukan"
                tracelog_errorBPBBKL(tmpDt.Rows.Item(0)("tgl_pb").ToString, docno, supco, prdcd, qtyinput, tmpDt.Rows.Item(0)("FRACTION_PCS"), tmpDt.Rows.Item(0)("MINOR"), tmpDt.Rows.Item(0)("qty").ToString, tmpDt.Rows.Item(0)("sj_qty").ToString, result.Feedback, namauser)
            Else
                result.BKL = New ClsBKL
                result.BKL.Prdcd = tmpDt.Rows.Item(0)("prdcd").ToString
                result.BKL.Qty = tmpDt.Rows.Item(0)("qty").ToString
                result.BKL.sjQty = tmpDt.Rows.Item(0)("sj_qty").ToString
                result.BKL.Desc = tmpDt.Rows.Item(0)("NAMA").ToString
                result.BKL.fraction_pcs = tmpDt.Rows.Item(0)("FRACTION_PCS")
                tempQtytotal = qtyinput * tmpDt.Rows.Item(0)("FRACTION_PCS")
                hasilMod = tempQtytotal Mod tmpDt.Rows.Item(0)("MINOR")
                'result.BKL.Docno = tmpDt.Rows.Item(0)("docno").ToString
                'result.BKL.Toko = tmpDt.Rows.Item(0)("toko").ToString
                'MsgBox("reoder =" & tmpDt.Rows.Item(0)("FRACTION_PCS") & ", qty total = " & tempQtytotal & ", qty bpb = " & result.BKL.Qty)

                If qtyinput = 0 Then
                    result.StatusQTY = "4"
                    result.Feedback = "QTY 0"
                    tracelog_errorBPBBKL(tmpDt.Rows.Item(0)("tgl_pb").ToString, docno, supco, prdcd, qtyinput, tmpDt.Rows.Item(0)("FRACTION_PCS"), tmpDt.Rows.Item(0)("MINOR"), tmpDt.Rows.Item(0)("qty").ToString, tmpDt.Rows.Item(0)("sj_qty").ToString, result.Feedback, namauser)
                ElseIf tempQtytotal > result.BKL.sjQty Then
                    'mcom.CommandText = "Update bpbbkl_wdcp set qty_bpbw = '" & qtyinput & "' where prdcd = '" & result.BKL.Prdcd & "' AND supco = '" & supco & "'"
                    'mcom.ExecuteNonQuery()
                    result.StatusQTY = "3"
                    result.Feedback = "QTY melebihi"
                    tracelog_errorBPBBKL(tmpDt.Rows.Item(0)("tgl_pb").ToString, docno, supco, prdcd, qtyinput, tmpDt.Rows.Item(0)("FRACTION_PCS"), tmpDt.Rows.Item(0)("MINOR"), tmpDt.Rows.Item(0)("qty").ToString, tmpDt.Rows.Item(0)("sj_qty").ToString, result.Feedback, namauser)
                ElseIf tempQtytotal <= result.BKL.sjQty Then
                    If hasilMod = 0 Then
                        tgl_exp_convert = DateTime.ParseExact(tgl, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)

                        'temp_tgl = tgl_exp_convert.ToString("yyyy-MM-dd")
                        temp_tgl = tgl
                        '30/1/23
                        'Perubahan baca parameter dari DOCNO menjadi GROUP DOCNO
                        If main_groupdocno = False Then
                            mcom.CommandText = "Update bpbbkl_wdcp set tgl_bpbw=NOW(), qty_bpbw = '" & tempQtytotal & "' ,TGL_EPW = '" & temp_tgl & "' where prdcd = '" & result.BKL.Prdcd & "' AND supco = '" & supco & "' AND DOCNO = '" & docno & "' "
                        Else
                            mcom.CommandText = "Update bpbbkl_wdcp set tgl_bpbw=NOW(), qty_bpbw = '" & tempQtytotal & "' ,TGL_EPW = '" & temp_tgl & "' where prdcd = '" & result.BKL.Prdcd & "' AND supco = '" & supco & "' AND GROUP_DOCNO = '" & docno & "' "
                        End If
                        mcom.ExecuteNonQuery()

                        result.StatusQTY = "2"
                        result.Feedback = "Berhasil update QTY"
                        result.BKL.totalqty = tempQtytotal
                        tracelog_errorBPBBKL(tmpDt.Rows.Item(0)("tgl_pb").ToString, docno, supco, prdcd, qtyinput, tmpDt.Rows.Item(0)("FRACTION_PCS"), tmpDt.Rows.Item(0)("MINOR"), tmpDt.Rows.Item(0)("qty").ToString, tmpDt.Rows.Item(0)("sj_qty").ToString, result.Feedback, namauser)

                    Else
                        result.StatusQTY = "5"
                        result.Feedback = "QTY tdk sesuai minor"
                        tracelog_errorBPBBKL(tmpDt.Rows.Item(0)("tgl_pb").ToString, docno, supco, prdcd, qtyinput, tmpDt.Rows.Item(0)("FRACTION_PCS"), tmpDt.Rows.Item(0)("MINOR"), tmpDt.Rows.Item(0)("qty").ToString, tmpDt.Rows.Item(0)("sj_qty").ToString, result.Feedback, namauser)
                    End If

                End If
            End If

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "cekQTYBKL", conn)
        Finally
            conn.Close()
        End Try

        Return result
    End Function

    Public Function cekQTYNPS(ByVal qtyinput As String, ByVal prdcd As String, ByVal NOPO As String, ByVal tgl As String, ByVal namauser As String) As ClsBPBNPS
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim tmpDt As New DataTable
        Dim result As New ClsBPBNPS
        Dim mcom As New MySqlCommand("", conn)
        Dim tempQtytotal As Integer
        Dim tgl_exp_convert As Date
        Dim temp_tgl As String
        Dim hasilMod As Integer
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            mcom.CommandText = "SELECT DISTINCT d.*, NAMA FROM pos.BPBNPS_WDCP d  "
            mcom.CommandText &= "LEFT JOIN pos.barcode b ON d.PRDCD = b.PLU "
            mcom.CommandText &= "  LEFT JOIN pos.prodmast p ON p.prdcd = b.plu "
            mcom.CommandText &= "  WHERE d.NOPO = '" & NOPO & "' "
            mcom.CommandText &= "  AND (b.BARCD ='" & prdcd & "'  OR d.PRDCD = '" & prdcd & "')"
            IDM.Fungsi.TraceLog("BPB_NPS CekQTYNPS : " & mcom.CommandText)
            Dim sDap As New MySqlDataAdapter(mcom)
            sDap.Fill(tmpDt)

            If tmpDt.Rows.Count = 0 Then
                result.StatusQTY = "1"
                result.Feedback = "Tidak ditemukan"
                'IDM.Fungsi.TraceLog("TRACELOG BPBNPS : " & tmpDt.Rows.Item(0)("tgl_po").ToString & "-" & NOPO & "-" & prdcd &
                '                    "-" & qtyinput & "-" & tmpDt.Rows.Item(0)("FRACTION_PCS") & "-" & tmpDt.Rows.Item(0)("MINOR") & "-" &
                '                    tmpDt.Rows.Item(0)("qty").ToString & "-" & result.Feedback & "-" & namauser)
            Else
                result.NPS = New CLSNPS
                result.NPS.Prdcd = tmpDt.Rows.Item(0)("prdcd").ToString
                result.NPS.Qty = tmpDt.Rows.Item(0)("qty").ToString
                'result.NPS.sjQty = tmpDt.Rows.Item(0)("sj_qty").ToString
                result.NPS.Desc = tmpDt.Rows.Item(0)("NAMA").ToString

                tempQtytotal = qtyinput * tmpDt.Rows.Item(0)("FRACTION_PCS")

                hasilMod = tempQtytotal Mod tmpDt.Rows.Item(0)("MINOR")

                'result.BKL.Docno = tmpDt.Rows.Item(0)("docno").ToString
                'result.BKL.Toko = tmpDt.Rows.Item(0)("toko").ToString
                'MsgBox("reoder =" & tmpDt.Rows.Item(0)("FRACTION_PCS") & ", qty total = " & tempQtytotal & ", qty bpb = " & result.BKL.Qty)
                '29/11/2021 - permintaan pak Yuyun, input qty 0 bisa diterima
                'If qtyinput = 0 Then
                '    result.StatusQTY = "4"
                '    result.Feedback = "QTY 0"

                'IDM.Fungsi.TraceLog("TRACELOG BPBNPS : " & tmpDt.Rows.Item(0)("tgl_po").ToString & "-" & NOPO & "-" & prdcd &
                '                "-" & qtyinput & "-" & tmpDt.Rows.Item(0)("FRACTION_PCS") & "-" & tmpDt.Rows.Item(0)("MINOR") & "-" &
                '                tmpDt.Rows.Item(0)("qty").ToString & "-" & result.Feedback & "-" & namauser)

                If tempQtytotal > result.NPS.Qty Then

                    result.StatusQTY = "3"
                    result.Feedback = "QTY melebihi"
                    result.NPS.sjQty = tempQtytotal
                    'IDM.Fungsi.TraceLog("TRACELOG BPBNPS : " & tmpDt.Rows.Item(0)("tgl_po").ToString & "-" & NOPO & "-" & prdcd &
                    '                "-" & qtyinput & "-" & tmpDt.Rows.Item(0)("FRACTION_PCS") & "-" & tmpDt.Rows.Item(0)("MINOR") & "-" &
                    '                tmpDt.Rows.Item(0)("qty").ToString & "-" & result.Feedback & "-" & namauser)

                ElseIf tempQtytotal <= result.NPS.Qty Then

                    If hasilMod = 0 Then
                        tgl_exp_convert = DateTime.ParseExact(tgl, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)

                        temp_tgl = tgl

                        mcom.CommandText = "Update bpbnps_wdcp set tgl_bpbw=NOW(), 
                                            qty_bpbw = '" & tempQtytotal & "' ,
                                            TGL_EPW = '" & temp_tgl & "' 
                                            where prdcd = '" & result.NPS.Prdcd & "' AND nopo = '" & NOPO & "'"
                        mcom.ExecuteNonQuery()

                        result.StatusQTY = "2"

                        result.Feedback = "Berhasil update QTY"

                        result.NPS.totalqty = tempQtytotal

                        'IDM.Fungsi.TraceLog("TRACELOG BPBNPS : " & tmpDt.Rows.Item(0)("tgl_po").ToString & "-" & NOPO & "-" & prdcd &
                        '         "-" & qtyinput & "-" & tmpDt.Rows.Item(0)("FRACTION_PCS") & "-" & tmpDt.Rows.Item(0)("MINOR") & "-" &
                        '         tmpDt.Rows.Item(0)("qty").ToString & "-" & result.Feedback & "-" & namauser)

                    Else
                        result.StatusQTY = "5"
                        result.Feedback = "QTY tdk sesuai minor"
                        'IDM.Fungsi.TraceLog("TRACELOG BPBNPS : " & tmpDt.Rows.Item(0)("tgl_po").ToString & "-" & NOPO & "-" & prdcd &
                        '            "-" & qtyinput & "-" & tmpDt.Rows.Item(0)("FRACTION_PCS") & "-" & tmpDt.Rows.Item(0)("MINOR") & "-" &
                        '            tmpDt.Rows.Item(0)("qty").ToString & "-" & result.Feedback & "-" & namauser)
                    End If

                End If

            End If

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "cekQTYBKL", conn)
        Finally
            conn.Close()
        End Try

        Return result
    End Function

    Public Function inputTGL_EPW(ByVal Barcode_Plu As String, ByVal supplier As String, ByVal QTY As String, ByVal tgl As String) As ClsBPBBKL
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mcom As New MySqlCommand("", conn)
        Dim mda As New MySqlDataAdapter("", conn)
        Dim DtPlu As New DataTable
        Dim Result As New ClsBPBBKL

        Dim FINISHW As String = ""
        Dim tgl_exp As String
        Dim tgl_exp_convert As Date
        Dim temp_tgl As String
        Dim hari As Integer
        Dim hari2 As Integer
        Dim bulan As Integer
        Dim bulan2 As Integer
        Dim tahun As Integer
        Dim tahun2 As Integer
        Dim tanggalBPB As Date
        Dim compre As Integer
        Dim tanggalLayak As Date
        Dim stat As Boolean

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            tgl_exp = tgl
            mcom.CommandText = "SELECT DISTINCT d.*, NAMA FROM pos.bpbbkl_Wdcp d " &
                                   "LEFT JOIN pos.barcode b ON d.PRDCD = b.PLU " &
                                   "LEFT JOIN pos.prodmast p ON p.prdcd = b.plu " &
                                   "WHERE d.SUPCO = '" & supplier & "' " &
                                   "AND b.BARCD = '" & Barcode_Plu & "' OR d.PRDCD = '" & Barcode_Plu & "';"
            mda.SelectCommand = mcom
            mda.Fill(DtPlu)
            Result.BKL = New ClsBKL

            Result.BKL.Prdcd = DtPlu.Rows(0)("prdcd").ToString
            Result.BKL.Desc = DtPlu.Rows(0)("NAMA").ToString
            Result.BKL.fraction_pcs = DtPlu.Rows(0)("fraction_pcs")

            If tgl_exp.Length = "6" Or tgl_exp = "00" Then

                If tgl_exp = "00" Then
                    tgl_exp = DateTime.Now.ToString("ddMMyy")
                    stat = True
                End If
                hari = Integer.Parse(tgl_exp.Substring(0, 1))

                hari2 = Integer.Parse(tgl_exp.Substring(1, 1))
                bulan = Integer.Parse(tgl_exp.Substring(2, 1))
                bulan2 = Integer.Parse(tgl_exp.Substring(3, 1))
                tahun = Integer.Parse(tgl_exp.Substring(4, 1))
                tahun2 = Integer.Parse(tgl_exp.Substring(5, 1))

                tanggalBPB = DtPlu.Rows(0)("TGL_PB")

                tanggalLayak = DtPlu.Rows(0)("BATAS_KELAYAKAN")
                'tempBPB = DateTime.ParseExact(tgl_exp, "ddMMyy", System.Globalization.CultureInfo.InvariantCulture)
                'compre = DateTime.Compare(tempBPB, tanggalBPB)

                'If compre < 0 Then
                '    Result.StatusExp = "3"
                '    Result.Feedback = "Tgl EXP < Tgl BPB"
                'Else

                Try
                    'Dim cbtgl As Date
                    'cbtgl = New Date("2019", "09", "10")

                    tgl_exp_convert = New Date("20" & tahun & tahun2, bulan & bulan2, hari & hari2)

                    'tgl_exp_convert = DateTime.ParseExact(tgl_exp, "ddMMyy", System.Globalization.CultureInfo.InvariantCulture)
                    compre = Date.Compare(tgl_exp_convert, tanggalLayak)

                    If stat = True Then
                        tanggalLayak = tanggalLayak.AddDays(1)

                        compre = Date.Compare(tgl_exp_convert, tanggalLayak)

                        If compre <> 0 Then
                            tanggalLayak = tanggalLayak.AddDays(-1)

                            compre = Date.Compare(tgl_exp_convert, tanggalLayak)

                            If compre <> 0 Then

                                Result.StatusExp = "3"
                                Result.Feedback = "TglEXP ditolak"
                            Else

                                temp_tgl = tgl_exp_convert.ToString("yyyy-MM-dd")

                                'mcom.CommandText = "Update bpbbkl_wdcp set tgl_bpbw=NOW(), qty_bpbw = '" & QTY & "' ,TGL_EPW = '" & temp_tgl & "' where prdcd = '" & Result.BKL.Prdcd & "' AND supco = '" & supplier & "'"
                                'mcom.ExecuteNonQuery()

                                Result.BKL.TglEXP = temp_tgl
                                Result.StatusExp = "2"
                                Result.Feedback = "Berhasil Update"
                            End If

                        Else
                            temp_tgl = tgl_exp_convert.ToString("yyyy-MM-dd")

                            'mcom.CommandText = "Update bpbbkl_wdcp set tgl_bpbw=NOW(), qty_bpbw = '" & QTY & "' ,TGL_EPW = '" & temp_tgl & "' where prdcd = '" & Result.BKL.Prdcd & "' AND supco = '" & supplier & "'"
                            'mcom.ExecuteNonQuery()

                            Result.BKL.TglEXP = temp_tgl
                            Result.StatusExp = "2"
                            Result.Feedback = "Berhasil Update"

                        End If
                    Else

                        If compre < 0 Then
                            Result.StatusExp = "3"
                            Result.Feedback = "TglEXP<BatasLayak"

                        Else

                            temp_tgl = tgl_exp_convert.ToString("yyyy-MM-dd")

                            'mcom.CommandText = "Update bpbbkl_wdcp set tgl_bpbw=NOW(), qty_bpbw = '" & QTY & "' ,TGL_EPW = '" & temp_tgl & "' where prdcd = '" & Result.BKL.Prdcd & "' AND supco = '" & supplier & "'"
                            'mcom.ExecuteNonQuery()

                            Result.BKL.TglEXP = temp_tgl
                            Result.StatusExp = "2"
                            Result.Feedback = "Berhasil Update"

                        End If
                    End If

                Catch ex As Exception
                    Result.StatusExp = "1"
                    Result.Feedback = "Format Salah"
                End Try
            Else
                Result.StatusExp = "1"
                Result.Feedback = "Format Salah"
            End If
        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "inputTGL_EPW", conn)
        Finally
            conn.Close()
        End Try
        Return Result
    End Function

    Public Function inputTGL_EPW_NPS(ByVal Barcode_Plu As String, ByVal nopo As String, ByVal QTY As String, ByVal tgl As String) As ClsBPBNPS
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mcom As New MySqlCommand("", conn)
        Dim mda As New MySqlDataAdapter("", conn)
        Dim DtPlu As New DataTable
        Dim Result As New ClsBPBNPS

        Dim FINISHW As String = ""
        Dim tgl_exp As String
        Dim tgl_exp_convert As Date
        Dim temp_tgl As String
        Dim hari As Integer
        Dim hari2 As Integer
        Dim bulan As Integer
        Dim bulan2 As Integer
        Dim tahun As Integer
        Dim tahun2 As Integer
        Dim compre As Integer
        Dim tanggalLayak As Date
        Dim stat As Boolean

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            tgl_exp = tgl
            mcom.CommandText = "SELECT DISTINCT d.*, NAMA FROM pos.bpbnps_wdcp d " &
                                   "LEFT JOIN pos.barcode b ON d.PRDCD = b.PLU " &
                                   "LEFT JOIN pos.prodmast p ON p.prdcd = b.plu " &
                                   "WHERE d.NOPO = '" & nopo & "' " &
                                   "AND b.BARCD = '" & Barcode_Plu & "' OR d.PRDCD = '" & Barcode_Plu & "';"
            mda.SelectCommand = mcom
            mda.Fill(DtPlu)
            Result.NPS = New CLSNPS

            Result.NPS.Prdcd = DtPlu.Rows(0)("prdcd").ToString
            Result.NPS.Desc = DtPlu.Rows(0)("NAMA").ToString

            If tgl_exp.Length = "6" Or tgl_exp = "00" Then

                If tgl_exp = "00" Then
                    tgl_exp = DateTime.Now.ToString("ddMMyy")
                    stat = True
                End If
                hari = Integer.Parse(tgl_exp.Substring(0, 1))

                hari2 = Integer.Parse(tgl_exp.Substring(1, 1))
                bulan = Integer.Parse(tgl_exp.Substring(2, 1))
                bulan2 = Integer.Parse(tgl_exp.Substring(3, 1))
                tahun = Integer.Parse(tgl_exp.Substring(4, 1))
                tahun2 = Integer.Parse(tgl_exp.Substring(5, 1))

                'tanggalBPB = DtPlu.Rows(0)("TGL_PB")

                tanggalLayak = DtPlu.Rows(0)("BATAS_KELAYAKAN")
                'tempBPB = DateTime.ParseExact(tgl_exp, "ddMMyy", System.Globalization.CultureInfo.InvariantCulture)
                'compre = DateTime.Compare(tempBPB, tanggalBPB)

                'If compre < 0 Then
                '    Result.StatusExp = "3"
                '    Result.Feedback = "Tgl EXP < Tgl BPB"
                'Else

                Try
                    'Dim cbtgl As Date
                    'cbtgl = New Date("2019", "09", "10")

                    tgl_exp_convert = New Date("20" & tahun & tahun2, bulan & bulan2, hari & hari2)

                    'tgl_exp_convert = DateTime.ParseExact(tgl_exp, "ddMMyy", System.Globalization.CultureInfo.InvariantCulture)
                    compre = Date.Compare(tgl_exp_convert, tanggalLayak)

                    If stat = True Then
                        tanggalLayak = tanggalLayak.AddDays(1)

                        compre = Date.Compare(tgl_exp_convert, tanggalLayak)

                        If compre <> 0 Then
                            tanggalLayak = tanggalLayak.AddDays(-1)

                            compre = Date.Compare(tgl_exp_convert, tanggalLayak)

                            If compre <> 0 Then

                                Result.StatusExp = "3"
                                Result.Feedback = "TglEXP ditolak"
                            Else

                                temp_tgl = tgl_exp_convert.ToString("yyyy-MM-dd")

                                'mcom.CommandText = "Update bpbbkl_wdcp set tgl_bpbw=NOW(), qty_bpbw = '" & QTY & "' ,TGL_EPW = '" & temp_tgl & "' where prdcd = '" & Result.BKL.Prdcd & "' AND supco = '" & supplier & "'"
                                'mcom.ExecuteNonQuery()

                                Result.NPS.TglEXP = temp_tgl
                                Result.StatusExp = "2"
                                Result.Feedback = "Berhasil Update"
                            End If

                        Else
                            temp_tgl = tgl_exp_convert.ToString("yyyy-MM-dd")

                            'mcom.CommandText = "Update bpbbkl_wdcp set tgl_bpbw=NOW(), qty_bpbw = '" & QTY & "' ,TGL_EPW = '" & temp_tgl & "' where prdcd = '" & Result.BKL.Prdcd & "' AND supco = '" & supplier & "'"
                            'mcom.ExecuteNonQuery()

                            Result.NPS.TglEXP = temp_tgl
                            Result.StatusExp = "2"
                            Result.Feedback = "Berhasil Update"
                        End If
                    Else

                        If compre < 0 Then
                            Result.StatusExp = "3"
                            Result.Feedback = "TglEXP<BatasLayak"

                        Else

                            temp_tgl = tgl_exp_convert.ToString("yyyy-MM-dd")

                            'mcom.CommandText = "Update bpbbkl_wdcp set tgl_bpbw=NOW(), qty_bpbw = '" & QTY & "' ,TGL_EPW = '" & temp_tgl & "' where prdcd = '" & Result.BKL.Prdcd & "' AND supco = '" & supplier & "'"
                            'mcom.ExecuteNonQuery()

                            Result.NPS.TglEXP = temp_tgl
                            Result.StatusExp = "2"
                            Result.Feedback = "Berhasil Update"
                        End If
                    End If

                Catch ex As Exception
                    Result.StatusExp = "1"
                    Result.Feedback = "Format Salah"
                End Try
            Else
                Result.StatusExp = "1"
                Result.Feedback = "Format Salah"
            End If
        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "inputTGL_EPW_NPS", conn)
        Finally
            conn.Close()
        End Try
        Return Result
    End Function
    ''' <summary>
    ''' Load kode gudang yang belum proses cek BPB
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LoadKodeGudang() As DataTable
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mda As New MySqlDataAdapter("", conn)
        Dim Rtn As New DataTable

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            'Hanya mengambil docno yang belum di scan #Kukuh270723
            mda.SelectCommand.CommandText = "SELECT DISTINCT b.kirim AS KODEGUDANG, CONCAT(b.kirim,'-',d.type_dc) AS INFO " &
                                            "FROM dcp_boxplu b " &
                                            "JOIN dcmast d " &
                                            "ON b.kirim=d.kode_dc WHERE b.RECID <> '1';"
            TraceLog("LoadKodeGudang: " & mda.SelectCommand.CommandText)
            mda.Fill(Rtn)
        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "LoadKodeGudang", conn)
        Finally
            conn.Close()
        End Try

        Return Rtn
    End Function
    Public Function LoadKodeGudang_Docno() As DataTable
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mda As New MySqlDataAdapter("", conn)
        Dim Rtn As New DataTable

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            'PROGRAM CEK BARANG BISA PILIH PER DOCNO (NO NPB) - 438/03-23/E/PMO
            'Perubahan sebelumnya pilihan KODEGUDANG menjadi KODEGUDANG - DOCNO
            mda.SelectCommand.CommandText = "SELECT DISTINCT CONCAT(b.kirim,'-',CONVERT(b.DOCNO,CHAR)) AS KODEGUDANG, CONCAT(b.kirim,'-',d.type_dc,'-',CONVERT(b.DOCNO,CHAR)) AS INFO " &
                                            "FROM dcp_boxplu b " &
                                            "JOIN dcmast d " &
                                            "ON b.kirim=d.kode_dc WHERE b.RECID <> '1';"
            TraceLog("LoadKodeGudang_Docno: " & mda.SelectCommand.CommandText)
            mda.Fill(Rtn)
        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "LoadKodeGudang", conn)
        Finally
            conn.Close()
        End Try

        Return Rtn
    End Function

    Public Function LoadKodeSupplier(ByVal supco As String, ByVal docno As String) As DataTable
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mda As New MySqlDataAdapter("", conn)
        Dim Rtn As New DataTable

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            '30/1/23
            'Perubahan baca parameter dari DOCNO menjadi GROUP DOCNO
            If supco <> "" And docno <> "" Then
                If main_groupdocno = False Then
                    mda.SelectCommand.CommandText = "SELECT DISTINCT supco AS KODESUPPLIER, CONCAT(SUPCO, ' - ',docno) AS INFO " &
                                                "FROM BPBBKL_WDCP where supco = '" & supco & "' AND docno = '" & docno & "'"
                Else
                    mda.SelectCommand.CommandText = "SELECT DISTINCT supco AS KODESUPPLIER, CONCAT(SUPCO, ' - ',docno) AS INFO " &
                                                    "FROM BPBBKL_WDCP where supco = '" & supco & "' AND GROUP_DOCNO = '" & docno & "'"
                End If

                TraceLog("LoadKodeSupplier: " & mda.SelectCommand.CommandText)

                mda.Fill(Rtn)
            Else
                mda.SelectCommand.CommandText = "SELECT DISTINCT supco AS KODESUPPLIER, SUPCO AS INFO " &
                                                "FROM supmast "
                TraceLog("LoadKodeSupplier: " & mda.SelectCommand.CommandText)
                mda.Fill(Rtn)
            End If

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "LoadKodeSupplierBKL", conn)
        Finally
            conn.Close()
        End Try

        Return Rtn
    End Function

    Public Function LoadNomorPO(ByVal noPO As String) As DataTable
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mda As New MySqlDataAdapter("", conn)
        Dim Rtn As New DataTable

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            Console.WriteLine(noPO)

            mda.SelectCommand.CommandText = "SELECT DISTINCT NOPO ,NOPO AS INFO " &
                                                "FROM BPBNPS_WDCP where NOPO = '" & noPO & "'"
            TraceLog("LoadNomorPO: " & mda.SelectCommand.CommandText)

            mda.Fill(Rtn)

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "LoadKodeSupplierBKL", conn)
        Finally
            conn.Close()
        End Try

        Return Rtn
    End Function

    ''' <summary>
    ''' Delete Barang yang NPB nya lbh kecil dari NPB terakir
    ''' </summary>
    ''' <param name="kodeDC"></param>
    ''' <returns>true or false</returns>
    ''' <remarks></remarks>
    Public Function DeleteDcpBoxPlu(ByVal kodeDC As String) As Boolean
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mcom As New MySqlCommand("", conn)
        Dim tmpDateMax As Date
        Dim Rtn As Boolean
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            mcom.CommandText = "SELECT DATE(MAX(tanggal)) AS tanggal FROM dcp_boxplu WHERE kirim='" & kodeDC & "';"
            tmpDateMax = mcom.ExecuteScalar

            mcom.CommandText = "DELETE FROM dcp_boxplu " &
                               "WHERE DATE(tanggal) < '" & Format(tmpDateMax, "yyyy-MM-dd") & "' " &
                               "AND kirim='" & kodeDC & "';"
            mcom.ExecuteNonQuery()
            Rtn = True
        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "DeleteDcpBoxPlu", conn)
            Rtn = False
        Finally
            conn.Close()
        End Try
        Return Rtn
    End Function

    ''' <summary>
    ''' Load DataTable proses cek BPB
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetTableCekPB(ByVal KodeDC As String, Optional ByVal docno As String = "") As DataTable
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mda As New MySqlDataAdapter("", conn)
        Dim Rtn As New DataTable

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            mda.SelectCommand.CommandText = "SELECT DUS_NO, Prdcd, Nama, Qty FROM dcp_boxplu " &
                                "WHERE Recid <> '1' " &
                                "AND (DPDID IS NULL OR DPDID = '') "
            If KodeDC <> "" Then
                mda.SelectCommand.CommandText &= "AND KIRIM = '" & KodeDC & "' "

            End If
            'PROGRAM CEK BARANG BISA PILIH PER DOCNO (NO NPB) - 438/03-23/E/PMO
            If docno <> "" Then
                mda.SelectCommand.CommandText &= " AND DOCNO = '" & docno & "' "
            End If

            TraceLog("GetTableCekPB: " & mda.SelectCommand.CommandText)

            mda.Fill(Rtn)
        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetTableCekPB", conn)
        Finally
            conn.Close()
        End Try

        Return Rtn
    End Function

    Public Function CekSplitCtn(ByRef Msg As String, ByVal kodeCabang As String) As DataTable
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mda As New MySqlDataAdapter("", conn)
        Dim mcom As New MySqlCommand("", conn)
        Dim Rtn As New DataTable

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            mcom.CommandText = "SELECT count(*) FROM DCP_Boxplu " &
                               "WHERE Recid <> '1' " &
                               "AND (DPDID = '' OR DPDID is null) " &
                               "AND KIRIM = '" & kodeCabang & "'"
            If mcom.ExecuteScalar = 0 Then
                Msg = "Data Tidak Tersedia"
            Else
                mda.SelectCommand.CommandText = "SELECT Dus_No, Prdcd, Nama, Qty, Qtyqc FROM DCP_Boxplu " &
                                   "WHERE (DPDID = '' OR DPDID is null) " &
                                   "AND KIRIM = '" & kodeCabang & "'"
                mda.Fill(Rtn)
            End If
        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "CekSplitCtn", conn)
            Msg = "Data Tidak Tersedia"
        Finally
            conn.Close()
        End Try

        Return Rtn
    End Function

    Public Function FindPLU(ByRef PLU As String, ByVal kodeCabang As String) As DataTable
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mda As New MySqlDataAdapter("", conn)
        Dim Rtn As New DataTable

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            mda.SelectCommand.CommandText = "SELECT prdcd, nama, dus_no FROM DCP_Boxplu " &
                               "WHERE prdcd = '" & PLU & "' " &
                               "AND Recid <> 1 " &
                               "AND KIRIM = '" & kodeCabang & "';"
            TraceLog("FindPLU: " & mda.SelectCommand.CommandText)
            mda.Fill(Rtn)

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "FindPLU", conn)
        Finally
            conn.Close()
        End Try

        Return Rtn
    End Function

    Public Function Bersihin_DCP_Boxplu() As Boolean
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", conn)
        Dim NIC As String = ""
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            Mcom.CommandText = "SHOW TABLES LIKE 'DCP_Boxplu'"
            If Mcom.ExecuteScalar & "" <> "" Then
                Mcom.CommandText = "Delete From DCP_Boxplu Where date(`update`) < DATE_ADD(curdate(), INTERVAL -20 DAY)"
                Mcom.ExecuteNonQuery()
            End If
        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "Bersihin_DCP_Boxplu", conn)
        Finally
            conn.Close()
        End Try
        Return True
    End Function

    Public Function Bersihin_BPBBKL_WDCP() As Boolean
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", conn)
        Dim NIC As String = ""
        'Dim Toko24Jam As Boolean
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            Mcom.CommandText = "SHOW TABLES LIKE 'BPBBKL_WDCP'"
            If Mcom.ExecuteScalar & "" <> "" Then
                Mcom.CommandText = "Delete From BPBBKL_WDCP Where date(`TGL_PB`) < DATE_ADD(curdate(), INTERVAL -30 DAY)"
                Mcom.ExecuteNonQuery()
            End If
            Mcom.CommandText = "SHOW TABLES LIKE 'bpbbkl_wdcp_errorlog'"
            If Mcom.ExecuteScalar & "" <> "" Then
                Mcom.CommandText = "Delete From bpbbkl_wdcp_errorlog Where tglscan < DATE_ADD(curdate(), INTERVAL -10 DAY)"
                Mcom.ExecuteNonQuery()
            End If

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "Bersihin_BPBBKL_WDCP", conn)
        Finally
            conn.Close()
        End Try
        Return True
    End Function
    Public Function GetTableCekPBBKL(ByVal KodeDC As String, ByVal docno As String) As DataTable
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mda As New MySqlDataAdapter("", conn)
        Dim Rtn As New DataTable

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            mda.SelectCommand.CommandText = "SELECT b.Prdcd as PRDCD, Nama FROM BPBBKL_wDCP b " &
                                            "LEFT JOIN prodmast p ON b.prdcd = p.prdcd WHERE FINISHW IS NULL AND QTY_BPBW IS NULL "
            If KodeDC <> "" Then
                mda.SelectCommand.CommandText &= "AND B.SUPCO = '" & KodeDC & "' "
            End If
            '30/1/23
            'Perubahan baca dari docno menjadi GROUP_DOCNO
            If docno <> "" Then
                If main_groupdocno = False Then
                    mda.SelectCommand.CommandText &= "AND B.DOCNO = '" & docno & "' "
                Else
                    mda.SelectCommand.CommandText &= "AND B.GROUP_DOCNO = '" & docno & "' "
                End If
            End If

            'mda.SelectCommand.CommandText = "SELECT  DUS_NO, Prdcd, Nama, Qty FROM dcp_boxplu " & _
            '                    "WHERE Recid <> '1' " & _
            '                    "AND (DPDID IS NULL OR DPDID = '') " & _
            '                    "AND Tgl_Scan = CURDATE();"
            TraceLog("GetTableCekPBBKL: " & mda.SelectCommand.CommandText)
            mda.Fill(Rtn)

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetTableCekPBbkl", conn)
        Finally
            conn.Close()
        End Try

        Return Rtn
    End Function

    Public Function GetTableCekPBNPS(ByVal nopo As String) As DataTable
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mda As New MySqlDataAdapter("", conn)
        Dim Rtn As New DataTable

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            mda.SelectCommand.CommandText = "SELECT b.Prdcd as PRDCD, Nama FROM BPBNPS_WDCP b " &
                                            "LEFT JOIN prodmast p ON b.prdcd = p.prdcd WHERE FINISHW IS NULL OR TGL_BPBW IS NULL "
            mda.SelectCommand.CommandText &= "AND B.NOPO = '" & nopo & "' "
            Console.WriteLine(mda.SelectCommand.CommandText)
            mda.Fill(Rtn)

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetTableCekPBbkl", conn)
        Finally
            conn.Close()
        End Try

        Return Rtn
    End Function


    Public Function GetDataReportBPB(ByVal KodeGudang As String, Optional ByVal docno As String = "") As DataTable
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mda As New MySqlDataAdapter("", conn)
        Dim Rtn As New DataTable

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            If KodeGudang <> "All" Then
                'PROGRAM CEK BARANG BISA PILIH PER DOCNO (NO NPB) - 438/03-23/E/PMO

                mda.SelectCommand.CommandText = "SELECT Dus_no,kirim, CAST(Docno AS CHAR) AS Docno, Prdcd, Nama, Qty, Qtyqc, DPDID, '' AS Ket " &
                                "FROM DCP_Boxplu " &
                                "WHERE (DPDID NOT IN ('@','=') " &
                                "OR DPDID IS NULL) " &
                                "AND Kirim='" & KodeGudang & "' AND DOCNO = '" & docno & "';"
            Else
                mda.SelectCommand.CommandText = "SELECT Dus_no,kirim, CAST(Docno AS CHAR) AS Docno, Prdcd, Nama, Qty, Qtyqc, DPDID, '' AS Ket " &
                               "FROM DCP_Boxplu " &
                               "WHERE (DPDID NOT IN ('@','=') " &
                               "OR DPDID IS NULL);"
            End If

            TraceLog("GetDataReportBPB: " & mda.SelectCommand.CommandText)
            mda.Fill(Rtn)

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetDataReportBPB", conn)
        Finally
            conn.Close()
        End Try

        Return Rtn
    End Function

    Public Function GetDataReportDeviasi(ByVal KodeGudang As String, Optional ByVal docno As String = "") As DataTable
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim mda As New MySqlDataAdapter("", conn)
        Dim Rtn As New DataTable

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            '438/03-23/E/PMO
            'mda.SelectCommand.CommandText = "SELECT Docno, Prdcd, Nama, Dus_No, Qty, Qtyqc, '' AS deviasi " &
            '                   "FROM DCP_Boxplu WHERE Tgl_Scan = NOW() " &
            '                   "AND qtyqc <> qty AND (DPDID <> '') OR (DPDID IS NULL);"

            If KodeGudang <> "All" Then
                '438/03-23/E/PMO

                mda.SelectCommand.CommandText = "Select Docno, Prdcd, Nama, Dus_No, Qty, Qtyqc, '' as deviasi,kirim " &
                                                "From DCP_Boxplu " &
                                                "Where Tgl_Scan = CURDATE() " &
                                                "AND KIRIM='" & KodeGudang & "' AND DOCNO  = '" & docno & "' " &
                                                "And qtyqc <> qty " &
                                                "And (DPDID <> '' " &
                                                "OR DPDID is null);"
            Else
                mda.SelectCommand.CommandText = "Select Prdcd, Nama, Dus_No, Qty, Qtyqc, '' as deviasi,kirim " &
                                                "From DCP_Boxplu " &
                                                "Where Tgl_Scan = now() " &
                                                "And qtyqc <> qty " &
                                                "And (DPDID <> '') " &
                                                "OR (DPDID is null);"
            End If

            TraceLog("GetDataReportDeviasi: " & mda.SelectCommand.CommandText)
            mda.Fill(Rtn)

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetDataReportDeviasi", conn)
        Finally
            conn.Close()
        End Try

        Return Rtn
    End Function

    Public Function tracelog_errorBPBBKL(ByVal tglpb As String, ByVal docno As String, ByVal supco As String, ByVal prdcd As String, ByVal qtyinput As String, ByVal frc_pcs As String, ByVal minor As String, ByVal qtypb As String, ByVal sj_qty As String, ByVal keterangan As String, ByVal user As String) As Boolean
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", conn)
        Dim tglscan As String = ""
        'Dim Toko24Jam As Boolean
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            '30/1/23
            'Perubahan baca parameter dari DOCNO menjadi GROUP DOCNO


            Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`bpbbkl_wdcp_errorlog` ("
            Mcom.CommandText &= "                `IDTRACELOG` BIGINT(20) NOT NULL AUTO_INCREMENT, "
            Mcom.CommandText &= "                `TGLSCAN` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP, "
            Mcom.CommandText &= "                `AppName` VARCHAR(50), "
            Mcom.CommandText &= "                `user` VARCHAR(99), "

            Mcom.CommandText &= "                `DOCNO` VARCHAR(150), "
            Mcom.CommandText &= "                `GROUP_DOCNO` VARCHAR(150), "
            Mcom.CommandText &= "                 `supco` VARCHAR(10), "
            Mcom.CommandText &= "                   `prdcd` VARCHAR(10), "
            Mcom.CommandText &= "                `tglpb` VARCHAR(50), "

            Mcom.CommandText &= "                `qty_input` VARCHAR(10),  "
            Mcom.CommandText &= "                `fraction_pcs` VARCHAR(10), "
            Mcom.CommandText &= "                `minor` VARCHAR(10), "
            Mcom.CommandText &= "                 `qty` VARCHAR(10), "
            Mcom.CommandText &= "                  `sj_qty` VARCHAR(10), "

            Mcom.CommandText &= "                `keterangan` VARCHAR(50), "
            Mcom.CommandText &= "                PRIMARY KEY (`IDTRACELOG`) ) "
            Mcom.CommandText &= "                ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci;"
            Mcom.ExecuteNonQuery()

            Mcom.CommandText = "INSERT INTO `pos`.`bpbbkl_wdcp_errorlog` "
            Mcom.CommandText &= "  (`AppName`, `user`, `GROUP_DOCNO`, `supco`, `prdcd`, `tglpb`, `qty_input`, `fraction_pcs`,`minor`, `qty`, `sj_qty`, `keterangan`) "

            Mcom.CommandText &= " VALUES ('" & Application.ProductName & " " & Application.ProductVersion & "',"
            Mcom.CommandText &= "'" & user & "','" & docno & "','" & supco & "','" & prdcd & "','" & tglpb & "',"
            Mcom.CommandText &= "'" & qtyinput & "','" & frc_pcs & "','" & minor & "','" & qtypb & "','" & sj_qty & "','" & keterangan & "'); "
            Mcom.ExecuteNonQuery()

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "tracelog_errorBPBBKL", conn)
        Finally
            conn.Close()
        End Try
        Return True
    End Function
    Public Function inputTGL_EXP_Bazar(ByVal tabel As String, ByVal Barcode_Plu As String, ByVal tgl As String) As ClsBazar
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)
        Dim Sdap As New MySqlDataAdapter("", Scon)
        Dim DtPlu As New DataTable
        Dim Results As New ClsBazar

        Dim FINISHW As String = ""
        Dim tgl_exp As String
        Dim tgl_exp_sz As String

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            tgl_exp = tgl
            Scom.CommandText = "Select plu from barcode where barcd = '" & Barcode_Plu & "'"
            Dim plu As String = Scom.ExecuteScalar
            Scom.CommandText = "SELECT prdcd, singkat, bulan_exp from " & tabel & " where (barcode = '" & Barcode_Plu & "' OR PRDCD = '" & Barcode_Plu & "');"

            Sdap.SelectCommand = Scom
            Sdap.Fill(DtPlu)

            For i As Integer = 0 To DtPlu.Rows.Count - 1
                Results.BZR = New ClsBazar
                Results.BZR.BarcodePlu = Barcode_Plu
                Results.BZR.PRDCD = DtPlu.Rows(i)("prdcd").ToString
                Results.BZR.Deskripsi = DtPlu.Rows(i)("singkat").ToString
                Results.BZR.Tgl_exp = DtPlu.Rows(i)("bulan_exp").ToString
                tgl_exp_sz = Results.BZR.Tgl_exp
                If tgl_exp.Length = "6" Or tgl_exp = "00" Then
                    Try
                        If tgl_exp = tgl_exp_sz Then
                            Results.StatusExp = "2"
                            Results.Feedback = "TglExp ditemukan"
                        Else
                            Results.StatusExp = "2"
                            Results.Feedback = "TglExp Baru"
                        End If

                    Catch ex As Exception
                        Results.StatusExp = "1"
                        Results.Feedback = "Format Salah"
                    End Try
                Else
                    Results.StatusExp = "1"
                    Results.Feedback = "Format Salah"
                End If
            Next
        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "inputTGL_EPW_NPS", Scon)
        Finally
            Scon.Close()
        End Try
        Return Results
    End Function

    Public Function isMainGroupDocno() As Boolean
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim tmpDt As New DataTable
        Dim result As Boolean = False
        Dim mcom As New MySqlCommand("", conn)

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            '30/1/23
            'Perubahan baca parameter dari DOCNO menjadi GROUP DOCNO
            mcom.CommandText = "SELECT COUNT(*) FROM vir_bacaprod WHERE jenis='BEBASPPN_PERSUBBKP' AND program='PosBPB' AND `filter`='ON';"
            If mcom.ExecuteScalar > 0 Then
                result = True
            End If

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "isMainGroupDocno", conn)
        Finally
            conn.Close()
        End Try

        Return result
    End Function


End Class
