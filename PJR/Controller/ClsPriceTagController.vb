Imports MySql.Data.MySqlClient
Public Class ClsPriceTagController
    Private utility As New Utility

    Public Function CekTablePriceTag_wdcp(ByVal tablename As String, ByVal confirm As Boolean) As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Mda As New MySqlDataAdapter("", Conn)

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`" & tablename & "` ( "
            Mcom.CommandText &= "`RECID` DECIMAL(1), "
            Mcom.CommandText &= "`PRDCD` VARCHAR(8) NOT NULL DEFAULT '', "
            Mcom.CommandText &= "`SINGKAT` VARCHAR(45) NOT NULL, "
            Mcom.CommandText &= "`KODEPROMO` VARCHAR(10) DEFAULT '', "
            Mcom.CommandText &= "`KEYPROMOSI` VARCHAR(100) DEFAULT '', "
            Mcom.CommandText &= "`PRICE` DECIMAL(20,6), "
            Mcom.CommandText &= "`LT` VARCHAR(30) NOT NULL DEFAULT '', "
            Mcom.CommandText &= "`RAK` INT(11) Not NULL, "
            Mcom.CommandText &= "`BAR` INT(11) NOT NULL, "
            Mcom.CommandText &= "`UNIT` VARCHAR(4), "
            Mcom.CommandText &= "`CAT_COD` VARCHAR(8), "
            Mcom.CommandText &= "`PROMOSI` DECIMAL(20,6), "
            Mcom.CommandText &= "`TGL_AWL` DATE, "
            Mcom.CommandText &= "`TGL_AKH` DATE, "
            Mcom.CommandText &= "`BARCODE` VARCHAR(15), "
            Mcom.CommandText &= "`FCETAK` CHAR(1) DEFAULT '1', "
            Mcom.CommandText &= "`PLUMD` VARCHAR(8), "
            Mcom.CommandText &= "`EXPIRED` VARCHAR(50), "
            Mcom.CommandText &= "`HMIN1` VARCHAR(1), "
            Mcom.CommandText &= "PRIMARY KEY (`PRDCD`, `LT`, `RAK`, `BAR`) ) "
            Mcom.CommandText &= "ENGINE=INNODB CHARSET=latin1 COLLATE=latin1_swedish_ci;"
            Mcom.ExecuteNonQuery()

            Mcom.CommandText = "ALTER TABLE `pos`.`ptag_wdcp` CHANGE `RECID` `RECID` CHAR(1) DEFAULT '' NULL"
            Mcom.ExecuteNonQuery()

            If confirm = True Then
                Mcom.CommandText = "DELETE FROM `" & tablename & "`"
                Mcom.ExecuteNonQuery()
                utility.Tracelog("Query", "Confirm = True : " & Mcom.CommandText, "CekTablePlano", Conn)

            End If


        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "CekTablePlano", Conn)
        Finally
            Conn.Close()
        End Try

        Return True
    End Function


    Public Function GetDeskripsiPriceTag(ByVal tabel_name As String, ByVal barcode_plu As String, ByVal confirm As String) As ClsPriceTag
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Result As New ClsPriceTag
        Dim dt As New DataTable
        If Conn Is Nothing Then
            utility.TraceLogTxt("Error - GetDeskripsiPriceTag (connection Nothing) " & vbCrLf & "PLU:" & barcode_plu)
            Return Result
            Exit Function
        End If

        SyncLock Conn
            Try
                If Conn.State = ConnectionState.Closed Then
                    Conn.Open()
                End If

                Mcom.CommandText = "SELECT prdcd,singkat FROM ptag_old WHERE prdcd = '" & barcode_plu & "' OR barcode = '" & barcode_plu & "';"
                Dim sDap As New MySqlDataAdapter(Mcom)
                sDap.Fill(dt)
              
                If dt.Rows.Count <= 0 Then
                    Mcom.CommandText = "SELECT prdcd,singkat FROM ptag WHERE prdcd = '" & barcode_plu & "' OR barcode = '" & barcode_plu & "';"
                    dt.Clear()
                    sDap = New MySqlDataAdapter(Mcom)
                    sDap.Fill(dt)

                    If dt.Rows.Count <= 0 Then
                        Mcom.CommandText = "DROP TABLE IF EXISTS `temp_ptag_wdcp`"
                        Mcom.ExecuteNonQuery()

                        Mcom.CommandText = "CREATE TABLE `temp_ptag_wdcp` Select ''AS RECID,P.PRDCD,P.PLUMD,P.SINGKATAN AS SINGKAT,P.PRICE,"
                        Mcom.CommandText &= "R.NAMA_RAK AS LT,R.NOSHELF AS RAK,R.KIRIKANAN AS BAR,P.UNIT,"
                        Mcom.CommandText &= "CONCAT(P.DIVISI,P.DEPART,P.KATEGORI) AS CAT_COD,"
                        Mcom.CommandText &= "IF(NOT PRM.PROMOSI IS NULL,PRM.PROMOSI,IF((PRM_OLD.TGL_AWL<=CURDATE() AND PRM_OLD.TGL_AKH>=CURDATE()),PRM_OLD.PROMOSI ,0)) AS PROMOSI "
                        Mcom.CommandText &= ",cast(IF(NOT PRM.PROMOSI IS NULL,PRM.TGL_AWL,IF((PRM_OLD.TGL_AWL<=CURDATE() AND PRM_OLD.TGL_AKH>=CURDATE()),PRM_OLD.TGL_AWL,NULL)) as date) AS TGL_AWL "
                        Mcom.CommandText &= ",cast(IF(NOT PRM.PROMOSI IS NULL,PRM.TGL_AKH,IF((PRM_OLD.TGL_AWL<=CURDATE() AND PRM_OLD.TGL_AKH>=CURDATE()),PRM_OLD.TGL_AKH,NULL)) as date) AS TGL_AKH "
                        Mcom.CommandText &= ",B.BARCD AS BARCODE"
                        Mcom.CommandText &= ",'1' AS FCETAK "
                        Mcom.CommandText &= ",R.TKIRIKANAN "
                        Mcom.CommandText &= ",R.TATASBAWAH "
                        Mcom.CommandText &= ",R.TDEPANBLK "
                        Mcom.CommandText &= ",CAST(CONCAT(IFNULL(P.STATUS_RETUR,''), "
                        Mcom.CommandText &= "IF(CONCAT(',(E-',IFNULL(PF.MAX_RET_TOKO2DCI,''),IFNULL(PF.MAX_RET_TOKO2DCI_S,''),')')=',(E-)','', "
                        Mcom.CommandText &= "CONCAT(',(E-',IFNULL(PF.MAX_RET_TOKO2DCI,''),IFNULL(PF.MAX_RET_TOKO2DCI_S,''),')')) "
                        Mcom.CommandText &= ") AS CHAR) AS EXPIRED, "
                        Mcom.CommandText &= "CASE WHEN P.flagprod LIKE '%mdr=1%' THEN '1' ELSE "
                        Mcom.CommandText &= "CASE WHEN P.flagprod LIKE '%mdr=2%' THEN '2' ELSE "
                        Mcom.CommandText &= "CASE WHEN P.flagprod LIKE '%mdr=3%' THEN '3' ELSE "
                        Mcom.CommandText &= "CASE WHEN P.flagprod LIKE '%mdr=4%' THEN '4' ELSE '' "
                        Mcom.CommandText &= "END END END END AS ITEMMANDIRI "
                        Mcom.CommandText &= ",PRM.HMIN1  "
                        Mcom.CommandText &= "FROM PRODMAST P LEFT JOIN RAK R ON P.PRDCD=R.PRDCD "
                        Mcom.CommandText &= " LEFT JOIN (SELECT PLU,BARCD FROM BARCODE WHERE PLU='" & barcode_plu & "' OR BARCD = '" & barcode_plu & "' LIMIT 1) B ON P.PRDCD=B.PLU "
                        Mcom.CommandText &= " LEFT JOIN PTAG PRM ON P.PRDCD=PRM.PRDCD "
                        Mcom.CommandText &= " LEFT JOIN BATAS_RETUR PF ON P.PRDCD=PF.FMKODE "
                        Mcom.CommandText &= " LEFT JOIN PTAG_OLD PRM_OLD ON P.PRDCD=PRM_OLD.PRDCD "
                        Mcom.CommandText &= " WHERE (P.PTAG NOT IN('N','R') OR PTAG IS NULL) "
                        Mcom.CommandText &= " AND (P.PRDCD='" & barcode_plu & "' OR B.BARCD = '" & barcode_plu & "') "
                        Mcom.CommandText &= " AND (P.RECID<>'1' OR P.RECID IS NULL) "
                        Mcom.CommandText &= " AND R.NORAK IS NOT NULL "
                        Mcom.CommandText &= " AND P.PRICE>0 "
                        Mcom.CommandText &= " GROUP BY PRDCD,LT,RAK,BAR "
                        Mcom.CommandText &= " ORDER BY CONCAT(P.DIVISI,P.DEPART,P.KATEGORI); "
                        'IDM.Fungsi.TraceLog(Mcom.CommandText)
                        Mcom.ExecuteNonQuery()



                        Mcom.CommandText = "SELECT prdcd,singkat FROM temp_ptag_wdcp WHERE prdcd = '" & barcode_plu & "' OR barcode = '" & barcode_plu & "';"

                        dt.Clear()
                        sDap = New MySqlDataAdapter(Mcom)
                        sDap.Fill(dt)
                        'IDM.Fungsi.TraceLog(Mcom.CommandText)
                        If dt.Rows.Count <= 0 Then
                            Result.Prdcd = ""
                            Result.Desc = "PLU Tidak Aktif"
                            Result.Keterangan = ""

                        Else
                            Mcom.CommandText = "SELECT prdcd from ptag_Wdcp where prdcd = '" & barcode_plu & "' OR barcode = '" & barcode_plu & "' group by prdcd"
                            If Mcom.ExecuteScalar = "" Then
                                If confirm = "1" Then

                                    Mcom.CommandText = "INSERT ignore INTO PTAG_WDCP SELECT "
                                    Mcom.CommandText &= "`RECID`,`PRDCD`, `SINGKAT` , '', '', "
                                    Mcom.CommandText &= "`PRICE`, `LT`, `RAK`, `BAR`, `UNIT`, `CAT_COD`, `PROMOSI`, "
                                    Mcom.CommandText &= "`TGL_AWL`, `TGL_AKH`, `BARCODE`, `FCETAK`, `PLUMD` , `EXPIRED` , `HMIN1`"
                                    Mcom.CommandText &= "FROM temp_ptag_wdcp WHERE prdcd = '" & barcode_plu & "' OR barcode = '" & barcode_plu & "'"
                                    Mcom.ExecuteNonQuery()
                                    Result.Prdcd = ""
                                    Result.Desc = ""
                                    Result.Keterangan = ""
                                    'IDM.Fungsi.TraceLog(Mcom.CommandText)
                                    'Jika konfirm NO
                                ElseIf confirm = "2" Then
                                    Result.Prdcd = ""
                                    Result.Desc = ""
                                    Result.Keterangan = ""

                                    'Jika konfirm SELAIN 1/2
                                Else
                                    Result.Prdcd = barcode_plu
                                    Result.Desc = dt.Rows(0).Item("SINGKAT")
                                    Result.Keterangan = "INSERT"

                                End If
                            Else
                                If confirm = "1" Then
                                    Mcom.CommandText = "DELETE FROM PTAG_wDCP WHERE prdcd = '" & barcode_plu & "' OR barcode = '" & barcode_plu & "'"
                                    Mcom.ExecuteNonQuery()
                                    Result.Prdcd = ""
                                    Result.Desc = ""
                                    Result.Keterangan = ""
                                ElseIf confirm = "2" Then
                                    Result.Prdcd = ""
                                    Result.Desc = ""
                                    Result.Keterangan = ""
                                Else
                                    Result.Prdcd = barcode_plu
                                    Result.Desc = dt.Rows(0).Item("SINGKAT")
                                    Result.Keterangan = "HAPUS"
                                End If
                            End If
                            
                        End If





                    Else
                        Mcom.CommandText = "SELECT prdcd from ptag_Wdcp where prdcd = '" & barcode_plu & "' OR barcode = '" & barcode_plu & "' group by prdcd"

                        If Mcom.ExecuteScalar = "" Then

                            'Jika konfirm YES
                            If confirm = "1" Then

                                Mcom.CommandText = "INSERT INTO PTAG_WDCP SELECT"
                                Mcom.CommandText &= "`RECID`,`PRDCD`, `SINGKAT` , `KODEPROMO`, `KEYPROMOSI`, "
                                Mcom.CommandText &= "`PRICE`, `LT`, `RAK`, `BAR`, `UNIT`, `CAT_COD`, `PROMOSI`, "
                                Mcom.CommandText &= "`TGL_AWL`, `TGL_AKH`, `BARCODE`, `FCETAK`, `PLUMD` , `EXPIRED` , `HMIN1`"
                                Mcom.CommandText &= "FROM ptag WHERE prdcd = '" & barcode_plu & "' OR barcode = '" & barcode_plu & "'"
                                Mcom.ExecuteNonQuery()
                                Result.Prdcd = ""
                                Result.Desc = ""
                                Result.Keterangan = ""

                                'Jika konfirm NO
                            ElseIf confirm = "2" Then
                                Result.Prdcd = ""
                                Result.Desc = ""
                                Result.Keterangan = ""

                                'Jika konfirm SELAIN 1/2
                            Else
                                Result.Prdcd = barcode_plu
                                Result.Desc = dt.Rows(0).Item("SINGKAT")
                                Result.Keterangan = "INSERT"

                            End If


                        Else

                            If confirm = "1" Then
                                Mcom.CommandText = "DELETE FROM PTAG_wDCP WHERE prdcd = '" & barcode_plu & "' OR barcode = '" & barcode_plu & "'"
                                Mcom.ExecuteNonQuery()
                                Result.Prdcd = ""
                                Result.Desc = ""
                                Result.Keterangan = ""
                            ElseIf confirm = "2" Then
                                Result.Prdcd = ""
                                Result.Desc = ""
                                Result.Keterangan = ""
                            Else
                                Result.Prdcd = barcode_plu
                                Result.Desc = dt.Rows(0).Item("SINGKAT")
                                Result.Keterangan = "HAPUS"
                            End If


                        End If
                    End If
                Else
                    Mcom.CommandText = "SELECT prdcd from ptag_Wdcp where prdcd = '" & barcode_plu & "' OR barcode = '" & barcode_plu & "'  group by prdcd"


                    If Mcom.ExecuteScalar = "" Then

                        If confirm = "1" Then
                            Mcom.CommandText = "INSERT INTO PTAG_WDCP SELECT"
                            Mcom.CommandText &= "`RECID`,`PRDCD`, `SINGKAT` , `KODEPROMO`, `KEYPROMOSI`, "
                            Mcom.CommandText &= "`PRICE`, `LT`, `RAK`, `BAR`, `UNIT`, `CAT_COD`, `PROMOSI`, "
                            Mcom.CommandText &= "`TGL_AWL`, `TGL_AKH`, `BARCODE`, `FCETAK`, `PLUMD` , `EXPIRED` , `HMIN1`"
                            Mcom.CommandText &= "FROM ptag_old WHERE prdcd = '" & barcode_plu & "' OR barcode = '" & barcode_plu & "'"
                            Mcom.ExecuteNonQuery()
                            Result.Prdcd = ""
                            Result.Desc = ""
                            'Jika konfirm NO
                        ElseIf confirm = "2" Then
                            Result.Prdcd = ""
                            Result.Desc = ""
                            Result.Keterangan = ""
                            'Jika konfirm SELAIN 1/2
                        Else
                            Result.Prdcd = barcode_plu
                            Result.Desc = dt.Rows(0).Item("SINGKAT")
                            Result.Keterangan = "INSERT"
                        End If

                    Else
                        If confirm = "1" Then
                            Mcom.CommandText = "DELETE FROM PTAG_wDCP WHERE prdcd = '" & barcode_plu & "' OR barcode = '" & barcode_plu & "'"
                            Mcom.ExecuteNonQuery()
                            Result.Prdcd = ""
                            Result.Desc = ""
                            Result.Keterangan = ""
                        ElseIf confirm = "2" Then
                            Result.Prdcd = ""
                            Result.Desc = ""
                            Result.Keterangan = ""
                        Else
                            Result.Prdcd = barcode_plu
                            Result.Desc = dt.Rows(0).Item("SINGKAT")
                            Result.Keterangan = "HAPUS"
                        End If
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
End Class
