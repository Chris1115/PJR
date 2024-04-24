Imports IDM.Fungsi
Imports MySql.Data.MySqlClient

Public Class ClsSOICController
    Public Function isItemBKL(ByVal plu As String) As Boolean
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SELECT kons FROM prodmast WHERE "
            Scom.CommandText &= "PRDCD='" & plu & "';"
            TraceLog("isItemBKL-Q1: " & Scom.CommandText)

            If Scom.ExecuteScalar = "k" Then
                isItemBKL = True
            Else
                isItemBKL = False
            End If
        Catch ex As Exception
            TraceLog("Last Query: " & Scom.CommandText)
            TraceLog("isItemBKL Error:  " & ex.ToString)
            MsgBox("Error: " & ex.Message & ex.StackTrace, MsgBoxStyle.Exclamation, "Error isItemBKL")
        Finally
            Scon.Close()
        End Try

        Return isItemBKL
    End Function
    Public Function isBarangPutus(ByVal plu As String) As Boolean
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)
        Dim tmpDt As New DataTable

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SELECT status_retur FROM prodmast WHERE prdcd='" & plu & "';"
            TraceLog("isBarangPutus-Q1: " & Scom.CommandText)
            Dim flagprod As String = Scom.ExecuteScalar

            If flagprod.Contains("PT") Then
                isBarangPutus = True
            Else
                isBarangPutus = False
            End If
        Catch ex As Exception
            TraceLog("Last Query: " & Scom.CommandText)
            TraceLog("isBarangPutus Error:  " & ex.ToString)
            MsgBox("Error: " & ex.Message & ex.StackTrace, MsgBoxStyle.Exclamation, "Error isBarangPutus")
        Finally
            Scon.Close()
        End Try

        Return isBarangPutus
    End Function
    Public Function IsItemBAP(ByVal plu As String) As Boolean
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)
        Dim tmpDt As New DataTable

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SELECT flagprod FROM prodmast WHERE prdcd='" & plu & "';"
            TraceLog("IsItemBAP-Q1: " & Scom.CommandText)
            Dim flagprod As String = Scom.ExecuteScalar

            If flagprod.Contains("BAP=Y") Then
                IsItemBAP = True
            Else
                IsItemBAP = False
            End If
        Catch ex As Exception
            TraceLog("Last Query: " & Scom.CommandText)
            TraceLog("IsItemBAP Error:  " & ex.ToString)
            MsgBox("Error: " & ex.Message & ex.StackTrace, MsgBoxStyle.Exclamation, "Error IsItemBAP")
        Finally
            Scon.Close()
        End Try

        Return IsItemBAP
    End Function
    Public Function IsValidNonBAP(ByVal plu As String) As Boolean
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SELECT flagprod FROM prodmast WHERE prdcd='" & plu & "';"
            TraceLog("IsValidNonBAP-Q1: " & Scom.CommandText)
            Dim flagprod As String = Scom.ExecuteScalar

            If flagprod.Contains("TBR=N") Or flagprod.Contains("TBF=N") Or flagprod.Contains("TBP=N") Then
                IsValidNonBAP = True
            Else
                IsValidNonBAP = False
            End If
        Catch ex As Exception
            TraceLog("Last Query: " & Scom.CommandText)
            TraceLog("IsValidNonBAP Error:  " & ex.ToString)
            MsgBox("Error: " & ex.Message & ex.StackTrace, MsgBoxStyle.Exclamation, "Error IsValidNonBAP")
        Finally
            Scon.Close()
        End Try

        Return IsValidNonBAP
    End Function
    Public Sub cekTabelBarangRusak()
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "DROP TABLE IF EXISTS temp_DataBarangRusak;"
            TraceLog("cekTabelBarangRusak-Q1: " & Scom.CommandText)
            Scom.ExecuteNonQuery()

            Scom.CommandText = "CREATE TABLE temp_DataBarangRusak("
            Scom.CommandText &= "`RECID` VARCHAR(8) DEFAULT NULL,"
            Scom.CommandText &= "`PRDCD` VARCHAR(8) NOT NULL,"
            Scom.CommandText &= "`QTY` DECIMAL NOT NULL,"
            Scom.CommandText &= "`ALASAN` VARCHAR(50) NOT NULL,"
            Scom.CommandText &= "`TABEL` VARCHAR(10) NOT NULL,"
            Scom.CommandText &= "`ADDTIME` DATETIME NOT NULL,"
            Scom.CommandText &= "PRIMARY KEY(PRDCD, ALASAN));"
            TraceLog("cekTabelBarangRusak-Q2: " & Scom.CommandText)
            Scom.ExecuteNonQuery()
        Catch ex As Exception
            TraceLog("Last Query: " & Scom.CommandText)
            TraceLog("cekTabelBarangRusak Error:  " & ex.ToString)
            MsgBox("Error: " & ex.Message & ex.StackTrace, MsgBoxStyle.Exclamation, "Error cekTabelBarangRusak")
        Finally
            Scon.Close()
        End Try
    End Sub
    Public Sub insertOrUpdate_barangRusak(ByVal plu As String, ByVal qty As String,
                                          ByVal tabel As String, ByVal alasan As String)

        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SELECT COUNT(1) FROM temp_DataBarangRusak WHERE PRDCD='" & plu & "' AND Alasan='" & alasan & "' AND RECID IS NULL;"
            If Scom.ExecuteScalar = 0 Then
                Scom.CommandText = "INSERT IGNORE INTO temp_DataBarangRusak(PRDCD, QTY, Alasan, Tabel, ADDTIME) "
                Scom.CommandText &= "VALUES('" & plu & "', " & qty & ", '" & alasan & "', '" & tabel & "', NOW());"
                TraceLog("insertOrUpdate_barangRusak-Q1: " & Scom.CommandText)
            Else
                Scom.CommandText = "UPDATE temp_DataBarangRusak SET qty=" & qty & ""
                Scom.CommandText &= " WHERE PRDCD='" & plu & "' AND Alasan='" & alasan & "'"
                Scom.CommandText &= " AND RECID IS NULL;"
                TraceLog("insertOrUpdate_barangRusak-Q1: " & Scom.CommandText)
            End If

            Scom.ExecuteNonQuery()

        Catch ex As Exception
            TraceLog("Last Query: " & Scom.CommandText)
            TraceLog("insertOrUpdate_barangRusak Error:  " & ex.ToString)
            MsgBox("Error: " & ex.Message & ex.StackTrace, MsgBoxStyle.Exclamation, "Error insertOrUpdate_barangRusak")
        Finally
            Scon.Close()
        End Try
    End Sub
    Public Function isItemActive(ByVal plu As String) As Boolean
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SELECT COUNT(1) FROM PRODMAST WHERE "
            Scom.CommandText &= "PRDCD='" & plu & "' AND "
            Scom.CommandText &= "(RECID='' OR RECID=NULL) AND CTGR<>'99';"
            TraceLog("isItemActive-Q1: " & Scom.CommandText)

            If Scom.ExecuteScalar = 0 Then
                isItemActive = False
            Else
                isItemActive = True
            End If
        Catch ex As Exception
            TraceLog("Last Query: " & Scom.CommandText)
            TraceLog("isItemActive Error:  " & ex.ToString)
            MsgBox("Error: " & ex.Message & ex.StackTrace, MsgBoxStyle.Exclamation, "Error isItemActive")
        Finally
            Scon.Close()
        End Try

        Return isItemActive
    End Function
    Public Function cekFisikStmast(ByVal plu As String) As Double
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SELECT QTY FROM STMAST WHERE "
            Scom.CommandText &= "PRDCD='" & plu & "';"
            TraceLog("cekFisikSTMAST-Q1: " & Scom.CommandText)

            cekFisikStmast = Scom.ExecuteScalar

        Catch ex As Exception
            TraceLog("Last Query: " & Scom.CommandText)
            TraceLog("cekFisikSTMAST Error:  " & ex.ToString)
            MsgBox("Error: " & ex.Message & ex.StackTrace, MsgBoxStyle.Exclamation, "Error cekFisikSTMAST")
        Finally
            Scon.Close()
        End Try

        Return cekFisikStmast
    End Function
    Public Function cekAcostItem(ByVal plu As String) As Double
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SELECT ACOST FROM PRODMAST WHERE "
            Scom.CommandText &= "PRDCD='" & plu & "';"
            TraceLog("cekAcostItem-Q1: " & Scom.CommandText)

            cekAcostItem = Scom.ExecuteScalar

        Catch ex As Exception
            TraceLog("Last Query: " & Scom.CommandText)
            TraceLog("cekAcostItem Error:  " & ex.ToString)
            MsgBox("Error: " & ex.Message & ex.StackTrace, MsgBoxStyle.Exclamation, "Error cekAcostItem")
        Finally
            Scon.Close()
        End Try

        Return cekAcostItem
    End Function
End Class
