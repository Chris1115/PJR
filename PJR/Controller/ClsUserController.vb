Imports MySql.Data.MySqlClient
Imports IDM.Fungsi
Public Class ClsUserController
    Private utility As New Utility

    ''' <summary>
    ''' untuk proses pindah user dari table PASSTOKO ke USERSODCP
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CopyPassTokToUserSO() As Boolean
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim result As Boolean = True
        Dim tmpUser As New DataTable
        Dim mcom As New MySqlCommand("", conn)
        Dim sql As String = ""
        Dim Count As String = ""

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            mcom.CommandText = "SELECT COUNT(COLUMN_NAME) FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE"
            mcom.CommandText &= " WHERE TABLE_SCHEMA = 'pos' AND TABLE_NAME = 'USERSODCP'"
            mcom.CommandText &= " AND CONSTRAINT_NAME='PRIMARY';"
            Count = mcom.ExecuteScalar & ""
            If Count <> "" Then
                If Convert.ToUInt64(Count) = 0 Then
                    mcom.CommandText = "ALTER TABLE pos.USERSODCP ADD PRIMARY KEY (ID);"
                    mcom.ExecuteNonQuery()
                End If
            End If

            mcom.CommandText = "SELECT * FROM passtoko;"
            Dim sDap As New MySqlDataAdapter(mcom)
            sDap.Fill(tmpUser)

            If tmpUser.Rows.Count > 0 Then
                Dim PassUser As String = ""
                sql = "INSERT IGNORE INTO USERSODCP (`NAMA`,`GROUP`,`ID`,`PASSWORD`) VALUES"
                For Each dr As DataRow In tmpUser.Rows
                    'If IsNumeric(dr("PASS")) Then
                    '    PassUser = dr("PASS").ToString
                    'Else
                    '    PassUser = ""
                    'End If
                    sql &= " ('" & dr("NAMA") & "','TOKO','" & dr("NIK") & "', '" & PassUser & "'),"
                Next
                sql = sql.Remove(sql.Length - 1, 1)
                sql &= ";"
                mcom.CommandText = sql
                mcom.ExecuteNonQuery()
            End If

        Catch ex As Exception
            result = False
        Finally
            conn.Close()
        End Try
        Return result

    End Function

    ''' <summary>
    ''' untuk mereset ulang data user SO
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ResetUserSO() As Boolean
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim result As Boolean = True
        Dim mcom As New MySqlCommand("", conn)

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            mcom.CommandText = "UPDATE pos.UserSODCP SET inused = '', IpUser='' WHERE IPUSER NOT LIKE '%DCPSOVB%';"
            mcom.ExecuteNonQuery()
           
            result = True

        Catch ex As Exception
            result = False
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "ResetUserSO", conn)
        Finally
            conn.Close()
        End Try

        Return result
    End Function

    ''' <summary>
    ''' untuk pengecekan apakah sudah ada user di USERSODCP
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CekUserSO() As Boolean
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim result As Boolean = True
        Dim mcom As New MySqlCommand("", conn)

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            mcom.CommandText = "select count(*) from UserSODCP;"
            If mcom.ExecuteScalar > 0 Then
                result = True
            Else
                result = False
            End If
        Catch ex As Exception
            result = False
        Finally
            conn.Close()
        End Try
        Return result

    End Function

    ''' <summary>
    ''' untuk mendapatkan user login
    ''' </summary>
    ''' <param name="id"></param>
    ''' <param name="password"></param>
    ''' <param name="ip_client"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetUserLogin(ByVal id As String, ByVal password As String, ByVal ip_client As String, ByVal jenis_so As String) As ClsLogin
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)
        Dim tmpDt As New DataTable
        Dim result As New ClsLogin
        Dim cpjr As New ClsPJRController
        Dim NikToko As String = ""
        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SELECT ID, NAMA, PASSWORD,Inused,IpUser FROM USERSODCP "
            Scom.CommandText &= "WHERE id='" & id & "'"
            If jenis_so.ToUpper = "BPB" Or jenis_so.ToUpper = "PLANOGRAM" Or jenis_so = "BPBBKL" Or jenis_so = "ptag_wdcp" Or jenis_so = "Khusus" Or jenis_so.ToLower = "bpbnps" Or jenis_so.ToLower = "bazar" Or jenis_so = "Kesegaran" Or jenis_so.ToUpper = "TINDAKLBTD_BAPJR" Or jenis_so.ToUpper = "CEKDISPLAY" Or jenis_so.ToUpper = "EXPIRED" Then
                Scom.CommandText &= " AND `GROUP` = 'TOKO'"
            ElseIf jenis_so.ToUpper = "CEKPJR" Or jenis_so.ToUpper = "TINDAKLBTD" Then
                NikToko = cpjr.getConstNIKPJR
                Scom.CommandText &= " AND `GROUP` = 'TOKO' AND ID = '" & NikToko & "'"
            Else
                If (FormMain.Toko.Kode.StartsWith("P") Or FormMain.Toko.Kode.StartsWith("B") Or FormMain.Toko.Kode.StartsWith("Y")) And FormMain.tabel_name.ToLower.Contains("sbe") Then
                    Scom.CommandText &= " AND `GROUP` = 'TOKO'"
                    Scom.CommandText &= " AND id IN (SELECT menoin FROM SOPPAGENT.ABSPEGAWAIMST 
                                          WHERE JABATAN IN('KEPALA TOKO','KEPALA TOKO (SS)','ASISTEN KEPALA TOKO','ASISTEN KEPALA TOKO (SS)','MERCHANDISER',
                                          'MERCHANDISER (SS)','STORE JR. LEADER','STORE JR. LEADER (SS)','STORE SR. LEADER','STORE SR. LEADER (SS)','CHIEF OF STORE',
                                          'CHIEF OF STORE (SS)')) "
                Else
                    Scom.CommandText &= " AND `GROUP` = 'BIC'"
                End If
            End If

            TraceLog("GetUserLogin-Q1: " & Scom.CommandText)

            Dim sDap As New MySqlDataAdapter(Scom)
            sDap.Fill(tmpDt)
            If tmpDt.Rows.Count = 0 Then
                result.Status = "2"
                result.Message = "NIK Tidak Terdaftar!"
            Else
                result.User = New ClsUser
                result.User.ID = tmpDt.Rows.Item(0)("ID").ToString
                result.User.Nama = tmpDt.Rows.Item(0)("NAMA").ToString
                result.User.Password = tmpDt.Rows.Item(0)("PASSWORD").ToString
                result.User.Status = tmpDt.Rows.Item(0)("Inused").ToString
                result.User.IpAddress = tmpDt.Rows.Item(0)("IpUser").ToString

                If result.User.Password.Trim.Length = 0 Then
                    result.User.Status = "5"
                    Return result
                    Exit Function
                End If
                If result.User.ID = id And result.User.Password = password Then
                    If result.User.IpAddress = ip_client.ToString Or result.User.IpAddress = "" Then
                        Scom.CommandText = "Update UserSODCP set Inused = '1',IpUser='" & ip_client & "' where ID = '" & id & "'"
                        TraceLog("GetUserLogin-Q2: " & Scom.CommandText)
                        Scom.ExecuteNonQuery()

                        result.Status = "1"
                        result.Message = "Login sukses Sebagai = " + tmpDt.Rows.Item(0)(1).ToString
                    Else
                        result.Status = "3"
                        result.Message = "IP sudah dipakai!"
                    End If
                ElseIf result.User.Password <> password Then
                    result.Status = "4"
                    result.Message = "Password salah!"
                End If
            End If

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetUserLogin", Scon)
        Finally
            Scon.Close()
        End Try

        Return result
    End Function

    ''' <summary>
    ''' untuk mencari user ID SO
    ''' </summary>
    ''' <param name="id"></param>
    ''' <param name="menu"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CariUserSO(ByVal id As String, ByVal menu As String) As ClsUser
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim tmpDt As New DataTable
        Dim result As New ClsUser
        Dim mcom As New MySqlCommand("", conn)

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            If menu = "Edit" Then
                mcom.CommandText = "SELECT ID,NAMA,PASSWORD,`GROUP` FROM USERSODCP where id = '" & id & "'"
            ElseIf menu = "Hapus" Or menu = "Tambah" Then
                mcom.CommandText = "SELECT ID,NAMA,PASSWORD,`GROUP` FROM USERSODCP where id = '" & id & "' AND `GROUP` = 'TOKO' "
            End If

            Dim sDap As New MySqlDataAdapter(mcom)
            sDap.Fill(tmpDt)

            If tmpDt.Rows.Count > 0 Then
                result.ID = tmpDt.Rows.Item(0)("ID").ToString
                result.Nama = tmpDt.Rows.Item(0)("NAMA").ToString
                result.Password = tmpDt.Rows.Item(0)("PASSWORD").ToString
                result.Group = tmpDt.Rows.Item(0)("GROUP").ToString
            Else
                result = Nothing
            End If

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "CariUserSO", conn)
        Finally
            conn.Close()
        End Try

        Return result

    End Function

    ''' <summary>
    ''' Insert user SO baru
    ''' </summary>
    ''' <param name="id"></param>
    ''' <param name="nama"></param>
    ''' <param name="password"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InsertUserSO(ByVal id As String, ByVal nama As String, ByVal password As String) As Boolean
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim result As Boolean = True
        Dim mcom As New MySqlCommand("", conn)

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            mcom.CommandText = "INSERT USERSODCP SET ID = '" & id & "', NAMA = '" & nama & "',"
            mcom.CommandText &= " PASSWORD = '" & password & "', `GROUP` = 'TOKO';"
            mcom.ExecuteNonQuery()

            result = True

        Catch ex As Exception
            result = False
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "InsertUserSO", conn)
        Finally
            conn.Close()
        End Try

        Return result
    End Function

    ''' <summary>
    ''' Update User SO
    ''' </summary>
    ''' <param name="id"></param>
    ''' <param name="nama"></param>
    ''' <param name="password"></param>
    ''' <param name="id_cari"></param>
    ''' <param name="group"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UpdateUserSO(ByVal id As String, ByVal nama As String, ByVal password As String, ByVal id_cari As String, ByVal group As String) As Boolean
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim result As Boolean = True
        Dim mcom As New MySqlCommand("", conn)

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If


            mcom.CommandText = "Update UserSODCP SET ID = '" & id & "', NAMA = '" & nama & "', "
            mcom.CommandText += " PASSWORD = '" & password & "' WHERE ID = '" & id_cari & "'"
            mcom.CommandText += " AND `GROUP` = '" & group & "' "
            mcom.ExecuteNonQuery()

            result = True

        Catch ex As Exception
            result = False
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "UpdateUserSO", conn)
        Finally
            conn.Close()
        End Try

        Return result
    End Function

    ''' <summary>
    ''' Update Password User SO
    ''' </summary>
    ''' <param name="id"></param>
    ''' <param name="nama"></param>
    ''' <param name="password"></param>
    ''' <param name="id_cari"></param>
    ''' <param name="group"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UpdatePassUserSO(ByVal id As String, ByVal newpassword As String) As Boolean
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim result As Boolean = True
        Dim mcom As New MySqlCommand("", conn)

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If


            mcom.CommandText = "Update UserSODCP SET PASSWORD = '" & newpassword & "' "
            mcom.CommandText += " WHERE ID = '" & id & "'"
            mcom.ExecuteNonQuery()

            result = True

        Catch ex As Exception
            result = False
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "UpdateUserSO", conn)
        Finally
            conn.Close()
        End Try

        Return result
    End Function

    ''' <summary>
    ''' Hapus User SO
    ''' </summary>
    ''' <param name="id"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DeleteUserSO(ByVal id As String) As Boolean
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim result As Boolean = True
        Dim mcom As New MySqlCommand("", conn)

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            mcom.CommandText = "DELETE FROM USERSODCP where ID = '" & id & "' AND `GROUP` = 'TOKO'"
            mcom.ExecuteNonQuery()

            result = True
        Catch ex As Exception
            result = False
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "DeleteUserSO", conn)
        Finally
            conn.Close()
        End Try

        Return result
    End Function

    ''' <summary>
    ''' Cek Const LTF
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsConstLTF(ByRef MsgBox As String) As Boolean
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim result As Boolean = True
        Dim mcom As New MySqlCommand("", conn)

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            mcom.CommandText = "SELECT jenis FROM pos.const WHERE rkey = 'LTF';"
            Dim Var = mcom.ExecuteScalar()
            If Not IsNothing(Var) And Not IsDBNull(Var) Then
                If Var.ToString.ToUpper = "Y" Then
                    result = True
                Else
                    result = False
                End If
            Else
                result = False
                MsgBox = "Harap setting mode LTF."
            End If

        Catch ex As Exception
            result = False
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "IsConstLTF", conn)
        Finally
            conn.Close()
        End Try

        Return result
    End Function


    ''' <summary>
    ''' Cek IsTutupHarian
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsTutupHarian() As Boolean
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim result As Boolean = True
        Dim mcom As New MySqlCommand("", conn)
        Dim tglSO, tglTtpHr As Date

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            mcom.CommandText = "SELECT period1 FROM const WHERE rkey = 'SO4';"
            TraceLog("IsTutupHarian-Q1: " & mcom.CommandText)
            Dim tempTglSO = mcom.ExecuteScalar()
            If Not IsNothing(tempTglSO) And Not IsDBNull(tempTglSO) Then
                tglSO = Date.Parse(tempTglSO.ToString)
                mcom.CommandText = "SELECT MAX(tanggal) AS tanggal FROM initial WHERE RECID = 'C';"
                Dim tempTglTtpHr = mcom.ExecuteScalar()
                If Not IsNothing(tempTglTtpHr) And Not IsDBNull(tempTglTtpHr) Then
                    tglTtpHr = Date.Parse(tempTglTtpHr.ToString)
                End If
            End If

            If tglSO = tglTtpHr Then
                result = True
            Else
                result = False
            End If

        Catch ex As Exception
            result = False
            TraceLog("Error IsTutupHarian" & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        Finally
            conn.Close()
        End Try

        Return result
    End Function

End Class
