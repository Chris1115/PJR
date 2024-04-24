Imports MySql.Data.MySqlClient
Imports System.Reflection
Imports System.Security.Cryptography
Imports System.Text									

Public Class Utility


    ''' <summary>
    ''' Tracelog SQL
    ''' </summary>
    ''' <param name="Type"></param>
    ''' <param name="MessageLog"></param>
    ''' <param name="Location"></param>
    ''' <param name="conn"></param>
    ''' <remarks></remarks>
    Public Sub Tracelog(ByVal Type As String, ByVal MessageLog As String, ByVal Location As String, ByVal conn As MySqlConnection)
        Dim appName As String = getAppNameVersion()

        MessageLog = MessageLog.Replace("'", "''").Replace("""", """""")
        MessageLog = MessageLog.Replace("''.", "'.").Trim

        If MessageLog.EndsWith("\") Then
            MessageLog = MessageLog.Substring(0, MessageLog.Length - 1)
        End If
        If MessageLog.Length > 4000 Then
            MessageLog = MessageLog.Substring(0, 4000)
        End If

        While MessageLog.EndsWith("'") And Not MessageLog.EndsWith("''")
            MessageLog = MessageLog.Substring(0, MessageLog.Length - 1)
        End While

        Dim mcom As New MySqlCommand("", conn)
        Try
            mcom.CommandText = "INSERT INTO Tracelog(TGL,Tipe,AppName,`Log`) VALUES (NOW(),'" & Type & "','" & appName & "','" & Location & " - " & MessageLog & "');"
            mcom.ExecuteNonQuery()

        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' Tarcelog txt
    ''' </summary>
    ''' <param name="Teks"></param>
    ''' <remarks></remarks>
    Public Sub TraceLogTxt(ByVal Teks As String)
        Try
            'directory untuk tracelog bentuk txt
            Dim FolderName As String = Application.StartupPath & "\TracelogHandheldSO"

            If (Not System.IO.Directory.Exists(FolderName)) Then
                System.IO.Directory.CreateDirectory(FolderName)
            End If

            'nama file untuk tracelog bentuk txt
            Dim FileName As String = FolderName & "\" & "HandheldSO_" & Format(Date.Now(), "yyyyMMdd") & ".txt"

            'tulis tracelog bentuk txt
            Dim sw As IO.StreamWriter = Nothing
            If Not (IO.File.Exists(FileName)) Then
                sw = IO.File.CreateText(FileName)
            Else
                sw = IO.File.AppendText(FileName)
            End If
            sw.WriteLine(Format(Date.Now(), "yyyy-MM-dd HH:mm:ss") & vbCrLf & Teks & vbCrLf)
            sw.Close()

        Catch ex As Exception
            'ErrorTryCatch(ex)
        End Try
    End Sub

    ''' <summary>
    ''' get nama aplikasi dan versi applikasi
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getAppNameVersion() As String

        Dim AppName As String = ""
        Dim AppVersion As String = ""
        Dim result As String = ""

        Try
            AppName = System.IO.Path.GetFileName(Application.ExecutablePath)
            AppVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString

            result = AppName & " - V." & AppVersion
        Catch ex As Exception

        End Try

        Return result

    End Function

    Public Function ExecuteScalar(ByVal query As String) As Object
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim result As New Object
        Dim mcom As New MySqlCommand("", conn)

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            mcom.CommandText = query
            result = mcom.ExecuteScalar
            Console.WriteLine(result)
            Console.WriteLine(query)
        Catch ex As Exception
            Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "ExecuteScalar", conn)
        Finally
            conn.Close()
        End Try

        Return result
    End Function

    Public Function ExecuteNonQuery(ByVal query As String) As String
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim result As String = ""
        Dim mcom As New MySqlCommand("", conn)

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            mcom.CommandText = query
            mcom.ExecuteNonQuery()

        Catch ex As Exception
            Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "ExecuteNonQuery", conn)
        Finally
            conn.Close()
        End Try

        Return result
    End Function

    Public Function BrcdKontainerOrBronjong(ByVal Brcd As String) As String
        Dim tmpBrcd As String = ""

        If Strings.Left(brcd, 1) = "G" Then
            tmpBrcd = Brcd
        Else
            tmpBrcd = Strings.Left(Brcd, 12)
        End If

        Return tmpBrcd
    End Function
Public Shared Function AES128_Decrypt(ByVal textToDecrypt As String, ByVal pass As String) As String
        Dim hasil As String = ""
        Try
            Dim rijndaelCipher As New RijndaelManaged()
            rijndaelCipher.Mode = CipherMode.CBC
            rijndaelCipher.Padding = PaddingMode.PKCS7

            rijndaelCipher.KeySize = &H80
            rijndaelCipher.BlockSize = &H80
            Dim encryptedData As Byte() = Convert.FromBase64String(textToDecrypt)
            Dim pwdBytes As Byte() = Encoding.UTF8.GetBytes(pass)
            Dim keyBytes As Byte() = New Byte(15) {}
            Dim len As Integer = pwdBytes.Length
            If len > keyBytes.Length Then
                len = keyBytes.Length
            End If
            Array.Copy(pwdBytes, keyBytes, len)
            rijndaelCipher.Key = keyBytes
            rijndaelCipher.IV = keyBytes
            Dim plainText As Byte() = rijndaelCipher.CreateDecryptor().TransformFinalBlock(encryptedData, 0, encryptedData.Length)
            hasil = Encoding.UTF8.GetString(plainText)
        Catch ex As Exception
            'tracelog_1("Error di AES128_Decrypt " & ex.Message & vbCrLf & ex.StackTrace)
            'hasil = "Err|" & ex.Message & "|" & ex.StackTrace
        End Try
        Return hasil
    End Function

    Public Shared Function Close_InterfaceWifi()
        Try
            Dim proc As Process() = Process.GetProcessesByName("InterfaceWIFI")
            For Each p As Process In proc
                Console.WriteLine(p.ProcessName & " / " & p.Id.ToString)
                p.CloseMainWindow()
            Next
            'Shell("taskkill /f /im InterfaceWIFI.exe")
        Catch ex As Exception

        End Try
        Return True
    End Function

    'Penambahan fungsi alter table 01/06/2023
    Public Sub alterTable(ByVal tipe As Integer, ByVal tableName As String, ByVal columnName As String, ByVal syntax As String, Optional ByVal syntaxTambahan As String = "")
        Dim Mcon As MySqlConnection
        Dim Scom As New MySqlCommand
        Dim Sdap As New MySqlDataAdapter
        Dim dt As New DataTable

        Try
            Mcon = ClsConnection.GetConnection.Clone
            If Mcon.State = ConnectionState.Closed Then
                Mcon.Open()
            End If
            Scom.Connection = Mcon

            Try
                If tipe = 1 Then 'cek tabel dl baru kolom
                    Scom.CommandText = "SELECT COUNT(*) FROM Information_schema.tables WHERE TABLE_SCHEMA='pos' AND Table_Name='" & tableName.ToUpper & "'; "
                    If Scom.ExecuteScalar > 0 Then
                        Scom.CommandText = "Select column_type From Information_schema.Columns Where TABLE_SCHEMA='pos' AND Table_Name='" & tableName.ToUpper & "' And Column_Name='" & columnName.ToUpper & "' "
                        If Scom.ExecuteScalar & "" = "" Then
                            Scom.CommandText = "ALTER TABLE " & tableName.ToUpper & " " & syntax
                            Scom.ExecuteNonQuery()

                            If syntaxTambahan & "" <> "" Then
                                Scom.CommandText = "" & syntaxTambahan
                                Scom.ExecuteNonQuery()
                            End If
                        End If
                    End If
                Else  'cek kolom dl baru tabel
                    Scom.CommandText = "Select column_type From Information_schema.Columns Where TABLE_SCHEMA='pos' AND Table_Name='" & tableName.ToUpper & "' And Column_Name='" & columnName.ToUpper & "' "
                    If Scom.ExecuteScalar & "" = "" Then
                        Scom.CommandText = "SELECT COUNT(*) FROM Information_schema.tables WHERE TABLE_SCHEMA='pos' AND Table_Name='" & tableName.ToUpper & "'; "
                        If Scom.ExecuteScalar > 0 Then
                            Scom.CommandText = "ALTER TABLE " & tableName.ToUpper & " " & syntax
                            Scom.ExecuteNonQuery()

                            If syntaxTambahan & "" <> "" Then
                                Scom.CommandText = "" & syntaxTambahan
                                Scom.ExecuteNonQuery()
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                Dim sw As New IO.StreamWriter(Application.StartupPath & "\DEBUGPOSIDM.TXT", False)
                sw.Write("Error alter table : " & tableName & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & Scom.CommandText)
                sw.Flush()
                sw.Close()
                Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "ExecuteNonQuery", Mcon)
            End Try
        Catch ex As Exception
            Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "ExecuteNonQuery", Mcon)
        End Try
    End Sub
End Class
