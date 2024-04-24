Imports System.Threading
Imports MySql.Data.MySqlClient
Imports IDM.Fungsi
Imports System.Net.NetworkInformation

Public Class ClsInterfaceWifi

    Public Shared userMk As String = ""
    Public Shared passMk As String = ""
    Public Shared hostMk As String = ""
    Public Shared portMk As String = ""
    Public Shared status As Boolean = True
    Public Shared portwdcp As String = ""

    Public Shared Function initialInterfaceWifi() As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Try
            If Conn.State <> ConnectionState.Open Then
                Conn.Open()
            End If

            Mcom.CommandText = "SELECT `DESC` FROM CONST WHERE RKEY = 'AWK'"
            If Mcom.ExecuteScalar = "" Then
                Mcom.CommandText = "INSERT IGNORE INTO const (`rkey`, `desc`, `period`, `jenis`) VALUES('AWK', 'Flag Main RB', '2023-03-27', 'N')"
                Mcom.ExecuteNonQuery()
            End If
            Mcom.CommandText = "SELECT `JENIS` FROM CONST WHERE RKEY = 'AWK'"

            If Mcom.ExecuteScalar = "N" Then
            Else
                Dim wky As String()

                '    Public Shared userMk As String = "admin"
                'Public Shared passMk As String = "admin"
                'Public Shared hostMk As String = "172.31.31.14"
                'Public Shared portMk As String = "8278"
                Mcom.CommandText = "SELECT `DESC` FROM CONST WHERE RKEY = 'WKY'"
                If Mcom.ExecuteScalar = "" Then
                    Mcom.CommandText = "INSERT IGNORE INTO const (`rkey`, `desc`, `period`) VALUES('WKY', '172.31.31.14;8278;admin;admin', '2022-10-10')"
                    Mcom.ExecuteNonQuery()
                End If
                Mcom.CommandText = "SELECT `DESC` FROM CONST WHERE RKEY = 'WKY'"
                wky = Mcom.ExecuteScalar.ToString.Split(";")

                hostMk = wky(0)
                portMk = wky(1)
                userMk = wky(2)
                passMk = wky(3)


                status = cekKoneksi()

                If status = False Then
                    'Console.WriteLine("1")

                    FormMain.ckstatus = False
                    'Console.WriteLine("2")
                Else
                    panggilInterface()
                End If


            End If



        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
        End Try
        Return True
    End Function


    Public Shared Function getPortWDCP() As String
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Try


            If Conn.State <> ConnectionState.Open Then
                Conn.Open()
            End If
            Mcom.CommandText = "SELECT `DESC` FROM CONST WHERE RKEY = 'AWK'"
            If Mcom.ExecuteScalar = "" Then
                Mcom.CommandText = "INSERT IGNORE INTO const (`rkey`, `desc`, `period`, `jenis`) VALUES('AWK', 'Flag Main RB', '2023-03-27', 'N')"
                Mcom.ExecuteNonQuery()
            End If
            Mcom.CommandText = "SELECT `JENIS` FROM CONST WHERE RKEY = 'AWK'"

            If Mcom.ExecuteScalar = "N" Then
                portwdcp = 9400
            Else
                Thread.Sleep(2000)
                Dim filepath As String = Application.StartupPath & "\portwdcp.txt"
                Dim x As Short = 0
ulang:
                Try
                    x = x + 1
                    If System.IO.File.Exists(filepath) Then
                        Dim filereader As New System.IO.StreamReader(Application.StartupPath & "\portwdcp.txt")
                        portwdcp = filereader.ReadLine
                        portwdcp = Utility.AES128_Decrypt(portwdcp, "indomar3t")
                    End If

                Catch ex As Exception
                    TraceLog("Proses ulang getPortWDCP (" & x & ")")
                    GoTo ulang
                End Try
            End If

        Catch ex As Exception

        End Try
        Return portwdcp
    End Function

    Public Shared Function cekKoneksi() As Boolean
        Dim ping As New Ping
        Dim rep As PingReply = ping.Send(hostMk)
        Dim ret As Boolean = True
        'Console.WriteLine(rep.Status.ToString.ToLower.Trim)
        If rep.Status.ToString.ToLower.Trim <> "success" Then
            ret = False
        End If
        'Console.WriteLine(ret)
        Return ret
    End Function
    Public Shared Sub panggilInterface()
        Try
            FormMain.panggil.FileName = Application.StartupPath & "\InterfaceWIFI.exe"
            FormMain.panggil.Arguments = "-Panggil_Interface_WIFI -" & hostMk & " -" & portMk & " -" & userMk & " -" & passMk

            Console.WriteLine(FormMain.panggil.Arguments)
            Try
                Process.Start(FormMain.panggil)

            Catch ex As Exception
                FormMain.Close()

            End Try
        Catch ex As Exception
            MsgBox("Komputer belum terhubung dengan RB !" & ex.Message & vbCrLf & ex.StackTrace)
            FormMain.Close()

        End Try
    End Sub
End Class
