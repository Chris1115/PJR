Imports MySql.Data.MySqlClient

Public Class ClsConnection
    'Private MasterConn As MySqlConnection = Nothing
    Private Shared MasterConn As New MySqlConnection
    Private Server As String = ""
    Private UserId As String
    Private Password As String
    Private Database As String
    Private Port As String
    Dim ConnString As String = ""
    Dim Baca As New Regedit.Reg
    Public Shared isSector As Boolean = True
    Public Shared sector As New IDM.Sector
    Public Shared MyKey As String = "89AE46EB70BC2A7B6A1BB141F34C0BC5"

    Public Sub New()
        Dim Scon As New MySqlConnection

        Server = Baca.BacaRegistry("SOFTWARE\Indomaret\POS.NET\Database", "server")
        Port = Baca.BacaRegistry("SOFTWARE\Indomaret\POS.NET\Database", "port")
        Database = "pos"
        UserId = "root"

        Password = "WWukix1wU5QjZo2bVL6lKF/tofYCfQkmQ=MlNO1dHvF4"

        ConnString = "server=" + Server + ";" &
                     "port=" + Port + ";" &
                     "pooling=true;" &
                     "user id=" + UserId + ";" &
                     "password=" + Password + ";" &
                     "connection timeout=15;" &
                     "database=" + Database + ";" &
                     "Allow User Variables=True"
        Try
            'wisnu
            If isSector Then
                Scon = sector.GetVersionV2(MyKey, Application.StartupPath & "\PJR.exe".ToUpper, "kasir")
            Else
                Scon = New MySqlConnection(ConnString)
            End If

            If Not IsNothing(Scon) Then
                MasterConn = Scon.Clone
            End If

        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
        End Try

    End Sub
    Public Shared ReadOnly Property GetConnection() As MySqlConnection
        Get
            Return MasterConn
        End Get
    End Property

End Class
