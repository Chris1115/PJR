Imports MySql.Data.MySqlClient

Public Class FrmRKCB
    Private Sub FrmRKCB_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Rtn As New Boolean
        Dim DtCP As New DataTable
        Dim sqltampung As String = ""
        Dim ds1 As New DataSet
        Dim Rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Rpt = New rptListingRKCB
        Dim user As String = ""

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Mcom.CommandText = "SELECT KASIR_NAME FROM INITIAL WHERE STATION = '" & IDM.Fungsi.Get_Station() & "' AND TANGGAL = CURDATE()"
            user = Mcom.ExecuteScalar
            If Mcom.ExecuteScalar <> "" Then
                Mcom.CommandText = "SHOW TABLES LIKE 'TEMP_DRAFT_KESEGARAN'"
                If Mcom.ExecuteScalar = "" Then
                    MsgBox("Mohon tekan Menu Scan Barang terlebih dahulu !")
                    Me.Close()
                    Exit Sub
                Else

                    sqltampung = "SELECT a.prdcd as `PLU`,singkatan AS `DESC`,status_retur AS `STATUSPTRT`,tanggal_Exp_terakhir AS `ExpDate`,b.NAMA_RAK AS MODIS FROM TEMP_DRAFT_KESEGARAN a 
                            LEFT JOIN rak b ON a.prdcd = b.prdcd
                            WHERE a.prdcd IS NOT NULL AND singkatan IS NOT NULL AND status_retur IS NOT NULL AND tanggal_Exp_terakhir IS NOT NULL AND a.nama_rak IS NOT NULL
                           GROUP BY  b.prdcd,tanggal_Exp_terakhir,b.nama_rak

                            ORDER BY noshelf,kirikanan"

                    IDM.Fungsi.TraceLog("WDCP_FrmRKCB_Load : " & sqltampung)
                    Dim Mda As New MySqlDataAdapter(sqltampung, Conn)

                    Mda.Fill(DtCP)

                    Rpt.SetDataSource(DtCP)

                    Rpt.SetParameterValue("toko", FormMain.Toko.Kode & " - " & FormMain.Toko.Nama)
                    Rpt.SetParameterValue("user", user)



                    CrystalReportViewer1.ReportSource = Rpt

                End If

            End If
        Catch ex As Exception
            IDM.Fungsi.ShowError("Error WDCP ", ex.Message & ex.StackTrace)
            IDM.Fungsi.TraceLog("WDCP : " & ex.Message & ex.StackTrace)

        Finally
            Conn.Close()
        End Try


    End Sub
End Class