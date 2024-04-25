Imports IDM.Fungsi
Imports CrystalDecisions.Shared
Imports MySql.Data.MySqlClient

Public Class FrmRptDCP

    Public KodeGudang As String
    Private Scon As MySqlConnection
    Private Sdap As New MySqlDataAdapter
    Private Scom As New MySqlCommand
    Private Mrdr As MySqlDataReader
    Public modis As String

    Private Sub FrmRptDCP_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If FormMain.isBazar Or FormMain.isExpiredDate Then
            Dim cfileSO As String
            Dim Rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Scon = ClsConnection.GetConnection.Clone
            Sdap = New MySqlDataAdapter("", Scon)
            Scom = New MySqlCommand("", Scon)
            Dim dtCP As New DataTable

            If FormMain.isBazar Then
                cfileSO = "SZ" & Format(Date.Now, "yyMM") & Mid(FormMain.Toko.Kode, 1, 1)
            Else
                cfileSO = "SE" & Format(Date.Now, "yyMM") & Mid(FormMain.Toko.Kode, 1, 1)
            End If

            Rpt = New rptBarangSOBazar

            If FormMain.isBazar Then
                Scom.CommandText = " SELECT PRDCD, SINGKAT as DESKRIPSI, BULAN_EXP AS BULANEXP,(CAST(TTL AS CHAR(6))) AS QTY FROM " & cfileSO
            Else
                Scom.CommandText = " SELECT PRDCD, DESKRIPSI, BULAN_EXP AS BULANEXP,(CAST(TTL AS CHAR(6))) AS QTY FROM " & cfileSO
            End If

            If FormMain.isBazar Then
                Scom.CommandText &= " WHERE RECID = 'P';"
            Else
                Scom.CommandText &= " WHERE STATUS = 'Mendekati Expired' AND TGL_AKHIR >= NOW();"
            End If

            TraceLog("FrmRptDCP_Load-Q1: " & Scom.CommandText)

            Sdap = New MySqlDataAdapter(Scom)
            Sdap.Fill(dtCP)

            Rpt.SetDataSource(dtCP)

            Rpt.SetParameterValue("Tgl", Date.Now)

            If FormMain.isBazar Then
                Rpt.SetParameterValue("rptTitle", "LIST ITEM YANG AKAN DI BAZAR")
            Else
                Rpt.SetParameterValue("rptTitle", "LIST ITEM YANG MENDEKATI EXPIRED")
            End If

            CRV.ReportSource = Rpt
            CRV.Zoom(1)
            Rpt.PrintToPrinter(1, False, 0, 0)

            If FormMain.isBazar Then
                FormMain.isBazar = False
            Else
                FormMain.isExpiredDate = False
            End If

        ElseIf FormMain.isCekDisplay = True Then
            Dim cfileSO As String = ""
            Dim Rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Scon = ClsConnection.GetConnection.Clone
            Sdap = New MySqlDataAdapter("", Scon)
            Scom = New MySqlCommand("", Scon)
            Dim dtCP As New DataTable

            Rpt = New rptListingDisplayProduk
            Scom.CommandText = "SELECT kodemodis AS Modis,CAST(CONCAT(norak,'-',norak,'-',kirikanan) AS CHAR) AS kode_rak, plu, 
                                deskripsi, kap_disp AS kap_dis, qty_disp AS qty_rak, IF((qty_disp - kap_disp)> 0,0,(qty_disp - kap_disp)) AS qty_kebutuhan FROM cekdisplay
                                WHERE DATE(tglscan) = CURDATE() and NIK = '" & FormMain.user_cekdisplay & "' AND kodemodis = '" & modis & "' AND PLU IN (SELECT PLU FROM TEMP_CEKDISPLAY)"
            TraceLog("ListingDisplay_Cetak : " & Scom.CommandText)
            Sdap = New MySqlDataAdapter(Scom)
            Sdap.Fill(dtCP)

            Rpt.SetDataSource(dtCP)

            Rpt.SetParameterValue("toko", FormMain.Toko.Kode)
            Rpt.SetParameterValue("user", FormMain.user_cekdisplay)
            CRV.ReportSource = Rpt
            CRV.Zoom(1)
            FormMain.isCekDisplay = False

        Else
            Try
                ''PROGRAM CEK BARANG BISA PILIH PER DOCNO (NO NPB) - 438/03-23/E/PMO
                Dim kodeDC As String = KodeGudang.Split("-")(0).ToString
                Dim DOCNO As String = KodeGudang.Split("-")(1).ToString

                Dim cBPB As New ClsBPBController
                Dim Rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
                Rpt = New RptDCPToko

                Dim DtReport As New DataTable
                Dim DpdID As String
                'PROGRAM CEK BARANG BISA PILIH PER DOCNO (NO NPB) - 438/03-23/E/PMO
                DtReport = cBPB.GetDataReportBPB(kodeDC, DOCNO)
                'DtReport = cBPB.GetDataReportBPB(KodeGudang)
                If DtReport.Rows.Count > 0 Then
                    For i As Integer = 0 To DtReport.Rows.Count - 1
                        DpdID = DtReport.Rows(i).Item("DPDID") & ""
                        Select Case DpdID
                            Case "="
                                DtReport.Rows(i).Item("Ket") = "Barang Sama Dengan NPB"
                            Case "+"
                                DtReport.Rows(i).Item("Ket") = "Barang Lebih dari NPB"
                            Case "-"
                                DtReport.Rows(i).Item("Ket") = "Barang Kurang dari NPB"
                            Case ""
                                DtReport.Rows(i).Item("Ket") = "Barang Tidak di Scan"
                        End Select
                    Next
                End If
                Rpt.SetDataSource(DtReport)

                Dim paramFields As New ParameterFields
                Dim paramField As New ParameterField
                Dim discreteVal As New ParameterDiscreteValue

                paramField = New ParameterField
                paramField.ParameterFieldName = "cTok"
                discreteVal = New ParameterDiscreteValue
                discreteVal.Value = FormMain.Toko.Kode
                paramField.CurrentValues.Add(discreteVal)
                paramFields.Add(paramField)

                paramField = New ParameterField
                paramField.ParameterFieldName = "cLok"
                discreteVal = New ParameterDiscreteValue
                discreteVal.Value = FormMain.Toko.Lokasi
                paramField.CurrentValues.Add(discreteVal)
                paramFields.Add(paramField)

                paramField = New ParameterField
                paramField.ParameterFieldName = "Gudang"
                discreteVal = New ParameterDiscreteValue
                discreteVal.Value = KodeGudang
                paramField.CurrentValues.Add(discreteVal)
                paramFields.Add(paramField)

                CRV.ParameterFieldInfo = paramFields

                CRV.ReportSource = Rpt
                CRV.ShowPrintButton = True
                CRV.ShowExportButton = False
                CRV.Zoom(75)
                CRV.Refresh()

            Catch ex As Exception
                MsgBox(ex.Message & vbCrLf & ex.StackTrace, MessageBoxButtons.OK)
                Me.Close()
                Exit Sub
            End Try
        End If

    End Sub

End Class