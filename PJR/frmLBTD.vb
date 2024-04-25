Imports PJR.clsFinger
Imports PJR.ClsPJRController

Public Class frmLBTD
    Public Shared nik As String = ""
    Private Sub frmLBTD_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim cpjr As New ClsPJRController
        Dim dt As New DataTable
        Dim dtHari As New DataTable
        Dim rakpjr As String = ""
        rakpjr = cpjr.getConstRakPJR
        Dim NikToko As String = ""
        nik = cpjr.getConstNIKPJR

        dt = cpjr.HASILBTD(Now.ToString("yyyy-MM-dd"), nik, rakpjr.Split("-")(0).Trim, rakpjr.Split("-")(2).Trim)

        dgvHASILLBTD.DataSource = dt


        dgvHASILLBTD.ReadOnly = True
        dgvHASILLBTD.Refresh()
    End Sub

    Private Sub btnAprvLBTD_Click(sender As Object, e As EventArgs) Handles btnAprvLBTD.Click
        Dim Result As DialogResult
        Dim cPJR As New ClsPJRController
        Dim tanggal As String
        Dim rakpjr As String = ""
        Dim norak As String = ""
        Dim nik As String
        Dim tes As String() = {"a", "b", "c"}

        nik = cPJR.getConstNIKPJR
        rakpjr = cPJR.getConstRakPJR.Split("-")(0).Trim
        tanggal = cPJR.getConstRakPJR.Split("-")(1).Trim
        norak = cPJR.getConstRakPJR.Split("-")(2).Trim

        If cPJR.cekApproveLBTD(tanggal.Replace("/", "-"), nik, rakpjr) = False Then
            MsgBox("Rak " & rakpjr & " telah di Approve !", MsgBoxStyle.Information)
            Me.Close()

        Else
            Result = MessageBox.Show("Apakah Anda yakin akan Approve Hasil LBTD?", "Approval Hasil LBTD Selesai..", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If Result = Windows.Forms.DialogResult.Yes Then
                MsgBox("Proses Approval Hasil LBTD menggunakan Scan FINGER", MsgBoxStyle.Information)
                tes = Panggil_CekFingerprintV3("frmLBTD", "WDCP_PJR 2")
                If tes(0) = "" And tes(1) = "" And tes(2) = "" Then
                    Debug.WriteLine("Password Salah")
                Else
                    ConstNIKPJR(tes(2).Split("|")(0).Trim)

                    cPJR.approveLBTD("Y", tanggal.Replace("/", "-"), nik, rakpjr, norak)
                    MsgBox("Berhasil Approve")

                    Me.Close()
                End If
            Else
                MsgBox("Proses Approval Tolak PJR menggunakan Scan FINGER", MsgBoxStyle.Information)
                tes = Panggil_CekFingerprintV3("frmLBTD", "WDCP_PJR 2")
                If tes(0) = "" And tes(1) = "" And tes(2) = "" Then
                    Debug.WriteLine("Password Salah")
                Else
                    ConstNIKPJR(tes(2).Split("|")(0).Trim)

                    cPJR.approveLBTD("N", tanggal.Replace("/", "-"), nik, rakpjr, norak)
                    MsgBox("Berhasil Tolak")
                    Me.Close()

                End If
            End If
        End If

    End Sub
End Class