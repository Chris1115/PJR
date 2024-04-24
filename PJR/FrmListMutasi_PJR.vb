Imports MySql.Data.MySqlClient

Public Class FrmListMutasi_PJR
    Private Sub FrmListMutasi_PJR_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Madp As New MySqlDataAdapter("", Conn)
        Dim Rtn As New DataTable
        Dim Mcom As New MySqlCommand("", Conn)
        Dim jabatan As String = ""
        Dim cPJR As New ClsPJRController

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            jabatan = cPJR.getJabatanVirbacaprod

            If lblJudul.Text = "LIST PERUBAHAN MUTASI PERSONIL TOKO" Then

                Madp.SelectCommand.CommandText = "SELECT DISTINCT MENOIN AS NIK, MENAME AS NAMA, JABATAN, 'BELUM DIDAFTARKAN' AS KETERANGAN FROM SOPPAGENT.ABSPEGAWAIMST WHERE JABATAN IN (" & jabatan & ") 
                                                    AND MENOIN NOT IN (SELECT DISTINCT nik FROM JADWAL_PENANGGUNGJAWABRAK  ) AND PINJAMAN = 0"

                Madp.Fill(Rtn)

                Madp.SelectCommand.CommandText = "SELECT DISTINCT NIK, NAMA, JABATAN, 'MUTASI' AS KETERANGAN FROM JADWAL_PENANGGUNGJAWABRAK WHERE NIK NOT IN 
                                                (SELECT MENOIN FROM SOPPAGENT.ABSPEGAWAIMST WHERE JABATAN IN 
                                                (" & jabatan & ") AND PINJAMAN = 0 );"
                Madp.Fill(Rtn)


                DataGridView1.DataSource = Rtn

                DataGridView1.Columns(0).Width = 80
                DataGridView1.Columns(1).Width = 200
                DataGridView1.Columns(2).Width = 120
                DataGridView1.Columns(3).Width = 143


                DataGridView1.ReadOnly = True
                DataGridView1.Refresh()
            ElseIf lblJudul.Text = "LIST PERUBAHAN DATA MODISP TOKO" Then

                'Memo 447/cps/23
                'PJR hnya FJP=Y

                Madp.SelectCommand.CommandText = "SELECT NIK,NAMA, KODE_MODIS, MODIS AS KET_RAK, 'PENGURANGAN MODIS' AS KETERANGAN FROM JADWAL_PENANGGUNGJAWABRAK  WHERE kode_modis NOT IN (SELECT  DISTINCT kodemodis FROM RAK WHERE flagprod LIKE '%FJP=Y%') GROUP BY NIK,KODE_MODIS"

                Madp.Fill(Rtn)


                'Memo 447/cps/23
                'PJR hnya FJP=Y


                Madp.SelectCommand.CommandText = "SELECT '' AS NIK,'' AS NAMA,KODEMODIS AS KODE_MODIS, KET_RAK,'BELUM DIDAFTARKAN' AS KETERANGAN FROM RAK WHERE flagprod LIKE '%FJP=Y%' AND KODEMODIS  
                                                    NOT IN (SELECT DISTINCT kode_modis FROM JADWAL_PENANGGUNGJAWABRAK)"

                'Madp.SelectCommand.CommandText = "SELECT '' AS NIK,'' AS NAMA,KODEMODIS AS KODE_MODIS, KET_RAK,'BELUM DIDAFTARKAN' AS KETERANGAN FROM RAK WHERE flagprod NOT LIKE '%FJP=N%' AND KODEMODIS  
                '                                    NOT IN (SELECT DISTINCT kode_modis FROM JADWAL_PENANGGUNGJAWABRAK)"

                Madp.Fill(Rtn)


                DataGridView1.DataSource = Rtn
                DataGridView1.Columns(0).Width = 80
                DataGridView1.Columns(1).Width = 130
                DataGridView1.Columns(2).Width = 80
                DataGridView1.Columns(3).Width = 200
                DataGridView1.Columns(4).Width = 193

            End If


        Catch ex As Exception

        End Try

    End Sub

End Class