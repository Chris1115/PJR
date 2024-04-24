<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmRegistPJR
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.dgvJadwalPJR = New System.Windows.Forms.DataGridView()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.cbHariBuka = New System.Windows.Forms.ComboBox()
        Me.btnHapusJadwal = New System.Windows.Forms.Button()
        Me.cmbNorak = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.TextBoxTo = New System.Windows.Forms.TextBox()
        Me.btnCari_byNIK = New System.Windows.Forms.Button()
        Me.btnTambahPJR = New System.Windows.Forms.Button()
        Me.TextBoxFrom = New System.Windows.Forms.TextBox()
        Me.txtNamaPersonil = New System.Windows.Forms.TextBox()
        Me.cmbHari = New System.Windows.Forms.ComboBox()
        Me.txtNamaModis = New System.Windows.Forms.TextBox()
        Me.cmbNIK = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cmbModis = New System.Windows.Forms.ComboBox()
        Me.lblHeaderPJR = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.btnSimpanPJR = New System.Windows.Forms.Button()
        CType(Me.dgvJadwalPJR, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvJadwalPJR
        '
        Me.dgvJadwalPJR.AllowUserToResizeColumns = False
        Me.dgvJadwalPJR.AllowUserToResizeRows = False
        Me.dgvJadwalPJR.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvJadwalPJR.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvJadwalPJR.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvJadwalPJR.Location = New System.Drawing.Point(274, 102)
        Me.dgvJadwalPJR.MinimumSize = New System.Drawing.Size(500, 326)
        Me.dgvJadwalPJR.Name = "dgvJadwalPJR"
        Me.dgvJadwalPJR.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvJadwalPJR.Size = New System.Drawing.Size(842, 326)
        Me.dgvJadwalPJR.TabIndex = 108
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(25, 102)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(70, 44)
        Me.Label11.TabIndex = 107
        Me.Label11.Text = "Hari Buka Toko"
        '
        'cbHariBuka
        '
        Me.cbHariBuka.Enabled = False
        Me.cbHariBuka.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbHariBuka.FormattingEnabled = True
        Me.cbHariBuka.Location = New System.Drawing.Point(116, 102)
        Me.cbHariBuka.Name = "cbHariBuka"
        Me.cbHariBuka.Size = New System.Drawing.Size(114, 23)
        Me.cbHariBuka.TabIndex = 106
        '
        'btnHapusJadwal
        '
        Me.btnHapusJadwal.BackColor = System.Drawing.SystemColors.Control
        Me.btnHapusJadwal.Enabled = False
        Me.btnHapusJadwal.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnHapusJadwal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnHapusJadwal.Location = New System.Drawing.Point(38, 482)
        Me.btnHapusJadwal.Name = "btnHapusJadwal"
        Me.btnHapusJadwal.Size = New System.Drawing.Size(90, 29)
        Me.btnHapusJadwal.TabIndex = 105
        Me.btnHapusJadwal.Text = "Hapus"
        Me.btnHapusJadwal.UseVisualStyleBackColor = False
        '
        'cmbNorak
        '
        Me.cmbNorak.Enabled = False
        Me.cmbNorak.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbNorak.FormattingEnabled = True
        Me.cmbNorak.Location = New System.Drawing.Point(116, 402)
        Me.cmbNorak.Name = "cmbNorak"
        Me.cmbNorak.Size = New System.Drawing.Size(114, 23)
        Me.cmbNorak.TabIndex = 104
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(165, 438)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(28, 21)
        Me.Label1.TabIndex = 103
        Me.Label1.Text = "s/d"
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(25, 439)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(70, 21)
        Me.Label9.TabIndex = 102
        Me.Label9.Text = "No. Shelf "
        '
        'TextBoxTo
        '
        Me.TextBoxTo.Enabled = False
        Me.TextBoxTo.Location = New System.Drawing.Point(199, 439)
        Me.TextBoxTo.Name = "TextBoxTo"
        Me.TextBoxTo.Size = New System.Drawing.Size(31, 20)
        Me.TextBoxTo.TabIndex = 101
        '
        'btnCari_byNIK
        '
        Me.btnCari_byNIK.Location = New System.Drawing.Point(38, 482)
        Me.btnCari_byNIK.Name = "btnCari_byNIK"
        Me.btnCari_byNIK.Size = New System.Drawing.Size(75, 23)
        Me.btnCari_byNIK.TabIndex = 91
        Me.btnCari_byNIK.Text = "Cari"
        Me.btnCari_byNIK.UseVisualStyleBackColor = True
        Me.btnCari_byNIK.Visible = False
        '
        'btnTambahPJR
        '
        Me.btnTambahPJR.BackColor = System.Drawing.SystemColors.Control
        Me.btnTambahPJR.Enabled = False
        Me.btnTambahPJR.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnTambahPJR.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTambahPJR.Location = New System.Drawing.Point(156, 482)
        Me.btnTambahPJR.Name = "btnTambahPJR"
        Me.btnTambahPJR.Size = New System.Drawing.Size(90, 29)
        Me.btnTambahPJR.TabIndex = 92
        Me.btnTambahPJR.Text = "Tambah"
        Me.btnTambahPJR.UseVisualStyleBackColor = False
        '
        'TextBoxFrom
        '
        Me.TextBoxFrom.Enabled = False
        Me.TextBoxFrom.Location = New System.Drawing.Point(117, 438)
        Me.TextBoxFrom.Name = "TextBoxFrom"
        Me.TextBoxFrom.Size = New System.Drawing.Size(32, 20)
        Me.TextBoxFrom.TabIndex = 100
        '
        'txtNamaPersonil
        '
        Me.txtNamaPersonil.Enabled = False
        Me.txtNamaPersonil.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNamaPersonil.Location = New System.Drawing.Point(116, 178)
        Me.txtNamaPersonil.Multiline = True
        Me.txtNamaPersonil.Name = "txtNamaPersonil"
        Me.txtNamaPersonil.Size = New System.Drawing.Size(114, 57)
        Me.txtNamaPersonil.TabIndex = 90
        '
        'cmbHari
        '
        Me.cmbHari.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbHari.FormattingEnabled = True
        Me.cmbHari.Location = New System.Drawing.Point(116, 256)
        Me.cmbHari.Name = "cmbHari"
        Me.cmbHari.Size = New System.Drawing.Size(114, 23)
        Me.cmbHari.TabIndex = 94
        '
        'txtNamaModis
        '
        Me.txtNamaModis.Enabled = False
        Me.txtNamaModis.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNamaModis.Location = New System.Drawing.Point(117, 332)
        Me.txtNamaModis.Multiline = True
        Me.txtNamaModis.Name = "txtNamaModis"
        Me.txtNamaModis.Size = New System.Drawing.Size(114, 57)
        Me.txtNamaModis.TabIndex = 99
        '
        'cmbNIK
        '
        Me.cmbNIK.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append
        Me.cmbNIK.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbNIK.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbNIK.FormattingEnabled = True
        Me.cmbNIK.Location = New System.Drawing.Point(117, 146)
        Me.cmbNIK.Name = "cmbNIK"
        Me.cmbNIK.Size = New System.Drawing.Size(114, 23)
        Me.cmbNIK.TabIndex = 87
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(25, 402)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(70, 21)
        Me.Label3.TabIndex = 96
        Me.Label3.Text = "Nomor Rak"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(25, 178)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 21)
        Me.Label4.TabIndex = 89
        Me.Label4.Text = "Nama"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(25, 298)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(70, 21)
        Me.Label2.TabIndex = 95
        Me.Label2.Text = "Modis"
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(25, 256)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(70, 21)
        Me.Label6.TabIndex = 93
        Me.Label6.Text = "Hari"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(25, 334)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 21)
        Me.Label5.TabIndex = 98
        Me.Label5.Text = "Nama Modis"
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(25, 146)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(70, 21)
        Me.Label7.TabIndex = 88
        Me.Label7.Text = "NIK"
        '
        'cmbModis
        '
        Me.cmbModis.Enabled = False
        Me.cmbModis.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbModis.FormattingEnabled = True
        Me.cmbModis.Location = New System.Drawing.Point(117, 296)
        Me.cmbModis.Name = "cmbModis"
        Me.cmbModis.Size = New System.Drawing.Size(114, 23)
        Me.cmbModis.TabIndex = 97
        '
        'lblHeaderPJR
        '
        Me.lblHeaderPJR.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblHeaderPJR.AutoSize = True
        Me.lblHeaderPJR.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHeaderPJR.Location = New System.Drawing.Point(375, 9)
        Me.lblHeaderPJR.Name = "lblHeaderPJR"
        Me.lblHeaderPJR.Size = New System.Drawing.Size(267, 29)
        Me.lblHeaderPJR.TabIndex = 113
        Me.lblHeaderPJR.Text = "Registrasi Personil PJR"
        Me.lblHeaderPJR.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(452, 450)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(13, 13)
        Me.Label10.TabIndex = 117
        Me.Label10.Text = "()"
        Me.Label10.Visible = False
        '
        'Label8
        '
        Me.Label8.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(271, 488)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(194, 13)
        Me.Label8.TabIndex = 116
        Me.Label8.Text = "Harap tunggu... Proses sedang berjalan"
        Me.Label8.Visible = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(274, 450)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(460, 23)
        Me.ProgressBar1.TabIndex = 115
        '
        'btnSimpanPJR
        '
        Me.btnSimpanPJR.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSimpanPJR.BackColor = System.Drawing.SystemColors.Control
        Me.btnSimpanPJR.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSimpanPJR.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSimpanPJR.Location = New System.Drawing.Point(769, 450)
        Me.btnSimpanPJR.Name = "btnSimpanPJR"
        Me.btnSimpanPJR.Size = New System.Drawing.Size(90, 29)
        Me.btnSimpanPJR.TabIndex = 114
        Me.btnSimpanPJR.Text = "Simpan"
        Me.btnSimpanPJR.UseVisualStyleBackColor = False
        '
        'FrmRegistPJR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1128, 573)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.btnSimpanPJR)
        Me.Controls.Add(Me.lblHeaderPJR)
        Me.Controls.Add(Me.dgvJadwalPJR)
        Me.Controls.Add(Me.cmbModis)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.cbHariBuka)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnHapusJadwal)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cmbNorak)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmbNIK)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.txtNamaModis)
        Me.Controls.Add(Me.TextBoxTo)
        Me.Controls.Add(Me.cmbHari)
        Me.Controls.Add(Me.btnCari_byNIK)
        Me.Controls.Add(Me.txtNamaPersonil)
        Me.Controls.Add(Me.btnTambahPJR)
        Me.Controls.Add(Me.TextBoxFrom)
        Me.Name = "FrmRegistPJR"
        Me.Text = "FrmRegistPJR"
        CType(Me.dgvJadwalPJR, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Timer1 As Timer
    Friend WithEvents dgvJadwalPJR As DataGridView
    Friend WithEvents Label11 As Label
    Friend WithEvents cbHariBuka As ComboBox
    Friend WithEvents btnHapusJadwal As Button
    Friend WithEvents cmbNorak As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents TextBoxTo As TextBox
    Friend WithEvents btnCari_byNIK As Button
    Friend WithEvents btnTambahPJR As Button
    Friend WithEvents TextBoxFrom As TextBox
    Friend WithEvents txtNamaPersonil As TextBox
    Friend WithEvents cmbHari As ComboBox
    Friend WithEvents txtNamaModis As TextBox
    Friend WithEvents cmbNIK As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents cmbModis As ComboBox
    Friend WithEvents lblHeaderPJR As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents btnSimpanPJR As Button
End Class
