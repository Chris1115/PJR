<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmCPJR
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCPJR))
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.dtpTglAwal = New System.Windows.Forms.DateTimePicker()
        Me.lblTgl = New System.Windows.Forms.Label()
        Me.btnProses = New System.Windows.Forms.Button()
        Me.btnKeluar = New System.Windows.Forms.Button()
        Me.btnCetak = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.lblJenisLaporan = New System.Windows.Forms.Label()
        Me.cmbJenisLap = New System.Windows.Forms.ComboBox()
        Me.lblNIK = New System.Windows.Forms.Label()
        Me.cbNik = New System.Windows.Forms.ComboBox()
        Me.Panel1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        Me.CrystalReportViewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrystalReportViewer1.Cursor = System.Windows.Forms.Cursors.Default
        Me.CrystalReportViewer1.DisplayGroupTree = False
        Me.CrystalReportViewer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(0, 0)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.SelectionFormula = ""
        Me.CrystalReportViewer1.ShowCloseButton = False
        Me.CrystalReportViewer1.ShowPrintButton = False
        Me.CrystalReportViewer1.ShowRefreshButton = False
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(916, 507)
        Me.CrystalReportViewer1.TabIndex = 5
        Me.CrystalReportViewer1.ViewTimeSelectionFormula = ""
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.cbNik)
        Me.Panel1.Controls.Add(Me.lblNIK)
        Me.Panel1.Controls.Add(Me.cmbJenisLap)
        Me.Panel1.Controls.Add(Me.lblJenisLaporan)
        Me.Panel1.Controls.Add(Me.GroupBox3)
        Me.Panel1.Controls.Add(Me.lblTgl)
        Me.Panel1.Controls.Add(Me.dtpTglAwal)
        Me.Panel1.Location = New System.Drawing.Point(1, 398)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(915, 109)
        Me.Panel1.TabIndex = 4
        '
        'dtpTglAwal
        '
        Me.dtpTglAwal.Location = New System.Drawing.Point(116, 62)
        Me.dtpTglAwal.Name = "dtpTglAwal"
        Me.dtpTglAwal.Size = New System.Drawing.Size(208, 20)
        Me.dtpTglAwal.TabIndex = 0
        '
        'lblTgl
        '
        Me.lblTgl.Location = New System.Drawing.Point(30, 62)
        Me.lblTgl.Name = "lblTgl"
        Me.lblTgl.Size = New System.Drawing.Size(80, 20)
        Me.lblTgl.TabIndex = 1
        Me.lblTgl.Text = "Tanggal"
        '
        'btnProses
        '
        Me.btnProses.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnProses.BackColor = System.Drawing.SystemColors.Control
        Me.btnProses.Image = CType(resources.GetObject("btnProses.Image"), System.Drawing.Image)
        Me.btnProses.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnProses.Location = New System.Drawing.Point(6, 16)
        Me.btnProses.Name = "btnProses"
        Me.btnProses.Size = New System.Drawing.Size(60, 45)
        Me.btnProses.TabIndex = 0
        Me.btnProses.Text = "&Proses"
        Me.btnProses.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnProses.UseVisualStyleBackColor = False
        '
        'btnKeluar
        '
        Me.btnKeluar.BackColor = System.Drawing.SystemColors.Control
        Me.btnKeluar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnKeluar.Image = CType(resources.GetObject("btnKeluar.Image"), System.Drawing.Image)
        Me.btnKeluar.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnKeluar.Location = New System.Drawing.Point(138, 16)
        Me.btnKeluar.Name = "btnKeluar"
        Me.btnKeluar.Size = New System.Drawing.Size(60, 45)
        Me.btnKeluar.TabIndex = 2
        Me.btnKeluar.Text = "&Keluar"
        Me.btnKeluar.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnKeluar.UseVisualStyleBackColor = False
        '
        'btnCetak
        '
        Me.btnCetak.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCetak.BackColor = System.Drawing.SystemColors.Control
        Me.btnCetak.Enabled = False
        Me.btnCetak.Image = CType(resources.GetObject("btnCetak.Image"), System.Drawing.Image)
        Me.btnCetak.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnCetak.Location = New System.Drawing.Point(72, 16)
        Me.btnCetak.Name = "btnCetak"
        Me.btnCetak.Size = New System.Drawing.Size(60, 45)
        Me.btnCetak.TabIndex = 3
        Me.btnCetak.Text = "&Cetak"
        Me.btnCetak.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnCetak.UseVisualStyleBackColor = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.btnCetak)
        Me.GroupBox3.Controls.Add(Me.btnKeluar)
        Me.GroupBox3.Controls.Add(Me.btnProses)
        Me.GroupBox3.Location = New System.Drawing.Point(684, 9)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(206, 72)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        '
        'lblJenisLaporan
        '
        Me.lblJenisLaporan.Location = New System.Drawing.Point(30, 13)
        Me.lblJenisLaporan.Name = "lblJenisLaporan"
        Me.lblJenisLaporan.Size = New System.Drawing.Size(80, 20)
        Me.lblJenisLaporan.TabIndex = 4
        Me.lblJenisLaporan.Text = "Jenis Laporan"
        '
        'cmbJenisLap
        '
        Me.cmbJenisLap.FormattingEnabled = True
        Me.cmbJenisLap.Items.AddRange(New Object() {"Laporan Jadwal Penanggung Jawab Rak (PJR)", "Jadwal Penanggung Jawab Rak (PJR)", "Laporan Jadwal Penanggung Jawab Rak (PJR) dengan Estimasi Waktu (Menit)", "Laporan Jadwal Penanggung Jawab Rak (PJR) - Per Rak", "Laporan Jadwal Penanggung Jawab Rak (PJR) - Per Kary. Toko Idm", "Laporan Jadwal Penanggung Jawab Rak (PJR) - Per Tanggal", "Laporan Final Barang Dagangan Tidak Pajang"})
        Me.cmbJenisLap.Location = New System.Drawing.Point(116, 10)
        Me.cmbJenisLap.Name = "cmbJenisLap"
        Me.cmbJenisLap.Size = New System.Drawing.Size(432, 21)
        Me.cmbJenisLap.TabIndex = 5
        '
        'lblNIK
        '
        Me.lblNIK.Location = New System.Drawing.Point(30, 36)
        Me.lblNIK.Name = "lblNIK"
        Me.lblNIK.Size = New System.Drawing.Size(80, 20)
        Me.lblNIK.TabIndex = 6
        Me.lblNIK.Text = "NIK"
        Me.lblNIK.Visible = False
        '
        'cbNik
        '
        Me.cbNik.FormattingEnabled = True
        Me.cbNik.Location = New System.Drawing.Point(116, 35)
        Me.cbNik.Name = "cbNik"
        Me.cbNik.Size = New System.Drawing.Size(121, 21)
        Me.cbNik.TabIndex = 7
        Me.cbNik.Visible = False
        '
        'frmCPJR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(916, 507)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.CrystalReportViewer1)
        Me.Name = "frmCPJR"
        Me.Text = "frmCPJR"
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Private WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Panel1 As Panel
    Friend WithEvents cbNik As ComboBox
    Friend WithEvents lblNIK As Label
    Friend WithEvents cmbJenisLap As ComboBox
    Friend WithEvents lblJenisLaporan As Label
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents btnCetak As Button
    Friend WithEvents btnKeluar As Button
    Friend WithEvents btnProses As Button
    Friend WithEvents lblTgl As Label
    Friend WithEvents dtpTglAwal As DateTimePicker
End Class
