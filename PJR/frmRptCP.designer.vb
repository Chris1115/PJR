<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRptCP
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRptCP))
        Me.CRVCP = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.txtNik = New System.Windows.Forms.TextBox
        Me.lblUser = New System.Windows.Forms.Label
        Me.cmbJenisLap = New System.Windows.Forms.ComboBox
        Me.lblJenisLaporan = New System.Windows.Forms.Label
        Me.dtpTglAkhir = New System.Windows.Forms.DateTimePicker
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.lblTglAkhir = New System.Windows.Forms.Label
        Me.lblTglAwal = New System.Windows.Forms.Label
        Me.dtpTglAwal = New System.Windows.Forms.DateTimePicker
        Me.btnProses = New System.Windows.Forms.Button
        Me.btnKeluar = New System.Windows.Forms.Button
        Me.btnCetak = New System.Windows.Forms.Button
        Me.Panel1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'CRVCP
        '
        Me.CRVCP.ActiveViewIndex = -1
        Me.CRVCP.AutoSize = True
        Me.CRVCP.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CRVCP.DisplayGroupTree = False
        Me.CRVCP.Location = New System.Drawing.Point(0, 0)
        Me.CRVCP.Name = "CRVCP"
        Me.CRVCP.SelectionFormula = ""
        Me.CRVCP.ShowCloseButton = False
        Me.CRVCP.ShowExportButton = False
        Me.CRVCP.ShowGotoPageButton = False
        Me.CRVCP.ShowGroupTreeButton = False
        Me.CRVCP.ShowPrintButton = False
        Me.CRVCP.ShowRefreshButton = False
        Me.CRVCP.Size = New System.Drawing.Size(836, 469)
        Me.CRVCP.TabIndex = 0
        Me.CRVCP.ViewTimeSelectionFormula = ""
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.txtNik)
        Me.Panel1.Controls.Add(Me.lblUser)
        Me.Panel1.Controls.Add(Me.cmbJenisLap)
        Me.Panel1.Controls.Add(Me.lblJenisLaporan)
        Me.Panel1.Controls.Add(Me.dtpTglAkhir)
        Me.Panel1.Controls.Add(Me.GroupBox3)
        Me.Panel1.Controls.Add(Me.lblTglAkhir)
        Me.Panel1.Controls.Add(Me.lblTglAwal)
        Me.Panel1.Controls.Add(Me.dtpTglAwal)
        Me.Panel1.Location = New System.Drawing.Point(0, 475)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(836, 124)
        Me.Panel1.TabIndex = 1
        '
        'txtNik
        '
        Me.txtNik.Location = New System.Drawing.Point(115, 38)
        Me.txtNik.Name = "txtNik"
        Me.txtNik.Size = New System.Drawing.Size(166, 20)
        Me.txtNik.TabIndex = 7
        Me.txtNik.Visible = False
        '
        'lblUser
        '
        Me.lblUser.Location = New System.Drawing.Point(30, 38)
        Me.lblUser.Name = "lblUser"
        Me.lblUser.Size = New System.Drawing.Size(80, 20)
        Me.lblUser.TabIndex = 6
        Me.lblUser.Text = "NIK"
        Me.lblUser.Visible = False
        '
        'cmbJenisLap
        '
        Me.cmbJenisLap.FormattingEnabled = True
        Me.cmbJenisLap.Items.AddRange(New Object() {"Laporan Item Tidak Terdisplay", "Laporan Trend Item Tidak Terdisplay", "Listing Periode Retur", "Rekapitulasi Laporan RLBTD"})
        Me.cmbJenisLap.Location = New System.Drawing.Point(116, 10)
        Me.cmbJenisLap.Name = "cmbJenisLap"
        Me.cmbJenisLap.Size = New System.Drawing.Size(252, 21)
        Me.cmbJenisLap.TabIndex = 5
        '
        'lblJenisLaporan
        '
        Me.lblJenisLaporan.Location = New System.Drawing.Point(30, 13)
        Me.lblJenisLaporan.Name = "lblJenisLaporan"
        Me.lblJenisLaporan.Size = New System.Drawing.Size(80, 20)
        Me.lblJenisLaporan.TabIndex = 4
        Me.lblJenisLaporan.Text = "Jenis Laporan"
        '
        'dtpTglAkhir
        '
        Me.dtpTglAkhir.Location = New System.Drawing.Point(116, 93)
        Me.dtpTglAkhir.Name = "dtpTglAkhir"
        Me.dtpTglAkhir.Size = New System.Drawing.Size(165, 20)
        Me.dtpTglAkhir.TabIndex = 3
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.btnCetak)
        Me.GroupBox3.Controls.Add(Me.btnKeluar)
        Me.GroupBox3.Controls.Add(Me.btnProses)
        Me.GroupBox3.Location = New System.Drawing.Point(605, 24)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(206, 72)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        '
        'lblTglAkhir
        '
        Me.lblTglAkhir.Location = New System.Drawing.Point(30, 93)
        Me.lblTglAkhir.Name = "lblTglAkhir"
        Me.lblTglAkhir.Size = New System.Drawing.Size(80, 20)
        Me.lblTglAkhir.TabIndex = 2
        Me.lblTglAkhir.Text = "Tanggal Akhir"
        '
        'lblTglAwal
        '
        Me.lblTglAwal.Location = New System.Drawing.Point(30, 64)
        Me.lblTglAwal.Name = "lblTglAwal"
        Me.lblTglAwal.Size = New System.Drawing.Size(80, 20)
        Me.lblTglAwal.TabIndex = 1
        Me.lblTglAwal.Text = "Tanggal Awal"
        '
        'dtpTglAwal
        '
        Me.dtpTglAwal.Location = New System.Drawing.Point(116, 64)
        Me.dtpTglAwal.Name = "dtpTglAwal"
        Me.dtpTglAwal.Size = New System.Drawing.Size(165, 20)
        Me.dtpTglAwal.TabIndex = 0
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
        'frmRptCP
        '
        Me.AcceptButton = Me.btnProses
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnKeluar
        Me.ClientSize = New System.Drawing.Size(836, 598)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.CRVCP)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frmRptCP"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cetak Laporan Planogram"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CRVCP As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents btnKeluar As System.Windows.Forms.Button
    Friend WithEvents btnProses As System.Windows.Forms.Button
    Friend WithEvents cmbJenisLap As System.Windows.Forms.ComboBox
    Friend WithEvents lblJenisLaporan As System.Windows.Forms.Label
    Friend WithEvents dtpTglAkhir As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblTglAkhir As System.Windows.Forms.Label
    Friend WithEvents lblTglAwal As System.Windows.Forms.Label
    Friend WithEvents dtpTglAwal As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnCetak As System.Windows.Forms.Button
    Friend WithEvents lblUser As System.Windows.Forms.Label
    Friend WithEvents txtNik As System.Windows.Forms.TextBox
End Class
