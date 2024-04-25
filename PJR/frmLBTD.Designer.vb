<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLBTD
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
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.dgvHASILLBTD = New System.Windows.Forms.DataGridView()
        Me.btnAprvLBTD = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        CType(Me.dgvHASILLBTD, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.dgvHASILLBTD)
        Me.Panel1.Location = New System.Drawing.Point(12, 54)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(570, 96)
        Me.Panel1.TabIndex = 3
        '
        'dgvHASILLBTD
        '
        Me.dgvHASILLBTD.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvHASILLBTD.Location = New System.Drawing.Point(3, 3)
        Me.dgvHASILLBTD.Name = "dgvHASILLBTD"
        Me.dgvHASILLBTD.Size = New System.Drawing.Size(562, 90)
        Me.dgvHASILLBTD.TabIndex = 0
        '
        'btnAprvLBTD
        '
        Me.btnAprvLBTD.Location = New System.Drawing.Point(521, 168)
        Me.btnAprvLBTD.Name = "btnAprvLBTD"
        Me.btnAprvLBTD.Size = New System.Drawing.Size(75, 23)
        Me.btnAprvLBTD.TabIndex = 5
        Me.btnAprvLBTD.Text = "Approve"
        Me.btnAprvLBTD.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(204, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(228, 20)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "HASIL TINDAK LANJUT LBTD"
        '
        'frmLBTD
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(608, 212)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.btnAprvLBTD)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmLBTD"
        Me.Text = "frmLBTD"
        Me.Panel1.ResumeLayout(False)
        CType(Me.dgvHASILLBTD, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents dgvHASILLBTD As DataGridView
    Friend WithEvents btnAprvLBTD As Button
    Friend WithEvents Label1 As Label
End Class
