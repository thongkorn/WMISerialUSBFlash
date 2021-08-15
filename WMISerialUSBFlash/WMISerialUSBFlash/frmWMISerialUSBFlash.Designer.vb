<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWMISerialUSBFlash
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
        Me.lvwData = New System.Windows.Forms.ListView()
        Me.SuspendLayout()
        '
        'lvwData
        '
        Me.lvwData.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvwData.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lvwData.Location = New System.Drawing.Point(0, 0)
        Me.lvwData.Name = "lvwData"
        Me.lvwData.Size = New System.Drawing.Size(667, 304)
        Me.lvwData.TabIndex = 2
        Me.lvwData.UseCompatibleStateImageBehavior = False
        '
        'frmWMISerialUSBFlash
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(667, 304)
        Me.Controls.Add(Me.lvwData)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmWMISerialUSBFlash"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "WMI Serial USB Flash Drive with Added or Removed Event. (Auto Detect)"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lvwData As System.Windows.Forms.ListView

End Class
