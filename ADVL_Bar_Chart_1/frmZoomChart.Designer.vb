<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmZoomChart
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
        Me.rbLockInterval = New System.Windows.Forms.RadioButton()
        Me.rbLockTo = New System.Windows.Forms.RadioButton()
        Me.rbLockFrom = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rbY2Axis = New System.Windows.Forms.RadioButton()
        Me.rbYAxis = New System.Windows.Forms.RadioButton()
        Me.rbX2Axis = New System.Windows.Forms.RadioButton()
        Me.rbXAxis = New System.Windows.Forms.RadioButton()
        Me.chkAutoUpdate = New System.Windows.Forms.CheckBox()
        Me.HScrollBar3 = New System.Windows.Forms.HScrollBar()
        Me.HScrollBar2 = New System.Windows.Forms.HScrollBar()
        Me.HScrollBar1 = New System.Windows.Forms.HScrollBar()
        Me.Label73 = New System.Windows.Forms.Label()
        Me.txtAxisZoomInterval = New System.Windows.Forms.TextBox()
        Me.txtAxisZoomTo = New System.Windows.Forms.TextBox()
        Me.Label72 = New System.Windows.Forms.Label()
        Me.txtAxisZoomFrom = New System.Windows.Forms.TextBox()
        Me.Label71 = New System.Windows.Forms.Label()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnApply = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmbAreaName = New System.Windows.Forms.ComboBox()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'rbLockInterval
        '
        Me.rbLockInterval.AutoSize = True
        Me.rbLockInterval.Location = New System.Drawing.Point(203, 144)
        Me.rbLockInterval.Name = "rbLockInterval"
        Me.rbLockInterval.Size = New System.Drawing.Size(49, 17)
        Me.rbLockInterval.TabIndex = 332
        Me.rbLockInterval.TabStop = True
        Me.rbLockInterval.Text = "Lock"
        Me.rbLockInterval.UseVisualStyleBackColor = True
        '
        'rbLockTo
        '
        Me.rbLockTo.AutoSize = True
        Me.rbLockTo.Location = New System.Drawing.Point(203, 118)
        Me.rbLockTo.Name = "rbLockTo"
        Me.rbLockTo.Size = New System.Drawing.Size(49, 17)
        Me.rbLockTo.TabIndex = 331
        Me.rbLockTo.TabStop = True
        Me.rbLockTo.Text = "Lock"
        Me.rbLockTo.UseVisualStyleBackColor = True
        '
        'rbLockFrom
        '
        Me.rbLockFrom.AutoSize = True
        Me.rbLockFrom.Location = New System.Drawing.Point(203, 92)
        Me.rbLockFrom.Name = "rbLockFrom"
        Me.rbLockFrom.Size = New System.Drawing.Size(49, 17)
        Me.rbLockFrom.TabIndex = 330
        Me.rbLockFrom.TabStop = True
        Me.rbLockFrom.Text = "Lock"
        Me.rbLockFrom.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rbY2Axis)
        Me.GroupBox1.Controls.Add(Me.rbYAxis)
        Me.GroupBox1.Controls.Add(Me.rbX2Axis)
        Me.GroupBox1.Controls.Add(Me.rbXAxis)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 40)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(258, 44)
        Me.GroupBox1.TabIndex = 329
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Axis to Zoom:"
        '
        'rbY2Axis
        '
        Me.rbY2Axis.AutoSize = True
        Me.rbY2Axis.Location = New System.Drawing.Point(192, 19)
        Me.rbY2Axis.Name = "rbY2Axis"
        Me.rbY2Axis.Size = New System.Drawing.Size(60, 17)
        Me.rbY2Axis.TabIndex = 296
        Me.rbY2Axis.TabStop = True
        Me.rbY2Axis.Text = "Y2 Axis"
        Me.rbY2Axis.UseVisualStyleBackColor = True
        '
        'rbYAxis
        '
        Me.rbYAxis.AutoSize = True
        Me.rbYAxis.Location = New System.Drawing.Point(132, 19)
        Me.rbYAxis.Name = "rbYAxis"
        Me.rbYAxis.Size = New System.Drawing.Size(54, 17)
        Me.rbYAxis.TabIndex = 295
        Me.rbYAxis.TabStop = True
        Me.rbYAxis.Text = "Y Axis"
        Me.rbYAxis.UseVisualStyleBackColor = True
        '
        'rbX2Axis
        '
        Me.rbX2Axis.AutoSize = True
        Me.rbX2Axis.Location = New System.Drawing.Point(66, 19)
        Me.rbX2Axis.Name = "rbX2Axis"
        Me.rbX2Axis.Size = New System.Drawing.Size(60, 17)
        Me.rbX2Axis.TabIndex = 294
        Me.rbX2Axis.TabStop = True
        Me.rbX2Axis.Text = "X2 Axis"
        Me.rbX2Axis.UseVisualStyleBackColor = True
        '
        'rbXAxis
        '
        Me.rbXAxis.AutoSize = True
        Me.rbXAxis.Location = New System.Drawing.Point(6, 19)
        Me.rbXAxis.Name = "rbXAxis"
        Me.rbXAxis.Size = New System.Drawing.Size(54, 17)
        Me.rbXAxis.TabIndex = 293
        Me.rbXAxis.TabStop = True
        Me.rbXAxis.Text = "X Axis"
        Me.rbXAxis.UseVisualStyleBackColor = True
        '
        'chkAutoUpdate
        '
        Me.chkAutoUpdate.AutoSize = True
        Me.chkAutoUpdate.Location = New System.Drawing.Point(120, 16)
        Me.chkAutoUpdate.Name = "chkAutoUpdate"
        Me.chkAutoUpdate.Size = New System.Drawing.Size(86, 17)
        Me.chkAutoUpdate.TabIndex = 328
        Me.chkAutoUpdate.Text = "Auto Update"
        Me.chkAutoUpdate.UseVisualStyleBackColor = True
        '
        'HScrollBar3
        '
        Me.HScrollBar3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.HScrollBar3.Location = New System.Drawing.Point(255, 143)
        Me.HScrollBar3.Name = "HScrollBar3"
        Me.HScrollBar3.Size = New System.Drawing.Size(532, 20)
        Me.HScrollBar3.TabIndex = 327
        '
        'HScrollBar2
        '
        Me.HScrollBar2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.HScrollBar2.Location = New System.Drawing.Point(255, 117)
        Me.HScrollBar2.Name = "HScrollBar2"
        Me.HScrollBar2.Size = New System.Drawing.Size(532, 20)
        Me.HScrollBar2.TabIndex = 326
        '
        'HScrollBar1
        '
        Me.HScrollBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.HScrollBar1.Location = New System.Drawing.Point(255, 91)
        Me.HScrollBar1.Name = "HScrollBar1"
        Me.HScrollBar1.Size = New System.Drawing.Size(532, 20)
        Me.HScrollBar1.TabIndex = 325
        '
        'Label73
        '
        Me.Label73.AutoSize = True
        Me.Label73.Location = New System.Drawing.Point(27, 143)
        Me.Label73.Name = "Label73"
        Me.Label73.Size = New System.Drawing.Size(45, 13)
        Me.Label73.TabIndex = 324
        Me.Label73.Text = "Interval:"
        '
        'txtAxisZoomInterval
        '
        Me.txtAxisZoomInterval.Location = New System.Drawing.Point(78, 143)
        Me.txtAxisZoomInterval.Name = "txtAxisZoomInterval"
        Me.txtAxisZoomInterval.Size = New System.Drawing.Size(119, 20)
        Me.txtAxisZoomInterval.TabIndex = 323
        '
        'txtAxisZoomTo
        '
        Me.txtAxisZoomTo.Location = New System.Drawing.Point(78, 117)
        Me.txtAxisZoomTo.Name = "txtAxisZoomTo"
        Me.txtAxisZoomTo.Size = New System.Drawing.Size(119, 20)
        Me.txtAxisZoomTo.TabIndex = 322
        '
        'Label72
        '
        Me.Label72.AutoSize = True
        Me.Label72.Location = New System.Drawing.Point(53, 120)
        Me.Label72.Name = "Label72"
        Me.Label72.Size = New System.Drawing.Size(19, 13)
        Me.Label72.TabIndex = 321
        Me.Label72.Text = "to:"
        '
        'txtAxisZoomFrom
        '
        Me.txtAxisZoomFrom.Location = New System.Drawing.Point(78, 91)
        Me.txtAxisZoomFrom.Name = "txtAxisZoomFrom"
        Me.txtAxisZoomFrom.Size = New System.Drawing.Size(119, 20)
        Me.txtAxisZoomFrom.TabIndex = 320
        '
        'Label71
        '
        Me.Label71.AutoSize = True
        Me.Label71.Location = New System.Drawing.Point(12, 94)
        Me.Label71.Name = "Label71"
        Me.Label71.Size = New System.Drawing.Size(60, 13)
        Me.Label71.TabIndex = 319
        Me.Label71.Text = "Zoom from:"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(66, 12)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(48, 22)
        Me.btnCancel.TabIndex = 318
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnApply
        '
        Me.btnApply.Location = New System.Drawing.Point(12, 12)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(48, 22)
        Me.btnApply.TabIndex = 317
        Me.btnApply.Text = "Apply"
        Me.btnApply.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(276, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 13)
        Me.Label1.TabIndex = 316
        Me.Label1.Text = "Chart area name:"
        '
        'cmbAreaName
        '
        Me.cmbAreaName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbAreaName.FormattingEnabled = True
        Me.cmbAreaName.Location = New System.Drawing.Point(370, 40)
        Me.cmbAreaName.Name = "cmbAreaName"
        Me.cmbAreaName.Size = New System.Drawing.Size(414, 21)
        Me.cmbAreaName.TabIndex = 315
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(736, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(48, 22)
        Me.btnExit.TabIndex = 314
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'frmZoomChart
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(796, 172)
        Me.Controls.Add(Me.rbLockInterval)
        Me.Controls.Add(Me.rbLockTo)
        Me.Controls.Add(Me.rbLockFrom)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.chkAutoUpdate)
        Me.Controls.Add(Me.HScrollBar3)
        Me.Controls.Add(Me.HScrollBar2)
        Me.Controls.Add(Me.HScrollBar1)
        Me.Controls.Add(Me.Label73)
        Me.Controls.Add(Me.txtAxisZoomInterval)
        Me.Controls.Add(Me.txtAxisZoomTo)
        Me.Controls.Add(Me.Label72)
        Me.Controls.Add(Me.txtAxisZoomFrom)
        Me.Controls.Add(Me.Label71)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnApply)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmbAreaName)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmZoomChart"
        Me.Text = "Zoom Chart"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents rbLockInterval As RadioButton
    Friend WithEvents rbLockTo As RadioButton
    Friend WithEvents rbLockFrom As RadioButton
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents rbY2Axis As RadioButton
    Friend WithEvents rbYAxis As RadioButton
    Friend WithEvents rbX2Axis As RadioButton
    Friend WithEvents rbXAxis As RadioButton
    Friend WithEvents chkAutoUpdate As CheckBox
    Friend WithEvents HScrollBar3 As HScrollBar
    Friend WithEvents HScrollBar2 As HScrollBar
    Friend WithEvents HScrollBar1 As HScrollBar
    Friend WithEvents Label73 As Label
    Friend WithEvents txtAxisZoomInterval As TextBox
    Friend WithEvents txtAxisZoomTo As TextBox
    Friend WithEvents Label72 As Label
    Friend WithEvents txtAxisZoomFrom As TextBox
    Friend WithEvents Label71 As Label
    Friend WithEvents btnCancel As Button
    Friend WithEvents btnApply As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents cmbAreaName As ComboBox
    Friend WithEvents btnExit As Button
End Class
