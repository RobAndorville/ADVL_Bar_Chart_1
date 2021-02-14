<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChart
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
        Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series1 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Me.txtChartFileName = New System.Windows.Forms.TextBox()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.btnZoomChart = New System.Windows.Forms.Button()
        Me.btnMarkerProps2 = New System.Windows.Forms.Button()
        Me.chkAutoDraw = New System.Windows.Forms.CheckBox()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnNew = New System.Windows.Forms.Button()
        Me.btnOpen = New System.Windows.Forms.Button()
        Me.btnDrawChart = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.Chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtChartFileName
        '
        Me.txtChartFileName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtChartFileName.Location = New System.Drawing.Point(98, 40)
        Me.txtChartFileName.Name = "txtChartFileName"
        Me.txtChartFileName.Size = New System.Drawing.Size(533, 20)
        Me.txtChartFileName.TabIndex = 301
        '
        'Label33
        '
        Me.Label33.AutoSize = True
        Me.Label33.Location = New System.Drawing.Point(12, 43)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(80, 13)
        Me.Label33.TabIndex = 300
        Me.Label33.Text = "Chart file name:"
        '
        'btnZoomChart
        '
        Me.btnZoomChart.Location = New System.Drawing.Point(418, 12)
        Me.btnZoomChart.Name = "btnZoomChart"
        Me.btnZoomChart.Size = New System.Drawing.Size(54, 22)
        Me.btnZoomChart.TabIndex = 299
        Me.btnZoomChart.Text = "Zoom"
        Me.btnZoomChart.UseVisualStyleBackColor = True
        '
        'btnMarkerProps2
        '
        Me.btnMarkerProps2.Location = New System.Drawing.Point(305, 12)
        Me.btnMarkerProps2.Name = "btnMarkerProps2"
        Me.btnMarkerProps2.Size = New System.Drawing.Size(107, 22)
        Me.btnMarkerProps2.TabIndex = 298
        Me.btnMarkerProps2.Text = "Marker Properties"
        Me.btnMarkerProps2.UseVisualStyleBackColor = True
        '
        'chkAutoDraw
        '
        Me.chkAutoDraw.AutoSize = True
        Me.chkAutoDraw.Location = New System.Drawing.Point(223, 16)
        Me.chkAutoDraw.Name = "chkAutoDraw"
        Me.chkAutoDraw.Size = New System.Drawing.Size(76, 17)
        Me.chkAutoDraw.TabIndex = 297
        Me.chkAutoDraw.Text = "Auto Draw"
        Me.chkAutoDraw.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(116, 12)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(46, 22)
        Me.btnSave.TabIndex = 296
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(64, 12)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(46, 22)
        Me.btnNew.TabIndex = 295
        Me.btnNew.Text = "New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnOpen
        '
        Me.btnOpen.Location = New System.Drawing.Point(12, 12)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.Size = New System.Drawing.Size(46, 22)
        Me.btnOpen.TabIndex = 294
        Me.btnOpen.Text = "Open"
        Me.btnOpen.UseVisualStyleBackColor = True
        '
        'btnDrawChart
        '
        Me.btnDrawChart.Location = New System.Drawing.Point(168, 12)
        Me.btnDrawChart.Name = "btnDrawChart"
        Me.btnDrawChart.Size = New System.Drawing.Size(49, 22)
        Me.btnDrawChart.TabIndex = 293
        Me.btnDrawChart.Text = "Draw"
        Me.btnDrawChart.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(583, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(48, 22)
        Me.btnExit.TabIndex = 292
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'Chart1
        '
        Me.Chart1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        ChartArea1.Name = "ChartArea1"
        Me.Chart1.ChartAreas.Add(ChartArea1)
        Legend1.Name = "Legend1"
        Me.Chart1.Legends.Add(Legend1)
        Me.Chart1.Location = New System.Drawing.Point(12, 66)
        Me.Chart1.Name = "Chart1"
        Series1.ChartArea = "ChartArea1"
        Series1.Legend = "Legend1"
        Series1.Name = "Series1"
        Me.Chart1.Series.Add(Series1)
        Me.Chart1.Size = New System.Drawing.Size(619, 430)
        Me.Chart1.TabIndex = 302
        Me.Chart1.Text = "Chart1"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'frmChart
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(643, 508)
        Me.Controls.Add(Me.Chart1)
        Me.Controls.Add(Me.txtChartFileName)
        Me.Controls.Add(Me.Label33)
        Me.Controls.Add(Me.btnZoomChart)
        Me.Controls.Add(Me.btnMarkerProps2)
        Me.Controls.Add(Me.chkAutoDraw)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnOpen)
        Me.Controls.Add(Me.btnDrawChart)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmChart"
        Me.Text = "frmChart"
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtChartFileName As TextBox
    Friend WithEvents Label33 As Label
    Friend WithEvents btnZoomChart As Button
    Friend WithEvents btnMarkerProps2 As Button
    Friend WithEvents chkAutoDraw As CheckBox
    Friend WithEvents btnSave As Button
    Friend WithEvents btnNew As Button
    Friend WithEvents btnOpen As Button
    Friend WithEvents btnDrawChart As Button
    Friend WithEvents btnExit As Button
    Friend WithEvents Chart1 As DataVisualization.Charting.Chart
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
End Class
