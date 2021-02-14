<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDesignQuery
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
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnApply = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.txtQuery = New System.Windows.Forms.TextBox()
        Me.btnMakeSqlStatement = New System.Windows.Forms.Button()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.chkDate2 = New System.Windows.Forms.CheckBox()
        Me.chkDate1 = New System.Windows.Forms.CheckBox()
        Me.DateTimePicker6 = New System.Windows.Forms.DateTimePicker()
        Me.txtSecondValue2 = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtValue2 = New System.Windows.Forms.TextBox()
        Me.DateTimePicker5 = New System.Windows.Forms.DateTimePicker()
        Me.cmbType2 = New System.Windows.Forms.ComboBox()
        Me.cmbConstraint2 = New System.Windows.Forms.ComboBox()
        Me.cmbField2 = New System.Windows.Forms.ComboBox()
        Me.DateTimePicker4 = New System.Windows.Forms.DateTimePicker()
        Me.DateTimePicker3 = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtSecondValue1 = New System.Windows.Forms.TextBox()
        Me.txtValue1 = New System.Windows.Forms.TextBox()
        Me.cmbType1 = New System.Windows.Forms.ComboBox()
        Me.cmbField1 = New System.Windows.Forms.ComboBox()
        Me.cmbConstraint1 = New System.Windows.Forms.ComboBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.txtDatabasePath = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnNone = New System.Windows.Forms.Button()
        Me.btnAll = New System.Windows.Forms.Button()
        Me.lstSelectFields = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lstTables = New System.Windows.Forms.ListBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(863, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(64, 22)
        Me.btnExit.TabIndex = 35
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnApply
        '
        Me.btnApply.Location = New System.Drawing.Point(12, 12)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(70, 22)
        Me.btnApply.TabIndex = 34
        Me.btnApply.Text = "Apply"
        Me.btnApply.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 40)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(915, 521)
        Me.TabControl1.TabIndex = 36
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.txtQuery)
        Me.TabPage1.Controls.Add(Me.btnMakeSqlStatement)
        Me.TabPage1.Controls.Add(Me.btnClear)
        Me.TabPage1.Controls.Add(Me.GroupBox4)
        Me.TabPage1.Controls.Add(Me.GroupBox3)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(907, 495)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Select"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'txtQuery
        '
        Me.txtQuery.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtQuery.Location = New System.Drawing.Point(6, 341)
        Me.txtQuery.Multiline = True
        Me.txtQuery.Name = "txtQuery"
        Me.txtQuery.Size = New System.Drawing.Size(895, 148)
        Me.txtQuery.TabIndex = 31
        '
        'btnMakeSqlStatement
        '
        Me.btnMakeSqlStatement.Location = New System.Drawing.Point(78, 313)
        Me.btnMakeSqlStatement.Name = "btnMakeSqlStatement"
        Me.btnMakeSqlStatement.Size = New System.Drawing.Size(128, 22)
        Me.btnMakeSqlStatement.TabIndex = 30
        Me.btnMakeSqlStatement.Text = "Make SQL Statement"
        Me.btnMakeSqlStatement.UseVisualStyleBackColor = True
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(6, 313)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(66, 22)
        Me.btnClear.TabIndex = 29
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox4.Controls.Add(Me.Label7)
        Me.GroupBox4.Controls.Add(Me.Label6)
        Me.GroupBox4.Controls.Add(Me.chkDate2)
        Me.GroupBox4.Controls.Add(Me.chkDate1)
        Me.GroupBox4.Controls.Add(Me.DateTimePicker6)
        Me.GroupBox4.Controls.Add(Me.txtSecondValue2)
        Me.GroupBox4.Controls.Add(Me.Label5)
        Me.GroupBox4.Controls.Add(Me.txtValue2)
        Me.GroupBox4.Controls.Add(Me.DateTimePicker5)
        Me.GroupBox4.Controls.Add(Me.cmbType2)
        Me.GroupBox4.Controls.Add(Me.cmbConstraint2)
        Me.GroupBox4.Controls.Add(Me.cmbField2)
        Me.GroupBox4.Controls.Add(Me.DateTimePicker4)
        Me.GroupBox4.Controls.Add(Me.DateTimePicker3)
        Me.GroupBox4.Controls.Add(Me.Label4)
        Me.GroupBox4.Controls.Add(Me.txtSecondValue1)
        Me.GroupBox4.Controls.Add(Me.txtValue1)
        Me.GroupBox4.Controls.Add(Me.cmbType1)
        Me.GroupBox4.Controls.Add(Me.cmbField1)
        Me.GroupBox4.Controls.Add(Me.cmbConstraint1)
        Me.GroupBox4.Location = New System.Drawing.Point(6, 171)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(895, 136)
        Me.GroupBox4.TabIndex = 28
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "General Constraints"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(112, 108)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(202, 13)
        Me.Label7.TabIndex = 42
        Me.Label7.Text = "Note: Access queries use US date format"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(112, 55)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(202, 13)
        Me.Label6.TabIndex = 41
        Me.Label6.Text = "Note: Access queries use US date format"
        '
        'chkDate2
        '
        Me.chkDate2.AutoSize = True
        Me.chkDate2.Location = New System.Drawing.Point(320, 107)
        Me.chkDate2.Name = "chkDate2"
        Me.chkDate2.Size = New System.Drawing.Size(49, 17)
        Me.chkDate2.TabIndex = 40
        Me.chkDate2.Text = "Date"
        Me.chkDate2.UseVisualStyleBackColor = True
        '
        'chkDate1
        '
        Me.chkDate1.AutoSize = True
        Me.chkDate1.Location = New System.Drawing.Point(320, 54)
        Me.chkDate1.Name = "chkDate1"
        Me.chkDate1.Size = New System.Drawing.Size(49, 17)
        Me.chkDate1.TabIndex = 39
        Me.chkDate1.Text = "Date"
        Me.chkDate1.UseVisualStyleBackColor = True
        '
        'DateTimePicker6
        '
        Me.DateTimePicker6.Location = New System.Drawing.Point(629, 104)
        Me.DateTimePicker6.Name = "DateTimePicker6"
        Me.DateTimePicker6.Size = New System.Drawing.Size(211, 20)
        Me.DateTimePicker6.TabIndex = 38
        '
        'txtSecondValue2
        '
        Me.txtSecondValue2.Location = New System.Drawing.Point(629, 78)
        Me.txtSecondValue2.Name = "txtSecondValue2"
        Me.txtSecondValue2.ShortcutsEnabled = False
        Me.txtSecondValue2.Size = New System.Drawing.Size(212, 20)
        Me.txtSecondValue2.TabIndex = 37
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(592, 81)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(30, 13)
        Me.Label5.TabIndex = 36
        Me.Label5.Text = "AND"
        '
        'txtValue2
        '
        Me.txtValue2.Location = New System.Drawing.Point(375, 78)
        Me.txtValue2.Name = "txtValue2"
        Me.txtValue2.Size = New System.Drawing.Size(211, 20)
        Me.txtValue2.TabIndex = 35
        '
        'DateTimePicker5
        '
        Me.DateTimePicker5.Location = New System.Drawing.Point(375, 104)
        Me.DateTimePicker5.Name = "DateTimePicker5"
        Me.DateTimePicker5.Size = New System.Drawing.Size(211, 20)
        Me.DateTimePicker5.TabIndex = 34
        '
        'cmbType2
        '
        Me.cmbType2.FormattingEnabled = True
        Me.cmbType2.Location = New System.Drawing.Point(293, 77)
        Me.cmbType2.Name = "cmbType2"
        Me.cmbType2.Size = New System.Drawing.Size(76, 21)
        Me.cmbType2.TabIndex = 33
        '
        'cmbConstraint2
        '
        Me.cmbConstraint2.FormattingEnabled = True
        Me.cmbConstraint2.Location = New System.Drawing.Point(9, 77)
        Me.cmbConstraint2.Name = "cmbConstraint2"
        Me.cmbConstraint2.Size = New System.Drawing.Size(80, 21)
        Me.cmbConstraint2.TabIndex = 32
        '
        'cmbField2
        '
        Me.cmbField2.FormattingEnabled = True
        Me.cmbField2.Location = New System.Drawing.Point(95, 77)
        Me.cmbField2.Name = "cmbField2"
        Me.cmbField2.Size = New System.Drawing.Size(192, 21)
        Me.cmbField2.TabIndex = 31
        '
        'DateTimePicker4
        '
        Me.DateTimePicker4.Location = New System.Drawing.Point(629, 51)
        Me.DateTimePicker4.Name = "DateTimePicker4"
        Me.DateTimePicker4.Size = New System.Drawing.Size(212, 20)
        Me.DateTimePicker4.TabIndex = 30
        '
        'DateTimePicker3
        '
        Me.DateTimePicker3.Location = New System.Drawing.Point(375, 51)
        Me.DateTimePicker3.Name = "DateTimePicker3"
        Me.DateTimePicker3.Size = New System.Drawing.Size(212, 20)
        Me.DateTimePicker3.TabIndex = 29
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(593, 28)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 13)
        Me.Label4.TabIndex = 28
        Me.Label4.Text = "AND"
        '
        'txtSecondValue1
        '
        Me.txtSecondValue1.Location = New System.Drawing.Point(629, 25)
        Me.txtSecondValue1.Name = "txtSecondValue1"
        Me.txtSecondValue1.Size = New System.Drawing.Size(212, 20)
        Me.txtSecondValue1.TabIndex = 27
        '
        'txtValue1
        '
        Me.txtValue1.Location = New System.Drawing.Point(375, 24)
        Me.txtValue1.Name = "txtValue1"
        Me.txtValue1.Size = New System.Drawing.Size(212, 20)
        Me.txtValue1.TabIndex = 26
        '
        'cmbType1
        '
        Me.cmbType1.FormattingEnabled = True
        Me.cmbType1.Location = New System.Drawing.Point(293, 24)
        Me.cmbType1.Name = "cmbType1"
        Me.cmbType1.Size = New System.Drawing.Size(76, 21)
        Me.cmbType1.TabIndex = 25
        '
        'cmbField1
        '
        Me.cmbField1.FormattingEnabled = True
        Me.cmbField1.Location = New System.Drawing.Point(95, 24)
        Me.cmbField1.Name = "cmbField1"
        Me.cmbField1.Size = New System.Drawing.Size(192, 21)
        Me.cmbField1.TabIndex = 24
        '
        'cmbConstraint1
        '
        Me.cmbConstraint1.FormattingEnabled = True
        Me.cmbConstraint1.Location = New System.Drawing.Point(9, 24)
        Me.cmbConstraint1.Name = "cmbConstraint1"
        Me.cmbConstraint1.Size = New System.Drawing.Size(80, 21)
        Me.cmbConstraint1.TabIndex = 23
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.txtDatabasePath)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.btnNone)
        Me.GroupBox3.Controls.Add(Me.btnAll)
        Me.GroupBox3.Controls.Add(Me.lstSelectFields)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Controls.Add(Me.lstTables)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Location = New System.Drawing.Point(6, 6)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(895, 159)
        Me.GroupBox3.TabIndex = 27
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Select Statement"
        '
        'txtDatabasePath
        '
        Me.txtDatabasePath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDatabasePath.Location = New System.Drawing.Point(515, 35)
        Me.txtDatabasePath.Multiline = True
        Me.txtDatabasePath.Name = "txtDatabasePath"
        Me.txtDatabasePath.Size = New System.Drawing.Size(369, 113)
        Me.txtDatabasePath.TabIndex = 14
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(512, 19)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 13)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Database path:"
        '
        'btnNone
        '
        Me.btnNone.Location = New System.Drawing.Point(9, 63)
        Me.btnNone.Name = "btnNone"
        Me.btnNone.Size = New System.Drawing.Size(45, 22)
        Me.btnNone.TabIndex = 12
        Me.btnNone.Text = "None"
        Me.btnNone.UseVisualStyleBackColor = True
        '
        'btnAll
        '
        Me.btnAll.Location = New System.Drawing.Point(9, 35)
        Me.btnAll.Name = "btnAll"
        Me.btnAll.Size = New System.Drawing.Size(45, 22)
        Me.btnAll.TabIndex = 11
        Me.btnAll.Text = "All"
        Me.btnAll.UseVisualStyleBackColor = True
        '
        'lstSelectFields
        '
        Me.lstSelectFields.FormattingEnabled = True
        Me.lstSelectFields.Location = New System.Drawing.Point(60, 19)
        Me.lstSelectFields.Name = "lstSelectFields"
        Me.lstSelectFields.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.lstSelectFields.Size = New System.Drawing.Size(163, 134)
        Me.lstSelectFields.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "SELECT"
        '
        'lstTables
        '
        Me.lstTables.FormattingEnabled = True
        Me.lstTables.Location = New System.Drawing.Point(229, 35)
        Me.lstTables.Name = "lstTables"
        Me.lstTables.Size = New System.Drawing.Size(280, 108)
        Me.lstTables.TabIndex = 8
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(229, 19)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(38, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "FROM"
        '
        'TabPage2
        '
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(907, 495)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Misc"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'frmDesignQuery
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(939, 573)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnApply)
        Me.Name = "frmDesignQuery"
        Me.Text = "Design Query"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnExit As Button
    Friend WithEvents btnApply As Button
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents txtQuery As TextBox
    Friend WithEvents btnMakeSqlStatement As Button
    Friend WithEvents btnClear As Button
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents chkDate2 As CheckBox
    Friend WithEvents chkDate1 As CheckBox
    Friend WithEvents DateTimePicker6 As DateTimePicker
    Friend WithEvents txtSecondValue2 As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents txtValue2 As TextBox
    Friend WithEvents DateTimePicker5 As DateTimePicker
    Friend WithEvents cmbType2 As ComboBox
    Friend WithEvents cmbConstraint2 As ComboBox
    Friend WithEvents cmbField2 As ComboBox
    Friend WithEvents DateTimePicker4 As DateTimePicker
    Friend WithEvents DateTimePicker3 As DateTimePicker
    Friend WithEvents Label4 As Label
    Friend WithEvents txtSecondValue1 As TextBox
    Friend WithEvents txtValue1 As TextBox
    Friend WithEvents cmbType1 As ComboBox
    Friend WithEvents cmbField1 As ComboBox
    Friend WithEvents cmbConstraint1 As ComboBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents txtDatabasePath As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents btnNone As Button
    Friend WithEvents btnAll As Button
    Friend WithEvents lstSelectFields As ListBox
    Friend WithEvents Label1 As Label
    Friend WithEvents lstTables As ListBox
    Friend WithEvents Label2 As Label
    Friend WithEvents TabPage2 As TabPage
End Class
