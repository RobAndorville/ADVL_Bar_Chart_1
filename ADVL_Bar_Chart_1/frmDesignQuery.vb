Public Class frmDesignQuery
    'This form is used to design an SQL query.

#Region " Variable Declarations - All the variables used in this form and this application." '-------------------------------------------------------------------------------------------------
#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Properties - All the properties used in this form and this application" '------------------------------------------------------------------------------------------------------------

    Private _databasePath As String = "" 'The path of the database.
    Property DatabasePath As String
        Get
            Return _databasePath
        End Get
        Set(value As String)
            _databasePath = value
            txtDatabasePath.Text = _databasePath
            lstSelectFields.Items.Clear()
            FillLstTables()
        End Set
    End Property

    Private _tableName As String = "" 'TableName stores the name of the Table selected for viewing.
    Public Property TableName As String
        Get
            Return _tableName
        End Get
        Set(value As String)
            _tableName = value
        End Set
    End Property

#End Region 'Properties -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Process XML files - Read and write XML files." '-------------------------------------------------------------------------------------------------------------------------------------

    Private Sub SaveFormSettings()
        'Save the form settings in an XML document.
        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <FormSettings>
                               <Left><%= Me.Left %></Left>
                               <Top><%= Me.Top %></Top>
                               <Width><%= Me.Width %></Width>
                               <Height><%= Me.Height %></Height>
                               <!---->
                           </FormSettings>

        'Add code to include other settings to save after the comment line <!---->

        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Main.Project.SaveXmlSettings(SettingsFileName, settingsData)
    End Sub

    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & ".xml"

        If Main.Project.SettingsFileExists(SettingsFileName) Then
            Dim Settings As System.Xml.Linq.XDocument
            Main.Project.ReadXmlSettings(SettingsFileName, Settings)

            If IsNothing(Settings) Then 'There is no Settings XML data.
                Exit Sub
            End If

            'Restore form position and size:
            If Settings.<FormSettings>.<Left>.Value <> Nothing Then Me.Left = Settings.<FormSettings>.<Left>.Value
            If Settings.<FormSettings>.<Top>.Value <> Nothing Then Me.Top = Settings.<FormSettings>.<Top>.Value
            If Settings.<FormSettings>.<Height>.Value <> Nothing Then Me.Height = Settings.<FormSettings>.<Height>.Value
            If Settings.<FormSettings>.<Width>.Value <> Nothing Then Me.Width = Settings.<FormSettings>.<Width>.Value

            'Add code to read other saved setting here:

            CheckFormPos()
        End If
    End Sub

    Private Sub CheckFormPos()
        'Check that the form can be seen on a screen.

        Dim MinWidthVisible As Integer = 192 'Minimum number of X pixels visible. The form will be moved if this many form pixels are not visible.
        Dim MinHeightVisible As Integer = 64 'Minimum number of Y pixels visible. The form will be moved if this many form pixels are not visible.

        Dim FormRect As New Rectangle(Me.Left, Me.Top, Me.Width, Me.Height)
        Dim WARect As Rectangle = Screen.GetWorkingArea(FormRect) 'The Working Area rectangle - the usable area of the screen containing the form.

        ''Check if the top of the form is less than zero:
        'If Me.Top < 0 Then Me.Top = 0

        'Check if the top of the form is above the top of the Working Area:
        If Me.Top < WARect.Top Then
            Me.Top = WARect.Top
        End If

        'Check if the top of the form is too close to the bottom of the Working Area:
        If (Me.Top + MinHeightVisible) > (WARect.Top + WARect.Height) Then
            Me.Top = WARect.Top + WARect.Height - MinHeightVisible
        End If

        'Check if the left edge of the form is too close to the right edge of the Working Area:
        If (Me.Left + MinWidthVisible) > (WARect.Left + WARect.Width) Then
            Me.Left = WARect.Left + WARect.Width - MinWidthVisible
        End If

        'Check if the right edge of the form is too close to the left edge of the Working Area:
        If (Me.Left + Me.Width - MinWidthVisible) < WARect.Left Then
            Me.Left = WARect.Left - Me.Width + MinWidthVisible
        End If
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message) 'Save the form settings before the form is minimised:
        If m.Msg = &H112 Then 'SysCommand
            If m.WParam.ToInt32 = &HF020 Then 'Form is being minimised
                SaveFormSettings()
            End If
        End If
        MyBase.WndProc(m)
    End Sub

#End Region 'Process XML Files ----------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Display Methods - Code used to display this form." '----------------------------------------------------------------------------------------------------------------------------

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        RestoreFormSettings()   'Restore the form settings

        cmbConstraint1.Items.Add("")
        cmbConstraint1.Items.Add("WHERE")
        cmbConstraint1.Items.Add("ORDER BY")
        cmbConstraint1.SelectedIndex = 0

        cmbConstraint2.Items.Add("")
        cmbConstraint2.Items.Add("WHERE")
        cmbConstraint2.Items.Add("AND")
        cmbConstraint2.Items.Add("OR")
        cmbConstraint2.Items.Add("ORDER BY")
        cmbConstraint2.SelectedIndex = 0

        cmbType1.Items.Add(">")
        cmbType1.Items.Add(">=")
        cmbType1.Items.Add("=")
        cmbType1.Items.Add("<=")
        cmbType1.Items.Add("<")
        cmbType1.Items.Add("BETWEEN")
        cmbType1.SelectedIndex = 2

        cmbType2.Items.Add(">")
        cmbType2.Items.Add(">=")
        cmbType2.Items.Add("=")
        cmbType2.Items.Add("<=")
        cmbType2.Items.Add("<")
        cmbType2.Items.Add("BETWEEN")
        cmbType2.SelectedIndex = 2

        'Disable the second value text box. (This is only used when the BETWEEN constraint is selected.)

        txtSecondValue1.Enabled = False
        Label4.Enabled = False
        chkDate1.Checked = False
        Label6.Enabled = False
        DateTimePicker3.Enabled = False
        DateTimePicker4.Enabled = False

        txtSecondValue2.Enabled = False
        Label5.Enabled = False
        chkDate2.Checked = False
        Label7.Enabled = False
        DateTimePicker5.Enabled = False
        DateTimePicker6.Enabled = False

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form
        Me.Close() 'Close the form
    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        Else
            'Dont save settings if form is minimised.
        End If
    End Sub

#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Open and Close Forms - Code used to open and close other forms." '-------------------------------------------------------------------------------------------------------------------
#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Methods - The main actions performed by this form." '---------------------------------------------------------------------------------------------------------------------------




    Private Sub FillLstTables()
        'Fill the cmbSelectTable listbox with the availalble tables in the selected database.

        If DatabasePath = "" Then
            Main.Message.AddWarning("No database has been selected.")
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        lstTables.Items.Clear()

        'Specify the connection string:
        'Access 2003
        'connectionString = "provider=Microsoft.Jet.OLEDB.4.0;" + _
        '"data source = " + txtDatabase.Text

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0; data source = " + DatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        'This error occurs on the above line (conn.Open()):
        'Additional information: The 'Microsoft.ACE.OLEDB.12.0' provider is not registered on the local machine.
        'Fix attempt: 
        'http://www.microsoft.com/en-us/download/confirmation.aspx?id=23734
        'Download AccessDatabaseEngine.exe
        'Run the file to install the 2007 Office System Driver: Data Connectivity Components.


        Dim restrictions As String() = New String() {Nothing, Nothing, Nothing, "TABLE"} 'This restriction removes system tables
        dt = conn.GetSchema("Tables", restrictions)

        'Fill lstSelectTable
        Dim dr As DataRow
        Dim I As Integer 'Loop index
        Dim MaxI As Integer

        MaxI = dt.Rows.Count
        For I = 0 To MaxI - 1
            dr = dt.Rows(0)
            lstTables.Items.Add(dt.Rows(I).Item(2).ToString)
        Next I

        conn.Close()

    End Sub

    Private Sub FillLstFields()
        'Fill the lstSelectField listbox with the availalble fields in the selected table.
        'Also fill the cmbField1 and cmbField2 comboboxes.

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim commandString As String 'Declare a command string - contains the query to be passed to the database.
        Dim ds As DataSet 'Declate a Dataset.
        Dim dt As DataTable

        If lstTables.SelectedIndex = -1 Then 'No item is selected
            lstSelectFields.Items.Clear()

        Else 'A table has been selected. List its fields:
            lstSelectFields.Items.Clear()
            cmbField1.Items.Clear()
            cmbField2.Items.Clear()

            'Specify the connection string (Access 2007):
            connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" + "data source = " + DatabasePath

            'Connect to the Access database:
            conn = New System.Data.OleDb.OleDbConnection(connectionString)
            conn.Open()

            'Specify the commandString to query the database:
            commandString = "SELECT TOP 500 * FROM " + lstTables.SelectedItem.ToString
            Dim dataAdapter As New System.Data.OleDb.OleDbDataAdapter(commandString, conn)
            ds = New DataSet
            dataAdapter.Fill(ds, "SelTable") 'ds was defined earlier as a DataSet
            dt = ds.Tables("SelTable")

            Dim NFields As Integer
            NFields = dt.Columns.Count
            Dim I As Integer
            For I = 0 To NFields - 1
                lstSelectFields.Items.Add(dt.Columns(I).ColumnName.ToString)
                cmbField1.Items.Add(dt.Columns(I).ColumnName.ToString)
                cmbField2.Items.Add(dt.Columns(I).ColumnName.ToString)
            Next

            conn.Close()

        End If
    End Sub

    Private Sub lstTables_Click(sender As Object, e As EventArgs) Handles lstTables.Click
        FillLstFields()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtQuery.Text = ""
    End Sub

    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click

        'Main.PointChart.InputQuery = txtQuery.Text
        Main.ChartInfo.InputQuery = txtQuery.Text
        Main.txtInputQuery.Text = txtQuery.Text
        Main.ApplyQuery()

    End Sub

    Private Sub btnAll_Click(sender As Object, e As EventArgs) Handles btnAll.Click
        'Select all the fields in lstSelectFields

        Dim I As Integer 'Loop index
        For I = 1 To lstSelectFields.Items.Count
            lstSelectFields.SetSelected(I - 1, True)
        Next
    End Sub

    Private Sub btnNone_Click(sender As Object, e As EventArgs) Handles btnNone.Click
        'Select none of the fields in lstSelectFields
        Dim I As Integer 'Loop index
        For I = 1 To lstSelectFields.Items.Count
            lstSelectFields.SetSelected(I - 1, False)
        Next
    End Sub

    Private Sub btnMakeSqlStatement_Click(sender As Object, e As EventArgs) Handles btnMakeSqlStatement.Click
        'Make the SQL statement

        txtQuery.Text = ""

        'Make the SELECT part of the statement: ----------------------------------------------------------------------
        If lstTables.SelectedItems.Count = 0 Then
            Main.Message.AddWarning("No table has been selected" & vbCrLf)
            Exit Sub
        End If

        If lstSelectFields.SelectedItems.Count = 0 Then 'No fields have been selected
            Main.Message.AddWarning("No fields have been selected" & vbCrLf)
            Exit Sub
        End If

        If lstSelectFields.Items.Count = lstSelectFields.SelectedItems.Count Then 'All the fields are selected
            txtQuery.Text = "SELECT * FROM " & lstTables.SelectedItem
        Else 'A subset of the fields are selected
            txtQuery.Text = "SELECT "
            Dim I As Integer 'Loop index
            txtQuery.Text = txtQuery.Text & lstSelectFields.SelectedItems(0)
            For I = 1 To lstSelectFields.SelectedItems.Count - 1
                txtQuery.Text = txtQuery.Text & ", " & lstSelectFields.SelectedItems(I)
            Next
            txtQuery.Text = txtQuery.Text & " FROM " & lstTables.SelectedItem
        End If

        'Add the first constraint to the statement: -------------------------------------------------------------------
        If cmbConstraint1.SelectedIndex = -1 Then
            'No item selected
        Else
            If cmbConstraint1.SelectedItem.ToString = "" Then
                Main.Message.AddWarning("No constraints specified" & vbCrLf)
                Exit Sub
            Else
                If cmbConstraint1.SelectedItem.ToString = "WHERE" Then
                    If cmbField1.SelectedItem.ToString = "" Then

                    Else
                        txtQuery.Text = txtQuery.Text & " WHERE " & cmbField1.SelectedItem.ToString & " " & cmbType1.SelectedItem.ToString & " " & txtValue1.Text
                        If cmbType1.SelectedItem.ToString = "BETWEEN" Then
                            txtQuery.Text = txtQuery.Text & " AND " & txtSecondValue1.Text
                        End If
                    End If
                End If
                If cmbConstraint2.SelectedIndex = -1 Then
                    'No item selected.             
                Else
                    If cmbConstraint2.SelectedItem.ToString = "" Then
                    Else
                        If cmbConstraint2.SelectedItem.ToString = "AND" Then
                            If cmbField2.SelectedItem.ToString = "" Then

                            Else
                                txtQuery.Text = txtQuery.Text & " AND " & cmbField2.SelectedItem.ToString & " " & cmbType2.SelectedItem.ToString & " " & txtValue2.Text
                                If cmbType2.SelectedItem.ToString = "BETWEEN" Then
                                    txtQuery.Text = txtQuery.Text & " AND " & txtSecondValue2.Text
                                End If
                            End If
                        End If

                    End If
                End If
            End If
        End If

    End Sub

    Private Sub chkDate1_CheckedChanged(sender As Object, e As EventArgs) Handles chkDate1.CheckedChanged
        If chkDate1.Checked = True Then
            DateTimePicker3.Enabled = True
            Label6.Enabled = True
            If cmbType1.SelectedItem.ToString = "BETWEEN" Then
                DateTimePicker4.Enabled = True
            End If
        Else
            DateTimePicker3.Enabled = False
            Label6.Enabled = False
            DateTimePicker4.Enabled = False
        End If
    End Sub

    Private Sub chkDate2_CheckedChanged(sender As Object, e As EventArgs) Handles chkDate2.CheckedChanged
        If chkDate2.Checked = True Then
            DateTimePicker5.Enabled = True
            Label7.Enabled = True
            If cmbType2.SelectedItem.ToString = "BETWEEN" Then
                DateTimePicker6.Enabled = True
            End If
        Else
            DateTimePicker5.Enabled = False
            Label7.Enabled = False
            DateTimePicker6.Enabled = False
        End If
    End Sub

    Private Sub cmbType1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbType1.SelectedIndexChanged
        If cmbType1.SelectedItem.ToString = "BETWEEN" Then
            Label4.Enabled = True
            txtSecondValue1.Enabled = True
            If chkDate1.Checked = True Then
                DateTimePicker4.Enabled = True
            End If
        Else
            Label4.Enabled = False
            txtSecondValue1.Enabled = False
            DateTimePicker4.Enabled = False
        End If
    End Sub

    Private Sub cmbType2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbType2.SelectedIndexChanged
        If cmbType2.SelectedItem.ToString = "BETWEEN" Then
            Label5.Enabled = True
            txtSecondValue2.Enabled = True
            If chkDate2.Checked = True Then
                DateTimePicker6.Enabled = True
            End If
        Else
            Label5.Enabled = False
            txtSecondValue2.Enabled = False
            DateTimePicker6.Enabled = False
        End If
    End Sub

    Private Sub DateTimePicker3_GotFocus(sender As Object, e As EventArgs) Handles DateTimePicker3.GotFocus
        txtValue1.Text = "#" & Format(DateTimePicker3.Value, "MM-dd-yyyy") & "#"
    End Sub

    Private Sub DateTimePicker3_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker3.ValueChanged
        txtValue1.Text = "#" & Format(DateTimePicker3.Value, "MM-dd-yyyy") & "#"
    End Sub

    Private Sub DateTimePicker4_GotFocus(sender As Object, e As EventArgs) Handles DateTimePicker4.GotFocus
        txtSecondValue1.Text = "#" & Format(DateTimePicker4.Value, "MM-dd-yyyy") & "#"
    End Sub

    Private Sub DateTimePicker4_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker4.ValueChanged
        txtSecondValue1.Text = "#" & Format(DateTimePicker4.Value, "MM-dd-yyyy") & "#"
    End Sub

    Private Sub DateTimePicker5_GotFocus(sender As Object, e As EventArgs) Handles DateTimePicker5.GotFocus
        txtValue2.Text = "#" & Format(DateTimePicker5.Value, "MM-dd-yyyy") & "#"
    End Sub

    Private Sub DateTimePicker5_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker5.ValueChanged
        txtValue2.Text = "#" & Format(DateTimePicker5.Value, "MM-dd-yyyy") & "#"
    End Sub

    Private Sub DateTimePicker6_GotFocus(sender As Object, e As EventArgs) Handles DateTimePicker6.GotFocus
        txtSecondValue2.Text = "#" & Format(DateTimePicker6.Value, "MM-dd-yyyy") & "#"
    End Sub

    Private Sub DateTimePicker6_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker6.ValueChanged
        txtSecondValue2.Text = "#" & Format(DateTimePicker6.Value, "MM-dd-yyyy") & "#"
    End Sub

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Events - Events that can be triggered by this form." '--------------------------------------------------------------------------------------------------------------------------
#End Region 'Form Events ----------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class