Public Class frmMarkerProperties
    'This form is used to specify the Chart Marker Properties.

#Region " Variable Declarations - All the variables used in this form and this application." '=================================================================================================
    'Public myChart As DataVisualization.Charting.Chart
#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - All the properties used in this form and this application" '============================================================================================================

    Private _chart As DataVisualization.Charting.Chart 'The Chart containing the Markers to be modified.
    Property Chart As DataVisualization.Charting.Chart
        Get
            Return _chart
        End Get
        Set(value As DataVisualization.Charting.Chart)
            _chart = value
            UpdateSeriesList()
        End Set
    End Property

#End Region 'Properties -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Process XML files - Read and write XML files." '=====================================================================================================================================

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


#Region " Form Display Methods - Code used to display this form." '============================================================================================================================

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        RestoreFormSettings()   'Restore the form settings

        cmbMarkerStyle.Items.Clear()

        'myChart.Series(0).MarkerStyle = DataVisualization.Charting.MarkerStyle.Circle

        For Each item In [Enum].GetNames(GetType(DataVisualization.Charting.MarkerStyle))
            cmbMarkerStyle.Items.Add(item)
        Next

        UpdateSeriesList()

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form
        Me.Close() 'Close the form
    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        Else
            'Dont save settings if the form is minimised.
        End If
    End Sub



#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Open and Close Forms - Code used to open and close other forms." '===================================================================================================================

#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Methods - The main actions performed by this form." '===========================================================================================================================

    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
        'Apply the specified marker properties:

        'If myChart Is Nothing Then
        If Chart Is Nothing Then
            Main.Message.AddWarning("THe Chart has not been specified." & vbCrLf)
            Exit Sub
        End If

        Dim SeriesName As String = cmbSeriesName.SelectedItem.ToString
        'Main.Chart1.Series(SeriesName).MarkerBorderColor = txtBorderColor.BackColor
        'Main.Chart1.Series(SeriesName).MarkerBorderWidth = Int(Val(txtBorderWidth.Text))
        'Main.Chart1.Series(SeriesName).MarkerColor = txtMarkerColor.BackColor
        'Main.Chart1.Series(SeriesName).MarkerSize = Int(Val(txtMarkerSize.Text))
        'Main.Chart1.Series(SeriesName).MarkerStep = Int(Val(txtMarkerStep.Text))
        'Main.Chart1.Series(SeriesName).MarkerStyle = [Enum].Parse(GetType(DataVisualization.Charting.MarkerStyle), cmbMarkerStyle.SelectedItem.ToString)

        'myChart.Series(SeriesName).MarkerBorderColor = txtBorderColor.BackColor
        'myChart.Series(SeriesName).MarkerBorderWidth = Int(Val(txtBorderWidth.Text))
        'myChart.Series(SeriesName).MarkerColor = txtMarkerColor.BackColor
        'myChart.Series(SeriesName).MarkerSize = Int(Val(txtMarkerSize.Text))
        'myChart.Series(SeriesName).MarkerStep = Int(Val(txtMarkerStep.Text))
        'myChart.Series(SeriesName).MarkerStyle = [Enum].Parse(GetType(DataVisualization.Charting.MarkerStyle), cmbMarkerStyle.SelectedItem.ToString)

        Chart.Series(SeriesName).MarkerBorderColor = txtBorderColor.BackColor
        Chart.Series(SeriesName).MarkerBorderWidth = Int(Val(txtBorderWidth.Text))
        Chart.Series(SeriesName).MarkerColor = txtMarkerColor.BackColor
        Chart.Series(SeriesName).MarkerSize = Int(Val(txtMarkerSize.Text))
        Chart.Series(SeriesName).MarkerStep = Int(Val(txtMarkerStep.Text))
        Chart.Series(SeriesName).MarkerStyle = [Enum].Parse(GetType(DataVisualization.Charting.MarkerStyle), cmbMarkerStyle.SelectedItem.ToString)

        'Chart.Series(SeriesName).BorderColor = txtLineColor.BackColor
        Chart.Series(SeriesName).Color = txtSeriesColor.BackColor
    End Sub

    Public Sub ShowCurrentProperties()
        'Show the current marker properties:

        If cmbSeriesName.SelectedIndex = -1 Then
            If cmbSeriesName.Items.Count > 0 Then
                cmbSeriesName.SelectedIndex = 0
            Else
                Exit Sub
            End If
        End If

        Dim SeriesName As String = cmbSeriesName.SelectedItem.ToString

        'If Main.Chart1.Series.IndexOf(SeriesName) = -1 Then
        'If myChart.Series.IndexOf(SeriesName) = -1 Then
        If Chart.Series.IndexOf(SeriesName) = -1 Then
            'Series no found.
        Else
            'txtBorderColor.BackColor = Main.Chart1.Series(SeriesName).MarkerBorderColor
            'txtBorderWidth.Text = Main.Chart1.Series(SeriesName).MarkerBorderWidth
            'txtMarkerColor.BackColor = Main.Chart1.Series(SeriesName).MarkerColor
            'txtMarkerSize.Text = Main.Chart1.Series(SeriesName).MarkerSize
            'txtMarkerStep.Text = Main.Chart1.Series(SeriesName).MarkerStep
            'cmbMarkerStyle.SelectedIndex = cmbMarkerStyle.FindStringExact(Main.Chart1.Series(SeriesName).MarkerStyle.ToString)

            'txtBorderColor.BackColor = myChart.Series(SeriesName).MarkerBorderColor
            'txtBorderWidth.Text = myChart.Series(SeriesName).MarkerBorderWidth
            'txtMarkerColor.BackColor = myChart.Series(SeriesName).MarkerColor
            'txtMarkerSize.Text = myChart.Series(SeriesName).MarkerSize
            'txtMarkerStep.Text = myChart.Series(SeriesName).MarkerStep
            'cmbMarkerStyle.SelectedIndex = cmbMarkerStyle.FindStringExact(myChart.Series(SeriesName).MarkerStyle.ToString)

            If Chart.Series(SeriesName).MarkerBorderColor = Color.FromArgb(0) Then Chart.Series(SeriesName).MarkerBorderColor = Color.Gray
            txtBorderColor.BackColor = Chart.Series(SeriesName).MarkerBorderColor
            txtBorderWidth.Text = Chart.Series(SeriesName).MarkerBorderWidth
            If Chart.Series(SeriesName).MarkerColor = Color.Transparent Then Chart.Series(SeriesName).MarkerColor = Color.Gray
            If Chart.Series(SeriesName).MarkerColor = Color.FromArgb(0) Then Chart.Series(SeriesName).MarkerColor = Color.Gray
            txtMarkerColor.BackColor = Chart.Series(SeriesName).MarkerColor
            txtMarkerSize.Text = Chart.Series(SeriesName).MarkerSize
            txtMarkerStep.Text = Chart.Series(SeriesName).MarkerStep
            cmbMarkerStyle.SelectedIndex = cmbMarkerStyle.FindStringExact(Chart.Series(SeriesName).MarkerStyle.ToString)

            'txtLineColor.BackColor = Chart.Series(SeriesName).BorderColor
            If Chart.Series(SeriesName).Color = Color.FromArgb(0) Then Chart.Series(SeriesName).Color = Color.LightGray
            txtSeriesColor.BackColor = Chart.Series(SeriesName).Color
        End If

    End Sub

    Public Sub SelectSeries(ByVal Name As String)
        'Select the series with the specified name.
        cmbSeriesName.SelectedIndex = cmbSeriesName.FindStringExact(Name)
        ShowCurrentProperties()
    End Sub

    Public Sub SelectSeries(ByVal SeriesNo As Integer)
        'Select the series with the specified number
        Dim SeriesCount As Integer = Chart.Series.Count
        If SeriesNo + 1 > SeriesCount Then
            RaiseEvent ErrorMessage("The selected series number doea not exist." & vbCrLf)
        Else
            cmbSeriesName.SelectedIndex = SeriesNo
            ShowCurrentProperties()
        End If
    End Sub


    Public Sub UpdateSeriesList()
        cmbSeriesName.Items.Clear()
        ''For Each item In Main.Chart1.Series
        'If myChart Is Nothing Then
        'For Each item In Main.Chart1.Series
        If Chart Is Nothing Then
        Else
            'For Each item In myChart.Series
            For Each item In Chart.Series
                cmbSeriesName.Items.Add(item.Name)
            Next
        End If

        If cmbSeriesName.Items.Count > 0 Then
            cmbSeriesName.SelectedIndex = 0
        End If

    End Sub

    Private Sub txtBorderColor_MouseClick(sender As Object, e As MouseEventArgs) Handles txtBorderColor.MouseClick
        ColorDialog1.Color = txtBorderColor.BackColor
        ColorDialog1.ShowDialog()
        txtBorderColor.BackColor = ColorDialog1.Color
    End Sub

    Private Sub txtMarkerColor_Click(sender As Object, e As EventArgs) Handles txtMarkerColor.Click
        ColorDialog1.Color = txtMarkerColor.BackColor
        ColorDialog1.ShowDialog()
        txtMarkerColor.BackColor = ColorDialog1.Color
    End Sub

    Private Sub txtLineColor_MouseClick(sender As Object, e As MouseEventArgs) Handles txtSeriesColor.MouseClick
        ColorDialog1.Color = txtSeriesColor.BackColor
        ColorDialog1.ShowDialog()
        txtSeriesColor.BackColor = ColorDialog1.Color
    End Sub


#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region "Events" '--------------------------------------------------------------------------------------------------------

    Event ErrorMessage(ByVal Message As String) 'Send an error message.
    Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------


End Class