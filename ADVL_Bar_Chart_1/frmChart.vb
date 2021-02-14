Imports System.Windows.Forms.DataVisualization.Charting
Public Class frmChart
    'Chart form - displays a Bar Chart.

#Region " Variable Declarations - All the variables used in this form and this application." '=================================================================================================

    Public WithEvents MarkerProperties As frmMarkerProperties 'Used to edit the Chart marker properties.
    Public WithEvents ZoomChart As frmZoomChart 'Used to zoom the chart view.
    Public WithEvents ChartInfo As New ChartInfo 'Stores information about the chart. Contains methods to Save, Load and Clear the chart.

    Dim WithEvents Zip As ADVL_Utilities_Library_1.ZipComp

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - All the properties used in this form and this application" '============================================================================================================

    'The FormNo property stores the number of the instance of this form.
    'This form can have multipe instances, which are stored in the ChartList ArrayList in the ADVL_Chart_1 Main form.
    'When this form is closed, the FormNo is used to update the ClosedFormNo property of the Main form.
    'ClosedFormNo is then used by a method to set the corresponding form element in SharePricesList to Nothing.

    Private _formNo As Integer
    Public Property FormNo As Integer
        Get
            Return _formNo
        End Get
        Set(ByVal value As Integer)
            _formNo = value
        End Set
    End Property

    Private _chartName As String = "" 'A name assigned to this Chart.
    Public Property ChartName As String
        Get
            Return _chartName
        End Get
        Set(value As String)
            _chartName = value
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
                               <AutoDrawChart><%= chkAutoDraw.Checked %></AutoDrawChart>
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
            If Settings.<FormSettings>.<AutoDrawChart>.Value <> Nothing Then chkAutoDraw.Checked = Settings.<FormSettings>.<AutoDrawChart>.Value

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
        ChartInfo.DataLocation = Main.Project.DataLocn

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form

        'Close the Zoom form if it is open.
        If IsNothing(ZoomChart) Then
        Else
            ZoomChart.Close()
        End If


        'Close the Marker for if it is open:
        If IsNothing(MarkerProperties) Then
        Else
            MarkerProperties.Close()
        End If

        Main.ClosedFormNo = FormNo 'The Main form property ClosedFormNo is set to this form number. This is used in the SharePricesFormClosed method to select the correct form to set to nothing.
        Me.Close() 'Close the form
    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        Else
            'Dont save settings if the form is minimised.
        End If
    End Sub

    Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click
        'Open Line Chart file.

        'Find and open a Line Chart file:
        Select Case Main.Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                'Select a Line Chart file from the project directory:
                OpenFileDialog1.InitialDirectory = Main.Project.DataLocn.Path
                OpenFileDialog1.Filter = "Line Chart files | *.LineChart"
                If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                    Dim FileName As String = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
                    txtChartFileName.Text = FileName
                    Try
                        ChartInfo.LoadFile(FileName, Chart1)
                    Catch ex As Exception
                        Main.Message.AddWarning("Chart Form: ChartInfo.LoadFile error. FileName = " & FileName & vbCrLf)
                        Main.Message.AddWarning(ex.Message & vbCrLf & vbCrLf)
                    End Try

                    ChartInfo.ApplyQuery()
                    'UpdateInputDataTabSettings()
                    'UpdateTitlesTabSettings()
                    'UpdateAreasTabSettings() 'Update Areas Tab before Series Tab. This will update the list of Chart Areas on the Series Tab.
                    'UpdateSeriesTabSettings()
                    If chkAutoDraw.Checked Then DrawLineChart()
                End If

            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                'Select a Line Chart file from the project archive:
                'Show the zip archive file selection form:
                Zip = New ADVL_Utilities_Library_1.ZipComp
                Zip.ArchivePath = Main.Project.DataLocn.Path
                Zip.SelectFile() 'Show the Select File form
                'Zip.SelectFileForm.ApplicationName = Project.ApplicationName
                Zip.SelectFileForm.ApplicationName = Main.Project.Application.Name
                Zip.SelectFileForm.SettingsLocn = Main.Project.SettingsLocn
                Zip.SelectFileForm.Show()
                Zip.SelectFileForm.RestoreFormSettings()
                Zip.SelectFileForm.FileExtensions = {".LineChart"}
                Zip.SelectFileForm.GetFileList()
                If Zip.SelectedFile <> "" Then
                    'A file has been selected
                    txtChartFileName.Text = Zip.SelectedFile
                    'ChartInfo.LoadFile(Zip.SelectedFile, Chart1)
                    Try
                        ChartInfo.LoadFile(Zip.SelectedFile, Chart1)
                    Catch ex As Exception
                        Main.Message.AddWarning("Chart Form: ChartInfo.LoadFile error. Zip.SelectedFile = " & Zip.SelectedFile & vbCrLf)
                        Main.Message.AddWarning(ex.Message & vbCrLf & vbCrLf)
                    End Try
                    ChartInfo.ApplyQuery()
                    'UpdateInputDataTabSettings()
                    'UpdateTitlesTabSettings()
                    'UpdateAreasTabSettings()  'Update Areas Tab before Series Tab. This will update the list of Chart Areas on the Series Tab.
                    'UpdateSeriesTabSettings()
                    If chkAutoDraw.Checked Then DrawLineChart()
                End If
        End Select
    End Sub

#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Open and Close Forms - Code used to open and close other forms." '===================================================================================================================

    Private Sub btnMarkerProps2_Click(sender As Object, e As EventArgs) Handles btnMarkerProps2.Click
        'Open the Marker Properties form.
        If IsNothing(MarkerProperties) Then
            MarkerProperties = New frmMarkerProperties
            MarkerProperties.Show()
            'MarkerProperties.myChart = Chart1
            MarkerProperties.Chart = Chart1
            MarkerProperties.SelectSeries(0)
            'MarkerProperties.ShowCurrentProperties()
            'MarkerProperties.SelectSeries(txtSeriesName.Text)
        Else
            MarkerProperties.Show()
            'MarkerProperties.ShowCurrentProperties()
            'MarkerProperties.SelectSeries(txtSeriesName.Text)
        End If
    End Sub

    Private Sub MarkerProperties_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MarkerProperties.FormClosed
        MarkerProperties = Nothing
    End Sub

    Private Sub btnZoomChart_Click(sender As Object, e As EventArgs) Handles btnZoomChart.Click
        'Open the Zoom Chart form.
        If IsNothing(ZoomChart) Then
            ZoomChart = New frmZoomChart
            ZoomChart.Show()
            ZoomChart.Chart = Chart1
            ZoomChart.SelectAxis(0, "X Axis")
        Else
            ZoomChart.Show()
            ZoomChart.SelectAxis(0, "X Axis")
        End If

    End Sub

    Private Sub ZoomChart_FormClosed(sender As Object, e As FormClosedEventArgs) Handles ZoomChart.FormClosed
        ZoomChart = Nothing
    End Sub

#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Methods - The main actions performed by this form." '===========================================================================================================================

    Private Sub DrawLineChart()
        'Draw the Line Chart:
        Try
            Dim SeriesName As String
            Dim ChartArea As String
            For Each item In Chart1.Series
                SeriesName = item.Name
                ChartArea = Chart1.Series(SeriesName).ChartArea
                If ChartInfo.dictAreaInfo(ChartArea).AutoXAxisMinimum Then Chart1.ChartAreas(ChartArea).AxisX.Minimum = Double.NaN
                If ChartInfo.dictAreaInfo(ChartArea).AutoXAxisMaximum Then Chart1.ChartAreas(ChartArea).AxisX.Maximum = Double.NaN
                If ChartInfo.dictAreaInfo(ChartArea).AutoXAxisMajorGridInterval Then Chart1.ChartAreas(ChartArea).AxisX.MajorGrid.Interval = Double.NaN
                Chart1.ChartAreas(ChartArea).AxisX.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
                If item.ChartType = DataVisualization.Charting.SeriesChartType.Line Then
                    Chart1.Series(SeriesName).Points.DataBindXY(ChartInfo.ds.Tables(0).DefaultView, ChartInfo.dictSeriesInfo(SeriesName).XValuesFieldName, ChartInfo.ds.Tables(0).DefaultView, ChartInfo.dictSeriesInfo(SeriesName).YValuesFieldName)
                End If
                'If ChartInfo.dictAreaInfo(ChartArea).AutoXAxisMinimum Then Chart1.ChartAreas(ChartArea).AxisX.Minimum = Double.NaN
                'If ChartInfo.dictAreaInfo(ChartArea).AutoXAxisMaximum Then Chart1.ChartAreas(ChartArea).AxisX.Maximum = Double.NaN
                'If ChartInfo.dictAreaInfo(ChartArea).AutoXAxisMajorGridInterval Then Chart1.ChartAreas(ChartArea).AxisX.MajorGrid.Interval = Double.NaN
                'Chart1.ChartAreas(ChartArea).AxisX.Interval = 0
                'Chart1.ChartAreas(ChartArea).AxisX.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
            Next
        Catch ex As Exception
            Main.Message.AddWarning(ex.Message & vbCrLf)
        End Try
    End Sub

    Private Sub btnDrawChart_Click(sender As Object, e As EventArgs) Handles btnDrawChart.Click
        DrawLineChart()
    End Sub

    Private Sub ChartInfo_Message(Message As String) Handles ChartInfo.Message
        Main.Message.Add(Message)
    End Sub

    Private Sub ChartInfo_ErrorMessage(Message As String) Handles ChartInfo.ErrorMessage
        Main.Message.AddWarning(Message)
    End Sub

    Private Sub Chart1_AxisViewChanged(sender As Object, e As ViewEventArgs) Handles Chart1.AxisViewChanged
        If IsNothing(ZoomChart) Then
            ' Message.Add("ZoomChart is Nothing" & vbCrLf)
        Else
            ZoomChart.UpdateSettings() 'Update the Zoom settings. These may have changed if the chart was scrolled.
            'Message.Add("ZoomChart.UpdateSettings()" & vbCrLf)
        End If
    End Sub







#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class