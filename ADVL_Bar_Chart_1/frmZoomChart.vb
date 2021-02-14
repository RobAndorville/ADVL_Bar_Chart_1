Public Class frmZoomChart
    'This form is used to zoom the chart view.

#Region " Variable Declarations - All the variables used in this form and this application." '=================================================================================================

    Dim AxisMinValue As Double 'The minimum value of the selected axis
    Dim AxisMaxValue As Double 'The maximum value of the selected axis.
    Dim AxisValueRange As Double 'The value range of the selected axis.

    'The scrollbars use integer values. The axis values are rescaled to integer values on the scrollbars.
    Dim ScrollbarMin As Integer = 0       'The minimum integer value for the scrollbars
    Dim ScrollbarMax As Integer = 1000000 'The maximum integer value for the scrollbars
    Dim ScrollbarRange As Integer = ScrollbarMax - ScrollbarMin

    Dim AxisZoomFrom As Double 'The minimum axis value in the Zoomed view.
    Dim AxisZoomTo As Double   'The maximum axis value for the Zoomed view.
    Dim AxisZoomInterval As Double 'The axis value range of the Zoomed view.

    'Calculation variables:
    Dim NewAxisZoomFrom As Double     'The New From Value calculated from the Zoom From scroll bar. This value can not be larger than the IntervalMinValue if the Interval is locked.
    Dim IntervalMinValue As Double 'The minimum possible value using the specified Zoom Interval

    Dim NewAxisZoomTo As Double       'The New To Value calculated from the Zoom To scroll bar. This value can not be smaller than the IntervalMaxValue if the interval is locked.
    Dim IntervalMaxValue As Double 'The maximum possible value using the specified Zoom Interval

    Dim NewAxisZoomInterval As Double
    Dim TempValue As Integer 'Temporary scrollbar value. This will be adjusted if it is outside the scrollbar minimum and maximum values.

    Enum ScrollBar
        ScrollFrom
        ScrollTo
        ScrollInterval
        None
    End Enum

    Dim SelScrollBar As ScrollBar = ScrollBar.None 'Indicates the selected scrollbar

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - All the properties used in this form and this application" '============================================================================================================

    Private _chart As DataVisualization.Charting.Chart 'The Chart containing the Markers to be modified.
    Property Chart As DataVisualization.Charting.Chart
        Get
            Return _chart
        End Get
        Set(value As DataVisualization.Charting.Chart)
            _chart = value
            UpdateAreaList() 'Update the list of chart areas
        End Set
    End Property

    Private _areaName As String = ""
    Property AreaName As String
        Get
            Return _areaName
        End Get
        Set(value As String)
            _areaName = value
        End Set
    End Property

    Private _selectedAxis As String = "X Axis"
    Property SelectedAxis As String
        Get
            Return _selectedAxis
        End Get
        Set(value As String)
            _selectedAxis = value
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
                               <AreaName><%= AreaName %></AreaName>
                               <SelectedAxis><%= SelectedAxis %></SelectedAxis>
                               <LockFromValue><%= rbLockFrom.Checked %></LockFromValue>
                               <LockToValue><%= rbLockTo.Checked %></LockToValue>
                               <LockIntervalValue><%= rbLockInterval.Checked %></LockIntervalValue>
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
        rbLockTo.Checked = True
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



    Public Sub SelectAxis(ByVal AreaName As String, ByVal AxisName As String)

        If Chart Is Nothing Then
            RaiseEvent ErrorMessage("A Chart has not been specified." & vbCrLf)
            Exit Sub
        End If
        cmbAreaName.SelectedIndex = cmbAreaName.FindStringExact(AreaName)
        Me.AreaName = AreaName

        HScrollBar1.Minimum = ScrollbarMin
        HScrollBar1.Maximum = ScrollbarMax

        HScrollBar2.Minimum = ScrollbarMin
        HScrollBar2.Maximum = ScrollbarMax

        HScrollBar3.Minimum = ScrollbarMin
        HScrollBar3.Maximum = ScrollbarMax

        Select Case AxisName
            Case "X Axis"
                SelectedAxis = "X Axis"
                rbXAxis.Checked = True

                AxisMinValue = Chart.ChartAreas(AreaName).AxisX.Minimum
                AxisMaxValue = Chart.ChartAreas(AreaName).AxisX.Maximum
                AxisValueRange = AxisMaxValue - AxisMinValue

                AxisZoomFrom = Chart.ChartAreas(AreaName).AxisX.ScaleView.ViewMinimum
                txtAxisZoomFrom.Text = AxisZoomFrom
                HScrollBar1.Value = Int(((AxisZoomFrom - AxisMinValue) / AxisValueRange) * ScrollbarRange)

                AxisZoomTo = Chart.ChartAreas(AreaName).AxisX.ScaleView.ViewMaximum
                txtAxisZoomTo.Text = AxisZoomTo
                HScrollBar2.Value = Int(((AxisZoomTo - AxisMinValue) / AxisValueRange) * ScrollbarRange)

                AxisZoomInterval = AxisZoomTo - AxisZoomFrom
                IntervalMinValue = AxisMaxValue - AxisZoomInterval
                txtAxisZoomInterval.Text = AxisZoomInterval
                HScrollBar3.Value = Int((AxisZoomInterval / AxisValueRange) * ScrollbarRange)

            Case "X2 Axis"
                SelectedAxis = "X2 Axis"
                rbX2Axis.Checked = True

                AxisMinValue = Chart.ChartAreas(AreaName).AxisX2.Minimum
                AxisMaxValue = Chart.ChartAreas(AreaName).AxisX2.Maximum
                AxisValueRange = AxisMaxValue - AxisMinValue

                AxisZoomFrom = Chart.ChartAreas(AreaName).AxisX2.ScaleView.ViewMinimum
                txtAxisZoomFrom.Text = AxisZoomFrom
                HScrollBar1.Value = Int(((AxisZoomFrom - AxisMinValue) / AxisValueRange) * ScrollbarRange)

                AxisZoomTo = Chart.ChartAreas(AreaName).AxisX2.ScaleView.ViewMaximum
                txtAxisZoomTo.Text = AxisZoomTo
                HScrollBar2.Value = Int(((AxisZoomTo - AxisMinValue) / AxisValueRange) * ScrollbarRange)

                AxisZoomInterval = AxisZoomTo - AxisZoomFrom
                IntervalMinValue = AxisMaxValue - AxisZoomInterval
                txtAxisZoomInterval.Text = AxisZoomInterval
                HScrollBar3.Value = Int((AxisZoomInterval / AxisValueRange) * ScrollbarRange)

            Case "Y Axis"
                SelectedAxis = "Y Axis"
                rbYAxis.Checked = True

                AxisMinValue = Chart.ChartAreas(AreaName).AxisY.Minimum
                AxisMaxValue = Chart.ChartAreas(AreaName).AxisY.Maximum
                AxisValueRange = AxisMaxValue - AxisMinValue

                AxisZoomFrom = Chart.ChartAreas(AreaName).AxisY.ScaleView.ViewMinimum
                txtAxisZoomFrom.Text = AxisZoomFrom
                HScrollBar1.Value = Int(((AxisZoomFrom - AxisMinValue) / AxisValueRange) * ScrollbarRange)

                AxisZoomTo = Chart.ChartAreas(AreaName).AxisY.ScaleView.ViewMaximum
                txtAxisZoomTo.Text = AxisZoomTo
                HScrollBar2.Value = Int(((AxisZoomTo - AxisMinValue) / AxisValueRange) * ScrollbarRange)

                AxisZoomInterval = AxisZoomTo - AxisZoomFrom
                IntervalMinValue = AxisMaxValue - AxisZoomInterval
                txtAxisZoomInterval.Text = AxisZoomInterval
                HScrollBar3.Value = Int((AxisZoomInterval / AxisValueRange) * ScrollbarRange)

            Case "Y2 Axis"
                SelectedAxis = "Y2 Axis"
                rbY2Axis.Checked = True

                AxisMinValue = Chart.ChartAreas(AreaName).AxisY2.Minimum
                AxisMaxValue = Chart.ChartAreas(AreaName).AxisY2.Maximum
                AxisValueRange = AxisMaxValue - AxisMinValue

                AxisZoomFrom = Chart.ChartAreas(AreaName).AxisY2.ScaleView.ViewMinimum
                txtAxisZoomFrom.Text = AxisZoomFrom
                HScrollBar1.Value = Int((((AxisZoomFrom - AxisMinValue) - AxisMinValue) / AxisValueRange) * ScrollbarRange)

                AxisZoomTo = Chart.ChartAreas(AreaName).AxisY2.ScaleView.ViewMaximum
                txtAxisZoomTo.Text = AxisZoomTo
                HScrollBar2.Value = Int(((AxisZoomTo - AxisMinValue) / AxisValueRange) * ScrollbarRange)

                AxisZoomInterval = AxisZoomTo - AxisZoomFrom
                IntervalMinValue = AxisMaxValue - AxisZoomInterval
                txtAxisZoomInterval.Text = AxisZoomInterval
                HScrollBar3.Value = Int((AxisZoomInterval / AxisValueRange) * ScrollbarRange)
        End Select

    End Sub

    Public Sub SelectAxis(ByVal AreaNo As Integer, ByVal AxisName As String)

        If Chart Is Nothing Then
            RaiseEvent ErrorMessage("A Chart has not been specified." & vbCrLf)
            Exit Sub
        End If

        Dim AreaCount As Integer = Chart.ChartAreas.Count
        If AreaNo + 1 > AreaCount Then
            RaiseEvent ErrorMessage("The selected Area number doea not exist." & vbCrLf)
            Exit Sub
        End If

        cmbAreaName.SelectedIndex = AreaNo
        Me.AreaName = AreaName

        HScrollBar1.Minimum = ScrollbarMin
        HScrollBar1.Maximum = ScrollbarMax

        HScrollBar2.Minimum = ScrollbarMin
        HScrollBar2.Maximum = ScrollbarMax

        HScrollBar3.Minimum = ScrollbarMin
        HScrollBar3.Maximum = ScrollbarMax

        Select Case AxisName
            Case "X Axis"
                SelectedAxis = "X Axis"
                rbXAxis.Checked = True

                AxisMinValue = Chart.ChartAreas(AreaName).AxisX.Minimum
                AxisMaxValue = Chart.ChartAreas(AreaName).AxisX.Maximum
                AxisValueRange = AxisMaxValue - AxisMinValue

                AxisZoomFrom = Chart.ChartAreas(AreaName).AxisX.ScaleView.ViewMinimum
                txtAxisZoomFrom.Text = AxisZoomFrom
                HScrollBar1.Value = Int(((AxisZoomFrom - AxisMinValue) / AxisValueRange) * ScrollbarRange)

                AxisZoomTo = Chart.ChartAreas(AreaName).AxisX.ScaleView.ViewMaximum
                txtAxisZoomTo.Text = AxisZoomTo
                HScrollBar2.Value = Int(((AxisZoomTo - AxisMinValue) / AxisValueRange) * ScrollbarRange)

                AxisZoomInterval = AxisZoomTo - AxisZoomFrom
                IntervalMinValue = AxisMaxValue - AxisZoomInterval
                txtAxisZoomInterval.Text = AxisZoomInterval
                HScrollBar3.Value = Int((AxisZoomInterval / AxisValueRange) * ScrollbarRange)

            Case "X2 Axis"
                SelectedAxis = "X2 Axis"
                rbX2Axis.Checked = True

                AxisMinValue = Chart.ChartAreas(AreaName).AxisX2.Minimum
                AxisMaxValue = Chart.ChartAreas(AreaName).AxisX2.Maximum
                AxisValueRange = AxisMaxValue - AxisMinValue

                AxisZoomFrom = Chart.ChartAreas(AreaName).AxisX2.ScaleView.ViewMinimum
                txtAxisZoomFrom.Text = AxisZoomFrom
                HScrollBar1.Value = Int(((AxisZoomFrom - AxisMinValue) / AxisValueRange) * ScrollbarRange)

                AxisZoomTo = Chart.ChartAreas(AreaName).AxisX2.ScaleView.ViewMaximum
                txtAxisZoomTo.Text = AxisZoomTo
                HScrollBar2.Value = Int(((AxisZoomTo - AxisMinValue) / AxisValueRange) * ScrollbarRange)

                AxisZoomInterval = AxisZoomTo - AxisZoomFrom
                IntervalMinValue = AxisMaxValue - AxisZoomInterval
                txtAxisZoomInterval.Text = AxisZoomInterval
                HScrollBar3.Value = Int((AxisZoomInterval / AxisValueRange) * ScrollbarRange)

            Case "Y Axis"
                SelectedAxis = "Y Axis"
                rbYAxis.Checked = True

                AxisMinValue = Chart.ChartAreas(AreaName).AxisY.Minimum
                AxisMaxValue = Chart.ChartAreas(AreaName).AxisY.Maximum
                AxisValueRange = AxisMaxValue - AxisMinValue

                AxisZoomFrom = Chart.ChartAreas(AreaName).AxisY.ScaleView.ViewMinimum
                txtAxisZoomFrom.Text = AxisZoomFrom
                HScrollBar1.Value = Int(((AxisZoomFrom - AxisMinValue) / AxisValueRange) * ScrollbarRange)

                AxisZoomTo = Chart.ChartAreas(AreaName).AxisY.ScaleView.ViewMaximum
                txtAxisZoomTo.Text = AxisZoomTo
                HScrollBar2.Value = Int(((AxisZoomTo - AxisMinValue) / AxisValueRange) * ScrollbarRange)

                AxisZoomInterval = AxisZoomTo - AxisZoomFrom
                IntervalMinValue = AxisMaxValue - AxisZoomInterval
                txtAxisZoomInterval.Text = AxisZoomInterval
                HScrollBar3.Value = Int((AxisZoomInterval / AxisValueRange) * ScrollbarRange)

            Case "Y2 Axis"
                SelectedAxis = "Y2 Axis"
                rbY2Axis.Checked = True

                AxisMinValue = Chart.ChartAreas(AreaName).AxisY2.Minimum
                AxisMaxValue = Chart.ChartAreas(AreaName).AxisY2.Maximum
                AxisValueRange = AxisMaxValue - AxisMinValue

                AxisZoomFrom = Chart.ChartAreas(AreaName).AxisY2.ScaleView.ViewMinimum
                txtAxisZoomFrom.Text = AxisZoomFrom
                HScrollBar1.Value = Int((((AxisZoomFrom - AxisMinValue) - AxisMinValue) / AxisValueRange) * ScrollbarRange)

                AxisZoomTo = Chart.ChartAreas(AreaName).AxisY2.ScaleView.ViewMaximum
                txtAxisZoomTo.Text = AxisZoomTo
                HScrollBar2.Value = Int(((AxisZoomTo - AxisMinValue) / AxisValueRange) * ScrollbarRange)

                AxisZoomInterval = AxisZoomTo - AxisZoomFrom
                IntervalMinValue = AxisMaxValue - AxisZoomInterval
                txtAxisZoomInterval.Text = AxisZoomInterval
                HScrollBar3.Value = Int((AxisZoomInterval / AxisValueRange) * ScrollbarRange)
        End Select

    End Sub

    Private Sub UpdateAreaList()

        If Chart Is Nothing Then
        Else
            cmbAreaName.Items.Clear()
            For Each item In Chart.ChartAreas
                cmbAreaName.Items.Add(item.Name)
            Next
        End If
    End Sub

    Private Sub HScrollBar1_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBar1.Scroll
        'The Zoom From Values has changed:

        If SelScrollBar = ScrollBar.ScrollFrom Then
            If rbLockFrom.Checked Then
                rbLockTo.Checked = True
            End If
            'This will leave the Interval locked is it is already locked.
        End If

        If rbLockFrom.Checked Then
            'The From value is locked. Do not change it.
        Else
            If rbLockInterval.Checked Then 'Adjust the Zoom To value
                NewAxisZoomFrom = AxisMinValue + (e.NewValue / ScrollbarRange) * AxisValueRange 'Calculate the new Zoom From value from the ScrollBar value.
                If NewAxisZoomFrom > IntervalMinValue Then 'Use IntervalMinValue:
                    AxisZoomFrom = IntervalMinValue
                    AxisZoomTo = AxisMaxValue 'When AxisZoomFrom = IntervalMinValue, AxisZoomTo = AxisMaxValue.
                    txtAxisZoomTo.Text = AxisZoomTo
                    TempValue = Int(((AxisZoomFrom - AxisMinValue) / AxisValueRange) * ScrollbarRange) 'Calculate the ScrollBar value from the new Zoom From value.
                    If TempValue < HScrollBar1.Minimum Then TempValue = HScrollBar1.Minimum
                    If TempValue > HScrollBar1.Maximum Then TempValue = HScrollBar1.Maximum
                    HScrollBar1.Value = TempValue 'Re-position the scrollbar
                Else 'OK to use NewAxisZoomFrom:
                    AxisZoomFrom = NewAxisZoomFrom
                    AxisZoomTo = AxisZoomFrom + AxisZoomInterval
                    txtAxisZoomTo.Text = AxisZoomTo
                    TempValue = Int(((AxisZoomTo - AxisMinValue) / AxisValueRange) * ScrollbarRange)
                    If TempValue < HScrollBar1.Minimum Then TempValue = HScrollBar1.Minimum
                    If TempValue > HScrollBar1.Maximum Then TempValue = HScrollBar1.Maximum
                    HScrollBar2.Value = TempValue 'Re-position the scrollbar
                End If

            Else 'Adjust the Zoom Interval value
                NewAxisZoomFrom = AxisMinValue + (e.NewValue / ScrollbarRange) * AxisValueRange
                If NewAxisZoomFrom > AxisZoomTo Then
                    AxisZoomFrom = AxisZoomTo
                    TempValue = Int(((AxisZoomFrom - AxisMinValue) / AxisValueRange) * ScrollbarRange)
                    If TempValue < HScrollBar1.Minimum Then TempValue = HScrollBar1.Minimum
                    If TempValue > HScrollBar1.Maximum Then TempValue = HScrollBar1.Maximum
                    HScrollBar1.Value = TempValue 'Re-position the scrollbar
                Else
                    AxisZoomFrom = NewAxisZoomFrom
                End If
                AxisZoomInterval = AxisZoomTo - AxisZoomFrom
                IntervalMinValue = AxisMaxValue - AxisZoomInterval
                IntervalMaxValue = AxisMinValue + AxisZoomInterval
                txtAxisZoomInterval.Text = AxisZoomInterval
                TempValue = Int((AxisZoomInterval / AxisValueRange) * ScrollbarRange)
                If TempValue < HScrollBar3.Minimum Then TempValue = HScrollBar3.Minimum
                If TempValue > HScrollBar3.Maximum Then TempValue = HScrollBar3.Maximum
                HScrollBar3.Value = TempValue
            End If

        End If
        txtAxisZoomFrom.Text = AxisZoomFrom
        ZoomChart()
    End Sub

    Private Sub HScrollBar2_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBar2.Scroll
        'The Zoom To Values has changed:

        If SelScrollBar = ScrollBar.ScrollTo Then
            If rbLockTo.Checked Then
                rbLockFrom.Checked = True
            End If
            'This will leave the Interval locked is it is already locked.
        End If

        If rbLockTo.Checked Then
            'The To value is locked. Do not change it.
        Else
            If rbLockInterval.Checked Then 'Adjust the Zoom From value
                NewAxisZoomTo = AxisMinValue + (e.NewValue / ScrollbarRange) * AxisValueRange
                If NewAxisZoomTo < IntervalMaxValue Then 'Use IntervalMaxValue:
                    AxisZoomTo = IntervalMaxValue
                    AxisZoomFrom = AxisMinValue 'When AxisZoomTo = IntervalMaxValue, AxisZoomFrom = AxisMinValue.
                    txtAxisZoomFrom.Text = AxisZoomFrom
                    TempValue = Int(((AxisZoomTo - AxisMinValue) / AxisValueRange) * ScrollbarRange)
                    If TempValue < HScrollBar2.Minimum Then TempValue = HScrollBar2.Minimum
                    If TempValue > HScrollBar2.Maximum Then TempValue = HScrollBar2.Maximum
                    HScrollBar2.Value = TempValue 'Re-position the scrollbar
                Else 'OK to use NewAxisZoomTo:
                    AxisZoomTo = NewAxisZoomTo
                    AxisZoomFrom = AxisZoomTo - AxisZoomInterval
                    txtAxisZoomFrom.Text = AxisZoomFrom
                    TempValue = Int(((AxisZoomFrom - AxisMinValue) / AxisValueRange) * ScrollbarRange)
                    If TempValue < HScrollBar2.Minimum Then TempValue = HScrollBar2.Minimum
                    If TempValue > HScrollBar2.Maximum Then TempValue = HScrollBar2.Maximum
                    HScrollBar1.Value = TempValue 'Re-position the scrollbar
                End If

            Else 'Adjust the Zoom Interval value
                NewAxisZoomTo = AxisMinValue + (e.NewValue / ScrollbarRange) * AxisValueRange
                If NewAxisZoomTo < AxisZoomFrom Then
                    AxisZoomTo = AxisZoomFrom
                    TempValue = Int(((AxisZoomTo - AxisMinValue) / AxisValueRange) * ScrollbarRange)
                    If TempValue < HScrollBar2.Minimum Then TempValue = HScrollBar2.Minimum
                    If TempValue > HScrollBar2.Maximum Then TempValue = HScrollBar2.Maximum
                    HScrollBar2.Value = TempValue 'Re-position the scrollbar
                Else
                    AxisZoomTo = NewAxisZoomTo
                End If
                AxisZoomInterval = AxisZoomTo - AxisZoomFrom
                IntervalMinValue = AxisMaxValue - AxisZoomInterval
                IntervalMaxValue = AxisMinValue + AxisZoomInterval
                txtAxisZoomInterval.Text = AxisZoomInterval
                TempValue = Int((AxisZoomInterval / AxisValueRange) * ScrollbarRange)
                If TempValue < HScrollBar3.Minimum Then TempValue = HScrollBar3.Minimum
                If TempValue > HScrollBar3.Maximum Then TempValue = HScrollBar3.Maximum
                HScrollBar3.Value = TempValue
            End If

        End If
        txtAxisZoomTo.Text = AxisZoomTo
        ZoomChart()
    End Sub

    Private Sub HScrollBar3_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBar3.Scroll
        'The Zoom Interval has changed:

        If SelScrollBar = ScrollBar.ScrollInterval Then
            If rbLockInterval.Checked Then
                rbLockFrom.Checked = True
            End If
        End If

        If rbLockInterval.Checked Then
            'The Interval value is locked. Do not change it.
        Else
            If rbLockFrom.Checked Then 'Adjust the Zoom To value: 
                NewAxisZoomInterval = (e.NewValue / ScrollbarRange) * AxisValueRange
                AxisZoomTo = AxisZoomFrom + NewAxisZoomInterval
                If AxisZoomTo > AxisMaxValue Then
                    AxisZoomTo = AxisMaxValue
                    AxisZoomInterval = AxisZoomTo - AxisZoomFrom
                Else
                    AxisZoomInterval = NewAxisZoomInterval
                End If
                txtAxisZoomTo.Text = AxisZoomTo
                TempValue = Int((AxisZoomInterval / AxisValueRange) * ScrollbarRange)
                If TempValue < HScrollBar3.Minimum Then TempValue = HScrollBar3.Minimum
                If TempValue > HScrollBar3.Maximum Then TempValue = HScrollBar3.Maximum
                HScrollBar3.Value = TempValue
                TempValue = Int(((AxisZoomTo - AxisMinValue) / AxisValueRange) * ScrollbarRange)
                If TempValue < HScrollBar2.Minimum Then TempValue = HScrollBar2.Minimum
                If TempValue > HScrollBar2.Maximum Then TempValue = HScrollBar2.Maximum
                HScrollBar2.Value = TempValue 'Re-position the scrollbar
            Else 'Adjust the Zoom From value:
                NewAxisZoomInterval = (e.NewValue / ScrollbarRange) * AxisValueRange
                AxisZoomFrom = AxisZoomTo - NewAxisZoomInterval
                If AxisZoomFrom < AxisMinValue Then
                    AxisZoomFrom = AxisMinValue
                    AxisZoomInterval = AxisZoomTo - AxisZoomFrom
                Else
                    AxisZoomInterval = NewAxisZoomInterval
                End If
                txtAxisZoomFrom.Text = AxisZoomFrom
                TempValue = Int((AxisZoomInterval / AxisValueRange) * ScrollbarRange)
                If TempValue < HScrollBar3.Minimum Then TempValue = HScrollBar3.Minimum
                If TempValue > HScrollBar3.Maximum Then TempValue = HScrollBar3.Maximum
                HScrollBar3.Value = TempValue
                TempValue = Int(((AxisZoomFrom - AxisMinValue) / AxisValueRange) * ScrollbarRange)
                If TempValue < HScrollBar1.Minimum Then TempValue = HScrollBar1.Minimum
                If TempValue > HScrollBar1.Maximum Then TempValue = HScrollBar1.Maximum
                HScrollBar1.Value = TempValue 'Re-position the scrollbar
            End If
        End If
        IntervalMinValue = AxisMaxValue - AxisZoomInterval
        IntervalMaxValue = AxisMinValue + AxisZoomInterval
        txtAxisZoomInterval.Text = AxisZoomInterval
        ZoomChart()
    End Sub

    Private Sub ZoomChart()
        'Apply the Zoom View settings to the chart:

        If Chart Is Nothing Then
            RaiseEvent ErrorMessage("A Chart has not been specified." & vbCrLf)
            Exit Sub
        End If

        Select Case SelectedAxis
            Case "X Axis"
                Chart.ChartAreas(AreaName).AxisX.ScaleView.Zoom(AxisZoomFrom, AxisZoomTo)
            Case "X2 Axis"
                Chart.ChartAreas(AreaName).AxisX2.ScaleView.Zoom(AxisZoomFrom, AxisZoomTo)
            Case "Y Axis"
                Chart.ChartAreas(AreaName).AxisY.ScaleView.Zoom(AxisZoomFrom, AxisZoomTo)
            Case "Y2 Axis"
                Chart.ChartAreas(AreaName).AxisY2.ScaleView.Zoom(AxisZoomFrom, AxisZoomTo)
        End Select
    End Sub


    Private Sub cmbAreaName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbAreaName.SelectedIndexChanged
        AreaName = cmbAreaName.SelectedItem.ToString
    End Sub

    Public Sub UpdateSettings()
        SelectAxis(AreaName, SelectedAxis)
    End Sub


    'Note: the following events do not work on the HScrollBars:
    '       MouseDown, MouseClick, Click, Enter, GotFocus
    'The MouseEnter and MouseLeave events are used to record the selected scrollbar in SelScrollBar (ScrollFrom, ScrollTo, ScrollInterval, None).

    Private Sub HScrollBar1_MouseEnter(sender As Object, e As EventArgs) Handles HScrollBar1.MouseEnter
        SelScrollBar = ScrollBar.ScrollFrom
    End Sub

    Private Sub HScrollBar2_MouseEnter(sender As Object, e As EventArgs) Handles HScrollBar2.MouseEnter
        SelScrollBar = ScrollBar.ScrollTo
    End Sub

    Private Sub HScrollBar3_MouseEnter(sender As Object, e As EventArgs) Handles HScrollBar3.MouseEnter
        SelScrollBar = ScrollBar.ScrollInterval
    End Sub

    Private Sub HScrollBar1_MouseLeave(sender As Object, e As EventArgs) Handles HScrollBar1.MouseLeave
        SelScrollBar = ScrollBar.None
    End Sub

    Private Sub HScrollBar2_MouseLeave(sender As Object, e As EventArgs) Handles HScrollBar2.MouseLeave
        SelScrollBar = ScrollBar.None
    End Sub

    Private Sub HScrollBar3_MouseLeave(sender As Object, e As EventArgs) Handles HScrollBar3.MouseLeave
        SelScrollBar = ScrollBar.None
    End Sub

    Private Sub rbXAxis_CheckedChanged(sender As Object, e As EventArgs) Handles rbXAxis.CheckedChanged
        If rbXAxis.Checked Then SelectAxis(AreaName, "X Axis")
    End Sub

    Private Sub rbX2Axis_CheckedChanged(sender As Object, e As EventArgs) Handles rbX2Axis.CheckedChanged
        If rbX2Axis.Checked Then SelectAxis(AreaName, "X2 Axis")
    End Sub

    Private Sub rbYAxis_CheckedChanged(sender As Object, e As EventArgs) Handles rbYAxis.CheckedChanged
        If rbYAxis.Checked Then SelectAxis(AreaName, "Y Axis")
    End Sub

    Private Sub rbY2Axis_CheckedChanged(sender As Object, e As EventArgs) Handles rbY2Axis.CheckedChanged
        If rbY2Axis.Checked Then SelectAxis(AreaName, "Y2 Axis")
    End Sub

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region "Events" '--------------------------------------------------------------------------------------------------------

    Event ErrorMessage(ByVal Message As String) 'Send an error message.
    Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------

End Class