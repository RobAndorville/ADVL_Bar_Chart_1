Public Class ChartInfo
    'The ChartInfo class stores information that is not stored within the Chart control.


    'Dataset used to hold points for plotting:
    Public ds As New DataSet

    Public dictSeriesInfo As New Dictionary(Of String, SeriesInfo) 'dictSeriesInfo is indexed using the Chart SeriesName. dictSeriesInfo contains information about each Series in the Chart: .XValuesFieldName, .YValuesFieldName, ChartArea. 

    Public dictAreaInfo As New Dictionary(Of String, AreaInfo) 'dictAreaInfo is indexed using the Chart Area Name. dictAreaInfo contains AutoMinimum, AutoMaximum and AutoMajorGridInterval settings for each axis in the ChartArea. (These are not stored in the chart control.)

    Public DataLocation As New ADVL_Utilities_Library_1.FileLocation 'Stores information about the data location in the Project - used to read the chart settings files.

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _fileName As String = "" 'The file name (with extension) of the chart settings. This file is stored in the Project.
    Property FileName As String
        Get
            Return _fileName
        End Get
        Set(value As String)
            _fileName = value
        End Set
    End Property

    Private _inputDataType As String = "Database" 'Database or Dataset
    Property InputDataType As String
        Get
            Return _inputDataType
        End Get
        Set(value As String)
            _inputDataType = value
        End Set
    End Property

    Private _inputDatabasePath As String = ""
    Property InputDatabasePath As String
        Get
            Return _inputDatabasePath
        End Get
        Set(value As String)
            _inputDatabasePath = value
        End Set
    End Property

    Private _inputQuery As String = ""
    Property InputQuery As String
        Get
            Return _inputQuery
        End Get
        Set(value As String)
            _inputQuery = value
        End Set
    End Property

    Private _inputDataDescr As String = "" 'A description of the data selected for charting.
    Property InputDataDescr As String
        Get
            Return _inputDataDescr
        End Get
        Set(value As String)
            _inputDataDescr = value
        End Set
    End Property


#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

    Public Sub LoadFile(ByRef myFileName As String, ByRef myChart As System.Windows.Forms.DataVisualization.Charting.Chart)
        'Load the Line Chart settings from the selected file.
        'This will update properties in ChartInfo and the myChart control.

        If myFileName.Trim = "" Then
            Exit Sub
        End If

        Dim XDoc As System.Xml.Linq.XDocument
        DataLocation.ReadXmlData(myFileName, XDoc)

        If XDoc Is Nothing Then
            RaiseEvent ErrorMessage("Xml list file is blank." & vbCrLf)
            Exit Sub
        End If

        'Restore Input Data settings:
        If XDoc.<ChartSettings>.<InputDataType>.Value <> Nothing Then InputDataType = XDoc.<ChartSettings>.<InputDataType>.Value
        If XDoc.<ChartSettings>.<InputDatabasePath>.Value <> Nothing Then InputDatabasePath = XDoc.<ChartSettings>.<InputDatabasePath>.Value
        If XDoc.<ChartSettings>.<InputQuery>.Value <> Nothing Then InputQuery = XDoc.<ChartSettings>.<InputQuery>.Value
        If XDoc.<ChartSettings>.<InputDataDescr>.Value <> Nothing Then InputDataDescr = XDoc.<ChartSettings>.<InputDataDescr>.Value

        'Restore Series Info: SeriesName, XValuesFieldName, YValuesFieldName:
        Dim SeriesInfo = From item In XDoc.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>
        Dim SeriesInfoName As String
        dictSeriesInfo.Clear() 'Clear the dictionary of Series Information. New Field entries will be added below.
        For Each item In SeriesInfo
            SeriesInfoName = item.<Name>.Value
            dictSeriesInfo.Add(SeriesInfoName, New SeriesInfo)
            dictSeriesInfo(SeriesInfoName).XValuesFieldName = item.<XValuesFieldName>.Value
            dictSeriesInfo(SeriesInfoName).YValuesFieldName = item.<YValuesFieldName>.Value
            If item.<ChartArea>.Value <> Nothing Then dictSeriesInfo(SeriesInfoName).ChartArea = item.<ChartArea>.Value
        Next

        'Restore Area Info: 
        Dim AreaInfo = From item In XDoc.<ChartSettings>.<AreaInfoList>.<AreaInfo>
        Dim AreaInfoName As String
        dictAreaInfo.Clear() 'Clear the dictionary of Chart Area Information. New Field entries will be added below.
        For Each item In AreaInfo
            AreaInfoName = item.<Name>.Value
            dictAreaInfo.Add(AreaInfoName, New AreaInfo)
            dictAreaInfo(AreaInfoName).AutoXAxisMinimum = item.<AutoXAxisMinimum>.Value
            dictAreaInfo(AreaInfoName).AutoXAxisMaximum = item.<AutoXAxisMaximum>.Value
            dictAreaInfo(AreaInfoName).AutoXAxisMajorGridInterval = item.<AutoXAxisMajorGridInterval>.Value
            dictAreaInfo(AreaInfoName).AutoX2AxisMinimum = item.<AutoX2AxisMinimum>.Value
            dictAreaInfo(AreaInfoName).AutoX2AxisMaximum = item.<AutoX2AxisMaximum>.Value
            dictAreaInfo(AreaInfoName).AutoX2AxisMajorGridInterval = item.<AutoX2AxisMajorGridInterval>.Value
            dictAreaInfo(AreaInfoName).AutoYAxisMinimum = item.<AutoYAxisMinimum>.Value
            dictAreaInfo(AreaInfoName).AutoYAxisMaximum = item.<AutoYAxisMaximum>.Value
            dictAreaInfo(AreaInfoName).AutoYAxisMajorGridInterval = item.<AutoYAxisMajorGridInterval>.Value
            dictAreaInfo(AreaInfoName).AutoY2AxisMinimum = item.<AutoY2AxisMinimum>.Value
            dictAreaInfo(AreaInfoName).AutoY2AxisMaximum = item.<AutoY2AxisMaximum>.Value
            dictAreaInfo(AreaInfoName).AutoY2AxisMajorGridInterval = item.<AutoY2AxisMajorGridInterval>.Value
        Next

        'Restore Titles:
        Dim TitlesInfo = From item In XDoc.<ChartSettings>.<TitlesCollection>.<Title>
        Dim TitleName As String
        Dim myFontStyle As FontStyle
        Dim myFontSize As Single
        myChart.Titles.Clear()
        For Each item In TitlesInfo
            TitleName = item.<Name>.Value
            myChart.Titles.Add(TitleName).Name = TitleName 'The name needs to be explicitly declared!
            myChart.Titles(TitleName).Text = item.<Text>.Value
            myChart.Titles(TitleName).TextOrientation = [Enum].Parse(GetType(DataVisualization.Charting.TextOrientation), item.<TextOrientation>.Value)
            myChart.Titles(TitleName).Alignment = [Enum].Parse(GetType(ContentAlignment), item.<Alignment>.Value)
            myChart.Titles(TitleName).ForeColor = Color.FromArgb(item.<ForeColor>.Value)
            myFontStyle = FontStyle.Regular
            If item.<Font>.<Bold>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Bold
            If item.<Font>.<Italic>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Italic
            If item.<Font>.<Strikeout>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Strikeout
            If item.<Font>.<Underline>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Underline
            myFontSize = item.<Font>.<Size>.Value
            myChart.Titles(TitleName).Font = New Font(item.<Font>.<Name>.Value, myFontSize, myFontStyle)
        Next

        'Restore Chart Series:
        Dim Series = From item In XDoc.<ChartSettings>.<SeriesCollection>.<Series>
        Dim SeriesName As String
        myChart.Series.Clear()
        For Each item In Series
            SeriesName = item.<Name>.Value
            myChart.Series.Add(SeriesName)
            'myChart.Series(SeriesName).ChartType = [Enum].Parse(GetType(DataVisualization.Charting.SeriesChartType), item.<Type>.Value)
            myChart.Series(SeriesName).ChartType = [Enum].Parse(GetType(DataVisualization.Charting.SeriesChartType), item.<ChartType>.Value)
            If item.<ChartArea>.Value <> Nothing Then myChart.Series(SeriesName).ChartArea = item.<ChartArea>.Value
            myChart.Series(SeriesName).Legend = item.<Legend>.Value
            myChart.Series(SeriesName).SetCustomProperty("EmptyPointValue", item.<EmptyPointValue>.Value)
            myChart.Series(SeriesName).SetCustomProperty("LabelStyle", item.<LabelStyle>.Value)
            myChart.Series(SeriesName).SetCustomProperty("PixelPointDepth", item.<PixelPointDepth>.Value)
            myChart.Series(SeriesName).SetCustomProperty("PixelPointGapDepth", item.<PixelPointGapDepth>.Value)
            myChart.Series(SeriesName).SetCustomProperty("ShowMarkerLines", item.<ShowMarkerLines>.Value)
            myChart.Series(SeriesName).AxisLabel = item.<AxisLabel>.Value
            myChart.Series(SeriesName).XAxisType = [Enum].Parse(GetType(DataVisualization.Charting.AxisType), item.<XAxisType>.Value)
            myChart.Series(SeriesName).YAxisType = [Enum].Parse(GetType(DataVisualization.Charting.AxisType), item.<YAxisType>.Value)
            If item.<XValueType>.Value <> Nothing Then myChart.Series(SeriesName).XValueType = [Enum].Parse(GetType(DataVisualization.Charting.ChartValueType), item.<XValueType>.Value)
            If item.<YValueType>.Value <> Nothing Then myChart.Series(SeriesName).YValueType = [Enum].Parse(GetType(DataVisualization.Charting.ChartValueType), item.<YValueType>.Value)
            If item.<Marker>.<BorderColor>.Value <> Nothing Then myChart.Series(SeriesName).MarkerBorderColor = Color.FromArgb(item.<Marker>.<BorderColor>.Value)
            If item.<Marker>.<BorderWidth>.Value <> Nothing Then myChart.Series(SeriesName).MarkerBorderWidth = item.<Marker>.<BorderWidth>.Value
            If item.<Marker>.<Color>.Value <> Nothing Then myChart.Series(SeriesName).MarkerColor = Color.FromArgb(item.<Marker>.<Color>.Value)
            If item.<Marker>.<Size>.Value <> Nothing Then myChart.Series(SeriesName).MarkerSize = item.<Marker>.<Size>.Value
            If item.<Marker>.<Step>.Value <> Nothing Then myChart.Series(SeriesName).MarkerStep = item.<Marker>.<Step>.Value
            If item.<Marker>.<Style>.Value <> Nothing Then myChart.Series(SeriesName).MarkerStyle = [Enum].Parse(GetType(DataVisualization.Charting.MarkerStyle), item.<Marker>.<Style>.Value)
            If item.<Color>.Value <> Nothing Then myChart.Series(SeriesName).Color = Color.FromArgb(item.<Color>.Value)
        Next

        'Restore Chart Areas:
        Dim Areas = From item In XDoc.<ChartSettings>.<ChartAreasCollection>.<ChartArea>
        Dim AreaName As String
        myChart.ChartAreas.Clear()
        For Each item In Areas
            AreaName = item.<Name>.Value
            myChart.ChartAreas.Add(AreaName)
            'AxisX Properties:
            myChart.ChartAreas(AreaName).AxisX.Title = item.<AxisX>.<Title>.<Text>.Value
            myChart.ChartAreas(AreaName).AxisX.TitleAlignment = [Enum].Parse(GetType(StringAlignment), item.<AxisX>.<Title>.<Alignment>.Value)
            myChart.ChartAreas(AreaName).AxisX.TitleForeColor = Color.FromArgb(item.<AxisX>.<Title>.<ForeColor>.Value)
            myFontStyle = FontStyle.Regular
            If item.<AxisX>.<Title>.<Font>.<Bold>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Bold
            If item.<AxisX>.<Title>.<Font>.<Italic>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Italic
            If item.<AxisX>.<Title>.<Font>.<Strikeout>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Strikeout
            If item.<AxisX>.<Title>.<Font>.<Underline>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Underline
            myFontSize = item.<AxisX>.<Title>.<Font>.<Size>.Value
            myChart.ChartAreas(AreaName).AxisX.TitleFont = New Font(item.<AxisX>.<Title>.<Font>.<Name>.Value, myFontSize, myFontStyle)
            If item.<AxisX>.<LabelStyleFormat>.Value <> Nothing Then myChart.ChartAreas(AreaName).AxisX.LabelStyle.Format = item.<AxisX>.<LabelStyleFormat>.Value
            myChart.ChartAreas(AreaName).AxisX.Minimum = item.<AxisX>.<Minimum>.Value

            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoXAxisMinimum Then myChart.ChartAreas(AreaName).AxisX.Minimum = Double.NaN 'Set to Auto Minimum
            End If


            myChart.ChartAreas(AreaName).AxisX.Maximum = item.<AxisX>.<Maximum>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoXAxisMaximum Then myChart.ChartAreas(AreaName).AxisX.Maximum = Double.NaN 'Set to Auto Maximum
            End If

            myChart.ChartAreas(AreaName).AxisX.LineWidth = item.<AxisX>.<LineWidth>.Value
            myChart.ChartAreas(AreaName).AxisX.Interval = item.<AxisX>.<Interval>.Value
            myChart.ChartAreas(AreaName).AxisX.IntervalOffset = item.<AxisX>.<IntervalOffset>.Value
            myChart.ChartAreas(AreaName).AxisX.Crossing = item.<AxisX>.<Crossing>.Value
            myChart.ChartAreas(AreaName).AxisX.MajorGrid.Interval = item.<AxisX>.<MajorGrid>.<Interval>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoXAxisMajorGridInterval Then myChart.ChartAreas(AreaName).AxisX.MajorGrid.Interval = 0 'Set to Auto Interval
            End If

            myChart.ChartAreas(AreaName).AxisX.MajorGrid.IntervalOffset = item.<AxisX>.<MajorGrid>.<IntervalOffset>.Value

            'AxisX2 Properties:
            myChart.ChartAreas(AreaName).AxisX2.Title = item.<AxisX2>.<Title>.<Text>.Value
            myChart.ChartAreas(AreaName).AxisX2.TitleAlignment = [Enum].Parse(GetType(StringAlignment), item.<AxisX2>.<Title>.<Alignment>.Value)
            myChart.ChartAreas(AreaName).AxisX2.TitleForeColor = Color.FromArgb(item.<AxisX2>.<Title>.<ForeColor>.Value)
            myFontStyle = FontStyle.Regular
            If item.<AxisX2>.<Title>.<Font>.<Bold>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Bold
            If item.<AxisX2>.<Title>.<Font>.<Italic>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Italic
            If item.<AxisX2>.<Title>.<Font>.<Strikeout>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Strikeout
            If item.<AxisX2>.<Title>.<Font>.<Underline>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Underline
            myFontSize = item.<AxisX2>.<Title>.<Font>.<Size>.Value
            myChart.ChartAreas(AreaName).AxisX2.TitleFont = New Font(item.<AxisX2>.<Title>.<Font>.<Name>.Value, myFontSize, myFontStyle)
            If item.<AxisX2>.<LabelStyleFormat>.Value <> Nothing Then myChart.ChartAreas(AreaName).AxisX2.LabelStyle.Format = item.<AxisX2>.<LabelStyleFormat>.Value
            myChart.ChartAreas(AreaName).AxisX2.Minimum = item.<AxisX2>.<Minimum>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoX2AxisMinimum Then myChart.ChartAreas(AreaName).AxisX2.Minimum = Double.NaN 'Set to Auto Minimum
            End If


            myChart.ChartAreas(AreaName).AxisX2.Maximum = item.<AxisX2>.<Maximum>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoX2AxisMaximum Then myChart.ChartAreas(AreaName).AxisX2.Maximum = Double.NaN 'Set to Auto Maximum
            End If

            myChart.ChartAreas(AreaName).AxisX2.LineWidth = item.<AxisX2>.<LineWidth>.Value
            myChart.ChartAreas(AreaName).AxisX2.Interval = item.<AxisX2>.<Interval>.Value
            myChart.ChartAreas(AreaName).AxisX2.IntervalOffset = item.<AxisX2>.<IntervalOffset>.Value
            myChart.ChartAreas(AreaName).AxisX2.Crossing = item.<AxisX2>.<Crossing>.Value
            myChart.ChartAreas(AreaName).AxisX2.MajorGrid.Interval = item.<AxisX2>.<MajorGrid>.<Interval>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoX2AxisMajorGridInterval Then myChart.ChartAreas(AreaName).AxisX2.MajorGrid.Interval = 0 'Set to Auto Interval
            End If

            myChart.ChartAreas(AreaName).AxisX2.MajorGrid.IntervalOffset = item.<AxisX2>.<MajorGrid>.<IntervalOffset>.Value

            'AxisY Properties:
            myChart.ChartAreas(AreaName).AxisY.Title = item.<AxisY>.<Title>.<Text>.Value
            myChart.ChartAreas(AreaName).AxisY.TitleAlignment = [Enum].Parse(GetType(StringAlignment), item.<AxisY>.<Title>.<Alignment>.Value)
            myChart.ChartAreas(AreaName).AxisY.TitleForeColor = Color.FromArgb(item.<AxisY>.<Title>.<ForeColor>.Value)
            myFontStyle = FontStyle.Regular
            If item.<AxisY>.<Title>.<Font>.<Bold>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Bold
            If item.<AxisY>.<Title>.<Font>.<Italic>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Italic
            If item.<AxisY>.<Title>.<Font>.<Strikeout>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Strikeout
            If item.<AxisY>.<Title>.<Font>.<Underline>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Underline
            myFontSize = item.<AxisY>.<Title>.<Font>.<Size>.Value
            myChart.ChartAreas(AreaName).AxisY.TitleFont = New Font(item.<AxisY>.<Title>.<Font>.<Name>.Value, myFontSize, myFontStyle)
            If item.<AxisY>.<LabelStyleFormat>.Value <> Nothing Then myChart.ChartAreas(AreaName).AxisY.LabelStyle.Format = item.<AxisY>.<LabelStyleFormat>.Value
            myChart.ChartAreas(AreaName).AxisY.Minimum = item.<AxisY>.<Minimum>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoYAxisMinimum Then myChart.ChartAreas(AreaName).AxisY.Minimum = Double.NaN 'Set to Auto Minimum
            End If


            myChart.ChartAreas(AreaName).AxisY.Maximum = item.<AxisY>.<Maximum>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoYAxisMaximum Then myChart.ChartAreas(AreaName).AxisY.Maximum = Double.NaN 'Set to Auto Maximum
            End If

            myChart.ChartAreas(AreaName).AxisY.LineWidth = item.<AxisY>.<LineWidth>.Value
            myChart.ChartAreas(AreaName).AxisY.Interval = item.<AxisY>.<Interval>.Value
            myChart.ChartAreas(AreaName).AxisY.IntervalOffset = item.<AxisY>.<IntervalOffset>.Value
            myChart.ChartAreas(AreaName).AxisY.Crossing = item.<AxisY>.<Crossing>.Value
            myChart.ChartAreas(AreaName).AxisY.MajorGrid.Interval = item.<AxisY>.<MajorGrid>.<Interval>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoYAxisMajorGridInterval Then myChart.ChartAreas(AreaName).AxisY.MajorGrid.Interval = 0 'Set to Auto Interval
            End If

            myChart.ChartAreas(AreaName).AxisY.MajorGrid.IntervalOffset = item.<AxisY>.<MajorGrid>.<IntervalOffset>.Value

            'AxisY2 Properties:
            myChart.ChartAreas(AreaName).AxisY2.Title = item.<AxisY2>.<Title>.<Text>.Value
            myChart.ChartAreas(AreaName).AxisY2.TitleAlignment = [Enum].Parse(GetType(StringAlignment), item.<AxisY2>.<Title>.<Alignment>.Value)
            myChart.ChartAreas(AreaName).AxisY2.TitleForeColor = Color.FromArgb(item.<AxisY2>.<Title>.<ForeColor>.Value)
            myFontStyle = FontStyle.Regular
            If item.<AxisY2>.<Title>.<Font>.<Bold>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Bold
            If item.<AxisY2>.<Title>.<Font>.<Italic>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Italic
            If item.<AxisY2>.<Title>.<Font>.<Strikeout>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Strikeout
            If item.<AxisY2>.<Title>.<Font>.<Underline>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Underline
            myFontSize = item.<AxisY2>.<Title>.<Font>.<Size>.Value
            myChart.ChartAreas(AreaName).AxisY2.TitleFont = New Font(item.<AxisY2>.<Title>.<Font>.<Name>.Value, myFontSize, myFontStyle)
            If item.<AxisY2>.<LabelStyleFormat>.Value <> Nothing Then myChart.ChartAreas(AreaName).AxisY2.LabelStyle.Format = item.<AxisY2>.<LabelStyleFormat>.Value
            myChart.ChartAreas(AreaName).AxisY2.Minimum = item.<AxisY2>.<Minimum>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoY2AxisMinimum Then myChart.ChartAreas(AreaName).AxisY2.Minimum = Double.NaN 'Set to Auto Minimum
            End If

            myChart.ChartAreas(AreaName).AxisY2.Maximum = item.<AxisY2>.<Maximum>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoY2AxisMaximum Then myChart.ChartAreas(AreaName).AxisY2.Maximum = Double.NaN 'Set to Auto Maximum
            End If

            myChart.ChartAreas(AreaName).AxisY2.LineWidth = item.<AxisY2>.<LineWidth>.Value
            myChart.ChartAreas(AreaName).AxisY2.Interval = item.<AxisY2>.<Interval>.Value
            myChart.ChartAreas(AreaName).AxisY2.IntervalOffset = item.<AxisY2>.<IntervalOffset>.Value
            myChart.ChartAreas(AreaName).AxisY2.Crossing = item.<AxisY2>.<Crossing>.Value
            myChart.ChartAreas(AreaName).AxisY2.MajorGrid.Interval = item.<AxisY2>.<MajorGrid>.<Interval>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoY2AxisMajorGridInterval Then myChart.ChartAreas(AreaName).AxisY2.MajorGrid.Interval = 0 'Set to Auto Interval
            End If

            myChart.ChartAreas(AreaName).AxisY2.MajorGrid.IntervalOffset = item.<AxisY2>.<MajorGrid>.<IntervalOffset>.Value
        Next

    End Sub

    Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument, ByRef myChart As System.Windows.Forms.DataVisualization.Charting.Chart)
        'Load the Line Chart settings from the XDocument.
        'This will update properties in ChartInfo and the myChart control.

        If XDoc Is Nothing Then
            RaiseEvent ErrorMessage("Xml list file is blank." & vbCrLf)
            Exit Sub
        End If

        'Restore Input Data settings:
        If XDoc.<ChartSettings>.<InputDataType>.Value <> Nothing Then InputDataType = XDoc.<ChartSettings>.<InputDataType>.Value
        If XDoc.<ChartSettings>.<InputDatabasePath>.Value <> Nothing Then InputDatabasePath = XDoc.<ChartSettings>.<InputDatabasePath>.Value
        If XDoc.<ChartSettings>.<InputQuery>.Value <> Nothing Then InputQuery = XDoc.<ChartSettings>.<InputQuery>.Value
        If XDoc.<ChartSettings>.<InputDataDescr>.Value <> Nothing Then InputDataDescr = XDoc.<ChartSettings>.<InputDataDescr>.Value

        'Restore Series Info: SeriesName, XValuesFieldName, YValuesFieldName:
        Dim SeriesInfo = From item In XDoc.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>
        Dim SeriesInfoName As String
        dictSeriesInfo.Clear() 'Clear the dictionary of Series Information. New Field entries will be added below.
        For Each item In SeriesInfo
            SeriesInfoName = item.<Name>.Value
            dictSeriesInfo.Add(SeriesInfoName, New SeriesInfo)
            dictSeriesInfo(SeriesInfoName).XValuesFieldName = item.<XValuesFieldName>.Value
            dictSeriesInfo(SeriesInfoName).YValuesFieldName = item.<YValuesFieldName>.Value
            If item.<ChartArea>.Value <> Nothing Then dictSeriesInfo(SeriesInfoName).ChartArea = item.<ChartArea>.Value
        Next

        'Restore Area Info: 
        Dim AreaInfo = From item In XDoc.<ChartSettings>.<AreaInfoList>.<AreaInfo>
        Dim AreaInfoName As String
        dictAreaInfo.Clear() 'Clear the dictionary of Chart Area Information. New Field entries will be added below.
        For Each item In AreaInfo
            AreaInfoName = item.<Name>.Value
            dictAreaInfo.Add(AreaInfoName, New AreaInfo)
            dictAreaInfo(AreaInfoName).AutoXAxisMinimum = item.<AutoXAxisMinimum>.Value
            dictAreaInfo(AreaInfoName).AutoXAxisMaximum = item.<AutoXAxisMaximum>.Value
            dictAreaInfo(AreaInfoName).AutoXAxisMajorGridInterval = item.<AutoXAxisMajorGridInterval>.Value
            dictAreaInfo(AreaInfoName).AutoX2AxisMinimum = item.<AutoX2AxisMinimum>.Value
            dictAreaInfo(AreaInfoName).AutoX2AxisMaximum = item.<AutoX2AxisMaximum>.Value
            dictAreaInfo(AreaInfoName).AutoX2AxisMajorGridInterval = item.<AutoX2AxisMajorGridInterval>.Value
            dictAreaInfo(AreaInfoName).AutoYAxisMinimum = item.<AutoYAxisMinimum>.Value
            dictAreaInfo(AreaInfoName).AutoYAxisMaximum = item.<AutoYAxisMaximum>.Value
            dictAreaInfo(AreaInfoName).AutoYAxisMajorGridInterval = item.<AutoYAxisMajorGridInterval>.Value
            dictAreaInfo(AreaInfoName).AutoY2AxisMinimum = item.<AutoY2AxisMinimum>.Value
            dictAreaInfo(AreaInfoName).AutoY2AxisMaximum = item.<AutoY2AxisMaximum>.Value
            dictAreaInfo(AreaInfoName).AutoY2AxisMajorGridInterval = item.<AutoY2AxisMajorGridInterval>.Value
        Next

        'Restore Titles:
        Dim TitlesInfo = From item In XDoc.<ChartSettings>.<TitlesCollection>.<Title>
        Dim TitleName As String
        Dim myFontStyle As FontStyle
        Dim myFontSize As Single
        myChart.Titles.Clear()
        For Each item In TitlesInfo
            TitleName = item.<Name>.Value
            myChart.Titles.Add(TitleName).Name = TitleName 'The name needs to be explicitly declared!
            myChart.Titles(TitleName).Text = item.<Text>.Value
            myChart.Titles(TitleName).TextOrientation = [Enum].Parse(GetType(DataVisualization.Charting.TextOrientation), item.<TextOrientation>.Value)
            myChart.Titles(TitleName).Alignment = [Enum].Parse(GetType(ContentAlignment), item.<Alignment>.Value)
            myChart.Titles(TitleName).ForeColor = Color.FromArgb(item.<ForeColor>.Value)
            myFontStyle = FontStyle.Regular
            If item.<Font>.<Bold>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Bold
            If item.<Font>.<Italic>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Italic
            If item.<Font>.<Strikeout>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Strikeout
            If item.<Font>.<Underline>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Underline
            myFontSize = item.<Font>.<Size>.Value
            myChart.Titles(TitleName).Font = New Font(item.<Font>.<Name>.Value, myFontSize, myFontStyle)
        Next

        'Restore Chart Series:
        Dim Series = From item In XDoc.<ChartSettings>.<SeriesCollection>.<Series>
        Dim SeriesName As String
        myChart.Series.Clear()
        For Each item In Series
            SeriesName = item.<Name>.Value
            myChart.Series.Add(SeriesName)
            'myChart.Series(SeriesName).ChartType = [Enum].Parse(GetType(DataVisualization.Charting.SeriesChartType), item.<Type>.Value)
            myChart.Series(SeriesName).ChartType = [Enum].Parse(GetType(DataVisualization.Charting.SeriesChartType), item.<ChartType>.Value)
            If item.<ChartArea>.Value <> Nothing Then myChart.Series(SeriesName).ChartArea = item.<ChartArea>.Value
            myChart.Series(SeriesName).Legend = item.<Legend>.Value
            myChart.Series(SeriesName).SetCustomProperty("EmptyPointValue", item.<EmptyPointValue>.Value)
            myChart.Series(SeriesName).SetCustomProperty("LabelStyle", item.<LabelStyle>.Value)
            myChart.Series(SeriesName).SetCustomProperty("PixelPointDepth", item.<PixelPointDepth>.Value)
            myChart.Series(SeriesName).SetCustomProperty("PixelPointGapDepth", item.<PixelPointGapDepth>.Value)
            myChart.Series(SeriesName).SetCustomProperty("ShowMarkerLines", item.<ShowMarkerLines>.Value)
            myChart.Series(SeriesName).AxisLabel = item.<AxisLabel>.Value
            myChart.Series(SeriesName).XAxisType = [Enum].Parse(GetType(DataVisualization.Charting.AxisType), item.<XAxisType>.Value)
            myChart.Series(SeriesName).YAxisType = [Enum].Parse(GetType(DataVisualization.Charting.AxisType), item.<YAxisType>.Value)
            If item.<XValueType>.Value <> Nothing Then myChart.Series(SeriesName).XValueType = [Enum].Parse(GetType(DataVisualization.Charting.ChartValueType), item.<XValueType>.Value)
            If item.<YValueType>.Value <> Nothing Then myChart.Series(SeriesName).YValueType = [Enum].Parse(GetType(DataVisualization.Charting.ChartValueType), item.<YValueType>.Value)
            If item.<Marker>.<BorderColor>.Value <> Nothing Then myChart.Series(SeriesName).MarkerBorderColor = Color.FromArgb(item.<Marker>.<BorderColor>.Value)
            If item.<Marker>.<BorderWidth>.Value <> Nothing Then myChart.Series(SeriesName).MarkerBorderWidth = item.<Marker>.<BorderWidth>.Value
            If item.<Marker>.<Color>.Value <> Nothing Then myChart.Series(SeriesName).MarkerColor = Color.FromArgb(item.<Marker>.<Color>.Value)
            If item.<Marker>.<Size>.Value <> Nothing Then myChart.Series(SeriesName).MarkerSize = item.<Marker>.<Size>.Value
            If item.<Marker>.<Step>.Value <> Nothing Then myChart.Series(SeriesName).MarkerStep = item.<Marker>.<Step>.Value
            If item.<Marker>.<Style>.Value <> Nothing Then myChart.Series(SeriesName).MarkerStyle = [Enum].Parse(GetType(DataVisualization.Charting.MarkerStyle), item.<Marker>.<Style>.Value)
            If item.<Color>.Value <> Nothing Then myChart.Series(SeriesName).Color = Color.FromArgb(item.<Color>.Value)
        Next

        'Restore Chart Areas:
        Dim Areas = From item In XDoc.<ChartSettings>.<ChartAreasCollection>.<ChartArea>
        Dim AreaName As String
        myChart.ChartAreas.Clear()
        For Each item In Areas
            AreaName = item.<Name>.Value
            myChart.ChartAreas.Add(AreaName)
            'AxisX Properties:
            myChart.ChartAreas(AreaName).AxisX.Title = item.<AxisX>.<Title>.<Text>.Value
            myChart.ChartAreas(AreaName).AxisX.TitleAlignment = [Enum].Parse(GetType(StringAlignment), item.<AxisX>.<Title>.<Alignment>.Value)
            myChart.ChartAreas(AreaName).AxisX.TitleForeColor = Color.FromArgb(item.<AxisX>.<Title>.<ForeColor>.Value)
            myFontStyle = FontStyle.Regular
            If item.<AxisX>.<Title>.<Font>.<Bold>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Bold
            If item.<AxisX>.<Title>.<Font>.<Italic>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Italic
            If item.<AxisX>.<Title>.<Font>.<Strikeout>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Strikeout
            If item.<AxisX>.<Title>.<Font>.<Underline>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Underline
            myFontSize = item.<AxisX>.<Title>.<Font>.<Size>.Value
            myChart.ChartAreas(AreaName).AxisX.TitleFont = New Font(item.<AxisX>.<Title>.<Font>.<Name>.Value, myFontSize, myFontStyle)
            If item.<AxisX>.<LabelStyleFormat>.Value <> Nothing Then myChart.ChartAreas(AreaName).AxisX.LabelStyle.Format = item.<AxisX>.<LabelStyleFormat>.Value
            myChart.ChartAreas(AreaName).AxisX.Minimum = item.<AxisX>.<Minimum>.Value

            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoXAxisMinimum Then myChart.ChartAreas(AreaName).AxisX.Minimum = Double.NaN 'Set to Auto Minimum
            End If


            myChart.ChartAreas(AreaName).AxisX.Maximum = item.<AxisX>.<Maximum>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoXAxisMaximum Then myChart.ChartAreas(AreaName).AxisX.Maximum = Double.NaN 'Set to Auto Maximum
            End If

            myChart.ChartAreas(AreaName).AxisX.LineWidth = item.<AxisX>.<LineWidth>.Value
            myChart.ChartAreas(AreaName).AxisX.Interval = item.<AxisX>.<Interval>.Value
            myChart.ChartAreas(AreaName).AxisX.IntervalOffset = item.<AxisX>.<IntervalOffset>.Value
            myChart.ChartAreas(AreaName).AxisX.Crossing = item.<AxisX>.<Crossing>.Value
            myChart.ChartAreas(AreaName).AxisX.MajorGrid.Interval = item.<AxisX>.<MajorGrid>.<Interval>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoXAxisMajorGridInterval Then myChart.ChartAreas(AreaName).AxisX.MajorGrid.Interval = 0 'Set to Auto Interval
            End If

            myChart.ChartAreas(AreaName).AxisX.MajorGrid.IntervalOffset = item.<AxisX>.<MajorGrid>.<IntervalOffset>.Value

            'AxisX2 Properties:
            myChart.ChartAreas(AreaName).AxisX2.Title = item.<AxisX2>.<Title>.<Text>.Value
            myChart.ChartAreas(AreaName).AxisX2.TitleAlignment = [Enum].Parse(GetType(StringAlignment), item.<AxisX2>.<Title>.<Alignment>.Value)
            myChart.ChartAreas(AreaName).AxisX2.TitleForeColor = Color.FromArgb(item.<AxisX2>.<Title>.<ForeColor>.Value)
            myFontStyle = FontStyle.Regular
            If item.<AxisX2>.<Title>.<Font>.<Bold>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Bold
            If item.<AxisX2>.<Title>.<Font>.<Italic>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Italic
            If item.<AxisX2>.<Title>.<Font>.<Strikeout>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Strikeout
            If item.<AxisX2>.<Title>.<Font>.<Underline>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Underline
            myFontSize = item.<AxisX2>.<Title>.<Font>.<Size>.Value
            myChart.ChartAreas(AreaName).AxisX2.TitleFont = New Font(item.<AxisX2>.<Title>.<Font>.<Name>.Value, myFontSize, myFontStyle)
            If item.<AxisX2>.<LabelStyleFormat>.Value <> Nothing Then myChart.ChartAreas(AreaName).AxisX2.LabelStyle.Format = item.<AxisX2>.<LabelStyleFormat>.Value
            myChart.ChartAreas(AreaName).AxisX2.Minimum = item.<AxisX2>.<Minimum>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoX2AxisMinimum Then myChart.ChartAreas(AreaName).AxisX2.Minimum = Double.NaN 'Set to Auto Minimum
            End If


            myChart.ChartAreas(AreaName).AxisX2.Maximum = item.<AxisX2>.<Maximum>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoX2AxisMaximum Then myChart.ChartAreas(AreaName).AxisX2.Maximum = Double.NaN 'Set to Auto Maximum
            End If

            myChart.ChartAreas(AreaName).AxisX2.LineWidth = item.<AxisX2>.<LineWidth>.Value
            myChart.ChartAreas(AreaName).AxisX2.Interval = item.<AxisX2>.<Interval>.Value
            myChart.ChartAreas(AreaName).AxisX2.IntervalOffset = item.<AxisX2>.<IntervalOffset>.Value
            myChart.ChartAreas(AreaName).AxisX2.Crossing = item.<AxisX2>.<Crossing>.Value
            myChart.ChartAreas(AreaName).AxisX2.MajorGrid.Interval = item.<AxisX2>.<MajorGrid>.<Interval>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoX2AxisMajorGridInterval Then myChart.ChartAreas(AreaName).AxisX2.MajorGrid.Interval = 0 'Set to Auto Interval
            End If

            myChart.ChartAreas(AreaName).AxisX2.MajorGrid.IntervalOffset = item.<AxisX2>.<MajorGrid>.<IntervalOffset>.Value


            'AxisY Properties:
            myChart.ChartAreas(AreaName).AxisY.Title = item.<AxisY>.<Title>.<Text>.Value
            myChart.ChartAreas(AreaName).AxisY.TitleAlignment = [Enum].Parse(GetType(StringAlignment), item.<AxisY>.<Title>.<Alignment>.Value)
            myChart.ChartAreas(AreaName).AxisY.TitleForeColor = Color.FromArgb(item.<AxisY>.<Title>.<ForeColor>.Value)
            myFontStyle = FontStyle.Regular
            If item.<AxisY>.<Title>.<Font>.<Bold>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Bold
            If item.<AxisY>.<Title>.<Font>.<Italic>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Italic
            If item.<AxisY>.<Title>.<Font>.<Strikeout>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Strikeout
            If item.<AxisY>.<Title>.<Font>.<Underline>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Underline
            myFontSize = item.<AxisY>.<Title>.<Font>.<Size>.Value
            myChart.ChartAreas(AreaName).AxisY.TitleFont = New Font(item.<AxisY>.<Title>.<Font>.<Name>.Value, myFontSize, myFontStyle)
            If item.<AxisY>.<LabelStyleFormat>.Value <> Nothing Then myChart.ChartAreas(AreaName).AxisY.LabelStyle.Format = item.<AxisY>.<LabelStyleFormat>.Value
            myChart.ChartAreas(AreaName).AxisY.Minimum = item.<AxisY>.<Minimum>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoYAxisMinimum Then myChart.ChartAreas(AreaName).AxisY.Minimum = Double.NaN 'Set to Auto Minimum
            End If


            myChart.ChartAreas(AreaName).AxisY.Maximum = item.<AxisY>.<Maximum>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoYAxisMaximum Then myChart.ChartAreas(AreaName).AxisY.Maximum = Double.NaN 'Set to Auto Maximum
            End If

            myChart.ChartAreas(AreaName).AxisY.LineWidth = item.<AxisY>.<LineWidth>.Value
            myChart.ChartAreas(AreaName).AxisY.Interval = item.<AxisY>.<Interval>.Value
            myChart.ChartAreas(AreaName).AxisY.IntervalOffset = item.<AxisY>.<IntervalOffset>.Value
            myChart.ChartAreas(AreaName).AxisY.Crossing = item.<AxisY>.<Crossing>.Value
            myChart.ChartAreas(AreaName).AxisY.MajorGrid.Interval = item.<AxisY>.<MajorGrid>.<Interval>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoYAxisMajorGridInterval Then myChart.ChartAreas(AreaName).AxisY.MajorGrid.Interval = 0 'Set to Auto Interval
            End If

            myChart.ChartAreas(AreaName).AxisY.MajorGrid.IntervalOffset = item.<AxisY>.<MajorGrid>.<IntervalOffset>.Value

            'AxisY2 Properties:
            myChart.ChartAreas(AreaName).AxisY2.Title = item.<AxisY2>.<Title>.<Text>.Value
            myChart.ChartAreas(AreaName).AxisY2.TitleAlignment = [Enum].Parse(GetType(StringAlignment), item.<AxisY2>.<Title>.<Alignment>.Value)
            myChart.ChartAreas(AreaName).AxisY2.TitleForeColor = Color.FromArgb(item.<AxisY2>.<Title>.<ForeColor>.Value)
            myFontStyle = FontStyle.Regular
            If item.<AxisY2>.<Title>.<Font>.<Bold>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Bold
            If item.<AxisY2>.<Title>.<Font>.<Italic>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Italic
            If item.<AxisY2>.<Title>.<Font>.<Strikeout>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Strikeout
            If item.<AxisY2>.<Title>.<Font>.<Underline>.Value = True Then myFontStyle = myFontStyle Or FontStyle.Underline
            myFontSize = item.<AxisY2>.<Title>.<Font>.<Size>.Value
            myChart.ChartAreas(AreaName).AxisY2.TitleFont = New Font(item.<AxisY2>.<Title>.<Font>.<Name>.Value, myFontSize, myFontStyle)
            If item.<AxisY2>.<LabelStyleFormat>.Value <> Nothing Then myChart.ChartAreas(AreaName).AxisY2.LabelStyle.Format = item.<AxisY2>.<LabelStyleFormat>.Value
            myChart.ChartAreas(AreaName).AxisY2.Minimum = item.<AxisY2>.<Minimum>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoY2AxisMinimum Then myChart.ChartAreas(AreaName).AxisY2.Minimum = Double.NaN 'Set to Auto Minimum
            End If

            myChart.ChartAreas(AreaName).AxisY2.Maximum = item.<AxisY2>.<Maximum>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoY2AxisMaximum Then myChart.ChartAreas(AreaName).AxisY2.Maximum = Double.NaN 'Set to Auto Maximum
            End If

            myChart.ChartAreas(AreaName).AxisY2.LineWidth = item.<AxisY2>.<LineWidth>.Value
            myChart.ChartAreas(AreaName).AxisY2.Interval = item.<AxisY2>.<Interval>.Value
            myChart.ChartAreas(AreaName).AxisY2.IntervalOffset = item.<AxisY2>.<IntervalOffset>.Value
            myChart.ChartAreas(AreaName).AxisY2.Crossing = item.<AxisY2>.<Crossing>.Value
            myChart.ChartAreas(AreaName).AxisY2.MajorGrid.Interval = item.<AxisY2>.<MajorGrid>.<Interval>.Value
            If dictAreaInfo.ContainsKey(AreaName) Then
                If dictAreaInfo(AreaName).AutoY2AxisMajorGridInterval Then myChart.ChartAreas(AreaName).AxisY2.MajorGrid.Interval = 0 'Set to Auto Interval
            End If

            myChart.ChartAreas(AreaName).AxisY2.MajorGrid.IntervalOffset = item.<AxisY2>.<MajorGrid>.<IntervalOffset>.Value
        Next

    End Sub

    Public Function ToXDoc(ByRef myChart As System.Windows.Forms.DataVisualization.Charting.Chart) As System.Xml.Linq.XDocument
        'Function to return the Line Chart settings in an XDocument.

        Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                   <!---->
                   <!--Line Chart Settings File-->
                   <ChartSettings>
                       <!--Input Data:-->
                       <InputDataType><%= InputDataType %></InputDataType>
                       <InputDatabasePath><%= InputDatabasePath %></InputDatabasePath>
                       <InputQuery><%= InputQuery %></InputQuery>
                       <InputDataDescr><%= InputDataDescr %></InputDataDescr>
                       <SeriesInfoList>
                           <%= From item In dictSeriesInfo
                               Select
                                   <SeriesInfo>
                                       <Name><%= item.Key %></Name>
                                       <XValuesFieldName><%= item.Value.XValuesFieldName %></XValuesFieldName>
                                       <YValuesFieldName><%= item.Value.YValuesFieldName %></YValuesFieldName>
                                       <ChartArea><%= item.Value.ChartArea %></ChartArea>
                                   </SeriesInfo> %>
                       </SeriesInfoList>
                       <AreaInfoList>
                           <%= From item In dictAreaInfo
                               Select
                                    <AreaInfo>
                                        <Name><%= item.Key %></Name>
                                        <AutoXAxisMinimum><%= item.Value.AutoXAxisMinimum %></AutoXAxisMinimum>
                                        <AutoXAxisMaximum><%= item.Value.AutoXAxisMaximum %></AutoXAxisMaximum>
                                        <AutoXAxisMajorGridInterval><%= item.Value.AutoXAxisMajorGridInterval %></AutoXAxisMajorGridInterval>
                                        <AutoX2AxisMinimum><%= item.Value.AutoX2AxisMinimum %></AutoX2AxisMinimum>
                                        <AutoX2AxisMaximum><%= item.Value.AutoX2AxisMaximum %></AutoX2AxisMaximum>
                                        <AutoX2AxisMajorGridInterval><%= item.Value.AutoX2AxisMajorGridInterval %></AutoX2AxisMajorGridInterval>
                                        <AutoYAxisMinimum><%= item.Value.AutoYAxisMinimum %></AutoYAxisMinimum>
                                        <AutoYAxisMaximum><%= item.Value.AutoYAxisMaximum %></AutoYAxisMaximum>
                                        <AutoYAxisMajorGridInterval><%= item.Value.AutoYAxisMajorGridInterval %></AutoYAxisMajorGridInterval>
                                        <AutoY2AxisMinimum><%= item.Value.AutoY2AxisMinimum %></AutoY2AxisMinimum>
                                        <AutoY2AxisMaximum><%= item.Value.AutoY2AxisMaximum %></AutoY2AxisMaximum>
                                        <AutoY2AxisMajorGridInterval><%= item.Value.AutoY2AxisMajorGridInterval %></AutoY2AxisMajorGridInterval>
                                    </AreaInfo> %>
                       </AreaInfoList>
                       <!--Chart Properties:-->
                       <TitlesCollection>
                           <%= From item In myChart.Titles
                               Select
                               <Title>
                                   <Name><%= item.Name %></Name>
                                   <Text><%= item.Text %></Text>
                                   <TextOrientation><%= item.TextOrientation %></TextOrientation>
                                   <Alignment><%= item.Alignment %></Alignment>
                                   <ForeColor><%= item.ForeColor.ToArgb.ToString %></ForeColor>
                                   <Font>
                                       <Name><%= item.Font.Name %></Name>
                                       <Size><%= item.Font.Size %></Size>
                                       <Bold><%= item.Font.Bold %></Bold>
                                       <Italic><%= item.Font.Italic %></Italic>
                                       <Strikeout><%= item.Font.Strikeout %></Strikeout>
                                       <Underline><%= item.Font.Underline %></Underline>
                                   </Font>
                               </Title> %>
                       </TitlesCollection>
                       <SeriesCollection>
                           <%= From item In myChart.Series
                               Select
                                   <Series>
                                       <Name><%= item.Name %></Name>
                                       <ChartType><%= item.ChartType %></ChartType>
                                       <ChartArea><%= item.ChartArea %></ChartArea>
                                       <Legend><%= item.Legend %></Legend>
                                       <EmptyPointValue><%= item.GetCustomProperty("EmptyPointValue") %></EmptyPointValue>
                                       <LabelStyle><%= item.GetCustomProperty("LabelStyle") %></LabelStyle>
                                       <PixelPointDepth><%= item.GetCustomProperty("PixelPointDepth") %></PixelPointDepth>
                                       <PixelPointGapDepth><%= item.GetCustomProperty("PixelPointGapDepth") %></PixelPointGapDepth>
                                       <ShowMarkerLines><%= item.GetCustomProperty("ShowMarkerLines") %></ShowMarkerLines>
                                       <AxisLabel><%= item.AxisLabel %></AxisLabel>
                                       <XAxisType><%= item.XAxisType %></XAxisType>
                                       <XValueType><%= item.XValueType %></XValueType>
                                       <YAxisType><%= item.YAxisType %></YAxisType>
                                       <YValueType><%= item.YValueType %></YValueType>
                                       <Marker>
                                           <BorderColor><%= item.MarkerBorderColor.ToArgb.ToString %></BorderColor>
                                           <BorderWidth><%= item.MarkerBorderWidth %></BorderWidth>
                                           <Color><%= item.MarkerColor.ToArgb.ToString %></Color>
                                           <Size><%= item.MarkerSize %></Size>
                                           <Step><%= item.MarkerStep %></Step>
                                           <Style><%= item.MarkerStyle %></Style>
                                       </Marker>
                                       <Color><%= item.Color.ToArgb.ToString %></Color>
                                   </Series> %>
                       </SeriesCollection>
                       <ChartAreasCollection>
                           <%= From item In myChart.ChartAreas
                               Select
                               <ChartArea>
                                   <Name><%= item.Name %></Name>
                                   <AxisX>
                                       <Title>
                                           <Text><%= item.AxisX.Title %></Text>
                                           <Alignment><%= item.AxisX.TitleAlignment %></Alignment>
                                           <ForeColor><%= item.AxisX.TitleForeColor.ToArgb.ToString %></ForeColor>
                                           <Font>
                                               <Name><%= item.AxisX.TitleFont.Name %></Name>
                                               <Size><%= item.AxisX.TitleFont.Size %></Size>
                                               <Bold><%= item.AxisX.TitleFont.Bold %></Bold>
                                               <Italic><%= item.AxisX.TitleFont.Italic %></Italic>
                                               <Strikeout><%= item.AxisX.TitleFont.Strikeout %></Strikeout>
                                               <Underline><%= item.AxisX.TitleFont.Underline %></Underline>
                                           </Font>
                                       </Title>
                                       <LabelStyleFormat><%= item.AxisX.LabelStyle.Format %></LabelStyleFormat>
                                       <Minimum><%= item.AxisX.Minimum %></Minimum>
                                       <Maximum><%= item.AxisX.Maximum %></Maximum>
                                       <LineWidth><%= item.AxisX.LineWidth %></LineWidth>
                                       <Interval><%= item.AxisX.Interval %></Interval>
                                       <IntervalOffset><%= item.AxisX.IntervalOffset %></IntervalOffset>
                                       <Crossing><%= item.AxisX.Crossing %></Crossing>
                                       <MajorGrid>
                                           <Interval><%= item.AxisX.MajorGrid.Interval %></Interval>
                                           <IntervalOffset><%= item.AxisX.MajorGrid.IntervalOffset %></IntervalOffset>
                                       </MajorGrid>
                                   </AxisX>
                                   <AxisX2>
                                       <Title>
                                           <Text><%= item.AxisX2.Title %></Text>
                                           <Alignment><%= item.AxisX2.TitleAlignment %></Alignment>
                                           <ForeColor><%= item.AxisX2.TitleForeColor.ToArgb.ToString %></ForeColor>
                                           <Font>
                                               <Name><%= item.AxisX2.TitleFont.Name %></Name>
                                               <Size><%= item.AxisX2.TitleFont.Size %></Size>
                                               <Bold><%= item.AxisX2.TitleFont.Bold %></Bold>
                                               <Italic><%= item.AxisX2.TitleFont.Italic %></Italic>
                                               <Strikeout><%= item.AxisX2.TitleFont.Strikeout %></Strikeout>
                                               <Underline><%= item.AxisX2.TitleFont.Underline %></Underline>
                                           </Font>
                                       </Title>
                                       <LabelStyleFormat><%= item.AxisX2.LabelStyle.Format %></LabelStyleFormat>
                                       <Minimum><%= item.AxisX2.Minimum %></Minimum>
                                       <Maximum><%= item.AxisX2.Maximum %></Maximum>
                                       <LineWidth><%= item.AxisX2.LineWidth %></LineWidth>
                                       <Interval><%= item.AxisX2.Interval %></Interval>
                                       <IntervalOffset><%= item.AxisX2.IntervalOffset %></IntervalOffset>
                                       <Crossing><%= item.AxisX2.Crossing %></Crossing>
                                       <MajorGrid>
                                           <Interval><%= item.AxisX2.MajorGrid.Interval %></Interval>
                                           <IntervalOffset><%= item.AxisX2.MajorGrid.IntervalOffset %></IntervalOffset>
                                       </MajorGrid>
                                   </AxisX2>
                                   <AxisY>
                                       <Title>
                                           <Text><%= item.AxisY.Title %></Text>
                                           <Alignment><%= item.AxisY.TitleAlignment %></Alignment>
                                           <ForeColor><%= item.AxisY.TitleForeColor.ToArgb.ToString %></ForeColor>
                                           <Font>
                                               <Name><%= item.AxisY.TitleFont.Name %></Name>
                                               <Size><%= item.AxisY.TitleFont.Size %></Size>
                                               <Bold><%= item.AxisY.TitleFont.Bold %></Bold>
                                               <Italic><%= item.AxisY.TitleFont.Italic %></Italic>
                                               <Strikeout><%= item.AxisY.TitleFont.Strikeout %></Strikeout>
                                               <Underline><%= item.AxisY.TitleFont.Underline %></Underline>
                                           </Font>
                                       </Title>
                                       <LabelStyleFormat><%= item.AxisY.LabelStyle.Format %></LabelStyleFormat>
                                       <Minimum><%= item.AxisY.Minimum %></Minimum>
                                       <Maximum><%= item.AxisY.Maximum %></Maximum>
                                       <LineWidth><%= item.AxisY.LineWidth %></LineWidth>
                                       <Interval><%= item.AxisY.Interval %></Interval>
                                       <IntervalOffset><%= item.AxisY.IntervalOffset %></IntervalOffset>
                                       <Crossing><%= item.AxisY.Crossing %></Crossing>
                                       <MajorGrid>
                                           <Interval><%= item.AxisY.MajorGrid.Interval %></Interval>
                                           <IntervalOffset><%= item.AxisY.MajorGrid.IntervalOffset %></IntervalOffset>
                                       </MajorGrid>
                                   </AxisY>
                                   <AxisY2>
                                       <Title>
                                           <Text><%= item.AxisY2.Title %></Text>
                                           <Alignment><%= item.AxisY2.TitleAlignment %></Alignment>
                                           <ForeColor><%= item.AxisY2.TitleForeColor.ToArgb.ToString %></ForeColor>
                                           <Font>
                                               <Name><%= item.AxisY2.TitleFont.Name %></Name>
                                               <Size><%= item.AxisY2.TitleFont.Size %></Size>
                                               <Bold><%= item.AxisY2.TitleFont.Bold %></Bold>
                                               <Italic><%= item.AxisY2.TitleFont.Italic %></Italic>
                                               <Strikeout><%= item.AxisY2.TitleFont.Strikeout %></Strikeout>
                                               <Underline><%= item.AxisY2.TitleFont.Underline %></Underline>
                                           </Font>
                                       </Title>
                                       <LabelStyleFormat><%= item.AxisY2.LabelStyle.Format %></LabelStyleFormat>
                                       <Minimum><%= item.AxisY2.Minimum %></Minimum>
                                       <Maximum><%= item.AxisY2.Maximum %></Maximum>
                                       <LineWidth><%= item.AxisY2.LineWidth %></LineWidth>
                                       <Interval><%= item.AxisY2.Interval %></Interval>
                                       <IntervalOffset><%= item.AxisY2.IntervalOffset %></IntervalOffset>
                                       <Crossing><%= item.AxisY2.Crossing %></Crossing>
                                       <MajorGrid>
                                           <Interval><%= item.AxisY2.MajorGrid.Interval %></Interval>
                                           <IntervalOffset><%= item.AxisY2.MajorGrid.IntervalOffset %></IntervalOffset>
                                       </MajorGrid>
                                   </AxisY2>
                               </ChartArea> %>
                       </ChartAreasCollection>
                   </ChartSettings>

        Return XDoc
    End Function

    Public Sub SaveFile(ByVal myFileName As String, ByRef myChart As System.Windows.Forms.DataVisualization.Charting.Chart)
        'Save the Point Chart settings in a file named FileName.
        If myFileName = "" Then 'No stock chart settings file has been selected.
            Exit Sub
        End If

        'Clean the AreaInfo and SeriesInfo dictionaries before saving:
        CleanAreaInfo(myChart)
        CleanSeriesInfo(myChart)

        DataLocation.SaveXmlData(myFileName, ToXDoc(myChart))
    End Sub

    Public Sub Clear(ByRef myChart As System.Windows.Forms.DataVisualization.Charting.Chart)
        'Clear the Line Chart settings and apply defaults.

        'Clear the myChart properties:
        myChart.ChartAreas.Clear()
        myChart.Series.Clear()

        'Clear the ChartInfo properties:
        FileName = ""
        InputDataType = "Database"
        InputDatabasePath = ""
        InputQuery = ""
        InputDataDescr = ""

        ds.Clear() 'Clear the dataset containin the points to be plotted in the line chart.
        dictSeriesInfo.Clear() 'Clear the dictionary of Series Information.
        dictAreaInfo.Clear()   'Clear the dictionary of Area Information
    End Sub

    Public Sub CleanSeriesInfo(ByRef myChart As System.Windows.Forms.DataVisualization.Charting.Chart)
        'Clean the SeriesInfo dictionary of Series that are no longer in the Chart:

        Dim list As New List(Of String)(dictSeriesInfo.Keys) 'Get the list of keys in the SeriesInfo dictionary.

        Dim KeyFound As Boolean = False 'If the SeriesInfo dictionary key is found in the Chart control, KeyFound is True.
        For Each KeyStr In list
            'Check if the dictionary key (Series name) is found in myChart:
            For Each item In myChart.Series
                If item.Name = KeyStr Then
                    KeyFound = True
                    Exit For 'Key found, stop looking.
                End If
            Next
            If KeyFound = False Then
                'Remove the entry from the dictionary:
                dictSeriesInfo.Remove(KeyStr)
            Else
                'The key was found - do not remove the dictionary entry.
                KeyFound = False 'Reset the flas to False before searching for the next key.
            End If
        Next
    End Sub

    Public Sub CleanAreaInfo(ByRef myChart As System.Windows.Forms.DataVisualization.Charting.Chart)
        'Clean the AreaInfo dictionary of Chart Areas that are no longer in the Chart:

        Dim list As New List(Of String)(dictAreaInfo.Keys) 'Get the list of keys in the AreaInfo dictionary.

        Dim KeyFound As Boolean = False 'If the AreaInfo dictionary key is found in the Chart control, KeyFound is True.
        For Each KeyStr In list
            'Check if the dictionary key (ChartArea name) is found in myChart:
            For Each item In myChart.ChartAreas
                If item.Name = KeyStr Then
                    KeyFound = True
                    Exit For 'Key found, stop looking.
                End If
            Next
            If KeyFound = False Then
                'Remove the entry from the dictionary:
                dictAreaInfo.Remove(KeyStr)
            Else
                'The key was found - do not remove the dictionary entry.
                KeyFound = False 'Reset the flas to False before searching for the next key.
            End If
        Next
    End Sub


    Public Sub ApplyQuery()
        'Use the Query to fill the ds dataset

        If InputDatabasePath = "" Then
            RaiseEvent ErrorMessage("InputDatabasePath is not defined!" & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim commandString As String 'Declare a command string - contains the query to be passed to the database.

        'Specify the connection string (Access 2007):
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + InputDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        'Specify the commandString to query the database:
        commandString = InputQuery
        Dim dataAdapter As New System.Data.OleDb.OleDbDataAdapter(commandString, conn)

        ds.Clear()
        ds.Reset()

        dataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey

        Try
            dataAdapter.Fill(ds, "SelTable")
            'UpdateChartQuery() 'NOT NEEDED??? 'This was originally used to set PointChart or StockChart .Input Query to the property InputQuery. (See the Chart app code.)
        Catch ex As Exception
            RaiseEvent ErrorMessage("Error applying query." & vbCrLf)
            RaiseEvent ErrorMessage(ex.Message & vbCrLf)
        End Try

        conn.Close()

    End Sub 'ApplyQuery

#End Region 'Methods -----------------------------------------------------------------------------------------------------


#Region "Events" '--------------------------------------------------------------------------------------------------------

    Event ErrorMessage(ByVal Message As String) 'Send an error message.
    Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------


End Class 'ChartInfo


Public Class SeriesInfo
    'Used to store the X Values Field Name and Y Values Field Name.
    'These are the names of the fields in a database table used for the X and Y values in a chart.

    Private _xValuesFieldName As String = "" 'The name of the table field containing the X Values
    Property XValuesFieldName As String
        Get
            Return _xValuesFieldName
        End Get
        Set(value As String)
            _xValuesFieldName = value
        End Set
    End Property

    Private _yValuesFieldName As String = "" 'The name of the table field containing the Y Values
    Property YValuesFieldName As String
        Get
            Return _yValuesFieldName
        End Get
        Set(value As String)
            _yValuesFieldName = value
        End Set
    End Property

    Private _chartArea As String = "" 'The name of the Chart Area used to display the series.
    Property ChartArea As String
        Get
            Return _chartArea
        End Get
        Set(value As String)
            _chartArea = value
        End Set
    End Property

End Class 'SeriesInfo


Public Class AreaInfo
    'Used to indicate if chart area parameters are determined automatically or not.
    'These parameters cannot be stored in the Chart.

    Private _autoXAxisMinimum As Boolean = False 'If True, the X Axis minimum value is determined automatically.
    Property AutoXAxisMinimum As Boolean
        Get
            Return _autoXAxisMinimum
        End Get
        Set(value As Boolean)
            _autoXAxisMinimum = value
        End Set
    End Property

    Private _autoXAxisMaximum As Boolean = False 'If True, the X Axis maximum value is determined automatically.
    Property AutoXAxisMaximum As Boolean
        Get
            Return _autoXAxisMaximum
        End Get
        Set(value As Boolean)
            _autoXAxisMaximum = value
        End Set
    End Property

    Private _autoXAxisMajorGridInterval As Boolean = False 'If True, the X Axis Major Grid Interval value is determined automatically.
    Property AutoXAxisMajorGridInterval As Boolean
        Get
            Return _autoXAxisMajorGridInterval
        End Get
        Set(value As Boolean)
            _autoXAxisMajorGridInterval = value
        End Set
    End Property

    Private _autoX2AxisMinimum As Boolean = False 'If True, the X2 Axis minimum value is determined automatically.
    Property AutoX2AxisMinimum As Boolean
        Get
            Return _autoX2AxisMinimum
        End Get
        Set(value As Boolean)
            _autoX2AxisMinimum = value
        End Set
    End Property

    Private _autoX2AxisMaximum As Boolean = False 'If True, the X2 Axis maximum value is determined automatically.
    Property AutoX2AxisMaximum As Boolean
        Get
            Return _autoX2AxisMaximum
        End Get
        Set(value As Boolean)
            _autoX2AxisMaximum = value
        End Set
    End Property

    Private _autoX2AxisMajorGridInterval As Boolean = False 'If True, the X2 Axis Major Grid Interval value is determined automatically.
    Property AutoX2AxisMajorGridInterval As Boolean
        Get
            Return _autoX2AxisMajorGridInterval
        End Get
        Set(value As Boolean)
            _autoX2AxisMajorGridInterval = value
        End Set
    End Property

    Private _autoYAxisMinimum As Boolean = False 'If True, the Y Axis minimum value is determined automatically.
    Property AutoYAxisMinimum As Boolean
        Get
            Return _autoYAxisMinimum
        End Get
        Set(value As Boolean)
            _autoYAxisMinimum = value
        End Set
    End Property

    Private _autoYAxisMaximum As Boolean = False 'If True, the Y Axis maximum value is determined automatically.
    Property AutoYAxisMaximum As Boolean
        Get
            Return _autoYAxisMaximum
        End Get
        Set(value As Boolean)
            _autoYAxisMaximum = value
        End Set
    End Property

    Private _autoYAxisMajorGridInterval As Boolean = False 'If True, the Y Axis Major Grid Interval value is determined automatically.
    Property AutoYAxisMajorGridInterval As Boolean
        Get
            Return _autoYAxisMajorGridInterval
        End Get
        Set(value As Boolean)
            _autoYAxisMajorGridInterval = value
        End Set
    End Property

    Private _autoY2AxisMinimum As Boolean = False 'If True, the Y2 Axis minimum value is determined automatically.
    Property AutoY2AxisMinimum As Boolean
        Get
            Return _autoY2AxisMinimum
        End Get
        Set(value As Boolean)
            _autoY2AxisMinimum = value
        End Set
    End Property

    Private _autoY2AxisMaximum As Boolean = False 'If True, the Y2 Axis maximum value is determined automatically.
    Property AutoY2AxisMaximum As Boolean
        Get
            Return _autoY2AxisMaximum
        End Get
        Set(value As Boolean)
            _autoY2AxisMaximum = value
        End Set
    End Property

    Private _autoY2AxisMajorGridInterval As Boolean = False 'If True, the Y2 Axis Major Grid Interval value is determined automatically.
    Property AutoY2AxisMajorGridInterval As Boolean
        Get
            Return _autoY2AxisMajorGridInterval
        End Get
        Set(value As Boolean)
            _autoY2AxisMajorGridInterval = value
        End Set
    End Property

End Class 'AreaInfo

