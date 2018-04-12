VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnInputYC_Click()
'Pls ensure the following functions
'1. pop up a window for user to select the file to load
'2. if the data of a particular date exists, a warning msg will be shown
'3.col A to col L will be populated with the data from yc.2012.MM.DD.csv
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Import yield curves"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Yield Curve Files", "*.csv"
        .Filters.Add "All Files", "*.*"
        .InitialFileName = ThisWorkbook.Path + "\yc.2013*.csv"
        '设置两个文件过滤器
        .Show
        Dim i%, j%, k%
        Dim thisFile, StartRow%, RowNum%, Blanks, BlankNum%
        i = 2
        While Not IsEmpty(Worksheets("Yield.Curve").Cells(i, 1))
            i = i + 1
        Wend
        StartRow = i
        For Each SelectedItem In .SelectedItems
            Open SelectedItem For Input As #1
                thisFile = Split(StrConv(InputB(LOF(1), 1), vbUnicode), vbCrLf)
                Close #1
            RowNum = UBound(thisFile)
            If RowNum = 0 Then
                
            Else
            For i = 0 To RowNum - 1
                Blanks = Split(thisFile(i), "|")
                BlankNum = UBound(Blanks)
                For j = 0 To BlankNum
                    Worksheets("Yield.Curve").Cells(StartRow, j + 1) = Blanks(j)
                Next
                If StartRow > 2 Then
                     For k = 2 To (StartRow - 1)
                        'MsgBox Cells(k, 1)
                        If Worksheets("Yield.Curve").Cells(k, 1) = Worksheets("Yield.Curve").Cells(StartRow, 1) Then
                            Worksheets("Yield.Curve").Range(Cells(StartRow, 1), Cells(StartRow, 12)).ClearContents
                            MsgBox Blanks(0) & "的数据已存在"
                            Exit For
                        End If
                    Next
                    If k <> StartRow Then
                       
                    Else
                        StartRow = StartRow + 1
                    End If
                Else
                    StartRow = StartRow + 1
                End If
            Next
            End If
        Next
    End With
    Range("A1:A25").Select
    ActiveWorkbook.Worksheets("Yield.Curve").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Yield.Curve").Sort.SortFields.Add Key:=Range("A1") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Yield.Curve").Sort
        .SetRange Range("A2:A25")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("O1").Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlLine
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(1).Name = "=Yield.Curve!$A$2"
    ActiveChart.SeriesCollection(1).Values = "=Yield.Curve!$B$2:$L$2"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).Name = "=Yield.Curve!$A$3"
    ActiveChart.SeriesCollection(2).Values = "=Yield.Curve!$B$3:$L$3"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(3).Name = "=Yield.Curve!$A$4"
    ActiveChart.SeriesCollection(3).Values = "=Yield.Curve!$B$4:$L$4"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(4).Name = "=Yield.Curve!$A$5"
    ActiveChart.SeriesCollection(4).Values = "=Yield.Curve!$B$5:$L$5"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(5).Name = "=Yield.Curve!$A$6"
    ActiveChart.SeriesCollection(5).Values = "=Yield.Curve!$B$6:$L$6"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(6).Name = "=Yield.Curve!$A$7"
    ActiveChart.SeriesCollection(6).Values = "=Yield.Curve!$B$7:$L$7"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(7).Name = "=Yield.Curve!$A$8"
    ActiveChart.SeriesCollection(7).Values = "=Yield.Curve!$B$8:$L$8"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(8).Name = "=Yield.Curve!$A$9"
    ActiveChart.SeriesCollection(8).Values = "=Yield.Curve!$B$9:$L$9"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(9).Name = "=Yield.Curve!$A$10"
    ActiveChart.SeriesCollection(9).Values = "=Yield.Curve!$B$10:$L$10"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(10).Name = "=Yield.Curve!$A$11"
    ActiveChart.SeriesCollection(10).Values = "=Yield.Curve!$B$11:$L$11"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(11).Name = "=Yield.Curve!$A$12"
    ActiveChart.SeriesCollection(11).Values = "=Yield.Curve!$B$12:$L$12"
    ActiveChart.SeriesCollection(11).XValues = "=Yield.Curve!$B$1:$L$1"
    ActiveChart.SetElement (msoElementChartTitleAboveChart)
    ActiveChart.ChartTitle.Text = "Yield Curve"
    ActiveChart.ChartArea.Select
    ActiveChart.Parent.Cut
    ActiveSheet.Range("A27").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Height = 377.0078740157
    ActiveChart.Parent.Width = 581.1023622047
End Sub
