VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub btnCalPosBPV_Click()
'Pls ensure the following functions
'1. pop up a window for user to select the file to load
'2. if the data of a particular date exists, a warning msg will be shown
'3.col F to col J will be populated
'4. Plot the whole book's Position BPV curve. The book includes 4 bond (A, B, C and D)
    Dim i%, j%, k%
    Dim a#, C#, R#, DP#, CP#, MD#, BPV#
    Dim YieldEnd%, RowEnd%
    Dim NextDate As Date, LastDate As Date
    NextDate = #9/30/2013#
    LastDate = #3/31/2013#
    i = 2
    While Not IsEmpty(Worksheets("Bond.Position").Cells(i, 1))
        i = i + 1
    Wend
    RowEnd = i - 1
    i = 2
    While Not IsEmpty(Worksheets("Yield.Curve").Cells(i, 1))
        i = i + 1
    Wend
    YieldEnd = i - 1
    
    Range("A2").Select
    ActiveWorkbook.Worksheets("Bond.Position").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Bond.Position").Sort.SortFields.Add Key:=Range("A1") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Bond.Position").Sort
        .SetRange Range(Selection, Selection.End(xlDown))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Dim daysleft As Integer
    For i = 2 To RowEnd
        DP = 0
        MD = 0
        daysleft = NextDate - Worksheets("Bond.Position").Cells(i, 1)
        a = daysleft / (NextDate - LastDate)
        C = Worksheets("Bond.Position").Cells(i, 3)
        Select Case (i - 2) Mod 4
        Case 0
            For j = 2 To YieldEnd
                If Worksheets("Yield.Curve").Cells(j, 1) = Worksheets("Bond.Position").Cells(i, 1) Then
                    R = Worksheets("Yield.Curve").Cells(j, 5) / 100
                    Exit For
                End If
            Next
            DP = (C / 2) / (1 + R / 2) ^ a + (C / 2 + 100) / (1 + R / 2) ^ (1 + a)
            MD = (1 / DP) * (a * (C / 2) / (1 + R / 2) ^ a + (a + 1) * (C / 2 + 100) / (1 + R / 2) ^ (1 + a))
            MD = MD / (1 + R / 2)
            BPV = (MD / 100) * (DP / 100)
        ' The following is the only Case clause that evaluates to True.
        Case 1
            For j = 2 To YieldEnd
                If Worksheets("Yield.Curve").Cells(j, 1) = Worksheets("Bond.Position").Cells(i, 1) Then
                    R = Worksheets("Yield.Curve").Cells(j, 8) / 100
                    Exit For
                End If
            Next
            For j = 0 To 6
                DP = DP + (C / 2) / ((1 + R / 2) ^ (a + j))
                MD = MD + (a + j) * (C / 2) / (1 + R / 2) ^ (a + j)
            Next
            DP = DP + (C / 2 + 100) / (1 + R / 2) ^ (a + 7)
            'MsgBox (C / 2 + 100) / (1 + R / 2) ^ (a + 7)
            MD = MD + (a + 7) * (C / 2 + 100) / (1 + R / 2) ^ (a + 7)
            MD = MD / DP / (1 + R / 2)
            BPV = (MD / 100) * (DP / 100)
            
        Case 2
            For j = 2 To YieldEnd
                If Worksheets("Yield.Curve").Cells(j, 1) = Worksheets("Bond.Position").Cells(i, 1) Then
                    R = Worksheets("Yield.Curve").Cells(j, 10) / 100
                    Exit For
                End If
            Next
            For j = 0 To 16
                DP = DP + (C / 2) / (1 + R / 2) ^ (a + j)
                MD = MD + (a + j) * (C / 2) / (1 + R / 2) ^ (a + j)
            Next
            DP = DP + (C / 2 + 100) / (1 + R / 2) ^ (a + 17)
            MD = MD + (a + 17) * (C / 2 + 100) / (1 + R / 2) ^ (a + 17)
            MD = MD / DP / (1 + R / 2)
            BPV = (MD / 100) * (DP / 100)
            
        Case 3
            For j = 2 To YieldEnd
                If Worksheets("Yield.Curve").Cells(j, 1) = Worksheets("Bond.Position").Cells(i, 1) Then
                    R = Worksheets("Yield.Curve").Cells(j, 12) / 100
                    Exit For
                End If
            Next
            For j = 0 To 56
                DP = DP + (C / 2) / (1 + R / 2) ^ (a + j)
                MD = MD + (a + j) * (C / 2) / (1 + R / 2) ^ (a + j)
            Next
            DP = DP + (C / 2 + 100) / (1 + R / 2) ^ (a + 57)
            MD = MD + (a + 57) * (C / 2 + 100) / (1 + R / 2) ^ (a + 57)
            MD = MD / DP / (1 + R / 2)
            BPV = (MD / 100) * (DP / 100)
        End Select
        Worksheets("Bond.Position").Cells(i, 7) = (1 - a) * C / 2
        CP = DP - (1 - a) * C / 2
        Worksheets("Bond.Position").Cells(i, 6) = CP
        Worksheets("Bond.Position").Cells(i, 8) = DP
        Worksheets("Bond.Position").Cells(i, 9) = MD
        Worksheets("Bond.Position").Cells(i, 10) = BPV * Worksheets("Bond.Position").Cells(i, 5) / 100
    Next
    Range("Q1").Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlLine
    ActiveChart.SeriesCollection
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(1).Name = "=Bond.Position!$B$2"
    ActiveChart.SeriesCollection(1).Values = _
        "=Bond.Position!$J$2,Bond.Position!$J$6,Bond.Position!$J$10,Bond.Position!$J$14,Bond.Position!$J$18,Bond.Position!$J$22"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).Name = "=Bond.Position!$B$3"
    ActiveChart.SeriesCollection(2).Values = _
        "=Bond.Position!$J$3,Bond.Position!$J$7,Bond.Position!$J$11,Bond.Position!$J$15,Bond.Position!$J$19,Bond.Position!$J$23"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(3).Name = "=Bond.Position!$B$4"
    ActiveChart.SeriesCollection(3).Values = _
        "=Bond.Position!$J$4,Bond.Position!$J$8,Bond.Position!$J$12,Bond.Position!$J$16,Bond.Position!$J$20,Bond.Position!$J$24"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(4).Name = "=Bond.Position!$B$5"
    ActiveChart.SeriesCollection(4).Values = _
        "=Bond.Position!$J$5,Bond.Position!$J$9,Bond.Position!$J$13,Bond.Position!$J$17,Bond.Position!$J$21,Bond.Position!$J$25"
    ActiveChart.SeriesCollection(4).XValues = _
        "=Bond.Position!$A$2,Bond.Position!$A$6,Bond.Position!$A$10,Bond.Position!$A$14,Bond.Position!$A$18,Bond.Position!$A$22"
    ActiveChart.SetElement (msoElementChartTitleAboveChart)
    ActiveChart.ChartTitle.Text = "Position BPV"
    ActiveChart.ChartArea.Select
    ActiveChart.Parent.Cut
    Sheets("Position.BPV.Curve").Select
    ActiveSheet.Range("A1").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Height = 377.0078740157
    ActiveChart.Parent.Width = 581.1023622047
End Sub

Private Sub btnInputPostion_Click()
'Pls ensure the following functions
'1. pop up a window for user to select the file to load
'2. if the data of a particular date exists, a warning msg will be shown
'3.col A to col E will be populated with the data from yc.2012.MM.DD.csv
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Import bond positions"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Yield Curve Files", "*.csv"
        .Filters.Add "All Files", "*.*"
        .InitialFileName = ThisWorkbook.Path + "\bond.postion.2013*.csv"
        '设置两个文件过滤器
        .Show
        Dim i%, j%, k%
        Dim thisFile, StartRow%, RowNum%, Blanks, BlankNum%
        i = 2
        While Not IsEmpty(Worksheets("Bond.Position").Cells(i, 1))
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
                    Worksheets("Bond.Position").Cells(StartRow, j + 1) = Blanks(j)
                Next
                If StartRow > 2 Then
                     For k = 2 To (StartRow - 1)
                        'MsgBox Cells(k, 1)
                        If Worksheets("Bond.Position").Cells(k, 1) = Worksheets("Bond.Position").Cells(StartRow, 1) And Worksheets("Bond.Position").Cells(k, 2) = Worksheets("Bond.Position").Cells(StartRow, 2) Then
                            Worksheets("Bond.Position").Range(Cells(StartRow, 1), Cells(StartRow, 5)).ClearContents
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
End Sub




