Sub CreateLineChart()
    Dim ws As Worksheet
    Dim cht As ChartObject
    Dim chartDataRange As Range
    Dim xLabelsRange As Range
    
    ' Assuming your x-axis data is in row 3 and y-axis data is in row 1
    Set ws = Application.ActiveWorkbook.Sheets("Sheet1") ' Replace "Sheet1" with your sheet name
    Set chartDataRange = ws.Range("B3:N3")
    Set xLabelsRange = ws.Range("B1:N1")
    
    ' Create a new chart object
    Set cht = ws.ChartObjects.Add(Left:=100, Width:=600, Top:=75, Height:=300)
    
    ' Set the chart data range
    cht.Chart.SetSourceData Source:=chartDataRange
    
    ' Set x-axis labels
With cht.Chart.Axes(xlCategory, xlPrimary)
    .CategoryType = xlCategoryScale
    .TickMarkSpacing = 5 ' Adjust the spacing of tick marks if needed
    .CategoryNames = xLabelsRange
End With
    
    ' Set chart type to Line Chart
    cht.Chart.ChartType = xlLine
    
    ' Set chart title
    cht.Chart.HasTitle = True
    cht.Chart.ChartTitle.Text = "My Line Chart"
End Sub


