
' This script makes a Packed Bars chart from a two-column Excel data in columns A and B.
' For general info, see my blog post at:
'      https://community.jmp.com/t5/JMP-Blog/Introducing-packed-bars-a-new-chart-form/ba-p/39972
' This is my first VBA script, so don't assume I know what I'm doing but otherwise use
' this code as you see fit.
'
' The data must be in the first two columns of sheet one with one row of header labels. A few
' tunable parameters are in columns D and E of the sheet:
'     the number of primary categories (and the number of rows in the chart)
'     value threshold for showing secondary labels (% or smallest primary bar value). Use 1.0 for none.
'     graph width
'     graph Height
'     gap adjustment, postive for bigger gaps, negative for smaller gaps
'
' What's working:
'    Primary bars ordered with prominent color and labels left aligned
'    Secondary bars randomish grays, some labeled
'    Hover labels on all bars
' Not working:
'    No UI to choose data range
'    No clipping/wrapping of labels that get too Long
'    The gap between bars is relative to the bar width, instead of being a fixed small size
'    The gap can be irregular if the bars don't fit into the graph height evenly
'    We not too careful about avoiding Excel's limit of 256 series in one chart
'    Some chart featurs are sensitive to current cell selection, requiring code to reset selection.
'
' Xan Gregg July 2017, developed on Excel for Mac 15.35
'
' The QuickSortArray function is from
'    https://stackoverflow.com/questions/4873182/sorting-a-multidimensionnal-array-in-vba/5104206#5104206


Sub MakePackedBars()

Dim sh As Worksheet
Set sh = ActiveWorkbook.Worksheets(1)

' Strangely, having a non-empty cell selected messes up the chart
sh.Range("$D$21").Select

nPrimaryBars = sh.Range("$E$8").Value
If nPrimaryBars <= 1 Then
    nPrimaryBars = 10
End If

' fraction of lowest primary value
secondaryLabelThreshold = sh.Range("$E$9").Value
If secondaryLabelThreshold <= 0 Then
    secondaryLabelThreshold = 0.8
End If

graphWidth = sh.Range("$E$10").Value
If graphWidth <= 100 Then
    graphWidth = 1000
End If

graphHeight = sh.Range("$E$11").Value
If graphHeight <= 100 Then
    graphHeight = 500
End If

gapAdjustment = sh.Range("$E$12").Value

Dim data As Variant

' read many rows and then trim out the empty ones
data = sh.Range("$A$2:$B$10001").Value

' count the non-empty category names
nCategories = 1
While data(nCategories, 1) <> ""
    nCategories = nCategories + 1
Wend

' Redim doesn't seem to work for this data, so just reread the smalled array from sheet
' ReDim Preserve data(1 To nCategories - 1, 1 To 2) As String
data = sh.Range("$A$2:$B$" & nCategories).Value

nCategories = UBound(data)

ChartTitle = "Top " & nPrimaryBars & " out of " & nCategories & " " & sh.Range("$A$1").Value & " " & sh.Range("$B$1").Value & " values"

' from https://stackoverflow.com/questions/4873182/sorting-a-multidimensionnal-array-in-vba/5104206#5104206
Call QuickSortArray(data, , , 2)

' categories and values for all the data
cat = Application.Index(data, 0, 1)
y = Application.Index(data, 0, 2)

' The general idea is to reshape the data into nPrimaryBars fake categories and many series.

' arrays for packed data

Dim pcat As Variant    ' [nrows] fake category name for each row
Dim psum As Variant    ' [nrows] running sum of values in each row
Dim pseries As Variant ' [nrows] next series to be used for each row
Dim py As Variant      ' [nrows x nSeries] values
Dim plabel As Variant  ' [nrows x nSeries] category labels

ReDim pcat(1 To nPrimaryBars) As String
ReDim psum(1 To nPrimaryBars) As Double
ReDim pseries(1 To nPrimaryBars) As Long
' over-estimate nSeries as nCategories since we don't yet know nSeries
ReDim py(1 To nPrimaryBars, 1 To nCategories) As Double
ReDim plabel(1 To nPrimaryBars, 1 To nCategories) As String

Randomize

For i = 1 To nPrimaryBars
    j = nPrimaryBars + 1 - i      ' because cat(1) comes out on the bottom row
    iall = nCategories + 1 - i    ' because it was sorted ascending
    pcat(j) = "row " & i
    py(j, 1) = y(iall, 1)
    plabel(j, 1) = cat(iall, 1)
    psum(j) = py(j, 1)
    pseries(j) = j + 1
Next i

secondaryLabelThreshold = secondaryLabelThreshold * psum(1)

nSeries = nPrimaryBars + 1
maxSum = psum(nPrimaryBars)

For i = nPrimaryBars + 1 To nCategories
    iall = nCategories + 1 - i
    ip = Application.Match(Application.Min(psum), psum, 0) ' smallest stack
    py(ip, pseries(ip)) = y(iall, 1)
    plabel(ip, pseries(ip)) = cat(iall, 1)
    If pseries(ip) > nSeries Then
        nSeries = pseries(ip)
    End If
    pseries(ip) = pseries(ip) + Int(2 * Rnd + 1)    ' some randomness to grays
    psum(ip) = psum(ip) + y(iall, 1)
    If psum(ip) > maxSum Then
        maxSum = psum(ip)
    End If
Next i

Dim ch As Chart

'ch = Charts.Add    ' creates a new chart sheet but then the chart can't be sized freely
Set ch = sh.Shapes.AddChart.Chart

ch.ChartType = xlBarStacked
For i = 1 To nSeries
    With ch.SeriesCollection.NewSeries
        .Values = Application.Index(py, 0, i)
        .XValues = Application.Index(plabel, 0, i) ' pcat
        If i = 1 Then
            .Name = "Primary"
            .HasDataLabels = True
            .DataLabels.ShowCategoryName = True
            .DataLabels.ShowValue = False
            .DataLabels.Position = xlLabelPositionInsideBase
            .DataLabels.Font.Color = RGB(255, 255, 255)
            .DataLabels.Font.Size = 16
        Else
            .Name = ""
            g = 240 - 3 * (i Mod 4)    ' a few light grays
            .Format.Fill.ForeColor.RGB = RGB(g, g, g)
            .HasDataLabels = True   ' but most are disabled individually
            .DataLabels.ShowValue = False
            .DataLabels.Font.Color = RGB(170, 170, 170)
            .DataLabels.Font.Size = 13
            For j = 1 To nPrimaryBars
                If py(j, i) >= secondaryLabelThreshold Then
                    .DataLabels(j).ShowCategoryName = True
                    .DataLabels(j).Text = plabel(j, i)
                End If
            Next j
           .DataLabels.ShowLegendKey = False
        End If
    End With
Next i

' %, we want the smallest visible gap, which unfortunative depends on the image height
gapSpacing = 100 * (gapAdjustment + 1.2 * nPrimaryBars / graphHeight)
If gapSpacing < 0 Then
    gapSpacing = 0
End If
ch.ChartGroups(1).GapWidth = gapSpacing

ch.HasAxis(xlCategory) = False   ' labeled inline instead
ch.HasLegend = False
ch.Axes(xlValue).TickLabels.Font.Size = 16
ch.Axes(xlValue).MajorTickMark = xlTickMarkOutside
ch.Axes(xlValue).HasMajorGridlines = False
mu = ch.Axes(xlValue).MajorUnit
ch.Axes(xlValue).MinimumScale = 0
ch.Axes(xlValue).MaximumScale = maxSum * 1.02
ch.Axes(xlValue).MajorUnit = mu
ch.Axes(xlValue).Border.Color = RGB(33, 33, 33)
ch.Axes(xlValue).HasTitle = True
ch.Axes(xlValue).AxisTitle.Text = sh.Range("$B$1").Value
ch.Axes(xlValue).AxisTitle.Font.Size = 16

ch.HasTitle = True
ch.ChartTitle.Text = ChartTitle
ch.ChartTitle.Font.Size = 16

ch.ChartArea.Select

With ActiveChart
    .ChartArea.Height = graphHeight
    .ChartArea.Width = graphWidth
    .ChartArea.Top = 20
   .ChartArea.Left = 300
End With

End Sub
