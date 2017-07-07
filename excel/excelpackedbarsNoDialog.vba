' MakePackedBar1 - Minor modifications by Jon Peltier
'                  '' explanations preceded by two single quotes
'                - Based on original by Xan Gregg

'' requires declaration of all variables
Option Explicit

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
'     graph height
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
'
' Xan Gregg July 2017, developed on Excel for Mac 15.35
'
' The QuickSortArray function is from
'    https://stackoverflow.com/questions/4873182/sorting-a-multidimensionnal-array-in-vba/5104206#5104206

Sub MakePackedBars1()
  Dim sh As Worksheet
  Dim nPrimaryBars As Long
  Dim SecondaryLabelThreshold As Double
  Dim GraphWidth As Double
  Dim GraphHeight As Double
  Dim GapAdjustment As Double
  Dim nCategories As Long, nSeries As Long
  Dim rData As Range
  Dim data As Variant
  Dim ch As Chart
  Dim ChartTitle As String
  Dim cat As Variant, y As Variant
  Dim pcat As Variant     ' [nrows] fake category name for each row
  Dim psum As Variant     ' [nrows] running sum of values in each row
  Dim pseries As Variant  ' [nrows] next series to be used for each row
  Dim py As Variant       ' [nrows x nSeries] values
  Dim plabel As Variant   ' [nrows x nSeries] category labels
  Dim maxSum As Double
  Dim i As Long, j As Long, iall As Long, ip As Long
  Dim g As Long
  Dim mu As Double

  Set sh = ActiveSheet

  nPrimaryBars = sh.Range("$E$8").Value
  If nPrimaryBars <= 1 Then
    nPrimaryBars = 10
  End If

  ' fraction of lowest primary value
  SecondaryLabelThreshold = sh.Range("$E$9").Value
  If SecondaryLabelThreshold <= 0 Then
    SecondaryLabelThreshold = 0.8
  End If

  GraphWidth = sh.Range("$E$10").Value
  If GraphWidth <= 100 Then
    GraphWidth = 1000  ' huge; Excel default = 360 (5")
  End If

  GraphHeight = sh.Range("$E$11").Value
  If GraphHeight <= 100 Then
    GraphHeight = 500  ' huge; Excel default = 211 (3")
  End If

  ‘ All the data in columns A and B
  Set rData = sh.Range("A1").CurrentRegion
  data = rData.Offset(1).Resize(rData.Rows.Count - 1).Value2

  nCategories = UBound(data)

  ChartTitle = "Top " & nPrimaryBars & " out of " & nCategories & " " & sh.Range("$A$1").Value & " " & sh.Range("$B$1").Value & " values"

  ' from https://stackoverflow.com/questions/4873182/sorting-a-multidimensionnal-array-in-vba/5104206#5104206
  Call QuickSortArray(data, , , 2)

  ' categories and values for all the data
  ' set row=0 to get entire column
  cat = Application.Index(data, 0, 1)
  y = Application.Index(data, 0, 2)

  ' The general idea is to reshape the data into nPrimaryBars fake categories and many series.

  ' arrays for packed data
  ReDim pcat(1 To nPrimaryBars) As String
  ReDim psum(1 To nPrimaryBars) As Double
  ReDim pseries(1 To nPrimaryBars) As Long
  ' over-estimate nSeries as nCategories since we don't yet know nSeries
  ReDim py(1 To nPrimaryBars, 1 To nCategories) As Double
  ReDim plabel(1 To nPrimaryBars, 1 To nCategories) As String

  Randomize

  For i = 1 To nPrimaryBars
    j = nPrimaryBars + 1 - i      ' because cat(1) comes out on the bottom row <-- ''easy to fix...
    iall = nCategories + 1 - i    ' because it was sorted ascending
    pcat(j) = "row " & i
    py(j, 1) = y(iall, 1)
    plabel(j, 1) = cat(iall, 1)
    psum(j) = py(j, 1)
    pseries(j) = j + 1
  Next i

  SecondaryLabelThreshold = SecondaryLabelThreshold * psum(1)

  nSeries = nPrimaryBars + 1
  maxSum = psum(nPrimaryBars)

  For i = nPrimaryBars + 1 To nCategories
    iall = nCategories + 1 - i
    ip = Application.Match(Application.Min(psum), psum, 0)  ' smallest stack
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

  Set ch = sh.Shapes.AddChart(xlBarStacked).Chart

  With ch
    .ChartArea.ClearContents
    
    For i = 1 To nSeries
      With .SeriesCollection.NewSeries
        .Values = Application.Index(py, 0, i)
        .XValues = Application.Index(plabel, 0, i)  ' pcat
        .Format.Line.ForeColor.RGB = vbWhite '' see note on gap width below
        If i = 1 Then
          .Name = "Primary"
          .HasDataLabels = True
          With .DataLabels
            .ShowCategoryName = True
            .ShowValue = False
            .Position = xlLabelPositionInsideBase
            .Font.Color = RGB(255, 255, 255)
            .Font.Size = 16
          End With
        Else
          .Name = ""
          g = 240 - 3 * (i Mod 4)    ' a few light grays
          .Format.Fill.ForeColor.RGB = RGB(g, g, g)
          .HasDataLabels = True   ' but most are disabled individually
          With .DataLabels
            .ShowValue = False
            .Font.Color = RGB(170, 170, 170)
            .Font.Size = 13
          End With
          For j = 1 To nPrimaryBars
            If py(j, i) >= SecondaryLabelThreshold Then
              .DataLabels(j).ShowCategoryName = True
              .DataLabels(j).Text = plabel(j, i)
            End If
          Next j
          .DataLabels.ShowLegendKey = False
        End If
      End With
    Next i

    '' why not gap width of 0 but thin white border on bars?
    ''  GapAdjustment = sh.Range("$E$12").Value
    ''  ' %, we want the smallest visible gap, which unfortunately depends on the image height
    ''  gapSpacing = 100 * (GapAdjustment + 1.2 * nPrimaryBars / GraphHeight)
    ''  If gapSpacing < 0 Then
    ''    gapSpacing = 0
    ''  End If
    ''  .ChartGroups(1).GapWidth = gapSpacing

    .ChartGroups(1).GapWidth = 0

    .HasAxis(xlCategory) = False   ' labeled inline instead
    .HasLegend = False
    .Axes(xlValue).TickLabels.Font.Size = 16
    .Axes(xlValue).MajorTickMark = xlTickMarkOutside
    .Axes(xlValue).HasMajorGridlines = False
    mu = .Axes(xlValue).MajorUnit
    .Axes(xlValue).MinimumScale = 0
    .Axes(xlValue).MaximumScale = maxSum * 1.02
    .Axes(xlValue).MajorUnit = mu
    .Axes(xlValue).Border.Color = RGB(33, 33, 33)
    .Axes(xlValue).HasTitle = True
    .Axes(xlValue).AxisTitle.Text = sh.Range("$B$1").Value
    .Axes(xlValue).AxisTitle.Font.Size = 16

    .HasTitle = True
    .ChartTitle.Text = ChartTitle
    .ChartTitle.Font.Size = 16

    .ChartArea.Select

    With .Parent
      .Height = GraphHeight
      .Width = GraphWidth
      .Top = 20
      .Left = 300
    End With

  End With

End Sub
