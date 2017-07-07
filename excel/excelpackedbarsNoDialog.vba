' MakePackedBar1 - Minor modifications by Jon Peltier
'                  '' explanations preceded by two single quotes
'                - Based on original by Xan Gregg

'' requires declaration of all variables
Option Explicit

Const msAPP_NAME As String = "Xan Gregg"
Const msPROJ_NAME As String = "Packed Bar Chart"

Sub MakePackedBarsDialog()
  Dim frmMakeBackedBar As FMakePackedBarDialog
  Dim bCanceled As Boolean
  Dim sh As Worksheet
  Dim nPrimaryBars As Long
  Dim iSecondaryLabelThreshold As Long '' note change to Long not Double
  Dim GraphWidth As Double
  Dim GraphHeight As Double
  Dim nCategories As Long, nSeries As Long
  Dim rData As Range
  Dim sDataAddress As String
  Dim vData As Variant
  Dim ch As Chart
  Dim sChartTitle As String
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

  ''Set sh = ActiveWorkbook.Worksheets(1)
  Set sh = ActiveSheet

  ' retrieve settings from last time (or default)
  nPrimaryBars = GetSetting(msAPP_NAME, msPROJ_NAME, "Primary Bars", 10)
  iSecondaryLabelThreshold = GetSetting(msAPP_NAME, msPROJ_NAME, "Secondary Label Threshold", 80)
  GraphWidth = GetSetting(msAPP_NAME, msPROJ_NAME, "Graph Width", 1000)
  GraphHeight = GetSetting(msAPP_NAME, msPROJ_NAME, "Graph Height", 500)

  ' find data range
  Set rData = sh.Range("A1").CurrentRegion
  Set rData = rData.Resize(, 2)
  sDataAddress = "'" & ActiveSheet.Name & "'!" & rData.Address
  
  ' open dialog
  Set frmMakeBackedBar = New FMakePackedBarDialog
  With frmMakeBackedBar
  
    ' send options to dialog
    .DataRangeAddress = sDataAddress
    .PrimaryCategories = nPrimaryBars
    .SecondaryLabelThreshold = iSecondaryLabelThreshold
    .ChartHeight = GraphHeight
    .ChartWidth = GraphWidth
    
    ' show dialog
    .Show
    
    ' was dialog canceled?
    bCanceled = .Cancel
    
    ' get options from dialog
    If Not bCanceled Then
      sDataAddress = .DataRangeAddress
      nPrimaryBars = .PrimaryCategories
      iSecondaryLabelThreshold = .SecondaryLabelThreshold
      GraphHeight = .ChartHeight
      GraphWidth = .ChartWidth
    End If
  
  End With
  Unload frmMakeBackedBar
  
  If bCanceled Then GoTo ExitSub
  
  ' validate options
  If nPrimaryBars <= 1 Then
    nPrimaryBars = 10
  End If
  If iSecondaryLabelThreshold <= 0 Then
    iSecondaryLabelThreshold = 80
  End If
  If GraphWidth <= 100 Then
    GraphWidth = 1000  ' huge; Excel default = 360 (5")
  End If
  If GraphHeight <= 100 Then
    GraphHeight = 500  ' huge; Excel default = 211 (3")
  End If
  
  ' save options for next time
  SaveSetting msAPP_NAME, msPROJ_NAME, "Primary Bars", nPrimaryBars
  SaveSetting msAPP_NAME, msPROJ_NAME, "Secondary Label Threshold", iSecondaryLabelThreshold
  SaveSetting msAPP_NAME, msPROJ_NAME, "Graph Width", GraphWidth
  SaveSetting msAPP_NAME, msPROJ_NAME, "Graph Height", GraphHeight

  ' check range
  Set rData = Range(sDataAddress)
  If IsNumeric(rData.Cells(1, 2).Value2) Then
    ' no header row
    vData = rData.Value2
    nCategories = UBound(vData)
    sChartTitle = "Top " & nPrimaryBars & " out of " & nCategories & " values"
  Else
    vData = rData.Offset(1).Resize(rData.Rows.Count - 1).Value2
    nCategories = UBound(vData)
    sChartTitle = "Top " & nPrimaryBars & " out of " & nCategories & " " & rData.Cells(1, 1).Value2 & " " & rData.Cells(1, 2).Value2 & " values"
  End If

  ' from https://stackoverflow.com/questions/4873182/sorting-a-multidimensionnal-array-in-vba/5104206#5104206
  QuickSortArray vData, , , 2

  ' categories and values for all the data
  ' set row=0 to get entire column
  cat = Application.Index(vData, 0, 1)
  y = Application.Index(vData, 0, 2)

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

  iSecondaryLabelThreshold = iSecondaryLabelThreshold * psum(1)

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
            .Font.Size = GraphHeight / 32
          End With
        Else
          .Name = ""
          g = 240 - 3 * (i Mod 4)    ' a few light grays
          .Format.Fill.ForeColor.RGB = RGB(g, g, g)
          .HasDataLabels = True   ' but most are disabled individually
          With .DataLabels
            .ShowValue = False
            .Font.Color = RGB(170, 170, 170)
            .Font.Size = GraphHeight / 40
          End With
         ' For j = 1 To nPrimaryBars
          '  If py(j, i) >= iSecondaryLabelThreshold / 100 Then
          '    .DataLabels(j).ShowCategoryName = True
          '    .DataLabels(j).Text = plabel(j, i)
         '   End If
         ' Next j
          For j = 1 To nPrimaryBars
            If py(j, i) >= iSecondaryLabelThreshold / 100 Then
              .Points(j).HasDataLabel = True
              With .DataLabels(j)
                .ShowValue = False
                .ShowCategoryName = True
                .Text = plabel(j, i)
                .Font.Color = RGB(170, 170, 170)
                .Font.Size = GraphHeight / 40
              End With
            End If
          Next j
          .DataLabels.ShowLegendKey = False
        End If
      End With
    Next i

    '' why not gap width of 0 but thin white border on bars?
    ''  GapAdjustment = sh.Range("$E$12").Value
    ''  ' %, we want the smallest visible gap, which unfortunative depends on the image height
    ''  gapSpacing = 100 * (GapAdjustment + 1.2 * nPrimaryBars / GraphHeight)
    ''  If gapSpacing < 0 Then
    ''    gapSpacing = 0
    ''  End If
    ''  .ChartGroups(1).GapWidth = gapSpacing

    .ChartGroups(1).GapWidth = 0

    .HasAxis(xlCategory) = False   ' labeled inline instead
    .HasLegend = False
    With .Axes(xlValue)
      .TickLabels.Font.Size = GraphHeight / 32
      .MajorTickMark = xlTickMarkOutside
      .HasMajorGridlines = False
      mu = .MajorUnit
      .MinimumScale = 0
      .MaximumScale = maxSum * 1.02
      .MajorUnit = mu
      .Border.Color = RGB(33, 33, 33)
      .HasTitle = True
      .AxisTitle.Text = sh.Range("$B$1").Value
      .AxisTitle.Font.Size = GraphHeight / 32
    End With

    .HasTitle = True
    .ChartTitle.Text = sChartTitle
    .ChartTitle.Font.Size = GraphHeight / 32

    .ChartArea.Select

    ''  With ActiveChart
    ''    .ChartArea.Height = GraphHeight
    ''    .ChartArea.Width = GraphWidth
    ''    .ChartArea.Top = 20
    ''    .ChartArea.Left = 300
    ''  End With
    With .Parent
      .Height = GraphHeight
      .Width = GraphWidth
      .Top = 20
      .Left = 300
    End With

  End With
  
ExitSub:

End Sub


