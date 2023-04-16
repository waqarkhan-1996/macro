Private Sub Generate_Timing_Blocks()


'------------------------------------------------------------------------------------------
' Program to create timing blocks to be used on bulletin responses.
' Blocks created are proportional and can be stretched to fit various timing slide scales.
'------------------------------------------------------------------------------------------


    '--- Counters used in For loops ---
    '
    Dim i As Integer, j As Integer, k As Integer, l As Integer, m As Integer, n As Integer
    Dim x As Integer, y As Integer, z As Integer, o As Integer, p As Integer, t As Integer
    
    '--- Used to create shapes ---
    '
    Dim s As Shape
    
    '--- Used to position shapes ---
    '
    Dim Shp1 As Shape, Shp2 As Shape
    Dim PShp1 As Shape, PShp2 As Shape
    Dim TShp1 As Shape, TShp2 As Shape
    Dim GShp1 As Shape, GShp2 As Shape
    Dim HShp1 As Shape, HShp2 As Shape
    Dim IShp1 As Shape, IShp2 As Shape
    Dim YShp1 As Shape, YShp2 As Shape
    Dim MShp1 As Shape, MShp2 As Shape
    
    '--- Used to clear old shapes ---
    '
    Dim c As Object
    
    '--- Used to format shapes ---
    '
    Dim d As Shape
    
    '--- Used to select and copy shapes ---
    '
    Dim R As Range
    Dim q As Shape

    '--- Define variables for number of activities input by the user
    '
    Dim NumTaskA As Integer, NumTaskB As Integer, NumTaskC As Integer, NumTaskD As Integer, NumTaskE As Integer, NumTaskF As Integer
    
    '--- Define variable for the scale selected by the user
    '
    Dim ScaleSlct As Integer
    
    '--- Define variable for back end calculations
    '
    Dim NumYears As Integer, FirstYear As Integer
    Dim MinDate As Date, MaxDate As Date
    Dim YearWidth As Double, MonthWidth As Double
    
    '---- Define array for Years and Months
    '
    Dim YearList(1 To 10) As Integer     'Tool will work for maximum 10-year span
    Dim MonthList(1 To 120) As String
    
    '---- Define Arrays for Level 1 ----
    '
    Dim LevelAText(1 To 13) As String
    Dim LevelAWidth(1 To 13) As Double
    Dim LevelAGap(1 To 13) As Double
    
    '---- Define Arrays for Level 2 ----
    '
    Dim LevelBText(1 To 13) As String
    Dim LevelBWidth(1 To 13) As Double
    Dim LevelBGap(1 To 13) As Double
    
    '---- Define Arrays for Level 3 ----
    '
    Dim LevelCText(1 To 13) As String
    Dim LevelCWidth(1 To 13) As Double
    Dim LevelCGap(1 To 13) As Double
    
    '---- Define Arrays for Level 4 ----
    '
    Dim LevelDText(1 To 13) As String
    Dim LevelDWidth(1 To 13) As Double
    Dim LevelDGap(1 To 13) As Double
    
    '---- Define Arrays for Level 5 ----
    '
    Dim LevelEText(1 To 13) As String
    Dim LevelEWidth(1 To 13) As Double
    Dim LevelEGap(1 To 13) As Double
    
    '---- Define Arrays for Level 6 ----
    '
    Dim LevelFText(1 To 13) As String
    Dim LevelFWidth(1 To 13) As Double
    Dim LevelFGap(1 To 13) As Double
    
    '--- Stop screen flicker while executing macro ---
    '
    Application.ScreenUpdating = False
    
    '---------------------------------------------------------------
    ' Clear timing blocks from previous runs of the file
    ' Do not delete the drop down menus and GO and RESET buttons
    '---------------------------------------------------------------
    
    For Each c In ActiveSheet.DrawingObjects
        If c.Name <> "GO_Arrow_1" And c.Name <> "Scale_drop_down" _
        And c.Name <> "Calendar_drop_down" And c.Name <> "Reset_Button" Then
           c.Delete
        End If
    Next c
    
    '------------------------------------------------------------------------
    ' Check if minimum required cells are populated.
    ' If they are empty, show a message to the user and exit the subroutine
    '------------------------------------------------------------------------
    
    If IsEmpty(Range("Timing!C6")) = True Or IsEmpty(Range("Timing!D6")) = True Or IsEmpty(Range("Timing!E6")) = True Or IsEmpty(Range("Timing!C22")) = True Or IsEmpty(Range("Timing!D22")) = True Or IsEmpty(Range("Timing!E22")) = True Then

        MsgBox "Minimum required input missing" & vbCrLf & vbCrLf & "For Level 1 : Activity, start date, and # of weeks are required", vbCritical, "Critical Error"
        Exit Sub

    End If
    
    '---- Calculate how many Activities are populated by the user ----
    '
    NumTaskA = WorksheetFunction.CountA(Range("TIMING!E6:E18"))
    NumTaskB = WorksheetFunction.CountA(Range("TIMING!K6:K18"))
    NumTaskC = WorksheetFunction.CountA(Range("TIMING!Q6:Q18"))
    NumTaskD = WorksheetFunction.CountA(Range("TIMING!E22:E34"))
    NumTaskE = WorksheetFunction.CountA(Range("TIMING!K22:K34"))
    NumTaskF = WorksheetFunction.CountA(Range("TIMING!Q22:Q34"))
    
    '---- Calculate number years of years to show on calendar
    '     This is calculated based on the earliest and latest dates for activities
    '
    NumYears = Year(WorksheetFunction.Max(Range("Timing!F6:F18"), Range("Timing!L6:L18"), Range("Timing!R6:R18"), Range("Timing!F22:R34"), Range("Timing!L22:L34"), Range("Timing!R22:R34"))) - Year(WorksheetFunction.Min(Range("Timing!D6:D18"), Range("Timing!J6:J18"), Range("Timing!P6:P18"), Range("Timing!D22:D34"), Range("Timing!J22:J34"), Range("Timing!P22:P34"))) + 1
    
    '---- Calculate first year value to show on calendar
    '
    FirstYear = Year(WorksheetFunction.Min(Range("Timing!D6:D18"), Range("Timing!J6:J18"), Range("Timing!P6:P18"), Range("Timing!D22:D34"), Range("Timing!J22:J34"), Range("Timing!P22:P34")))
    
    '---- Calculate Mix, Max dates for Calendar
    '
    MinDate = DateSerial(Year(WorksheetFunction.Min(Range("Timing!D6:D18"), Range("Timing!J6:J18"), Range("Timing!P6:P18"), Range("Timing!D22:D34"), Range("Timing!J22:J34"), Range("Timing!P22:P34"))), 1, 1)
    MaxDate = DateSerial(Year(WorksheetFunction.Max(Range("Timing!F6:F18"), Range("Timing!L6:L18"), Range("Timing!R6:R18"), Range("Timing!F22:F34"), Range("Timing!L22:L34"), Range("Timing!R22:R34"))), 1, 1)
    
        'Ref WorksheetFunction.Min
        'Ref WorksheetFunction.Max
    
    '--- Store Scale selected by the user
    '
    ScaleSlct = Range("TIMING!P2").Value       'Based on Standard/Double/Triple scale
    
    '---- Calculate width of the Year Box to show on calendar
    '
    YearWidth = 187.2 * ScaleSlct
    
    '---- Calculate width of the Month Box to show on calendar
    '
    MonthWidth = 15.6 * ScaleSlct
    
    '-------------------------------------------------------------------
    ' Level 1 or A (Primary Tasks)
    '-------------------------------------------------------------------
    
    For i = 1 To NumTaskA
        LevelAText(i) = Range("TIMING!C" & i + 5).Value & Chr(10) & Range("TIMING!E" & i + 5).Value & "w"
        LevelAWidth(i) = (Range("TIMING!E" & i + 5).Value * 3.6) * ScaleSlct
    Next i
    
    LevelAGap(1) = 0#   'set first value in array to zero
    
    For i = 2 To NumTaskA
        If Range("TIMING!F" & i + 4).Value > 0 Then
            LevelAGap(i) = (Range("TIMING!D" & i + 5).Value - Range("TIMING!F" & i + 4).Value) / 7 * 3.6 * ScaleSlct
        End If
    Next i
    
    '-------------------------------------------------------------------
    ' Level 2 or B (Secondary Tasks)
    '-------------------------------------------------------------------
    
    For i = 1 To NumTaskB
        LevelBText(i) = Range("TIMING!I" & i + 5).Value & Chr(10) & Range("TIMING!K" & i + 5).Value & "w"
        LevelBWidth(i) = (Range("TIMING!K" & i + 5).Value * 3.6) * ScaleSlct
    Next i
    
    LevelBGap(1) = (Range("TIMING!J6").Value - Range("TIMING!D6").Value) / 7 * 3.6 * ScaleSlct
    
    For i = 2 To NumTaskB
        If Range("TIMING!L" & i + 4).Value > 0 Then
            LevelBGap(i) = (Range("TIMING!J" & i + 5).Value - Range("TIMING!L" & i + 4).Value) / 7 * 3.6 * ScaleSlct
        End If
    Next i
    
    '-------------------------------------------------------------------
    ' Level 3 or C (Tertiary Tasks)
    '-------------------------------------------------------------------
    
    For i = 1 To NumTaskC
        LevelCText(i) = Range("TIMING!O" & i + 5).Value & Chr(10) & Range("TIMING!Q" & i + 5).Value & "w"
        LevelCWidth(i) = (Range("TIMING!Q" & i + 5).Value * 3.6) * ScaleSlct
    Next i
    
    LevelCGap(1) = (Range("TIMING!P6").Value - Range("TIMING!D6").Value) / 7 * 3.6 * ScaleSlct
    
    For i = 2 To NumTaskC
        If Range("TIMING!L" & i + 4).Value > 0 Then
            LevelCGap(i) = (Range("TIMING!P" & i + 5).Value - Range("TIMING!R" & i + 4).Value) / 7 * 3.6 * ScaleSlct
        End If
    Next i
    '-------------------------------------------------------------------
    ' Level 4 or D (Quaternary Tasks)
    '-------------------------------------------------------------------
    
    For i = 1 To NumTaskD
        LevelDText(i) = Range("TIMING!C" & i + 5).Value & Chr(10) & Range("TIMING!E" & i + 5).Value & "w"
        LevelDWidth(i) = (Range("TIMING!E" & i + 5).Value * 3.6) * ScaleSlct
    Next i
    
    LevelDGap(1) = 0#   'set first value in array to zero
    
    For i = 2 To NumTaskD
        If Range("TIMING!F" & i + 4).Value > 0 Then
            LevelDGap(i) = (Range("TIMING!D" & i + 5).Value - Range("TIMING!F" & i + 4).Value) / 7 * 3.6 * ScaleSlct
        End If
    Next i
    
    '-------------------------------------------------------------------
    ' Level 5 or E (Quinary Tasks)
    '-------------------------------------------------------------------
    
    For i = 1 To NumTaskE
        LevelEText(i) = Range("TIMING!I" & i + 5).Value & Chr(10) & Range("TIMING!K" & i + 5).Value & "w"
        LevelEWidth(i) = (Range("TIMING!K" & i + 5).Value * 3.6) * ScaleSlct
    Next i
    
    LevelEGap(1) = (Range("TIMING!J22").Value - Range("TIMING!D22").Value) / 7 * 3.6 * ScaleSlct
    
    For i = 2 To NumTaskE
        If Range("TIMING!L" & i + 4).Value > 0 Then
            LevelEGap(i) = (Range("TIMING!J" & i + 5).Value - Range("TIMING!L" & i + 4).Value) / 7 * 3.6 * ScaleSlct
        End If
    Next i
    
    '-------------------------------------------------------------------
    ' Level 6 or F (Senary Tasks)
    '-------------------------------------------------------------------
    
    For i = 1 To NumTaskF
        LevelFText(i) = Range("TIMING!O" & i + 5).Value & Chr(10) & Range("TIMING!Q" & i + 5).Value & "w"
        LevelFWidth(i) = (Range("TIMING!Q" & i + 5).Value * 3.6) * ScaleSlct
    Next i
    
    LevelFGap(1) = (Range("TIMING!P22").Value - Range("TIMING!D22").Value) / 7 * 3.6 * ScaleSlct
    
    For i = 2 To NumTaskF
        If Range("TIMING!L" & i + 4).Value > 0 Then
            LevelFGap(i) = (Range("TIMING!P" & i + 5).Value - Range("TIMING!R" & i + 4).Value) / 7 * 3.6 * ScaleSlct
        End If
    Next i
    
    '------------------------------------------------------------------------
    ' Populate Year List Array
    '------------------------------------------------------------------------
    
    YearList(1) = FirstYear
    
    For i = 2 To NumYears
        YearList(i) = YearList(i - 1) + 1
    Next i
    
    '------------------------------------------------------------------------
    ' Populate Month List Array 'Refer to DateAdd, Month, MonthName, Left
    '------------------------------------------------------------------------
    
    For i = 1 To (NumYears * 12)
        MonthList(i) = Left(MonthName(Month(DateAdd("m", (i - 1), MinDate))), 1)
    Next i

    '---------------------------------------------------------
    ' Add timing blocks for Primary tasks
    '---------------------------------------------------------
    '
    For i = 1 To NumTaskA

        'Create rounded rectangle
        Set s = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 50, 200, 50, 21.6)
        
        'Give the created shape a name
        s.Name = "Time_Shape_" & i
        
        'Format the shape
        ActiveSheet.Shapes.Range(Array("Time_Shape_" & i)).Select
        Selection.Width = LevelAWidth(i)
        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = LevelAText(i)

    Next i
 
    '---------------------------------------------------------
    ' Add timing blocks for Secondary tasks
    '---------------------------------------------------------
    '
    For j = 1 To NumTaskB

        'Create rounded rectangle
        Set s = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 50, 200, 50, 21.6)
        
        'Give the created shape a name
        s.Name = "Secondary_Shape_" & j
        
        'Format the shape
        ActiveSheet.Shapes.Range(Array("Secondary_Shape_" & j)).Select
        Selection.Width = LevelBWidth(j)
        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = LevelBText(j)
               
    Next j

    '---------------------------------------------------------
    ' Add timing blocks for Tertiary tasks
    '---------------------------------------------------------
    '
    For k = 1 To NumTaskC

        'Create rounded rectangle
        Set s = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 50, 200, 50, 21.6)
        
        'Give the created shape a name
        s.Name = "Tertiary_Shape_" & k
        
        'Format the shape
        ActiveSheet.Shapes.Range(Array("Tertiary_Shape_" & k)).Select
        Selection.Width = LevelCWidth(k)
        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = LevelCText(k)
               
    Next k
    
    For l = 1 To NumTaskD

        'Create rounded rectangle
        Set s = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 50, 200, 50, 21.6)
        
        'Give the created shape a name
        s.Name = "Quaternary_Shape_" & l
        
        'Format the shape
        ActiveSheet.Shapes.Range(Array("Quaternary_Shape_" & l)).Select
        Selection.Width = LevelDWidth(k)
        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = LevelDText(l)
               
    Next l

    For m = 1 To NumTaskE

        'Create rounded rectangle
        Set s = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 50, 200, 50, 21.6)
        
        'Give the created shape a name
        s.Name = "Quinary_Shape_" & l
        
        'Format the shape
        ActiveSheet.Shapes.Range(Array("Quinary_Shape_" & l)).Select
        Selection.Width = LevelEWidth(k)
        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = LevelEText(l)
               
    Next m
    
    For n = 1 To NumTaskF

        'Create rounded rectangle
        Set s = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 50, 200, 50, 21.6)
        
        'Give the created shape a name
        s.Name = "Senary_Shape_" & l
        
        'Format the shape
        ActiveSheet.Shapes.Range(Array("Senary_Shape_" & l)).Select
        Selection.Width = LevelFWidth(k)
        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = LevelFText(l)
    Next n
    
    '---------------------------------------------------------
    ' Format the added timing blocks
    '---------------------------------------------------------
    '
    For Each d In ActiveSheet.Shapes
        If d.AutoShapeType = msoShapeRoundedRectangle Then
            d.Select
            Selection.ShapeRange.TextFrame2.WordWrap = msoTrue
            Selection.ShapeRange.TextFrame2.MarginLeft = 0
            Selection.ShapeRange.TextFrame2.MarginRight = 0
            Selection.ShapeRange.TextFrame2.MarginTop = 0
            Selection.ShapeRange.TextFrame2.MarginBottom = 0
            Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
            Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
            Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            Selection.Placement = xlFreeFloating
            With Selection.ShapeRange.Line
                .Weight = 0.5
                .ForeColor.RGB = RGB(208, 208, 208)
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = -0.5
            End With
            With Selection.ShapeRange.Fill
                .TwoColorGradient msoGradientHorizontal, 1
                .ForeColor.RGB = RGB(194, 206, 230)
                .BackColor.RGB = RGB(227, 232, 246)
            End With
            With Selection.ShapeRange.TextFrame2.TextRange.Font
                .Name = "Arial"
                .Size = 6.5
            End With
            With Selection.ShapeRange.TextFrame2.TextRange.Font.Fill
                .Visible = msoTrue
                .ForeColor.RGB = RGB(1, 1, 1)
                .Transparency = 0
                .Solid
            End With
        End If
    Next d
        
    '*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
    ' Create year and month blocks only if user has selected that option
    '
    If Range("Timing!L2").Value = 1 Then
        
        '---------------------------------------------------------
        ' Add timing blocks for Years Blocks
        '---------------------------------------------------------
        
        For i = 1 To NumYears
    
            'Create rectangle
            '
            Set s = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 50, 200, 50, 16)
            
            'Give the created shape a name
            '
            s.Name = "Year_Shape_" & i
            
            'Format the shape
            '
            ActiveSheet.Shapes.Range(Array("Year_Shape_" & i)).Select
                 
            Selection.Width = YearWidth
            Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = YearList(i)
    
            With Selection.ShapeRange.TextFrame2.TextRange.Font
                .Name = "Arial"
                .Size = 9.5
                .Bold = msoTrue
                .Spacing = 0.6
            End With
               
            With Selection.ShapeRange.Fill
                .Visible = msoTrue
                .ForeColor.RGB = RGB(43, 63, 103)
                .Transparency = 0
                .Solid
            End With
     
        Next i
        
        '---------------------------------------------------------
        ' Add timing blocks for Months Blocks
        '---------------------------------------------------------
        
        For i = 1 To (NumYears * 12)
    
            'Create rectangle
            '
            Set s = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 50, 200, 50, 16)
            
            'Give the created shape a name
            '
            s.Name = "Month_Shape_" & i
            
            'Format the shape
            '
            ActiveSheet.Shapes.Range(Array("Month_Shape_" & i)).Select
                 
            Selection.Width = MonthWidth
            Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = MonthList(i)
     
            With Selection.ShapeRange.TextFrame2.TextRange.Font.Fill
                .Visible = msoTrue
                .ForeColor.RGB = RGB(0, 0, 0)
                .Transparency = 0
                .Solid
            End With
            
            With Selection.ShapeRange.TextFrame2.TextRange.Font
                .Name = "Arial"
                .Size = 9
            End With
               
            With Selection.ShapeRange.Fill
                .Visible = msoTrue
                .ForeColor.RGB = RGB(255, 255, 255)
                .Transparency = 0
                .Solid
            End With
     
        Next i
    
        '---------------------------------------------------------
        ' Format the added shapes for years and months
        '---------------------------------------------------------

        For Each d In ActiveSheet.Shapes
            If d.AutoShapeType = msoShapeRectangle Then
                d.Select
                Selection.ShapeRange.TextFrame2.WordWrap = msoTrue
                Selection.ShapeRange.TextFrame2.MarginLeft = 0
                Selection.ShapeRange.TextFrame2.MarginRight = 0
                Selection.ShapeRange.TextFrame2.MarginTop = 0
                Selection.ShapeRange.TextFrame2.MarginBottom = 0
                Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
                Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
                Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                Selection.Placement = xlFreeFloating
                With Selection.ShapeRange.Line
                    .Weight = 0.5
                    .ForeColor.RGB = RGB(255, 255, 255)
                    .ForeColor.TintAndShade = 0
                    .ForeColor.Brightness = -0.5
                End With
    
            End If
        Next d
        
        '--- Set position of First Year and First Month blocks ---
        '
            Set YShp1 = ActiveSheet.Shapes("Year_Shape_1")
            YShp1.Top = ActiveSheet.Range("Timing!C40").Top
            YShp1.Left = ActiveSheet.Range("Timing!C40").Left
            
        '
            If NumYears > 0 Then
                Set MShp1 = ActiveSheet.Shapes("Month_Shape_1")
                MShp1.Top = YShp1.Top + YShp1.Height
                MShp1.Left = YShp1.Left
            End If
            
        '--- Set position of the Remaining Year and Month blocks ---
        '
            For x = 2 To NumYears
                Set YShp1 = ActiveSheet.Shapes("Year_Shape_" & x - 1)
                Set YShp2 = ActiveSheet.Shapes("Year_Shape_" & x)
                YShp2.Top = YShp1.Top
                YShp2.Left = YShp1.Left + YShp1.Width
            Next x
            
            For y = 2 To (NumYears * 12)
      
                Set MShp1 = ActiveSheet.Shapes("Month_Shape_" & y - 1)
                Set MShp2 = ActiveSheet.Shapes("Month_Shape_" & y)
                MShp2.Top = MShp1.Top
                MShp2.Left = MShp1.Left + MShp1.Width
        
            Next y
        
    End If
    
    '
    '*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
        
    '---------------------------------------------------------
    ' Position the timing blocks
    '---------------------------------------------------------
    '
        Set Shp1 = ActiveSheet.Shapes("Time_Shape_1")
        Shp1.Top = ActiveSheet.Range("TIMING!C45").Top
        Shp1.Left = ActiveSheet.Range("TIMING!C45").Left + ((Range("Timing!D45") - MinDate) / 7) * 3.6 * ScaleSlct
        
        If NumTaskB > 0 Then
            Set PShp1 = ActiveSheet.Shapes("Secondary_Shape_1")
            PShp1.Top = Shp1.Top + Shp1.Height
            PShp1.Left = Shp1.Left + LevelBGap(1)
        End If
        
        If NumTaskC > 0 Then
            Set TShp1 = ActiveSheet.Shapes("Tertiary_Shape_1")
            TShp1.Top = PShp1.Top + PShp1.Height
            TShp1.Left = Shp1.Left + LevelCGap(1)
        End If
        
        If NumTaskD > 0 Then
            Set GShp1 = ActiveSheet.Shapes("Quaternary_Shape_1")
            GShp1.Top = TShp1.Top + TShp1.Height
            GShp1.Left = Shp1.Left + LevelDGap(1)
        End If
        
        If NumTaskE > 0 Then
            Set HShp1 = ActiveSheet.Shapes("Quinary_Shape_1")
            HShp1.Top = GShp1.Top + GShp1.Height
            HShp1.Left = Shp1.Left + LevelEGap(1)
        End If
        
        If NumTaskF > 0 Then
            Set IShp1 = ActiveSheet.Shapes("Senary_Shape_1")
            IShp1.Top = HShp1.Top + HShp1.Height
            IShp1.Left = Shp1.Left + LevelFGap(1)
        End If

    '------------- Primary Blocks -------------------------------
    '
        For x = 2 To NumTaskA
  
            Set Shp1 = ActiveSheet.Shapes("Time_Shape_" & x - 1)
            Set Shp2 = ActiveSheet.Shapes("Time_Shape_" & x)
            Shp2.Top = Shp1.Top
            Shp2.Left = Shp1.Left + Shp1.Width + LevelAGap(x)
    
        Next x
        
    '------------- Secondary Blocks ---------------------------
    '
        For y = 2 To NumTaskB
  
            Set PShp1 = ActiveSheet.Shapes("Secondary_Shape_" & y - 1)
            Set PShp2 = ActiveSheet.Shapes("Secondary_Shape_" & y)
            PShp2.Top = PShp1.Top
            PShp2.Left = PShp1.Left + PShp1.Width + LevelBGap(y)
    
        Next y
    
    '------------- Tertiary Blocks ---------------------------
    '
        For z = 2 To NumTaskC
  
            Set TShp1 = ActiveSheet.Shapes("Tertiary_Shape_" & z - 1)
            Set TShp2 = ActiveSheet.Shapes("Tertiary_Shape_" & z)
            TShp2.Top = TShp1.Top
            TShp2.Left = TShp1.Left + TShp1.Width + LevelCGap(z)
    
        Next z
    
    '------------- Quaternary Blocks ---------------------------
    '
        For o = 2 To NumTaskD
  
            Set GShp1 = ActiveSheet.Shapes("Quaternary_Shape_" & o - 1)
            Set GShp2 = ActiveSheet.Shapes("Quaternary_Shape_" & o)
            GShp2.Top = GShp1.Top
            GShp2.Left = GShp1.Left + GShp1.Width + LevelDGap(o)
    
        Next o
        
    '------------- Quinary Blocks ---------------------------
    '
        For p = 2 To NumTaskE
  
            Set HShp1 = ActiveSheet.Shapes("Quinary_Shape_" & p - 1)
            Set HShp2 = ActiveSheet.Shapes("Quinary_Shape_" & p)
            HShp2.Top = HShp1.Top
            HShp2.Left = HShp1.Left + HShp1.Width + LevelEGap(p)
    
        Next p
    
    '------------- Senary Blocks ---------------------------
    '
        For t = 2 To NumTaskF
  
            Set IShp1 = ActiveSheet.Shapes("Quinary_Shape_" & t - 1)
            Set IShp2 = ActiveSheet.Shapes("Quinary_Shape_" & t)
            IShp2.Top = IShp1.Top
            IShp2.Left = IShp1.Left + IShp1.Width + LevelFGap(t)
    
        Next t
        
    '--- Create a Group of the timing blocks and select the group ---
    
    Set R = Range("TIMING!C40:Z60")

    For Each q In ActiveSheet.Shapes
        If Not Intersect(Range(q.TopLeftCell, q.BottomRightCell), R) Is Nothing Then _
            q.Select Replace:=False
        Next q
    
    Selection.ShapeRange.Group.Select
    
    '--- Allow screen refresh ---
    Application.ScreenUpdating = True
    
End Sub

Private Sub Reset_Timing_Blocks()

    Dim Msg, Style, Title, Response

    Msg = "About to Reset - do you want to continue ?"
    Style = vbExclamation + vbYesNo + vbDefaultButton2
    Title = "W A R N I N G !"

    Response = MsgBox(Msg, Style, Title)

    If Response = vbNo Then
      Exit Sub
    End If


'--- Stop screen flicker while executing macro ---
    Application.ScreenUpdating = False
    

'---- Clear timing blocks from previous runs of the file ----
'
    For Each c In ActiveSheet.DrawingObjects
        If c.Name <> "GO_Arrow_1" And c.Name <> "Scale_drop_down" _
        And c.Name <> "Calendar_drop_down" And c.Name <> "Reset_Button" Then
           c.Delete
        End If
    Next c

'----- Clear Level 1 input cells ---------------------------
'
    Worksheets("Timing").Range("C6:C18,D6,E6:E18").Select
    Selection.ClearContents
   
'----- Clear Level 2 input cells ---------------------------
'
    Worksheets("Timing").Range("I6:I18,J6,K6:K18").Select
    Selection.ClearContents
    
'----- Clear Level 3 input cells ---------------------------
'
    Worksheets("Timing").Range("O6:O18,P6,Q6:Q18").Select
    Selection.ClearContents

'----- Clear Level 4 input cells ---------------------------
'
    Worksheets("Timing").Range("C22:C34,D22,E22:E34").Select
    Selection.ClearContents

'----- Clear Level 5 input cells ---------------------------
'
    Worksheets("Timing").Range("I22:I34,J22,K22:K34").Select
    Selection.ClearContents
    
'----- Clear Level 6 input cells ---------------------------
'
    Worksheets("Timing").Range("O22:O34,P22,Q22:Q34").Select
    Selection.ClearContents
    
'-----Add correct forcumals for Level 1 blocks --------------
'
    Worksheets("Timing").Range("D7").Select
    ActiveCell.FormulaR1C1 = "=IF(ISBLANK(RC[1]),"""",R[-1]C[2])"
    Worksheets("Timing").Range("D7").Select
    Selection.AutoFill Destination:=Range("D7:D18"), Type:=xlFillValues
    
    Worksheets("Timing").Range("F6").Select
    ActiveCell.FormulaR1C1 = "=IF(ISBLANK(RC[-1]),"""",RC[-2]+(RC[-1]*7))"
    Worksheets("Timing").Range("F6").Select
    Selection.AutoFill Destination:=Range("F6:F18"), Type:=xlFillValues

'-----Add correct forcumals for Level 2 blocks --------------
'
    Worksheets("Timing").Range("J7").Select
    ActiveCell.FormulaR1C1 = "=IF(ISBLANK(RC[1]),"""",R[-1]C[2])"
    Worksheets("Timing").Range("J7").Select
    Selection.AutoFill Destination:=Range("J7:J18"), Type:=xlFillValues
    
    Worksheets("Timing").Range("L6").Select
    ActiveCell.FormulaR1C1 = "=IF(ISBLANK(RC[-1]),"""",RC[-2]+(RC[-1]*7))"
    Worksheets("Timing").Range("L6").Select
    Selection.AutoFill Destination:=Range("L6:L18"), Type:=xlFillValues

'-----Add correct forcumals for Level 3 blocks --------------
'
    Worksheets("Timing").Range("P7").Select
    ActiveCell.FormulaR1C1 = "=IF(ISBLANK(RC[1]),"""",R[-1]C[2])"
    Worksheets("Timing").Range("P7").Select
    Selection.AutoFill Destination:=Range("P7:P18"), Type:=xlFillValues
    
    Worksheets("Timing").Range("R6").Select
    ActiveCell.FormulaR1C1 = "=IF(ISBLANK(RC[-1]),"""",RC[-2]+(RC[-1]*7))"
    Worksheets("Timing").Range("R6").Select
    Selection.AutoFill Destination:=Range("R6:R18"), Type:=xlFillValues
    
    '-----Add correct forcumals for Level 4 blocks --------------
'
    Worksheets("Timing").Range("D23").Select
    ActiveCell.FormulaR1C1 = "=IF(ISBLANK(RC[1]),"""",R[-1]C[2])"
    Worksheets("Timing").Range("D23").Select
    Selection.AutoFill Destination:=Range("D23:D34"), Type:=xlFillValues
    
    Worksheets("Timing").Range("F22").Select
    ActiveCell.FormulaR1C1 = "=IF(ISBLANK(RC[-1]),"""",RC[-2]+(RC[-1]*7))"
    Worksheets("Timing").Range("F22").Select
    Selection.AutoFill Destination:=Range("F22:F34"), Type:=xlFillValues

'-----Add correct forcumals for Level 5 blocks --------------
'
    Worksheets("Timing").Range("J23").Select
    ActiveCell.FormulaR1C1 = "=IF(ISBLANK(RC[1]),"""",R[-1]C[2])"
    Worksheets("Timing").Range("J23").Select
    Selection.AutoFill Destination:=Range("J23:J34"), Type:=xlFillValues
    
    Worksheets("Timing").Range("L22").Select
    ActiveCell.FormulaR1C1 = "=IF(ISBLANK(RC[-1]),"""",RC[-2]+(RC[-1]*7))"
    Worksheets("Timing").Range("L22").Select
    Selection.AutoFill Destination:=Range("L22:L34"), Type:=xlFillValues

'-----Add correct forcumals for Level 6 blocks --------------
'
    Worksheets("Timing").Range("J23").Select
    ActiveCell.FormulaR1C1 = "=IF(ISBLANK(RC[1]),"""",R[-1]C[2])"
    Worksheets("Timing").Range("P23").Select
    Selection.AutoFill Destination:=Range("P23:P34"), Type:=xlFillValues
    
    Worksheets("Timing").Range("R22").Select
    ActiveCell.FormulaR1C1 = "=IF(ISBLANK(RC[-1]),"""",RC[-2]+(RC[-1]*7))"
    Worksheets("Timing").Range("R22").Select
    Selection.AutoFill Destination:=Range("R22:R34"), Type:=xlFillValues

    Worksheets("Timing").Range("C6").Select

'--- Allow screen refresh ---
    Application.ScreenUpdating = True

End Sub


