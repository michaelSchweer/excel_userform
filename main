'funktion der checker om et ark med et givent navn eksisterer
Function sheetExists(sheetToFind As String) As Boolean
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function


'sub der kopierer rækker til Rapport-arket ud fra valg af år og måned på Forsiden.
Sub GetMonth()

'opdatering til skærmen slåes fra og faktura-arkets beskyttelse fjernes
Application.ScreenUpdating = False
Sheets("Tastede fakturaer").Unprotect

'hvis Rapport-arket eksisterer i forvejen slettes det
If sheetExists("Rapport") = True Then
    Application.DisplayAlerts = False
    Sheets("Rapport").Delete
    Application.DisplayAlerts = True
End If

'variablen aar defineres og det valgte år læses fra celle F2
Dim aar As Integer
aar = Range("F2").Value

'variablen maaned defineres og den valgte måned læses fra cellen F1 samt konverteres til et heltal (integer)
Dim maaned As Integer
If Range("F1").Value = "Januar" Then maaned = 1
If Range("F1").Value = "Februar" Then maaned = 2
If Range("F1").Value = "Marts" Then maaned = 3
If Range("F1").Value = "April" Then maaned = 4
If Range("F1").Value = "Maj" Then maaned = 5
If Range("F1").Value = "Juni" Then maaned = 6
If Range("F1").Value = "Juli" Then maaned = 7
If Range("F1").Value = "August" Then maaned = 8
If Range("F1").Value = "September" Then maaned = 9
If Range("F1").Value = "Oktober" Then maaned = 10
If Range("F1").Value = "November" Then maaned = 11
If Range("F1").Value = "December" Then maaned = 12

'dato-variablen defineres og sættes til den 1. i den valgte måned/år da Excel skal bruge et dato-objekt (dag/måned/år) til at
'sammenligne med i loopet nedenfor
Dim dato As Date
dato = DateSerial(aar, maaned, 1)

'vi definerer et par hjælpe-variable til brug i loopet nedenfor
Dim lastRow As Long
Dim currow As Long
Dim NextDest As Long

'vi definerer en 'ark'-variabel og sætter den til at pege på faktura-arket
Dim ws As Worksheet
Set ws = Sheets("Tastede fakturaer")

'finder sidste række med data (i E-kolonnen). bruges i loopet til at afgrænse området der itereres over
lastRow = ws.Range("E" & Rows.Count).End(xlUp).Row

'loop som:
'1. starter med række 7 (første række med data) og fortsætter til og med sidste række med data (lastRow)
'2. tester for om E-cellen er af datatypen Date. hvis ikke springes der til Else-delen af loopet og A-cellen farves rød
'3. måned og år læses ud fra datoen og sammenlignes med maaned og aar variablerne. hvis de matcher fortsætter loopet,
'hvis ikke går vi videre til næste række (Next CurRow)
'4. cellerne A til og med H i den nuværende række kopieres
'5. nederste celle med data i AA-kolonnen findes og der lægges en til (+ 1) for at finde den celle hvor det kopierede skal sættes ind
'6. det kopierede sættes ind
For currow = 7 To lastRow ' 1.
    If IsDate(ws.Range("E" & currow).Value) = True Then ' 2.
        If Month(ws.Range("E" & currow).Value) = Month(dato) And Year(ws.Range("E" & currow).Value) = Year(dato) Then ' 3.
            ws.Range("A" & currow & ":H" & currow).Copy ' 4.
            NextDest = ws.Range("AA" & Rows.Count).End(xlUp).Row + 1 ' 5.
            ws.Range("AA" & NextDest).PasteSpecial xlPasteValues ' 6.
        End If
    Else
        ws.Cells(currow, 1).Interior.Color = RGB(255, 0, 0)
    End If
Next currow

'rapport ark oprettes og formateres
Sheets.Add
ActiveSheet.Name = "Rapport"
Cells.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

'vi tilføjer nogle overskrifter
Range("A6").Value = "PSP"
Range("B6").Value = "Leverandør"
Range("C6").Value = "Faktura nr."
Range("D6").Value = "Beløb"
Range("E6").Value = "Dato"
Range("F6").Value = "Konteret på"
Range("H6").Value = "% Brugt"
Rows(6).Font.Bold = True

'vi tegner knapper, giver dem navne og tilkobler makroer (på Rapport-arket)
ActiveSheet.Buttons.Add(10, 16, 64, 32).Select
Selection.OnAction = "DeleteRapportSheet"
Selection.Characters.Text = "Tilbage til start"
ActiveSheet.Buttons.Add(100, 16, 64, 32).Select
Selection.OnAction = "DrawPie"
Selection.Characters.Text = "Pie!"

'de filtrerede poster fra loopet flyttes til rapport-arket
Set ws = Sheets("Tastede fakturaer")
lr = ws.Range("AA" & Rows.Count).End(xlUp).Row
ws.Range("AA8:AH" & lr).Cut Destination:=Sheets("Rapport").Range("A7")

'vi får de kopierede data til at se lidt pænere ud
Columns.AutoFit
Cells.Select
Selection.RowHeight = 16
Selection.HorizontalAlignment = xlLeft
Selection.VerticalAlignment = xlBottom
Range("D7:D5000").NumberFormat = "#,##0.00"
Range("D7:D5000").HorizontalAlignment = xlRight
Range("E7:E5000").NumberFormat = "dd-mm-yy"
Range("A7").Select

'ark låses igen og skærm-opdatering slåes til igen.
ActiveSheet.Protect
With Sheets("Tastede fakturaer")
    .Protect
    .EnableSelection = xlUnlockedCells
End With
Application.ScreenUpdating = True

End Sub

'sub der tegner lagkagediagram
Sub DrawPie()
    
    Application.ScreenUpdating = False
    
    If sheetExists("Pie") = True Then
    Application.DisplayAlerts = False
    Sheets("Pie").Delete
    Application.DisplayAlerts = True
    End If
    
    ' Summer-start
    Dim ws As Worksheet
    Set ws = Sheets("Rapport")
    ws.Unprotect
    
    lr = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Range("B6:B" & lr).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("AA6"), Unique:=True
    
    llr = ws.Cells(Rows.Count, 27).End(xlUp).Row
    
    For y = 7 To llr
        amt = 0
        For x = 7 To lr
            If ws.Cells(x, 2) = ws.Cells(y, 27) Then
                amt = amt + ws.Cells(x, 4)
            End If
        Next x
        ws.Cells(y, 28) = amt
    Next y
    ' Summer-slut

    lr = ws.Range("AA" & Rows.Count).End(xlUp).Row
    ws.Range("AA6:AB" & lr).Copy
    
    Sheets.Add
    ActiveSheet.Name = "Pie"
    Set ws = Sheets("Pie")
    
    Cells.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ws.Range("A5").PasteSpecial
    Application.CutCopyMode = False
    
    Application.DisplayAlerts = False
    Sheets("Rapport").Delete
    Application.DisplayAlerts = True
    
    Dim myrange As String
    lr = ws.Range("A" & Rows.Count).End(xlUp).Row
    myrange = "A5:B" & lr
    
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("B5:B" & lr), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange Range(myrange)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ws.Range("A:B").Columns.AutoFit
    ws.Range("B6:B5000").NumberFormat = "#,##0.00"
    ws.Range("B6:B5000").HorizontalAlignment = xlRight
    
    ws.Range("A1").Select
        
    ws.Shapes.AddChart2(255, xlDoughnut).Select
    ActiveChart.SetSourceData Source:=Range("Pie!$A$2:$B$" & lr)
    ActiveChart.ChartTitle.Select
    Selection.Delete
    ActiveChart.Legend.Select
    Selection.Delete
    ActiveChart.ChartGroups(1).DoughnutHoleSize = 50
    
    With ActiveChart.Parent
        .Height = 520
        .Width = 520
        .Top = 0
        .Left = 350
    End With
    
    ActiveSheet.Buttons.Add(10, 16, 64, 32).Select
    Selection.OnAction = "DeletePieSheet"
    Selection.Characters.Text = "Tilbage til start"
    
    Range("A111").Select
    Application.DisplayFullScreen = True
    Application.ScreenUpdating = True
    ws.Protect

End Sub

'sub der sletter Rapport-arket og vender tilbage til Forsiden
Sub DeleteRapportSheet()
Application.DisplayAlerts = False
Sheets("Rapport").Delete
Application.DisplayAlerts = True
Sheets("Budget").Select
End Sub

'sub der sletter Pie-arket, stopper fullscreen og vender tilbage til Forsiden
Sub DeletePieSheet()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Sheets("Pie").Delete
Application.DisplayAlerts = True
Sheets("Budget").Select
Application.DisplayFullScreen = False
Application.ScreenUpdating = True
End Sub

'sub der kalder UserForm_Initialize og viser indtast faktura formularen
Public Sub Knap4_Klik()
Call FakturaForm.UserForm_Initialize
FakturaForm.Show
End Sub
