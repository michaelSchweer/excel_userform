
Private Sub cbSletSidste_Click()

Dim svar As Integer
svar = MsgBox("Er du sikker?", vbYesNo + vbCritical + vbDefaultButton2, "Slet sidst bogført?")

If svar = vbNo Then
    Exit Sub
End If

Sheets("Tastede fakturaer").Unprotect
lr = Sheets("Tastede fakturaer").Range("A" & Rows.Count).End(xlUp).Row
Rows(lr).EntireRow.Delete
Sheets("Tastede fakturaer").Protect
Call UserForm_Initialize

End Sub

Private Sub cbAfslut_Click()
FakturaForm.Hide
Sheets("Tastede fakturaer").Protect
End Sub



Private Sub tbFakturaNr_Exit(ByVal Cancel As MSForms.ReturnBoolean)

Dim ws As Worksheet
Set ws = Sheets("Tastede fakturaer")
lr = ws.Range("A" & Rows.Count).End(xlUp).Row

Dim fakturaNr As String
fakturaNr = tbFakturaNr.Value

For x = 7 To lr
    If Range("C" & x).Value = fakturaNr Then
        MsgBox "Faktura nr. eksisterer!" & vbNewLine & Range("A" & x).Value & vbNewLine & Range("B" & x).Value & vbNewLine & Range("D" & x).Value & " kr." & vbNewLine & Range("E" & x).Value, vbInformation, "Mulig duplikat"
    End If
Next x

tbFakturaNr.SetFocus

End Sub

Private Sub cmbKonto_Exit(ByVal Cancel As MSForms.ReturnBoolean)

Dim ws As Worksheet
Set ws = Sheets("Budget")
lr = ws.Range("A" & Rows.Count).End(xlUp).Row

Dim kontoNr As String
kontoNr = cmbKonto.Text

For x = 4 To lr
    If ws.Range("A" & x).Value = kontoNr Then
        tilbage = ws.Range("F" & x).Value - ws.Range("G" & x).Value
        If ws.Range("F" & x).Value <> 0 Then
            procentBrugt = ws.Range("G" & x).Value / ws.Range("F" & x).Value * 100
        Else
            procentBrugt = ""
        End If
        Controls("lblKontoInfo").Caption = ws.Range("B" & x).Value & vbNewLine & Format(procentBrugt, "0.00") & "% brugt" & vbNewLine & tilbage & " kr. tilbage"
    
    End If
Next x

If kontoNr = "" Then
Controls("lblKontoInfo").Caption = ""
End If

End Sub

Private Sub cmbKonto_Click()

Dim ws As Worksheet
Set ws = Sheets("Budget")
lr = ws.Range("A" & Rows.Count).End(xlUp).Row

Dim kontoNr As String
kontoNr = cmbKonto.Text

For x = 4 To lr
    If ws.Range("A" & x).Value = kontoNr Then
        tilbage = ws.Range("F" & x).Value - ws.Range("G" & x).Value
        If ws.Range("F" & x).Value <> 0 Then
            procentBrugt = ws.Range("G" & x).Value / ws.Range("F" & x).Value * 100
        Else
            procentBrugt = ""
        End If
        Controls("lblKontoInfo").Caption = ws.Range("B" & x).Value & vbNewLine & Format(procentBrugt, "0.00") & "% brugt" & vbNewLine & Format(tilbage, "#,##0") & " kr. tilbage"
    End If
Next x

End Sub



Public Sub UserForm_Initialize()
    
    Application.ScreenUpdating = False
    
    'userform cleares for input
    Dim ctl As MSForms.Control
    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "TextBox"
            ctl.Text = ""
            Case "CheckBox", "OptionButton", "ToggleButton"
            ctl.Value = False
            Case "ComboBox", "ListBox"
            ctl.ListIndex = -1
        End Select
    Next ctl
    Controls("lblKontoInfo").Caption = ""
    
    'Leverandør-boksen udfyldes med unikke værdier fra B-kolonnon på faktura-arket
    Dim vStr, eStr
    Dim dObj As Object
    On Error Resume Next
    Set dObj = CreateObject("Scripting.Dictionary")
            
    lr = Sheets("Tastede fakturaer").Range("B" & Rows.Count).End(xlUp).Row
    vStr = Sheets("Tastede fakturaer").Range("B7:B" & lr)
    
    With dObj
        .comparemode = 1
        For Each eStr In vStr
            If Not .exists(eStr) And eStr <> "" Then .Add eStr, Nothing
        Next
        If .Count Then
            FakturaForm.cmbLev.List = WorksheetFunction.Transpose(.keys)
        End If
    End With
    
    'Konto-boksen udfyldes med unikke værdier fra A-kolonnon på budget-arket
    Dim kvStr, keStr
    Dim kdObj As Object
    On Error Resume Next
    Set kdObj = CreateObject("Scripting.Dictionary")
            
    lr = Sheets("Budget").Range("A" & Rows.Count).End(xlUp).Row
    kvStr = Sheets("Budget").Range("A10:A" & lr)
    
    With kdObj
        .comparemode = 1
        For Each keStr In kvStr
            If Not .exists(keStr) And keStr <> "" Then .Add keStr, Nothing
        Next
        If .Count Then
            FakturaForm.cmbKonto.List = WorksheetFunction.Transpose(.keys)
        End If
    End With
    
    'Alfabetisk sortering af leverandører og konti
    Dim i As Long
    Dim j As Long
    Dim sTemp As String
    Dim sTemp2 As String
    Dim LbList As Variant
    
    LbList = Me.cmbLev.List
    
    'Bubble sort the array on the first value
    For i = LBound(LbList, 1) To UBound(LbList, 1) - 1
        For j = i + 1 To UBound(LbList, 1)
            If LbList(i, 0) > LbList(j, 0) Then
                'Swap the first value
                sTemp = LbList(i, 0)
                LbList(i, 0) = LbList(j, 0)
                LbList(j, 0) = sTemp
                
                'Swap the second value
                sTemp2 = LbList(i, 1)
                LbList(i, 1) = LbList(j, 1)
                LbList(j, 1) = sTemp2
            End If
        Next j
    Next i

    Me.cmbLev.Clear
    Me.cmbLev.List = LbList
    
    LbList = Me.cmbKonto.List
    
    For i = LBound(LbList, 1) To UBound(LbList, 1) - 1
        For j = i + 1 To UBound(LbList, 1)
            If LbList(i, 0) > LbList(j, 0) Then
                sTemp = LbList(i, 0)
                LbList(i, 0) = LbList(j, 0)
                LbList(j, 0) = sTemp
                
                sTemp2 = LbList(i, 1)
                LbList(i, 1) = LbList(j, 1)
                LbList(j, 1) = sTemp2
            End If
        Next j
    Next i

    Me.cmbKonto.Clear
    Me.cmbKonto.List = LbList
    
    'Sidst faktureret udfyldes
    lr = Sheets("Tastede fakturaer").Range("A" & Rows.Count).End(xlUp).Row
    Controls("lblSidst").Caption = Sheets("Tastede fakturaer").Range("A" & lr).Value & ", " & Sheets("Tastede fakturaer").Range("B" & lr).Value & ", " & Sheets("Tastede fakturaer").Range("C" & lr).Value & ", " & Format(Sheets("Tastede fakturaer").Range("D" & lr).Value, "#,##0.00 kr.") & ", " & Sheets("Tastede fakturaer").Range("E" & lr).Value
    
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 25
    Me.Left = Application.Left + 25
    
    tbFakturaNr.SetFocus
    
    Application.ScreenUpdating = True
    
End Sub


Private Sub btnBogfor_Click()

Dim ws As Worksheet
Set ws = Sheets("Tastede fakturaer")
lr = ws.Range("A" & Rows.Count).End(xlUp).Row
inputRow = lr + 1

Dim t As Control
For Each t In Me.Controls
    If TypeName(t) = "TextBox" Or TypeName(t) = "ComboBox" Then
        If t.Text = vbNullString Then
            MsgBox "Der er noget der mangler..."
            Exit Sub
        End If
    End If
Next t

Sheets("Tastede fakturaer").Unprotect

On Error GoTo fejl

Dim fakturaDato As Date
aar = Me.tbAar.Text
maaned = Me.tbMaaned.Text
dag = Me.tbDag.Text
fakturaDato = DateSerial(aar, maaned, dag)

fakturaBelob = CDbl(Me.tbBelob.Value)

'If Not IsNumeric(fakturaBelob) Then
'GoTo fejl2
'End If

ws.Range("A" & inputRow) = Me.cmbKonto.Text
ws.Range("B" & inputRow) = Me.cmbLev.Text
ws.Range("C" & inputRow) = Me.tbFakturaNr.Text
ws.Range("D" & inputRow) = fakturaBelob
ws.Range("E" & inputRow) = fakturaDato

Sheets("Tastede fakturaer").Protect

Call UserForm_Initialize
Exit Sub

fejl:
MsgBox ("INVALID INPUT!!!" & vbNewLine & vbNewLine & "Tjek dato og/eller beløb.")
Exit Sub


End Sub
