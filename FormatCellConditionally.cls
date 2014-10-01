VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormatCellConditionally"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private fmtcon As String
Private bg As Integer
Private r As Integer
Private g As Integer
Private b As Integer

Property Let FormatRange(ByRef value As Range)
    Set rng = value
End Property

Property Let FormatCondition(ByRef value() As String)
    fmtcon = value
End Property

Property Let BGColor(ByRef value As Integer)
    Select Case Application.Version
        Case "14.0"
            If value = 1 Then
                'pink
                r = 255
                g = 153
                b = 204
            ElseIf value = 2 Then
                'grey
                r = 150
                g = 150
                b = 150
            End If
        Case "11.0"
            If value = 1 Then
                'pink
                bg = 38
            ElseIf value = 2 Then
                'grey
                bg = 48
            End If
    End Select
End Property

Public Sub ConditionalFormatMe(rng As Range, ConditionNumber As Integer)
    Dim i As Integer
    
    'Set format condition number
    If ConditionNumber = 0 Then
        i = 1
    Else
        i = ConditionNumber
    End If
    
    'Set conditional format based on application version
    Select Case Application.Version
        Case "11.0"
            With rng
                .FormatConditions.Add Type:=xlExpression, Formula1:=fmtcon
                .FormatConditions(i).Interior.ColorIndex = bg
            End With
        Case "14.0"
            With rng
                .FormatConditions.Add Type:=xlExpression, Formula1:=fmtcon
                '.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
                With .FormatConditions(i).Interior
                    .PatternColorIndex = xlAutomatic
                    .Color = RGB(r, g, b)
                    .TintAndShade = 0
                End With
                .FormatConditions(1).StopIfTrue = True
            End With
    End Select
End Sub

Public Sub SetConditionalFormat_WbOpen()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rblLastRow As Integer
    Set ws = ActiveWorkbook.Worksheets("Dashboard")
    
    'Set conditional format bg color to pink
    bg = 38
    r = 255
    g = 153
    b = 204
        
        'Clear Existing Formatting
        ws.Range("C2:C14").FormatConditions.Delete
        ws.Range("F9:F14").FormatConditions.Delete
        ws.Range("C19:C25").FormatConditions.Delete
        'NOTE: SchB & SchC conditional formatting will be
        ''deleted immeadiately before new formatting is applied
        ''due to the addresses not being absolute.
        
        'Practice Name
        fmtcon = "=ISBLANK(C2)"
        Call ConditionalFormatMe(ws.Range("C2"), 1)
        
        'Address
        fmtcon = "=ISBLANK(C4)"
        Call ConditionalFormatMe(ws.Range("C4"), 1)
        
        'City
        fmtcon = "=ISBLANK(C5)"
        Call ConditionalFormatMe(ws.Range("C5"), 1)
        
        'State
        fmtcon = "=ISBLANK(C6)"
        Call ConditionalFormatMe(ws.Range("C6"), 1)
        
        'Zip
        fmtcon = "=ISBLANK(C7)"
        Call ConditionalFormatMe(ws.Range("C7"), 1)
        
        'Phone
        fmtcon = "=MID(C8, 4, 1)<>""."""
        Call ConditionalFormatMe(ws.Range("C8"), 1)
        fmtcon = "=MID(C8, 8, 1)<>""."""
        Call ConditionalFormatMe(ws.Range("C8"), 2)
        
        'Alt Phone
        fmtcon = "=MID(C9, 4, 1)<>""."""
        Call ConditionalFormatMe(ws.Range("C9"), 1)
        fmtcon = "=MID(C9, 8, 1)<>""."""
        Call ConditionalFormatMe(ws.Range("C9"), 2)
        
        'Fax
        fmtcon = "=MID(C10, 4, 1)<>""."""
        Call ConditionalFormatMe(ws.Range("C10"), 1)
        fmtcon = "=MID(C10, 8, 1)<>""."""
        Call ConditionalFormatMe(ws.Range("C10"), 2)
        
        'Email
        fmtcon = "=ISBLANK(C11)"
        Call ConditionalFormatMe(ws.Range("C11"), 1)
        
        'Client Type
        fmtcon = "=ISBLANK(C12)"
        Call ConditionalFormatMe(ws.Range("C12"), 1)
        
        'Specialty
        fmtcon = "=ISBLANK(C13)"
        Call ConditionalFormatMe(ws.Range("C13"), 1)
        
        'Internal/External
        fmtcon = "=ISBLANK(C14)"
        Call ConditionalFormatMe(ws.Range("C14"), 1)
        
        'Invoice Date
        fmtcon = "=ISBLANK(F9)"
        Call ConditionalFormatMe(ws.Range("F9"), 1)
        
        'Schedule A Start Date
        fmtcon = "=ISBLANK(F10)"
        Call ConditionalFormatMe(ws.Range("F10"), 1)
        
        'Schedule A End Date
        fmtcon = "=ISBLANK(F11)"
        Call ConditionalFormatMe(ws.Range("F11"), 1)
        
        'Schedule B Start Date
        fmtcon = "=ISBLANK(F12)"
        Call ConditionalFormatMe(ws.Range("F12"), 1)
        
        'Schedule B End Date
        fmtcon = "=ISBLANK(F13)"
        Call ConditionalFormatMe(ws.Range("F13"), 1)
        
        'Quote Valid Until
        fmtcon = "=ISBLANK(F14)"
        Call ConditionalFormatMe(ws.Range("F14"), 1)
        
        'SchA Package/Product
        fmtcon = "=ISBLANK(C19)"
        Call ConditionalFormatMe(ws.Range("C19"), 1)
        
        'SchA Payment Terms & Payment/Term
        'MUST UPDATE fmtcon if value list for
        ''SchA Fee Type drop down is reordered.
        fmtcon = "=IF(AND(ISBLANK(C21), 'CELL REFERENCES'!AH3 = 3),TRUE,FALSE)"
        Call ConditionalFormatMe(ws.Range("C21"), 1)
        fmtcon = "=IF(AND(ISBLANK(C21), 'CELL REFERENCES'!AH3 = 1),TRUE,FALSE)"
        Call ConditionalFormatMe(ws.Range("C21"), 2)
        
        
        
    'Set conditional format bg color to grey
    bg = 48
    r = 150
    g = 150
    b = 150
        
        'SchA Payment Terms & Payment/Term
        'MUST UPDATE fmtcon if value list for
        ''SchA Fee Type drop down is reordered.
        fmtcon = "='CELL REFERENCES'!$AH$3=2"
        Call ConditionalFormatMe(ws.Range("C21"), 3)
        fmtcon = "='CELL REFERENCES'!$AH$3=2"
        Call ConditionalFormatMe(ws.Range("C22"), 1)
    
    'Set conditional format bg color to pink
    bg = 38
    r = 255
    g = 153
    b = 204
    
        'SchA Next Payment Due
        fmtcon = "=ISBLANK(C24)"
        Call ConditionalFormatMe(ws.Range("C24"), 1)
        
        'Find end of schedule A list object
        ''Need this row number to reference the Schedule B header
        Set tbl = ws.ListObjects("scha")
        tblLastRow = tbl.ListRows.Count
        t = tbl.DataBodyRange.Rows(tblLastRow).Row
        
        'SchB Package/Product, Monthly Fee, Upfront Payment,
        ''Next Payment Due, Total Due
        For i = 2 To 7
            If i <> 3 Then
                cell = ws.Cells(t + i, 3).AddressLocal
                Debug.Print "SchB - " & cell
                ws.Range(cell).FormatConditions.Delete
                fmtcon = "=ISBLANK(" & cell & ")"
                Call ConditionalFormatMe(ws.Range(cell), 1)
            End If
        Next i
        
        'Find end of schedule B list object
        ''Need this row number to reference the Schedule C header
        Set tbl = ws.ListObjects("schb")
        tblLastRow = tbl.ListRows.Count
        t = tbl.DataBodyRange.Rows(tblLastRow).Row
        
        'SchC Package/Product, Monthly Fee, Upfront Payment,
        ''Next Payment Due, Total Due
        For i = 2 To 4
            If i <> 3 Or i <> 4 Then
                cell = ws.Cells(t + i, 3).AddressLocal
                Debug.Print "SchC - " & cell
                ws.Range(cell).FormatConditions.Delete
                fmtcon = "=ISBLANK(" & cell & ")"
                Call ConditionalFormatMe(ws.Range(cell), 1)
            End If
        Next i
End Sub

Public Sub SetConditionalFormat_WbClose()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rblLastRow As Integer
    Set ws = ActiveWorkbook.Worksheets("Dashboard")
    
    'Clear Existing Formatting
        ws.Range("C2:C14").FormatConditions.Delete
        ws.Range("F9:F14").FormatConditions.Delete
        ws.Range("C19:C25").FormatConditions.Delete
        
    'Find end of schedule A list object
    ''Need this row number to reference the Schedule B header
    Set tbl = ws.ListObjects("scha")
    tblLastRow = tbl.ListRows.Count
    t = tbl.DataBodyRange.Rows(tblLastRow).Row
    
    'SchB Package/Product, Monthly Fee, Upfront Payment,
    ''Next Payment Due, Total Due
    For i = 2 To 7
        If i <> 3 Then
            cell = ws.Cells(t + i, 3).AddressLocal
            ws.Range(cell).FormatConditions.Delete
        End If
    Next i
    
    'Find end of schedule B list object
    ''Need this row number to reference the Schedule C header
    Set tbl = ws.ListObjects("schb")
    tblLastRow = tbl.ListRows.Count
    t = tbl.DataBodyRange.Rows(tblLastRow).Row
    
    'SchC Package/Product, Monthly Fee, Upfront Payment,
    ''Next Payment Due, Total Due
    For i = 2 To 4
        If i <> 3 Or i <> 4 Then
            cell = ws.Cells(t + i, 3).AddressLocal
            ws.Range(cell).FormatConditions.Delete
        End If
    Next i
End Sub


