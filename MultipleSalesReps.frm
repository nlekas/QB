VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MultipleSalesReps 
   Caption         =   "Multiple Sales Reps"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   OleObjectBlob   =   "MultipleSalesReps.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MultipleSalesReps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ComboBox1_Change()
    If gEnableErrorHandling Then On Error GoTo whoops
    Const PROC_NAME = "ComboBox1_Change"
    Dim FindString1 As String
    Dim rng As Range
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("CELL REFERENCES")
    FindString1 = ComboBox1.value
    If Trim(FindString1) <> vbNullString Then
        UnprotectAll
            Application.ScreenUpdating = False
            With Sheets("CELL REFERENCES").Range("BG:BG")
                On Error Resume Next
                Set rng = .Find(What:=FindString1, _
                    After:=.Cells(.Cells.Count), _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False)
            End With
            If Not rng Is Nothing Then
                TextBox1.value = ws.Cells(rng.Row, rng.Column + 2).value
                TextBox2.value = ws.Cells(rng.Row, rng.Column + 1).value
            End If
        ProtectAll
    End If
    Application.ScreenUpdating = True
    Exit Sub
whoops:
    ProtectAll
    MsgBox "There was a problem. Note the error information below and the steps to re-create the problem." _
    & vbNewLine & vbNewLine _
    & Err.Number & " - " & Err.Description & " in " & PROC_NAME, vbCritical, "Whoops!"
End Sub

Private Sub ComboBox2_Change()
    If gEnableErrorHandling Then On Error GoTo whoops
    Const PROC_NAME = "ComboBox2_Change"
    Dim FindString2 As String
    Dim rng As Range
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("CELL REFERENCES")
    FindString2 = ComboBox2.value
    If Trim(FindString2) <> vbNullString Then
        UnprotectAll
            Application.ScreenUpdating = False
            With Sheets("CELL REFERENCES").Range("BG:BG")
                On Error Resume Next
                Set rng = .Find(What:=FindString2, _
                    After:=.Cells(.Cells.Count), _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False)
            End With
            If Not rng Is Nothing Then
                TextBox4.value = ws.Cells(rng.Row, rng.Column + 2).value
                TextBox3.value = ws.Cells(rng.Row, rng.Column + 1).value
            End If
        ProtectAll
    End If
    Application.ScreenUpdating = True
    Exit Sub
whoops:
    ProtectAll
    MsgBox "There was a problem. Note the error information below and the steps to re-create the problem." _
    & vbNewLine & vbNewLine _
    & Err.Number & " - " & Err.Description & " in " & PROC_NAME, vbCritical, "Whoops!"
End Sub

Private Sub ComboBox3_Change()
    If gEnableErrorHandling Then On Error GoTo whoops
    Const PROC_NAME = "ComboBox3_Change"        'uses AcctMgrName_Change
    Dim FindString3 As String
    Dim rng As Range
    Dim ws As Worksheet
    Dim sh As Sheets
    Set ws = ActiveWorkbook.Worksheets("CELL REFERENCES")
        FindString3 = ComboBox3.value
    
    If Trim(FindString3) <> vbNullString Then
        UnprotectAll
            Application.ScreenUpdating = False
            With Sheets("CELL REFERENCES").Range("BG:BG")
                On Error Resume Next
                Set rng = .Find(What:=FindString3, _
                    After:=.Cells(.Cells.Count), _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False)
            End With
            If Not rng Is Nothing Then
                TextBox6.value = ws.Cells(rng.Row, rng.Column + 2).value
                TextBox5.value = ws.Cells(rng.Row, rng.Column + 1).value
            End If
        ProtectAll
    End If
    Application.ScreenUpdating = True
    Exit Sub
whoops:
    ProtectAll
    MsgBox "There was a problem. Note the error information below and the steps to re-create the problem." _
    & vbNewLine & vbNewLine _
    & Err.Number & " - " & Err.Description & " in " & PROC_NAME, vbCritical, "Whoops!"
End Sub
    
Private Sub CommandButton1_Click()
    If gEnableErrorHandling Then On Error GoTo whoops
    Const PROC_NAME = "CommandButton1_Click"
    UnprotectAll
    Me.Hide
    ProtectAll
    Exit Sub
whoops:
    ProtectAll
    MsgBox "There was a problem. Note the error information below and the steps to re-create the problem." _
    & vbNewLine & vbNewLine _
    & Err.Number & " - " & Err.Description & " in " & PROC_NAME, vbCritical, "Whoops!"
End Sub

Private Sub CommandButton2_Click()
        Unload Me
        
    Exit Sub
whoops:
    ProtectAll
    MsgBox "There was a problem. Note the error information below and the steps to re-create the problem." _
    & vbNewLine & vbNewLine _
    & Err.Number & " - " & Err.Description & " in " & PROC_NAME, vbCritical, "Whoops!"
End Sub



' TEST EDIT


Private Sub CommandButton3_Click()
    If gEnableErrorHandling Then On Error GoTo whoops
    Const PROC_NAME = "CommandButton3_Click"
    Dim cmbo1 As Boolean
    Dim cmbo2 As Boolean
    Dim cmbo3 As Boolean
    Dim y As Double
    Dim ws As Worksheet
    Dim varmsg As Variant
    Set ws = ActiveWorkbook.Worksheets("Dashboard")
    
    'Check if form is properly filled out
    If ComboBox1.value = vbNullString Or _
    TextBox1.value = vbNullString Or _
    TextBox2.value = vbNullString Then
        cmbo1 = False
    Else
        cmbo1 = True
        y = y + 1                           'THINK THAT THIS IS WHERE THAT BLANK SELECTION ERROR IS COMING FROM
    End If
    
    If ComboBox2.value = vbNullString Or _
    TextBox4.value = vbNullString Or _
    TextBox3.value = vbNullString Then
        cmbo2 = False
    Else
        cmbo2 = True
        y = y + 1
    End If
    
    If ComboBox3.value = vbNullString Or _
    TextBox5.value = vbNullString Or _
    TextBox6.value = vbNullString Then
        cmbo3 = False
    Else
        cmbo3 = True
        y = y + 1
    End If
    
    Select Case y
        Case Is >= 2
            'If more than one sales rep is present then
            'lock dashboard cells and hide form
            UnprotectAll
                'ws.Range("F2:F2").value = "MULTIPLE MGRS"
                Set rng = ws.Range("F2:F4")
                rng.Locked = True
                With ActiveWorkbook.Worksheets("CELL REFERENCES")
                    'F5 = name | F4 = email | F3 = ext
                    .Range("A17").value = ComboBox1.value
                    .Range("A18").value = TextBox1.value
                    .Range("A19").value = TextBox2.value
                    .Range("A21").value = ComboBox2.value
                    .Range("A22").value = TextBox4.value
                    .Range("A23").value = TextBox3.value
                    .Range("A25").value = ComboBox3.value
                    .Range("A26").value = TextBox5.value
                    .Range("A27").value = TextBox6.value
                    '.Range("A29").value = name
                    '.Range("A30").value = email
                    '.Range("A31").value = ext
                End With
                Me.Hide
            ProtectAll
        Case Is < 2
            'If no values are properly filled out, then warn and exit.
            varmsg = MsgBox("You must fill in at least 2 sales reps." & _
            vbNewLine & vbNewLine & _
            "If there is only 1 sales rep, use the fields on the Dashboard.", vbRetryCancel, "Whoops!")
            If varmsg = vbCancel Then
                CommandButton2_Click
                Me.Hide
            End If
    End Select
    Exit Sub
whoops:
    ProtectAll
    MsgBox "There was a problem. Note the error information below and the steps to re-create the problem." _
    & vbNewLine & vbNewLine _
    & Err.Number & " - " & Err.Description & " in " & PROC_NAME, vbCritical, "Whoops!"
    
End Sub

Private Sub ComboBoxAcctMgr_Click()
' Label3Click Macro
Const PROC_NAME = "ComboBoxAcctMgr"
End Sub
Private Sub ClearSchARep_Click()

End Sub

Private Sub CommandButton4_Click()
UnprotectAll
Application.ScreenUpdating = False
        Dim ws As Worksheet
        Set ws = ActiveWorkbook.Worksheets("Dashboard")
            With ActiveWorkbook.Worksheets("CELL REFERENCES")
            .Range("A17:A19").value = vbNullString
            
        End With
        ComboBox1.value = vbNullString
        TextBox1.value = vbNullString
        TextBox2.value = vbNullString
       
    ProtectAll
End Sub

Private Sub CommandButton5_Click()
UnprotectAll
Application.ScreenUpdating = False
        Dim ws As Worksheet
        Set ws = ActiveWorkbook.Worksheets("Dashboard")
        With ActiveWorkbook.Worksheets("CELL REFERENCES")
            .Range("A21:A23").value = vbNullString
        
        End With
        ComboBox2.value = vbNullString
        TextBox3.value = vbNullString
        TextBox4.value = vbNullString
       
    ProtectAll
End Sub

Private Sub CommandButton6_Click()
UnprotectAll
Application.ScreenUpdating = False
        Dim ws As Worksheet
        Set ws = ActiveWorkbook.Worksheets("Dashboard")
        With ActiveWorkbook.Worksheets("CELL REFERENCES")
            .Range("A25:A27").value = vbNullString
            
        End With
        ComboBox3.value = vbNullString
        TextBox5.value = vbNullString
        TextBox6.value = vbNullString
        
    ProtectAll
End Sub

'Private Sub UserForm_Click()
'CommandButton4.Caption = "ClearSchARep"
'End Sub


