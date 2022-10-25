Option Explicit


Private Sub CheckBox1_Click()

If CheckBox1.Value = True Then
    CommandButton4.BackColor = RGB(62, 159, 69)
Else
    CommandButton4.BackColor = RGB(255, 0, 0)
End If

End Sub



Private Sub ComboBox1_AfterUpdate()

Dim c As Range

Me.ListBox1.Value = ""
Me.ListBox1.Clear
Me.ListBox2.Clear
For Each c In ThisWorkbook.Sheets("Support Sheet").Range("F1:F20000")
    If c.Value = Me.ComboBox1.Value Then
        Me.ListBox1.AddItem c.Offset(0, 1).Value
    End If
Next

If Me.ComboBox1.Value = "Pensions" Then
    Me.ComboBox2.Enabled = True
    Me.ComboBox2.Value = ""
    Me.ComboBox2.List = Array("Aviva", "SLOC", "Generali", "Unum", "Hodgelife")
ElseIf Me.ComboBox1.Value = "Postal Share Dealing" Or Me.ComboBox1.Value = "Central Reconciliations" Or Me.ComboBox1.Value = "EQ Global" Or Me.ComboBox1.Value = "Customer Experience Centre" Then
    Me.ComboBox2.Enabled = False
    Me.ComboBox2.Value = "Investment Services"
ElseIf Me.ComboBox1.Value = "Employee Services" Then
    Me.ComboBox2.Enabled = False
    Me.ComboBox2.Value = "Employee Services"
ElseIf Me.ComboBox1.Value = "RDIR" Then
    Me.ComboBox2.Enabled = False
    Me.ComboBox2.Value = "RDIR"
ElseIf Me.ComboBox1.Value = "ProSearch" Then
    Me.ComboBox2.Enabled = False
    Me.ComboBox2.Value = "ProSearch"
Else
    Me.ComboBox2.Enabled = False
    Me.ComboBox2.Value = "Customer Processing"
End If


End Sub


Private Sub CommandButton3_Click()

    Dim myUserName As String
    Dim fso As Object
    Dim ts As Object
    Dim FilePath As String
    Dim i As Long
    Dim TransCount As Long
    Dim myComments As String
    Dim TempComment As String
    Dim j As Integer
    Dim k As Integer
    Dim Starttime As Date
    Dim Endtime As Date
    Dim Duration As Date
    Dim LenComment As Long
    Dim ProdValue As String
    Dim ProdTime As Date
    Dim drlist As Long
    
    On Error GoTo MyError
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    myUserName = Environ("username")
    
    TempComment = TextBox2.Text
    
    If Me.TextBox2.Value = "" Then
        myComments = "No Comments"
    Else
        myComments = TempComment
    End If
    
    FilePath = FileLocation & LCase(Environ("username")) & ".eqi"
    
    If Me.ComboBox3.Value = "" Then
        MsgBox "You haven't selected any activity", vbCritical, "Alert"
        Exit Sub
    End If
    
    drlist = Me.ComboBox3.ListIndex
    If drlist = -1 And Me.ComboBox3.Value <> "" Then
       Me.ComboBox3.SetFocus
       MsgBox "Please select correct option.", vbCritical, "Alert"
       Exit Sub
    End If
    
    Set ts = fso.OpenTextFile(FilePath, 8)
     
    If Me.CommandButton3.BackColor = RGB(255, 0, 0) Then
        Me.CommandButton3.BackColor = RGB(62, 159, 69)
        Me.ComboBox1.BackColor = RGB(229, 229, 229)
        Me.ComboBox2.BackColor = RGB(229, 229, 229)
        Me.ComboBox3.BackColor = RGB(229, 229, 229)
        Me.ListBox1.BackColor = RGB(229, 229, 229)
        Me.ListBox2.BackColor = RGB(229, 229, 229)
        Me.TextBox1.BackColor = RGB(229, 229, 229)
        Me.ComboBox1.Locked = True
        Me.ComboBox2.Locked = True
        Me.ComboBox3.Locked = True
        Me.ListBox1.Locked = True
        Me.ListBox2.Locked = True
        Me.TextBox1.Locked = True
        'Me.RightArrow_lbl.Locked = True
        'Me.LeftArrow_lbl.Locked = True
        Me.CommandButton4.Locked = True
        Me.CommandButton5.Locked = True
        Me.CommandButton6.Locked = True
        Me.CommandButton3.Caption = "Resume"
        Me.TextBox2.Value = ""
       
        Me.Label19.Caption = VBA.Format(Now, "dd/mm/yyyy hh:mm:ss")
        
        If Me.ComboBox3.Value = "Cross Training" Then
            Me.ComboBox4.Enabled = True
            Me.ComboBox4.Visible = True
        End If
    Else
    
        If Me.ComboBox3.Value = "Cross Training" Then
            If Me.ComboBox4.Value = "" Or Me.ComboBox4.ListIndex = -1 Then
                MsgBox "Please select cross training team.", vbInformation, "Alert"
                Exit Sub
            End If
        End If
        
        Me.CommandButton3.BackColor = RGB(255, 0, 0)
        Me.ComboBox1.BackColor = RGB(255, 255, 255)
        Me.ComboBox2.BackColor = RGB(255, 255, 255)
        Me.ComboBox3.BackColor = RGB(255, 255, 255)
        Me.ListBox1.BackColor = RGB(255, 255, 255)
        Me.ListBox2.BackColor = RGB(255, 255, 255)
        Me.TextBox1.BackColor = RGB(255, 255, 255)
        Me.ComboBox1.Locked = False
        Me.ComboBox2.Locked = False
        Me.ComboBox3.Locked = False
        Me.ListBox1.Locked = False
        Me.ListBox2.Locked = False
        Me.TextBox1.Locked = False
        'Me.RightArrow_lbl.Locked = False
        'Me.LeftArrow_lbl.Locked = False
        Me.CommandButton4.Locked = False
        Me.CommandButton5.Locked = False
        Me.CommandButton6.Locked = False
        Me.CommandButton3.Caption = "Pause"
        Me.TextBox2.Value = ""
        
        Endtime = VBA.Format(Now, "dd/mm/yyyy hh:mm:ss")
        
        ts.Write Me.Label19.Caption & "|" & Endtime & "|" & Environ("username") & "|" & Me.Label22.Caption & "|Non Production|" & Me.ComboBox3.Value & "|" & Me.ComboBox4.Value & "|||" & myComments & vbNewLine
        
        Me.Label21.Caption = Format((Now - TimeValue(Me.Label19.Caption)) + TimeValue(Me.Label21.Caption), "hh:mm:ss")
        
        Me.ComboBox3.Value = ""
        Me.Label19.Caption = ""
        Me.ComboBox4.Value = ""
        Me.ComboBox4.Enabled = False
        Me.ComboBox4.Visible = False
    
 End If
    
    ts.Close
    
    Set ts = Nothing
    Set fso = Nothing
    
    Exit Sub
    
MyError:
    
    MsgBox Err.Description, vbCritical, "Alert"
    
End Sub

Private Sub CommandButton4_Click()

    Dim myUserName As String
    Dim fso As Object
    Dim ts As Object
    Dim FilePath As String
    Dim i As Long
    Dim TransCount As Long
    Dim myComments As String
    Dim j As Integer
    Dim k As Integer
    Dim TransactionName As String
    Dim TempComment As String
    Dim LenComment As Long
    
    On Error GoTo MyError
    
    myUserName = Environ("username")
    
    i = Me.ListBox2.ListIndex
    
    If Me.ComboBox2.Value = "" Then
        MsgBox "Please capture all required inforamtion.", vbCritical, "Alert"
        Exit Sub
    End If
    
    If Me.TextBox1.Value <> "" Then
        If Me.TextBox1.Value < 1 Or Me.TextBox1.Value > 1000 Or IsNumeric(Me.TextBox1.Value) = False Then
            MsgBox "Invalid number", vbCritical, "Alert"
            Me.TextBox1.Value = ""
            Me.TextBox1.SetFocus
            Exit Sub
        End If
    End If
    
    TempComment = TextBox2.Text
    TempComment = Replace(TempComment, vbCr, "")
    TempComment = Replace(TempComment, vbLf, "")
    TempComment = Replace(TempComment, "|", "*")
    
    If Me.TextBox2.Value = "" Then
        myComments = "No Comments"
    Else
        myComments = Trim(TempComment)
    End If
    
    If Me.TextBox1.Value = "" Then
        TransCount = 1
    Else
        TransCount = Me.TextBox1.Value
    End If
      
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Rem Create do not open folder if it is not already available
    
    If Not fso.FolderExists(FileLocation) = True Then
        fso.CreateFolder (FileLocation)
    End If
    
    FilePath = FileLocation & LCase(Environ("username")) & ".eqi"
    
    'Create backup file if it is not already available
    
    If Not fso.FileExists(FilePath) = True Then
        Set ts = fso.CreateTextFile(FilePath)
    Else
        Set ts = fso.OpenTextFile(FilePath, 8)
    End If
    
    'Write login time for the users
    
    If i = -1 Then MsgBox "You haven't selected any transaction; Please check", vbCritical, "Alert": Exit Sub
    
    If CheckBox1.Value = True Then
        TransactionName = Me.ListBox2.Column(0, ListBox2.ListIndex) & "_QC"
    Else
        TransactionName = Me.ListBox2.Column(0, ListBox2.ListIndex)
    End If
    
    k = Me.ListBox3.ListCount - 1
    
    If CheckBox1.Value = True Then
        ts.Write Me.Label20.Caption & "|" & VBA.Format(Now, "dd/mm/yyyy hh:mm:ss") & "|" & Environ("username") & "|" & Me.Label22.Caption & "|" & Me.ListBox2.Column(0, i) & "_QC|" & Me.ComboBox1.Value & "|" & Me.ComboBox2.Value & "|" & TransCount & "|" & Me.Label21.Caption & "|" & myComments & vbNewLine
    Else
        ts.Write Me.Label20.Caption & "|" & VBA.Format(Now, "dd/mm/yyyy hh:mm:ss") & "|" & Environ("username") & "|" & Me.Label22.Caption & "|" & Me.ListBox2.Column(0, i) & "|" & Me.ComboBox1.Value & "|" & Me.ComboBox2.Value & "|" & TransCount & "|" & Me.Label21.Caption & "|" & myComments & vbNewLine
    End If
    
    Me.Label20.Caption = VBA.Format(Now, "dd/mm/yyyy hh:mm:ss")
    Me.Label21.Caption = "00:00:00"
    Me.TextBox1.Value = ""
    Me.TextBox2.Value = ""
    
    For j = 1 To k
        If Me.ListBox3.List(j, 0) = TransactionName Then
            Me.ListBox3.List(j, 1) = Me.ListBox3.List(j, 1) + TransCount
            Exit Sub
        End If
    Next j
    
    ListBox3.AddItem
    Me.ListBox3.List(k + 1, 0) = TransactionName
    Me.ListBox3.List(k + 1, 1) = TransCount
    
    ts.Close
    
    Set ts = Nothing
    Set fso = Nothing
    
    Exit Sub
    
MyError:
    
    MsgBox Err.Description, vbCritical, "Alert"

End Sub

Private Sub CommandButton5_Click()

    Dim myUserName As String
    Dim fso As Object
    Dim ts As Object
    Dim FilePath As String
    Dim i As Long
    Dim TransCount As Long
    Dim myComments As String
    Dim j As Integer
    Dim k As Integer
    Dim TransactionName As String
    Dim TempComment As String
    Dim LenComment As Long
    
    On Error GoTo MyError
    
    myUserName = Environ("username")
    
    i = Me.ListBox2.ListIndex
        
    If Me.ComboBox2.Value = "" Then
        MsgBox "Please capture all required inforamtion.", vbCritical, "Alert"
        Exit Sub
    End If
    
    If Me.TextBox1.Value <> "" Then
        MsgBox "You can't enter values in Diarise", vbCritical, "Alert"
        Exit Sub
    End If
    
    TransCount = 1
    
    TempComment = TextBox2.Text
    
    TempComment = Replace(TempComment, vbCr, "")
    TempComment = Replace(TempComment, vbLf, "")
    TempComment = Replace(TempComment, "|", "*")

    If Me.TextBox2.Value = "" Then
        myComments = "No Comments"
    Else
        myComments = Trim(TempComment)
    End If
        
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Create do not open folder if it is not already available
    
    If Not fso.FolderExists(FileLocation) = True Then
        fso.CreateFolder (FileLocation)
    End If
    
    FilePath = FileLocation & LCase(Environ("username")) & ".eqi"
    
    'Create backup file if it is not already available
    
    If Not fso.FileExists(FilePath) = True Then
        Set ts = fso.CreateTextFile(FilePath)
    Else
        Set ts = fso.OpenTextFile(FilePath, 8)
    End If
    
    'Write login time for the users
    
    If i = -1 Then MsgBox "You haven't selected any transaction; Please check", vbCritical, "Alert": Exit Sub
    
    'TransactionName = Me.ComboBox2.Column(1) & "-" & Me.ListBox2.Column(0, ListBox2.ListIndex)
    
    If Me.CheckBox1.Value = True Then
    ts.Write Me.Label20.Caption & "|" & VBA.Format(Now, "dd/mm/yyyy hh:mm:ss") & "|" & Environ("username") & "|" & Me.Label22.Caption & "|" & Me.ListBox2.Column(0, i) & "_QC|" & Me.ComboBox1.Value & "|" & Me.ComboBox2.Value & "|" & TransCount & "|" & Me.Label21.Caption & "|" & myComments & "|Diarise" & vbNewLine
    Else
    ts.Write Me.Label20.Caption & "|" & VBA.Format(Now, "dd/mm/yyyy hh:mm:ss") & "|" & Environ("username") & "|" & Me.Label22.Caption & "|" & Me.ListBox2.Column(0, i) & "|" & Me.ComboBox1.Value & "|" & Me.ComboBox2.Value & "|" & TransCount & "|" & Me.Label21.Caption & "|" & myComments & "|Diarise" & vbNewLine
    End If
    Me.Label20.Caption = VBA.Format(Now, "dd/mm/yyyy hh:mm:ss")
    Me.Label21.Caption = "00:00:00"
    Me.TextBox1.Value = ""
    Me.TextBox2.Value = ""
    
    ts.Close
    
    Set ts = Nothing
    Set fso = Nothing
     
    Exit Sub
    
MyError:
    
    MsgBox Err.Description, vbCritical, "Alert"
    
End Sub

Private Sub CommandButton6_Click()
    
    Dim FilePath As String
    Dim FileNum As Integer
    Dim Arr() As Variant
    Dim r As Long
    Dim Data As String
    Dim CustomDate As Date
    Dim fso As Object
    Dim ts As Object
    Dim i As Long
    Dim MsgResult As VbMsgBoxResult
    
    Dim ptxt As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    MsgResult = MsgBox("Do you want to Logout for the day?", vbCritical + vbYesNo, "Alert")
    
    If MsgResult = vbNo Then
        'Unload Me
        Exit Sub
    End If
    
    FilePath = FileLocation & LCase(Environ("username")) & ".eqi"

    FileNum = FreeFile
    r = 1
    
    Open FilePath For Input As #FileNum
    Do While Not EOF(FileNum)
        Line Input #FileNum, Data
        ReDim Preserve Arr(1 To r)
        Arr(r) = Data
        r = r + 1
    Loop
    Close #FileNum
    
    ptxt = Left(Arr(UBound(Arr)), 39)
    CustomDate = (Now - CDate(Left(ptxt, 19)))
    
    Set ts = fso.OpenTextFile(FilePath, 8)

    ts.Write Me.Label20.Caption & "|" & VBA.Format(Now, "dd/mm/yyyy hh:mm:ss") & "|" & Environ("username") & "|" & Me.Label22.Caption & "|Logout||||" & Me.Label21.Caption & vbNewLine
    
    ts.Close
    
    Set ts = Nothing
    Set fso = Nothing
    
    Unload Me
    
End Sub

Private Sub CommandButton7_Click()
   
   UserForm2.Show vbModeless
   
End Sub


Private Sub CommandButton8_Click()
    
    Dim FileNum As Long
    Dim r, s, t As Long
    Dim Arr()
    Dim MyList1()
    Dim MyList2()
    Dim Data As String
    Dim FilePath As String
    Dim SplitArr
    Dim Cmt As String
    Dim ArrCount As Long
    
    On Error GoTo MyError
    
    FilePath = FileLocation & LCase(Environ("username")) & ".eqi"
    
    FileNum = FreeFile
    
    Open FilePath For Input As #FileNum
    Do While Not EOF(FileNum)
        Line Input #FileNum, Data
        If Left(Data, 10) <> "welcome" Then
            If CDate(Left(Data, 10)) = Date Then
            r = r + 1
            ReDim Preserve Arr(1 To r)
            Arr(r) = Data
            End If
        End If
    Loop
    Close #FileNum
    
    '----------------------Production Count------------------------------------------------------------------------------
    
    For r = LBound(Arr) To UBound(Arr)
        SplitArr = Split(Arr(r), "|")
        ArrCount = UBound(SplitArr)
      If ArrCount < 10 Then
            If Trim(SplitArr(4)) <> "Login" And Trim(SplitArr(4)) <> "Downtime" And Trim(SplitArr(4)) <> "Non Production" And Trim(SplitArr(4)) <> "Logout" And ArrCount > 6 Then
                s = s + 1
                ReDim Preserve MyList1(1 To s)
                ReDim Preserve MyList2(1 To s)
                MyList1(s) = Trim(SplitArr(4))
                MyList2(s) = Trim(SplitArr(7))
            End If
        ElseIf ArrCount >= 10 Then
            If Trim(SplitArr(4)) <> "Login" And Trim(SplitArr(4)) <> "Downtime" And Trim(SplitArr(4)) <> "Non Production" And Trim(SplitArr(4)) <> "Logout" And Trim(SplitArr(10)) <> "Diarise" Then
                s = s + 1
                ReDim Preserve MyList1(1 To s)
                ReDim Preserve MyList2(1 To s)
                MyList1(s) = Trim(SplitArr(4))
                MyList2(s) = Trim(SplitArr(7))
            End If
        End If
    Next r
    
    Me.ListBox3.Clear
    
    With Me.ListBox3
        .ColumnCount = 2
        .ColumnWidths = "150;30"
        .AddItem
        .List(0, 0) = "Transaction"
        .List(0, 1) = "Count"
    End With
    
    For r = LBound(MyList1) To UBound(MyList1)
        t = Me.ListBox3.ListCount - 1
        For s = 1 To t
        If Me.ListBox3.List(s, 0) = MyList1(r) Then
            Me.ListBox3.List(s, 1) = Val(Me.ListBox3.List(s, 1)) + Val(MyList2(r))
            Cmt = "Update"
            Exit For
        End If
        Next s
        If Cmt <> "Update" Then
            Me.ListBox3.AddItem
            Me.ListBox3.List(t + 1, 0) = MyList1(r)
            Me.ListBox3.List(t + 1, 1) = Val(MyList2(r))
            Cmt = "New"
        End If
        Cmt = ""
    Next r
    
    Exit Sub
    
MyError:
    
       MsgBox "No data", vbInformation, "Alert": Exit Sub

End Sub



Private Sub LeftArrow_lbl_Click()

Dim i As Integer

For i = 0 To Me.ListBox2.ListCount - 1
    If ListBox2.Selected(i) Then
        ListBox1.AddItem Me.ListBox2.Column(0, ListBox2.ListIndex)
        ListBox2.RemoveItem (i)
    End If
Next i

End Sub


Private Sub RightArrow_lbl_Click()

Dim i As Integer

For i = 0 To Me.ListBox1.ListCount - 1
    If ListBox1.Selected(i) Then
        ListBox2.AddItem Me.ListBox1.Column(0, ListBox1.ListIndex)
        ListBox1.RemoveItem (i)
    End If
Next i

End Sub


Private Sub UserForm_Initialize()

Dim c As Range
Dim d As Range

Call FormatUserForm(Me.Caption)

Me.Label2.Caption = Format(Now, "dddd, dd mmmm yyyy")
Me.Label16.Caption = "Hi, " & UCase(Environ("Username"))

For Each d In ThisWorkbook.Sheets("Support Sheet").Range("A1:A20000")
    If UCase(d) = UCase(Environ("Username")) Then
        Me.Label22.Caption = d.Offset(0, 2).Value
    End If
Next

Me.Label20.Caption = VBA.Format(Now, "dd/mm/yyyy hh:mm:ss")
Me.Label21.Caption = "00:00:00"

Me.ComboBox1.Value = ""
Me.ComboBox1.Clear
Me.ComboBox4.Value = ""
Me.ComboBox4.Clear
For Each c In ThisWorkbook.Sheets("Support Sheet").Range("E1:E20000")
    If c = 1 Then
        Me.ComboBox1.AddItem c.Offset(0, 1).Value
        If c.Offset(0, 1) <> Me.Label22.Caption Then
            Me.ComboBox4.AddItem c.Offset(0, 1).Value
        End If
    End If
Next

Me.ComboBox4.Enabled = False
Me.ComboBox4.Visible = False
Me.ComboBox2.Value = ""

With Me.ListBox3
    .ColumnCount = 2
    .ColumnWidths = "150;30"
    .AddItem
    .List(0, 0) = "Transaction"
    .List(0, 1) = "Count"
End With

Me.ComboBox3.List = Array("Break", "Floor Support", "Team Meeting", "Process Training", "Softskill Training", "Cross Training", "Trainer's Hours", "One to One meeting", "Project Work", "Reports", "Workflow", "IT Issue", "Power Cut", "Fun @ Work", "Others")

End Sub

Private Sub Userform_QueryClose(Cancel As Integer, closemode As Integer)

If closemode = vbFormControlMenu Then
    MsgBox "Please use the Log Out button", vbInformation, "Alert"
    Cancel = True
End If

End Sub

