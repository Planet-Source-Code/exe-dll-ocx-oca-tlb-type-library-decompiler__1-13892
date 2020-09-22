VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Class manager"
   ClientHeight    =   7095
   ClientLeft      =   2025
   ClientTop       =   1065
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   7800
   Begin RichTextLib.RichTextBox RichTextBox3 
      Height          =   735
      Left            =   2900
      TabIndex        =   12
      Top             =   -10
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   1296
      _Version        =   393217
      BackColor       =   12632256
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   11
      Top             =   360
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   2895
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5400
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   11
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":00FB
            Key             =   "Propertys"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":035F
            Key             =   "Events"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":05C3
            Key             =   "Functions"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   7740
      TabIndex        =   5
      Top             =   6360
      Width           =   7800
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   255
         Left            =   20
         TabIndex        =   6
         Top             =   0
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   450
         _Version        =   393217
         BackColor       =   12632256
         BorderStyle     =   0
         Enabled         =   0   'False
         Appearance      =   0
         TextRTF         =   $"Form1.frx":0827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RichTextBox2 
         Height          =   190
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   344
         _Version        =   393217
         BackColor       =   12632256
         BorderStyle     =   0
         Enabled         =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"Form1.frx":0916
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   7250
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Help 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   420
         Width           =   6495
      End
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5340
      ItemData        =   "Form1.frx":0A05
      Left            =   0
      List            =   "Form1.frx":0A07
      TabIndex        =   2
      Top             =   1005
      Width           =   2250
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5340
      ItemData        =   "Form1.frx":0A09
      Left            =   2280
      List            =   "Form1.frx":0A0B
      TabIndex        =   0
      Top             =   1005
      Width           =   5525
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5160
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      InitDir         =   "c:\Windows\System"
   End
   Begin VB.Image ClassIcon 
      Height          =   480
      Left            =   7080
      Picture         =   "Form1.frx":0A0D
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image PropertyIcon 
      Height          =   480
      Left            =   7080
      Picture         =   "Form1.frx":0F8F
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image FunctionIcon 
      Height          =   480
      Left            =   6840
      Picture         =   "Form1.frx":13D1
      Top             =   5880
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "No open file!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2355
      TabIndex        =   9
      Top             =   735
      Width           =   4335
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   2295
      Picture         =   "Form1.frx":1953
      Top             =   720
      Width           =   5520
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   0
      Picture         =   "Form1.frx":6FD5
      Top             =   720
      Width           =   2250
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   8400
      TabIndex        =   4
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Label itsName 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Label ItemName 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8400
      TabIndex        =   1
      Top             =   5040
      Width           =   3015
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu opene 
         Caption         =   "&Open..."
      End
      Begin VB.Menu sperator 
         Caption         =   "-"
      End
      Begin VB.Menu quit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu s 
      Caption         =   "&Settings"
      Begin VB.Menu font 
         Caption         =   "&Font..."
      End
   End
   Begin VB.Menu helps 
      Caption         =   "&Help"
      Begin VB.Menu about 
         Caption         =   "A&bout..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Itemi As Integer
Dim mm As String
Dim mmm As String
Dim DoAs As String
Dim strParameterType
Dim strParameters 'As String
Dim strReturnType 'As String

Private Sub about_Click()
Form2.Show
End Sub

Private Sub Combo1_Change()
'Close #1
'Close #3
'Close #9
'Changed (Combo1.ListIndex + 1)


End Sub

Private Sub Combo2_Change()

Clicked2 Combo2.ListIndex

End Sub

Private Sub Form_Load()
'List1.ListIndex = 1
'Me.Hide
'Form2.Show
'List1_Click
RichTextBox3.Text = "No file selected."

    strFilter = "DLL Libraries (*.dll)|*.dll|"
    strFilter = strFilter & "OCX ActiveX control Files (*.ocx)|*.ocx|"
    strFilter = strFilter & "OCA Cached Type libraries (*.oca)|*.oca|"
    strFilter = strFilter & "Raw Type libraries (*.tlb)|*.tlb|"
    strFilter = strFilter & "Executable Files (*.exe)|*.exe|"
    strFilter = strFilter & "All Files (*.*)|*.*"
    CommonDialog1.Filter = strFilter
    'CommonDialog1.ShowOpen
    'Command1_Click
 '   Label1.Caption = "Members of '" & mm & "'"
'Form2.Hide
'Me.Show
End Sub

Private Sub Command1_Click()
    If CommonDialog1.Flags = 0 Then
        Exit Sub
    End If

Open "C:\Output.txt" For Output As #1
Open "C:\Output3.txt" For Output As #3
Open "C:\Output9.txt" For Output As #9
    
    Dim X As TypeLibInfo
    Dim Y As CoClasses
    Dim z As Interfaces
    Dim w As Members
    Dim u As MemberInfo
    Dim i As Integer, j As Integer, n As Integer, k As Integer
    Dim strFilter As String
    Dim strName As String, strMembers As String
    
    On Error Resume Next
    'CommonDialog1.ShowOpen
    List1.Clear
    List2.Clear
    Text1.Text = ""
    
    'Program ends if you click the Cancel button in the
    'file open dialog box
    
    'Get information from type library
    Set X = TypeLibInfoFromFile(CommonDialog1.FileName)
    Set Y = X.CoClasses

    'Show Type Library information in the List box
    For i = 1 To Y.Count
        If i <> 1 Then
 '           strName = ""
 '           List1.AddItem strName
        End If
        strName = Y.Item(i).Name
        List2.AddItem strName
        RichTextBox3.Text = "Library '" & Y.Item(i).Parent & "'" & vbCrLf & _
        CommonDialog1.FileName & vbCrLf & _
        Y.Item(i).Parent.HelpString & vbCrLf & _
        "Version " & Y.Item(i).Parent.MajorVersion & "." & Y.Item(i).Parent.MinorVersion & vbCrLf & _
        "GUID: " & Y.Item(i).Parent.Guid

        

        mm = Y.Item(i).Parent & "." & strName
        mmm = Y.Item(i).Parent
        Set z = Y.Item(i).Interfaces
        For n = 1 To z.Count
            Set w = z.Item(n).Members
            
            For k = 1 To w.Count
                Set u = w.Item(k)
                strMembers = u.Name
                                strParameters = "("
                For a = 1 To u.Parameters.Count
                    If a = u.Parameters.Count Then
                        strParameters = strParameters & u.Parameters(a) & ")"
                     
                    Exit For
                    End If
                strParameterType = u.Parameters(a).VarTypeInfo
                'u.Parameters(a).
                
                str1 = strParameterType
                If strParameterType = "1" Then
                strParameterType = " As Integer"
                
                ElseIf strParameterType = "2" Then
                strParameterType = " As Integer"
                
                ElseIf strParameterType = "3" Then
                strParameterType = " As Long"
                
                ElseIf strParameterType = "4" Then
                strParameterType = " As Single"
                
                ElseIf strParameterType = "5" Then
                strParameterType = " As Double"
                
                Else: strParameterType = "Unknown(" & str1 & ")"
                End If
                strParameters = strParameters & u.Parameters(a) & strParameterType & ", " '& u.Parameters(2) & "," & u.Parameters(3) & "," & u.Parameters(4) & "," & u.Parameters(5) & ")"
                                
                
                Next a
                If strParameters = "(" Then strParameters = "()"
                'Debug.Print u.HelpString
                
                
                strReturnType = u.ReturnType.TypeInfoNumber
                
                DoAs = " As "
                str2 = strReturnType
                If strReturnType = "2" Then
                strReturnType = "DisplayModeConstants"
                
                ElseIf strReturnType = "4" Then
                strReturnType = "BorderStyleConstants"
                
                ElseIf strReturnType = "3" Then
                strReturnType = "AppearanceConstants"
                
                ElseIf strReturnType = "1" Then
                strReturnType = "StateConstants"
                
                ElseIf strReturnType = "0" Then
                strReturnType = "WindowSizeConstants"
                
                ElseIf strReturnType = "35" Then
                strReturnType = "StdPicture"
                
                ElseIf strReturnType = "5" Then
                strReturnType = "Long"
                
                ElseIf strReturnType = "29" Then
                strReturnType = "FrmBackStyle"
                
                ElseIf strReturnType = "20" Then
                strReturnType = "frmMousePointer"
                
                Else: strReturnType = "Unknown(" & str2 & ")"
                End If
                
                If str2 = "-1" Then
                    strReturnType = ""
                    DoAs = ""
                End If
                
                Print #1, strMembers & "" & strParameters & DoAs & strReturnType
                Print #3, u.HelpString
                Print #9, u.InvokeKind 'u.i
               'If u.InvokeKind = INVOKE_EVENTFUNC Then
               ' ListView1.ListItems.Add i, "", strMembers, 2, 2
               ' End If
                
               'If u.InvokeKind = INVOKE_FUNC Then
               ' ListView1.ListItems.Add i, "", strMembers, 3, 3
               ' End If
                
               ' If u.InvokeKind = INVOKE_PROPERTYGET Or INVOKE_PROPERTYPUT Then
               ' ListView1.ListItems.Add i, "", strMecmbers, 1, 1
               ' End If
               
                'List1.RightToLeft = False
                
            'If Len(strMembers) < 20 Then GoTo skip
                List1.AddItem strMembers '& " " & strParameters & DoAs & strReturnType
'skip:
            Next
        Next
    Next
    Combo11
   Combo22
   Close #1
   Close #3
    Close #9
    Combo2.Text = List1.List(0)
    
'MsgBox "Closed"
    Set z = Nothing
    Set Y = Nothing
    Set X = Nothing
    Set w = Nothing

    'Display filename in the text box
Label1.Caption = "Members of '" & mm & "'"
    Text1.Text = CommonDialog1.FileName

    'If the file does not contain type library information
    'then display this error message.
    If mm = "" Then
        Dim strMsgTitle As String, strMsgError As String
        Dim intResponse
        strMsgTitle = "No Type Library"
        strMsgError = "You chose a file without a type library. "
        strMsgError = strMsgError & "Choose another file."
        Err.Clear
        intResponse = MsgBox(strMsgError, vbOKCancel, strMsgTitle)

        If intResponse = vbOK Then
            CommonDialog1.ShowOpen
            Command1_Click
        End If

    End If
End Sub

Private Sub Command2_Click()
    End

End Sub





Private Sub List1_Click()
'test.Text = List1.List(List1.ListIndex)



ItemName.Caption = List1.Text
Open "C:\Output.txt" For Input As #2
Do While Not EOF(2)
    Line Input #2, data
    Xx = Xx + 1
    If Xx = List1.ListIndex + 1 Then Exit Do
Loop
Close #2
itsName = data

Dim u As MemberInfo
Open "C:\Output3.txt" For Input As #5
Do While Not EOF(5)
    Line Input #5, Data2
    Xxx = Xxx + 1
    
    If Xxx = List1.ListIndex + 1 Then
        Help.Caption = Data2
        Exit Do
    End If
Loop
Close #5

Label2.Caption = "Member of " & mm
Label1.Caption = "Members of '" & mm & "'"

RichTextBox2.Text = "Member of " & mm
RichTextBox2.SelStart = 10
RichTextBox2.SelLength = 100
RichTextBox2.SelUnderline = True
RichTextBox2.SelColor = &H8000&
RichTextBox2.SelStart = 0


Open "C:\Output9.txt" For Input As #10
Do While Not EOF(10)
    Line Input #10, Data3
    Xxxx = Xxxx + 1
    
    If Xxxx = List1.ListIndex + 1 Then
        If Trim(Data3) = "1" Then
        Data3 = "Event"
        Image3.Picture = FunctionIcon.Picture
        End If
        If Trim(Data3) = "2" Then
        Data3 = "Property"
        Image3.Picture = PropertyIcon.Picture
        End If
        If Trim(Data3) = "4" Then Data3 = "Property"
        If Trim(Data3) = "8" Then Data3 = "Property"
        
        
        
        itsName = Data3 & " " & itsName
        RichTextBox1.Text = itsName
        Bold2 itsName
        Exit Do
    End If
Loop
Close #10

'Open "C:\x.txt" For Output As #99
'For b = 0 To List1.ListCount
'Print #99, List1.List(b)
'Next b
'Close #99
End Sub



Function DoBold(data)
'RichTextBox1.Text = ""
RichTextBox1.SelBold = False
RichTextBox1.SelLength = 100
RichTextBox1.SelStart = 0
        
Dim ReturnValue As String
Dim a As Integer
If Left(data, 8) = "Property" Then lastpos = 15
For a = lastpos To Len(data)
    If Mid(data, a, 1) = " " Then
        ReturnValue = Mid(data, lastpos, a - lastpos)
        Debug.Print ReturnValue
        'richtext
        'RichTextBox1.Text = Test
'        RichTextBox1.HideSelection = True
        'RichTextBox1.SelTexst
        
        RichTextBox1.SelStart = lastpos - 1
        RichTextBox1.SelLength = Len(ReturnValue)
        RichTextBox1.SelBold = True
        RichTextBox1.SelStart = 0
        RichTextBox1.SelBold = False
         lastpos = a + 1
        Exit For
    End If
Next a

DoBold = ReturnValue
''Msgbox ReturnValue
'a = 0

End Function



Function Bold2(data As String)
If Left(data, 8) = "Property" Then lastpos = 9
If Left(data, 5) = "Event" Then lastpos = 6
If Left(data, 5) = "Class" Then lastpos = 6
RichTextBox1.SelStart = lastpos

If InStr(1, data, "()") <> 0 Then
    RichTextBox1.SelLength = InStr(1, data, "()") + 1 - lastpos
GoTo pass2
End If

RichTextBox1.SelLength = InStr(1, data, "(") - 1 - lastpos

pass2:
RichTextBox1.SelBold = True
RichTextBox1.SelStart = 0
RichTextBox1.SelBold = False
 
End Function

Private Sub List2_Click()
'Close #1
'Close #3
'close #9
Changed List2.ListIndex + 1
RichTextBox1.Text = "Class " & List2.List(List2.ListIndex)
Bold3 ("Class " & List2.List(List2.ListIndex))
Image3.Picture = ClassIcon.Picture
Help.Caption = "This class contains " & List1.ListCount & " members."
Label1.Caption = "Member of '" & mmm & "'"
RichTextBox2.Text = "Member of " & mmm
RichTextBox2.SelStart = 10
RichTextBox2.SelLength = 100
RichTextBox2.SelUnderline = True
RichTextBox2.SelColor = &H8000&
RichTextBox2.SelStart = 0
Label1.Caption = "Members of '" & mm & "'"

End Sub


Function Changed(index As Long)
    
Open "C:\Output.txt" For Output As #1
Open "C:\Output3.txt" For Output As #3
Open "C:\Output9.txt" For Output As #9
    
    Dim X As TypeLibInfo
    Dim Y As CoClasses
    Dim z As Interfaces
    Dim w As Members
    Dim u As MemberInfo
    Dim i As Integer, j As Integer, n As Integer, k As Integer
    Dim strFilter As String
    Dim strName As String, strMembers As String
    
    On Error Resume Next
    'CommonDialog1.ShowOpen
    List1.Clear
    'List2.Clear
    Text1.Text = ""
Label1.Caption = "Members of '" & mm & "'"
    
    'Program ends if you click the Cancel button in the
    'file open dialog box
    
    'Get information from type library
    Set X = TypeLibInfoFromFile(CommonDialog1.FileName)
    Set Y = X.CoClasses
    
    'Show Type Library information in the List box
    i = index
        If i <> 1 Then
            'strName = ""
            'List1.AddItem strName
        End If
        
        strName = Y.Item(i).Name
        'List2.AddItem strName
        mm = Y.Item(i).Parent & "." & strName
        
        Set z = Y.Item(i).Interfaces
        For n = 1 To z.Count
            Set w = z.Item(n).Members
            
            For k = 1 To w.Count
                Set u = w.Item(k)
                strMembers = u.Name
                                strParameters = "("
                For a = 1 To u.Parameters.Count
                    If a = u.Parameters.Count Then
                        strParameters = strParameters & u.Parameters(a) & ")"
                     
                    Exit For
                    End If
                strParameterType = u.Parameters(a).VarTypeInfo
                str1 = strParameterType
                If strParameterType = "1" Then
                strParameterType = " As Integer"
                
                ElseIf strParameterType = "2" Then
                strParameterType = " As Integer"
                
                ElseIf strParameterType = "3" Then
                strParameterType = " As Long"
                
                ElseIf strParameterType = "4" Then
                strParameterType = " As Single"
                
                ElseIf strParameterType = "5" Then
                strParameterType = " As Double"
                
                Else: strParameterType = " As Unknown(" & str1 & ")"
                End If
                
                strParameters = strParameters & u.Parameters(a) & strParameterType & ", " '& u.Parameters(2) & "," & u.Parameters(3) & "," & u.Parameters(4) & "," & u.Parameters(5) & ")"
                
                
                Next a
                If strParameters = "(" Then strParameters = "()"
                'Debug.Print u.HelpString
                
                
                strReturnType = u.ReturnType.TypeInfoNumber
                
                DoAs = " As "
                str2 = strReturnType
                If strReturnType = "2" Then
                strReturnType = "DisplayModeConstants"
                
                ElseIf strReturnType = "4" Then
                strReturnType = "BorderStyleConstants"
                
                ElseIf strReturnType = "3" Then
                strReturnType = "AppearanceConstants"
                
                ElseIf strReturnType = "1" Then
                strReturnType = "StateConstants"
                
                ElseIf strReturnType = "0" Then
                strReturnType = "WindowSizeConstants"
                
                ElseIf strReturnType = "35" Then
                strReturnType = "StdPicture"
                
                ElseIf strReturnType = "5" Then
                strReturnType = "Long"
                
                ElseIf strReturnType = "29" Then
                strReturnType = "FrmBackStyle"
                
                ElseIf strReturnType = "20" Then
                strReturnType = "frmMousePointer"
                
                Else: strReturnType = "Unknown(" & str2 & ")"
                End If
                
                If str2 = "-1" Then
                    strReturnType = ""
                    DoAs = ""
                End If
                
                Print #1, strMembers & "" & strParameters & DoAs & strReturnType
                Print #3, u.HelpString
                Print #9, u.InvokeKind 'u.i
               'If u.InvokeKind = INVOKE_EVENTFUNC Then
               ' ListView1.ListItems.Add i, "", strMembers, 2, 2
               ' End If
                
               'If u.InvokeKind = INVOKE_FUNC Then
               ' ListView1.ListItems.Add i, "", strMembers, 3, 3
               ' End If
                
               ' If u.InvokeKind = INVOKE_PROPERTYGET Or INVOKE_PROPERTYPUT Then
               ' ListView1.ListItems.Add i, "", strMecmbers, 1, 1
               ' End If
               
                List1.RightToLeft = False
                
                If Trim(strMembers) = "" Then GoTo skip
                List1.AddItem strMembers '& " " & strParameters & DoAs & strReturnType
skip:
            Next
        Next
Combo11
Combo22
Close #1
   Close #3
    Close #9
'MsgBox "Closed"
    Set z = Nothing
    Set Y = Nothing
    Set X = Nothing
    Set w = Nothing

    'Display filename in the text box
    Text1.Text = CommonDialog1.FileName

    'If the file does not contain type library information
    'then display this error message.
    If Err.Number = 91 Then
        Dim strMsgTitle As String, strMsgError As String
        Dim intResponse
        strMsgTitle = "No Type Library"
        strMsgError = "You chose a file without a type library. "
        strMsgError = strMsgError & "Choose another file."
        Err.Clear
        intResponse = MsgBox(strMsgError, vbOKCancel, strMsgTitle)

        If intResponse = vbOK Then
            Command1_Click
        End If

    End If

End Function

Function Bold3(data As String)
If Left(data, 8) = "Property" Then lastpos = 9
If Left(data, 5) = "Event" Then lastpos = 6
If Left(data, 5) = "Class" Then lastpos = 6
RichTextBox1.SelStart = lastpos


RichTextBox1.SelLength = Len(data) - lastpos + 1

pass2:
RichTextBox1.SelBold = True
RichTextBox1.SelStart = 0
RichTextBox1.SelBold = False
 
End Function

Private Sub opene_Click()
CommonDialog1.ShowOpen
Command1_Click
End Sub


Function Combo11()
For X = 0 To List2.ListCount
If List2.List(X) = "" Then GoTo 1
Combo1.AddItem mmm & "." & List2.List(X)
1:
Next X
Combo1.Text = mm
End Function

Function Combo22()
For X = 0 To List1.ListCount
If List1.List(X) = "" Then GoTo 1
Combo2.AddItem List1.List(X)
1:
Next X
'Combo1.Text = mm
End Function

Function Clicked2(index As Integer)
'test.Text = List1.List(List1.ListIndex)



ItemName.Caption = List1.Text
Open "C:\Output.txt" For Input As #2
Do While Not EOF(2)
    Line Input #2, data
    Xx = Xx + 1
    If Xx = index + 1 Then Exit Do
Loop
Close #2
itsName = data

Dim u As MemberInfo
Open "C:\Output3.txt" For Input As #5
Do While Not EOF(5)
    Line Input #5, Data2
    Xxx = Xxx + 1
    
    If Xxx = index + 1 Then
        Help.Caption = Data2
        Exit Do
    End If
Loop
Close #5

Label2.Caption = "Member of " & mm
Label1.Caption = "Members of '" & mm & "'"

RichTextBox2.Text = "Member of " & mm
RichTextBox2.SelStart = 10
RichTextBox2.SelLength = 100
RichTextBox2.SelUnderline = True
RichTextBox2.SelColor = &H8000&
RichTextBox2.SelStart = 0


Open "C:\Output9.txt" For Input As #10
Do While Not EOF(10)
    Line Input #10, Data3
    Xxxx = Xxxx + 1
    
    If Xxxx = index + 1 Then
        If Trim(Data3) = "1" Then Data3 = "Event"
        If Trim(Data3) = "2" Then Data3 = "Property"
        If Trim(Data3) = "4" Then Data3 = "Property"
        If Trim(Data3) = "8" Then Data3 = "Property"
        
        
        
        itsName = Data3 & " " & itsName
        RichTextBox1.Text = itsName
        Bold2 itsName
        Exit Do
    End If
Loop
Close #10

'Open "C:\x.txt" For Output As #99
'For b = 0 To List1.ListCount
'Print #99, List1.List(b)
'Next b
'Close #99
End Function

