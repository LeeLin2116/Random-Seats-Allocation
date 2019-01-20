VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "RANDOM ALLOCATE SEATS"
   ClientHeight    =   7125
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form1.frx":0000
      Left            =   1080
      List            =   "Form1.frx":0010
      TabIndex        =   13
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "New Seat"
      Height          =   375
      Left            =   960
      TabIndex        =   11
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Build"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox Columns 
      Height          =   270
      Left            =   1680
      TabIndex        =   9
      Text            =   "7"
      Top             =   5880
      Width           =   255
   End
   Begin VB.TextBox Rows 
      Height          =   270
      Left            =   1320
      TabIndex        =   8
      Text            =   "6"
      Top             =   5880
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Include"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   4560
      ItemData        =   "Form1.frx":002E
      Left            =   120
      List            =   "Form1.frx":0030
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox UIndex 
      Height          =   270
      Left            =   1320
      TabIndex        =   3
      Text            =   "43"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox LIndex 
      Height          =   270
      Left            =   1320
      TabIndex        =   2
      Text            =   "1"
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   6855
      Left            =   2040
      ScaleHeight     =   6795
      ScaleWidth      =   9075
      TabIndex        =   12
      Top             =   120
      Width           =   9135
      Begin VB.PictureBox Seat 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   0
         Left            =   1080
         ScaleHeight     =   585
         ScaleWidth      =   705
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Teacher"
         Enabled         =   0   'False
         Height          =   3375
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Direction:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   14
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Rows,Columns:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "To        :"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Range From:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu Ap 
      Caption         =   "Applocate"
      Begin VB.Menu ABN 
         Caption         =   "Applocate By List"
      End
      Begin VB.Menu ABL 
         Caption         =   "Applocate By List And Name"
      End
   End
   Begin VB.Menu Save 
      Caption         =   "Save"
      Begin VB.Menu TakeI 
         Caption         =   "Take Image"
      End
      Begin VB.Menu CallPrinter 
         Caption         =   "Print"
      End
   End
   Begin VB.Menu ClickSeat 
      Caption         =   "ClickSeat"
      Visible         =   0   'False
      Begin VB.Menu appoint 
         Caption         =   "Appoint"
      End
      Begin VB.Menu delAppoint 
         Caption         =   "Cancel Appointment"
         Visible         =   0   'False
      End
      Begin VB.Menu delete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu ClickList 
      Caption         =   "ClickList"
      Visible         =   0   'False
      Begin VB.Menu Join 
         Caption         =   "Join"
      End
      Begin VB.Menu Remove 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringW" _
    (ByVal lpApplicationName As Long, _
    ByVal lpKeyName As Long, _
    ByVal lpDefault As Long, _
    ByVal lpReturnedString As Long, _
    ByVal nSize As Long, _
    ByVal lpFileName As Long) As Long
    
Private Declare Function PrintWindow Lib "user32" (ByVal Hwnd As Long, ByVal HDC As Long, ByVal nFlags As Long) As Long

Private Type Seats
    Belong As String
    isAppointed As Boolean
End Type

Dim RosWidth As Integer, RosHeight As Integer, FWidth As Long, FHeight As Long
Dim FocusID As Integer, Seats() As Seats

Private Function GetValueFromINIFile(ByVal SectionName As String, _
    ByVal KeyName As String, _
    ByVal IniFileName As String) As String
    
    Dim strBuf As String
    
    strBuf = String(128, 0)
    
    GetPrivateProfileString StrPtr(SectionName), _
    StrPtr(KeyName), _
    StrPtr(""), _
    StrPtr(strBuf), _
128, _
    StrPtr(IniFileName)
    
    strBuf = Replace(strBuf, Chr(0), "")
    GetValueFromINIFile = strBuf
End Function



Private Function ReadFile(filePath As String) As String
    If (Dir(filePath) <> "") Then
        Dim fileStr As String
        Open filePath For Input As #1
        
        Do While Not EOF(1)
            Line Input #1, tem
            fileStr = fileStr & tem & vbCrLf
        Loop
        Close #1
        ReadFile = fileStr
    Else
        MsgBox "File couldn't be found", 16
    End If
End Function

Function getID(ID) As Integer
    Dim i
    For i = 0 To List1.ListCount
        If List1.List(i) = ID Then getID = i: Exit Function
    Next
    getID = -1
End Function

Private Sub ABL_Click()
    m = UBound(Seats)
    For i = 1 To m
        If List1.List(0) = "" Then GoTo a1
        If Not Seats(i).isAppointed Then
            Randomize Timer
            g = Int(Rnd * List1.ListCount)
            Seats(i).Belong = List1.List(g)
            List1.RemoveItem g
        End If
        If i = m And List1.ListCount <> 0 Then MsgBox "No More Seat, and there are(is) still " & List1.ListCount & " persons without seats", 48: GoTo a1

    Next
a1:
    Dim d() As String
    d = Split(ReadFile(App.Path & IIf(Right(App.Path, "1") = "\", "", "\") & GetValueFromINIFile("Config", "studentList", App.Path & IIf(Right(App.Path, "1") = "\", "", "\") & "setting.ini")), vbCrLf)

    For i = 0 To UBound(Seats)
        With Seat(i)
            .Cls
            .Font.Size = (.Width - 120) / 60
            .CurrentX = 60
            .CurrentY = 60
            Seat(i).Print Seats(i).Belong
            If IsNumeric(Seats(i).Belong) Then Seat(i).Print d(Seats(i).Belong - 1)
        End With
    Next
    ReDim Seats(0)
End Sub

Private Sub appoint_Click()
    d = InputBox("Please Input His/Her Number:")
    If d <> "" Then
        c = getID(d)
        If c = -1 Then
            MsgBox "System couldn't find ID " & d, 48
        Else
            With Seats(FocusID)
                .Belong = d
                .isAppointed = True
            End With
            
            With Seat(FocusID)
                .Font.Size = (.Width - 120) / 60
                .CurrentX = 60
                .CurrentY = 60
                
                List1.RemoveItem c
                
                Seat(FocusID).Print d
            End With
        End If
    End If
End Sub

Private Sub Command2_Click()
    List1.Clear
    Dim i
    For Each i In Split(ReadFile(App.Path & IIf(Right(App.Path, "1") = "\", "", "\") & GetValueFromINIFile("Config", "studentList", App.Path & IIf(Right(App.Path, "1") = "\", "", "\") & "setting.ini")), vbCrLf)
        If i <> "" Then List1.AddItem i
    Next
End Sub

Private Sub Command3_Click()
    With Command6
        Columns = Columns + 1
        Select Case Combo1.ListIndex
        Case 0
            .Width = RosWidth
            .Height = RosHeight
            .Left = 120
            .Top = Picture1.Height / 2 - .Height / 2
            Seat(0).Width = (Picture1.Width - Command6.Width - 320 - Rows * 300) / Rows
            Seat(0).Height = (Picture1.Height - 320 - Colunmns * 300) / Columns
        Case 1
            .Width = RosWidth
            .Height = RosHeight
            .Left = Picture1.Width - .Width - 120
            .Top = Picture1.Height / 2 - .Height / 2
            Seat(0).Width = (Picture1.Width - Command6.Width - 320 - Rows * 300) / Rows
            Seat(0).Height = (Picture1.Height - 320 - Colunmns * 300) / Columns
        Case 2
            .Height = RosWidth
            .Width = RosHeight
            .Top = 120
            .Left = Picture1.Width / 2 - .Width / 2
            Seat(0).Width = (Picture1.Width - 320 - Rows * 300) / Rows
            Seat(0).Height = (Picture1.Height - Command6.Height - 320 - Colunmns * 300) / Columns
        Case 3
            .Height = RosWidth
            .Width = RosHeight
            .Top = Picture1.Height - 120 - .Height
            .Left = Picture1.Width / 2 - .Width / 2
            Seat(0).Width = (Picture1.Width - 320 - Rows * 300) / Rows
            Seat(0).Height = (Picture1.Height - Command6.Height - 320 - Colunmns * 300) / Columns
        End Select
        Seat(0).Width = Seat(0).Width * 3 / 4
        Seat(0).Height = Seat(0).Height * 3 / 4
        
    End With
    Columns = Columns - 1
    
    If Not (IsNumeric(Rows) And IsNumeric(Columns)) Then MsgBox "Wrong Parameters Set": Exit Sub
    If Rows >= 20 Or Columns >= 20 Then Exit Sub
    Dim SW As Long, SH As Long, i As Integer
    
    If Seat.UBound <> 0 Then
        For i = 1 To Seat.UBound
            Unload Seat(i)
        Next
        ReDim Seats(0)
    End If
    
    For SW = 1 To Rows
        For SH = 1 To Columns
            i = Seat.UBound + 1
            Load Seat(i)
            ReDim Seats(i)
            With Seat(i)
                .Font.Size = (.Width - 120) / 60
                .CurrentX = 60
                .CurrentY = 60
                .Left = IIf(Combo1.ListIndex = 0, Command6.Width + 520, 200) + (.Width + 230) * (SW - 1) * 4 / 3.3
                .Top = IIf(Combo1.ListIndex = 2, Command6.Height + 520, 200) + (.Height + 230) * (SH - 1) * 4 / 3.3
                .Visible = True
            End With
        Next
    Next
    
End Sub

Private Sub Command1_Click()
    List1.Clear
    Dim i
    If Not (IsNumeric(LIndex) And IsNumeric(UIndex)) Then MsgBox "Wrong Parameters Set": Exit Sub
    For i = LIndex To UIndex Step 1
        List1.AddItem i
    Next
End Sub

Private Sub ABN_Click()
    m = UBound(Seats)
    For i = 1 To m
        If List1.List(0) = "" Then GoTo a1
        If Not Seats(i).isAppointed Then
            Randomize Timer
            g = Int(Rnd * List1.ListCount)
            Seats(i).Belong = List1.List(g)
            List1.RemoveItem g
        End If
        If i = m And List1.ListCount <> 0 Then MsgBox "No More Seat, and there are(is) still " & List1.ListCount & " persons without seats", 48: GoTo a1

    Next
a1:
    For i = 0 To UBound(Seats)
        With Seat(i)
            .Cls
            .Font.Size = (.Width - 120) / 60
            .CurrentX = 60
            .CurrentY = 60
            Seat(i).Print Seats(i).Belong
        End With
    Next
    ReDim Seats(0)
End Sub



Private Sub Command5_Click()
    i = Seat.UBound + 1
    ReDim Seats(i)
    RH = IIf(i Mod Columns = 0, i / Columns - 1, i \ Columns)
    RM = i Mod Columns
    If RM = 0 Then RM = Columns
    Load Seat(i)
    ReDim Seats(i)
    With Seat(i)
        .Left = IIf(Combo1.ListIndex = 0, Command6.Width + 520, 200) + (.Width + 230) * ((RH) * 4 / 3.3)
        If .Left + .Width + Picture1.Left > Me.Width Then Me.Width = Me.Width + (230 + .Width) * 4 / 3.3
        If .Left + .Width >= Picture1.Width - 120 Then Me.Width = Me.Width + 120
        .Top = IIf(Combo1.ListIndex = 2, Command6.Height + 520, 200) + (.Height + 230) * (RM - 1) * 4 / 3.3
        .Visible = True
    End With
End Sub

Private Sub Command8_Click()

End Sub

Private Sub TakeI_Click()
    Dim PCI As String
    PCI = GetValueFromINIFile("Config", "saveTo", App.Path & IIf(Right(App.Path, "1") = "\", "", "\") & "setting.ini")
    If (Dir(PCI) <> "") Then Kill PCI
    PrintWindow Picture1.Hwnd, Picture1.HDC, 0
    SavePicture Picture1.Image, App.Path & IIf(Right(App.Path, "1") = "\", "", "\") & PCI
    Picture1.Cls
End Sub

Private Sub CallPrinter_Click()
    On Error GoTo Hm:
    PrintWindow Picture1.Hwnd, Picture1.HDC, 0
    Printer.PaintPicture Picture1.Image, 0, 0
    Printer.EndDoc
    Picture1.Cls
    Exit Sub
Hm:
    Picture1.Cls
    MsgBox "Printer Error", 16
End Sub

Private Sub delAppoint_Click()
    List1.AddItem Seats(FocusID).Belong
    Seats(FocusID).isAppointed = False
    Seats(FocusID).Belong = 0
    Seat(FocusID).Cls
End Sub

Private Sub delete_Click()
    If Seats(FocusID).isAppointed Then MsgBox "The seat appointed to someone could be delete": Exit Sub
    Seats(FocusID) = Seats(UBound(Seats))
    With Seat(FocusID)
        .Top = Seat(Seat.UBound).Top
        .Left = Seat(Seat.UBound).Left
    End With
    If Seats(FocusID).isAppointed Then
        Seat(FocusID).Print Seats(FocusID).Belong
    End If
    ReDim Preserve Seats(UBound(Seats) - 1)
    Unload Seat(Seat.UBound)
End Sub

Private Sub Form_Load()
    RosWidth = Command6.Width
    RosHeight = Command6.Height
    FWidth = Me.Width
    FHeight = Me.Height
    ReDim Seats(0)
    Combo1.ListIndex = 0
    
    Command3_Click
    If Dir(App.Path & IIf(Right(App.Path, "1") = "\", "", "\") & GetValueFromINIFile("Config", "studentList", App.Path & IIf(Right(App.Path, "1") = "\", "", "\") & "setting.ini")) <> "" Then UIndex.Text = UBound(Split(ReadFile(App.Path & IIf(Right(App.Path, "1") = "\", "", "\") & GetValueFromINIFile("Config", "studentList", App.Path & IIf(Right(App.Path, "1") = "\", "", "\") & "setting.ini")), vbCrLf))
End Sub

Private Sub Form_Resize()
    If Me.Width < FWidth Then Me.Width = FWidth
    If Me.Height < FHeight Then Me.Height = FHeight
    Picture1.Width = Me.Width - Picture1.Left - 120
    Picture1.Height = Me.Height - 700
End Sub

Private Sub Join_Click()
    d = InputBox("Please input his/her ID")
    If d <> "" Then
        Dim i
        For i = 0 To List1.ListCount
            If d = List1.List(i) Then
                If MsgBox("This ID already exists in the list" & vbCrLf & "Would you still like to join it?", 48 + vbYesNo) = vbYes Then List1.AddItem d
                Exit Sub
            End If
        Next
        List1.AddItem d
    End If
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Me.PopupMenu ClickList

End Sub

Private Sub Remove_Click()
    If List1.ListIndex <> -1 Then List1.RemoveItem List1.ListIndex Else MsgBox "You select nothing to remove from the list", 16
End Sub

Private Sub Seat_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    FocusID = Index
    If Button = 2 Then
        delAppoint.Visible = Seats(Index).isAppointed
        appoint.Visible = Not Seats(Index).isAppointed
        Me.PopupMenu ClickSeat
    End If
End Sub

Private Sub Seat_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static a As Single, b As Single, c As Single, m As Single
    With Seat(Index)
        If Button Then
            .Move .Left + Int((X - a) / (.Width + 230) * 4 / 4 * 3.3) / 4 * (.Width + 230) * 4 / 3.3, .Top + Int((Y - b) / (.Height + 230) * 4 / 4 * 3.3) / 4 * (.Height + 230) * 4 / 3.3
        Else
            a = X
            b = Y
        End If
    End With
End Sub
