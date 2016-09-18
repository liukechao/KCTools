VERSION 5.00
Begin VB.Form NumberAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "数字相加"
   ClientHeight    =   8070
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   6945
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   6945
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton TestExit 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   10
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton TestReset 
      Caption         =   "重置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   9
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Timer Timer 
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton TestPause 
      Caption         =   "暂停"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   8
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton TestStart 
      Caption         =   "开始"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   7
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox InputText 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   2
      Top             =   3480
      Width           =   4455
   End
   Begin VB.Label SumLabel 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   12
      Top             =   4560
      Width           =   4455
   End
   Begin VB.Label QLabel 
      Caption         =   "1.1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   30
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   11
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label WrongLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   5040
      TabIndex        =   6
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label RightLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   5
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label WLabel 
      Caption         =   "错误:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   4
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label RLabel 
      Caption         =   "正确:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   3
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label RandomLabel 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   1
      Top             =   2400
      Width           =   4455
   End
   Begin VB.Label TimeLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "NumberAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim minute As Integer
Dim second As Integer
Dim rightcount As Integer
Dim wrongcount As Integer
Dim question As Integer
Dim number As Integer
Dim rightsum As String
Dim inputsum As String

Private Sub Initialize()
    minute = 0
    second = 0

    question = 1
    number = 1
    
    rightsum = "0"
    inputsum = "0"

    rightcount = 0
    wrongcount = 0
    
    Timer.Interval = 1000
    Timer.Enabled = False

    TimeLabel.Caption = "00:00"
    
    RandomLabel.Caption = "0"

    InputText.Enabled = True
    InputText.Text = ""
    
    QLabel.Caption = Trim(Str(question)) + "." + Trim(Str(number))
    
    SumLabel.Caption = "0"
    
    RightLabel.Caption = "0"
    WrongLabel.Caption = "0"
End Sub

Private Sub GenRandomNumber()
    Dim digits As Integer
    Dim randomstr As String
    randomstr = ""

    Randomize
    digits = Int(Rnd * 7) + 4
    
    Randomize
    For i = 1 To digits
        randomstr = randomstr + Trim(Str(Int(Rnd * 9) + 1))
    Next i
    RandomLabel.Caption = randomstr
    InputText.Text = ""
End Sub


Private Sub Timer_Timer()
   If second = 0 Then
        second = 59
        minute = minute - 1
    Else
        second = second - 1
    End If

    Dim secondstr As String
    secondstr = "00" + Trim(Str(second))
    TimeLabel.Caption = Trim(Str(minute)) + ":" + Right(secondstr, 2)

    If minute = 0 And second = 0 Then
        Timer.Enabled = False
        InputText.Enabled = False
        TestStart.Caption = "开始"
        MsgBox ("时间到！")
    End If
End Sub

Private Sub Form_Load()
    Initialize
End Sub

Private Sub Form_Terminate()
    Timer.Enabled = False
End Sub

Private Sub TestStart_Click()
    If minute = 0 And second = 0 Then
        Initialize
        Dim inputstr As String
        inputstr = InputBox("请输入时间（分钟）", "输入时间", "10")
        If Len(inputstr) = 0 Then
            Exit Sub
        Else
            minute = Val(inputstr)
        End If
        second = 0
        TimeLabel.Caption = Trim(Str(minute)) + ":00"
        GenRandomNumber
    End If

    If TestStart.Caption = "继续" Then
        TestStart.Caption = "开始"
    End If

    Timer.Enabled = True
    
    InputText.SetFocus
End Sub

Private Sub TestPause_Click()
    Timer.Enabled = False
    TestStart.Caption = "继续"
End Sub

Private Sub TestReset_Click()
    Initialize
End Sub

Private Sub TestExit_Click()
    Unload Me
End Sub

Private Sub InputText_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If number <= 50 Then
            rightsum = Functions.sum(rightsum, Trim(RandomLabel.Caption))
            inputsum = Functions.sum(inputsum, Trim(InputText.Text))
            
            SumLabel.Caption = inputsum
            
            number = number + 1
        End If
        
        If number = 50 + 1 Then
            If rightsum = inputsum Then
                rightcount = rightcount + 1
            Else
                wrongcount = wrongcount + 1
            End If
            
            number = 1
            question = question + 1
        
            rightsum = "0"
            inputsum = "0"
                                
            RightLabel.Caption = rightcount
            WrongLabel.Caption = wrongcount
        End If
        
        QLabel.Caption = Trim(Str(question)) + "." + Trim(Str(number))

        GenRandomNumber
    End If
End Sub
