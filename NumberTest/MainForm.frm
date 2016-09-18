VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "数字测试"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   4560
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton ExitCommand 
      Caption         =   "退出测试"
      Height          =   735
      Left            =   1200
      TabIndex        =   2
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton AddCommand 
      Caption         =   "数字相加"
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton InputCommand 
      Caption         =   "数字录入"
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub InputCommand_Click()
    NumberInput.Show (1)
End Sub

Private Sub AddCommand_Click()
    NumberAdd.Show (1)
End Sub

Private Sub ExitCommand_Click()
    Unload Me
End Sub

