VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Toping Project"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13575
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   13575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "运行代码"
      Height          =   975
      Left            =   9240
      TabIndex        =   4
      Top             =   0
      Width           =   4335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "代码编程"
      Height          =   1095
      Left            =   9240
      TabIndex        =   3
      Top             =   960
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "快速执行"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "Toping Project的发展离不开开发者的努力"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   6255
   End
   Begin VB.Label Label2 
      Caption         =   "Beta1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Toping Project"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Open App.Path + "\Tmp.cde" For Output As #1
    Print #1, CodeWin.Text1.Text
Close #1
Open App.Path + "\Tmp.cde" For Input As #2
    While Not EOF(2)
        Dim TmpCode As String
        Line Input #2, TmpCode
        RunCode (TmpCode)
    Wend

Close #2
End Sub

Private Sub Command2_Click()
RunCode (Text1.Text)
End Sub

Private Sub Command3_Click()
CodeWin.Show
End Sub
