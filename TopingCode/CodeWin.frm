VERSION 5.00
Begin VB.Form CodeWin 
   Caption         =   "Code"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11280
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CodeWin.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7575
   ScaleWidth      =   11280
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   11295
   End
End
Attribute VB_Name = "CodeWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
Text1.Width = Me.ScaleWidth
Text1.Height = Me.ScaleHeight
End Sub
