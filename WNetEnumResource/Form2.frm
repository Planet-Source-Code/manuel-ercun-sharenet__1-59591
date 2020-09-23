VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LogErr"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6330
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Move Form1.Left + (Form1.Width / 2) - (Me.Width / 2), Form1.Top + (Form1.Height / 2) - (Me.Height / 2)

End Sub
