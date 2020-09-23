VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "options"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3195
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Show logerr"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   840
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "check ip"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   600
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Text            =   "2"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Text            =   "255"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "seg"
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "time"
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Socket"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Me.Move Form1.Left + (Form1.Width / 2) - (Me.Width / 2), Form1.Top + (Form1.Height / 2) - (Me.Height / 2)

End Sub
