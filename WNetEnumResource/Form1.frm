VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ShareDat"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Options"
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      Top             =   960
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   3360
      Top             =   3000
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   3720
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   495
      Left            =   6840
      TabIndex        =   16
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GO"
      Height          =   615
      Left            =   6960
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hostname LookUp"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3840
      TabIndex        =   8
      Top             =   120
      Width           =   2295
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   6000
      TabIndex        =   7
      Top             =   1440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   7011
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":98BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B940
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7011
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   2880
      Top             =   5640
      Width           =   3375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "READY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2880
      TabIndex        =   15
      Top             =   5640
      Width           =   3375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Share founds"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Potentional target"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3120
      Picture         =   "Form1.frx":C21A
      Top             =   840
      Width           =   480
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   120
      Top             =   5520
      Width           =   8055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   0
      X2              =   8160
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000A&
      BorderWidth     =   2
      X1              =   0
      X2              =   8160
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label4 
      Caption         =   "Stop IP"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Start IP"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Active Scan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6240
      TabIndex        =   2
      Top             =   1080
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Discovery share resource"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2685
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 
 Dim x As Long
 
Private Sub Command1_Click()
Text1 = Text3
Text2 = Text3
End Sub

Private Sub Command2_Click()

If Command2.Caption = "GO" Then
Form2.Show
sali = False
Y = 0

Calc Text1, Text2
Timer1.Interval = Val(Form3.Text2) * 1000
Text1.Enabled = False
Text2.Enabled = False
Command2.Caption = "STOP"
SRange Text1, Text2
If StrComp(Text1, Text2) <> 0 Then MsgBox "Scan Finish " & Text1 & "---" & Text2, vbInformation, "ShareDat"
x = 0
x1 = 0
ElseIf Command2.Caption = "STOP" Then
sali = True
Y = 0
Text1.Enabled = True
Text2.Enabled = True
Command2.Caption = "GO"

End If




End Sub

Private Sub Command4_Click()
Form3.Show vbModal
End Sub

Private Sub ListView1_DblClick()
Call ShareEnum("\\" & ListView1.SelectedItem.Text)
Label8.Caption = x1
End Sub

Private Sub Timer1_Timer()

salir
ser = True
Timer1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
salir
For i = 1 To Val(Form3.Text1)
Unload Winsock1(i)
Next i
End
End Sub

Private Sub salir()
For i = 1 To Val(Form3.Text1)
Winsock1(i).Close
Next i
End Sub

Private Sub TreeView1_DblClick()

Shell "explorer.exe file://" & Mid(TreeView1.SelectedItem.Root, 3) & "\" & Mid(TreeView1.SelectedItem.Text, 1, InStr(TreeView1.SelectedItem, "(") - 1), 1

End Sub

Private Sub Winsock1_Connect(Index As Integer)
If Winsock1(Index).State = sckConnected Then
Set b = ListView1.ListItems.Add(, , Winsock1(Index).RemoteHostIP, , 4)
x = x + 1
Label6.Caption = x

End If
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1(Index).Close
End Sub


Private Sub Form_Load()
For i = 1 To Val(Form3.Text1)
Load Winsock1(i)
Next i
TreeView1.ImageList = ImageList1
ListView1.View = lvwReport
Set ListView1.SmallIcons = ImageList1
ListView1.ColumnHeaders.Add 1, , "IP", 2100
Text1 = Left(Winsock1(0).LocalIP, InStrRev(Winsock1(0).LocalIP, ".")) & 1
Text2 = Left(Winsock1(0).LocalIP, InStrRev(Winsock1(0).LocalIP, ".")) & 255
Text3 = Text2
End Sub
