Attribute VB_Name = "Module2"
Option Explicit



Public i As Long, Y As Long
Public v() As String
Public ran() As Long, ran1() As Long
Public ser As Boolean, sali As Boolean
Public b As Object, an As String
Dim result As Long
Public a
Public x1 As Long


Private Function Uni(ip As String) As String
On Error Resume Next
v = Split(ip, ".")
ReDim ran(UBound(v))
For i = LBound(v) To UBound(v)
ran(i) = v(i)
Next i

Uni = ran(0) & "." & ran(1) & "." & ran(2) & "." & ran(3)
End Function

Private Function sumar(d0 As Long, d1 As Long, d2 As Long, d3 As Long) As String
On Error Resume Next
Dim ina As String
d3 = d3 + 1
If d3 = 255 Then
d3 = 0
d2 = d2 + 1
End If
If d2 = 255 Then
d3 = 0
d2 = 0
d1 = d1 + 1
End If
If d1 = 255 Then
d3 = 0
d2 = 0
d1 = 0
d0 = d0 + 1
End If
If d0 = 255 Then
d0 = 0
d1 = 0
d2 = 0
d3 = 0
End If

ina = d0 & "." & d1 & "." & d2 & "." & d3
Form1.Label9.Caption = ina
sumar = ina
Y = Y + 1


End Function

Public Sub SRange(ip As String, IPend As String)
On Error Resume Next
Uni ip
If StrComp(ip, IPend) = 0 Then
With Form1.Winsock1(0)
  .Close
  .Connect ip, 139
End With
With Form1
   .Text1.Enabled = True
   .Text2.Enabled = True
   .Command2.Caption = "GO"
End With
Form1.Label9.Caption = ip
Exit Sub
End If
Do

For i = 1 To Val(Form3.Text1)
Form1.Winsock1(i).Close
Form1.Winsock1(i).Connect sumar(ran(0), ran(1), ran(2), ran(3)), 139

Next i
ser = False

Form1.Timer1.Enabled = True
Do
DoEvents
Loop Until ser = True
Form1.Timer1.Enabled = False


If StrComp(Form1.Label9.Caption, Form1.Text2) = 0 Or Y >= result Or sali = True Then
sali = False
With Form1
   .Text1.Enabled = True
   .Text2.Enabled = True
   .Command2.Caption = "GO"
End With


Exit Sub

End If

Loop
End Sub



Public Sub Calc(ip As String, ip1 As String)
On Error Resume Next
If StrComp(ip, ip1) = 0 Then Exit Sub

Uni ip
Uni1 ip1
'(2^24)-2
result = ((ran1(0) - ran(0)) * (255 ^ 3)) + ((ran1(1) - ran(1)) * (255 ^ 2)) + ((ran1(2) - ran(2)) * 255) + (ran1(3) - ran(3))
If result < 0 Then MsgBox "ranges is misplaced", vbCritical, "ShareDat": Exit Sub


End Sub

Private Sub Uni1(ip As String)
On Error Resume Next
v = Split(ip, ".")
ReDim ran1(UBound(v))
For i = LBound(v) To UBound(v)
ran1(i) = v(i)
Next i

End Sub

