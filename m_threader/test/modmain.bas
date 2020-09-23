Attribute VB_Name = "modmain"

Sub loopthousand()
Dim i As Double
For i = 0 To 10000000
'
DoEvents
'Debug.Print i
frmtest.Text1 = i

Next

End Sub

Sub loop100()
Dim i As Double, j As Double
For i = 0 To 10000000
'
DoEvents
'For j = 0 To 1000
'Debug.Print i & "  " & j
frmtest.Text2 = i ' & "  " & j
'Next

Next

End Sub

Sub loop1000()
Dim i As Double ', j As Double
For i = 0 To 1000000
'
DoEvents
'For j = 0 To 1000
'Debug.Print i & "  " & j
frmtest.Text3 = i ' & "  " & j
'Next

Next

End Sub

Sub startnew()
MsgBox ""
End Sub
