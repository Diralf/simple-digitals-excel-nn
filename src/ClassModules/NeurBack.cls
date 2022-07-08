VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NeurBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Weight As Double


Sub ChangeWeight()
    LastWeight = Weight
    Weight = Weight + ((Rnd - 0.5) * 0.01)
End Sub

Sub BackWeight()
    Weight = LastWeight
End Sub

Function IsWeightRight(Right As Double) As Boolean
    If Right = 1 Then
        IsWeightRight = LastWeight <= Weight
    ElseIf Right = 0 Then
        IsWeightRight = LastWeight >= Weight
    End If
End Function

Function Ask(Quest As Double) As Boolean
    Ask = Abs(Weight - Quest) < 0.5
End Function

Function Education(Quest As Double, Answer As Boolean)
    Dim ans
    ans = Ask(Quest)
    Range("B3").Value = ans
    If Not ans = Answer Then ChangeWeight
End Function

Private Sub Class_Initialize()
    Weight = Rnd
    LastWeight = Weight
End Sub

