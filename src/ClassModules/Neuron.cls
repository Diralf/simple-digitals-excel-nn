VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Neuron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Weight As Double

Private Sub Class_Initialize()
    Weight = 0 '(Rnd * 2 - 1) * 5
End Sub

Function Ask(Quest)
    Ask = Weight * Quest
End Function

Sub Correct(IsTrAnswer)
    Weight = Weight + 1 * IsTrAnswer
End Sub
