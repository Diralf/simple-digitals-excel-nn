Attribute VB_Name = "Module1"
Dim Neur(150) As Neuron
Dim step As Integer
 
Function IsOXYValid(x As Integer, y As Integer) As Boolean
    IsOXYValid = x >= 0 And y >= 0
End Function

' need
Function SetValToCell(StartCell As String, x As Integer, y As Integer, newValue)
    If Not IsOXYValid(x, y) Then Exit Function
    Range(StartCell).Cells(y + 1, x + 1).Value = newValue
End Function

Function GetValFromCell(StartCell As String, x As Integer, y As Integer)
    If Not IsOXYValid(x, y) Then Exit Function
    GetValFromCell = Range(StartCell).Cells(y + 1, x + 1).Value
End Function

Sub StartLearn()
Attribute StartLearn.VB_ProcData.VB_Invoke_Func = "l\n14"
    For I = 0 To UBound(Neur()) - 1
        Dim x As Integer, y As Integer
        x = (I Mod 3) + 3 * Fix(I / 15)
        y = Fix(I / 3) Mod 5
        Set Neur(I) = New Neuron
        SetValToCell "C8", x, y, Neur(I).Weight
    Next I
    step = 0
End Sub

Function SeeNumber(StartNumber As Integer)
    For J = 0 To 9
        Dim SumW, AlignSum
        SumW = 0
        For I = 0 To 14
            Dim x As Integer, y As Integer
            x = (I Mod 3) '+ StartNumber * 3
            y = Fix(I / 3)
            SumW = SumW + Neur(I + J * 15).Ask(Int(GetValFromCell("N2", x, y)))
        Next I
        AlignSum = SumW / 15
        SetValToCell "C20", 1 + J * 3, 0, AlignSum
    Next J
End Function

Function Education(StartNumber As Integer)
    For J = 0 To 9
        Dim TrueNumber As Integer
        TrueNumber = J + 10 * Fix(StartNumber / 10)
        For I = 0 To 14
            Dim x As Integer, y As Integer, StartNumberCell, TrueNumberCell
            x = (I Mod 3)
            y = Fix(I / 3)
            StartNumberCell = Int(GetValFromCell("AG2", x + StartNumber * 3, y))
            TrueNumberCell = Sgn(Neur(I + J * 15).Weight) 'Int(GetValFromCell("AG2", x + TrueNumber * 3, y))
            
            
            
            If StartNumber = TrueNumber Then
                Neur(I + J * 15).Correct StartNumberCell * 1
            ElseIf StartNumberCell = TrueNumberCell Then
            Else
                'Neur(I + J * 15).Correct TrueNumberCell
            End If
            
            
            SetValToCell "C8", I Mod 3 + 3 * J, Fix(I / 3), Neur(I + J * 15).Weight
        Next I
    Next J
End Function

Sub OnceLearn()
Attribute OnceLearn.VB_ProcData.VB_Invoke_Func = "o\n14"
    SeeNumber Int(GetValFromCell("B4", 0, 0))
    'Education I Mod 10, 1
End Sub

Sub StepLearn()
Attribute StepLearn.VB_ProcData.VB_Invoke_Func = "p\n14"
    Education step Mod 40
    step = step + 1
End Sub

Sub UpdateLearn()
Attribute UpdateLearn.VB_ProcData.VB_Invoke_Func = "u\n14"
    Dim TrueNumber As Integer, StartNumber As Integer
    TrueNumber = 5

    For I = 0 To 39
        Dim Answer
        StartNumber = I ' Fix(Rnd * 40) ' I Mod 10
        'Answer = SeeNumber(StartNumber)
        Education StartNumber
    Next I
End Sub
