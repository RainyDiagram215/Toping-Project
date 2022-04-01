Attribute VB_Name = "CodeMod"
'Option Explicit

Function RunCode(ByVal Code As String)
    On Error Resume Next
    Dim CodeStr, Str As String
    Tmp1 = Split(Code, "(")
    CodeStr = Tmp1(0)
    For i = 1 To UBound(Tmp1)
        Str = Str + Tmp1(i)
    Next i
    OStr = Str
    Str = Left(Str, Len(Str) - 1)
    TStr = Right(OStr, 1)
    If TStr = ")" Then
        '大家就当无事发生
    Else
        MsgBox "Error"
        Exit Function
    End If
    Select Case CodeStr
        Case "InfoBox"
            Q1 = Right(Str, 1)
            Q2 = Left(Str, 1)
            If Q1 = """" And Q2 = """" Then
                MsgBox Replace(Str, """", "")
            Else
                Open App.Path + "\Var\" + Replace(Str, "$", "") + ".var" For Input As #1
                    Line Input #1, a
                    MsgBox a
                Close #1
            End If
        Case "Var"
            Q1 = Right(Str, 1)
            Q2 = Left(Str, 1)
            If Q1 = """" And Q2 = """" Then
                tmpvar = Split(Str, ",")
                VarName = tmpvar(0)
                VarName = Replace(VarName, """", "")
                For i = 1 To UBound(tmpvar)
                    VarData = VarData + tmpvar(i)
                Next i
                VarData = Replace(VarData, """", "")
                Open App.Path + "\Var\" + VarName + ".var" For Output As #1
  
                    Print #1, VarData
                Close #1
            End If
    End Select
End Function
