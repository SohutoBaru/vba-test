'A
Attribute VB_Name = "Module1"
Option Explicit
Private tryCnt As Long
Private side As Integer

Sub main()
    side = 9 
    Debug.Print Timer 
    Dim SuAry(1 To 9, 1 To 9) As Integer 
    Dim i1 As Integer 
    Dim i2 As Integer 
    
    tryCnt = 0
    Erase SuAry 
    For i1 = 1 To side
        For i2 = 1 To side
            If Cells(i1, i2) = "" Then
                Cells(i1, i2).Font.Color = vbBlue 
            Else
                SuAry(i1, i2) = Cells(i1, i2) 
            End If
        Next
    Next
    
    Call trySu(SuAry) 
    
    Range("A1:I9").Value = SuAry
    Debug.Print Timer
    
    If getBlank(SuAry(), i1, i2) = False Then
        MsgBox "�񓚐����B"
    Else
        MsgBox "�񓚕s�\�B"
    End If
End Sub
Function trySu(ByRef SuAry() As Integer) As Boolean
    Dim i1 As Integer
    Dim i2 As Integer
    Dim su As Integer
    If getBlank(SuAry(), i1, i2) = False Then
        trySu = True
        Exit Function
    End If
    For su = 1 To 9 
        If chkSu(SuAry(), i1, i2, su) = True Then
            SuAry(i1, i2) = su 
            tryCnt = tryCnt + 1
            Cells(i1, i2) = su 
            If trySu(SuAry) = True Then
                trySu = True
                Exit Function
            End If
            
        End If
    Next
    
    SuAry(i1, i2) = 0 
    Cells(i1, i2) = ""
    DoEvents
    trySu = False

End Function
Function getBlank(ByRef SuAry() As Integer, ByRef i1 As Integer, ByRef i2 As Integer) As Boolean
    For i1 = 1 To side
        For i2 = 1 To side
            If SuAry(i1, i2) = 0 Then
                getBlank = True
                Exit Function
            End If
        Next
    Next
    getBlank = False
End Function

Function chkSu(ByRef SuAry() As Integer, ByVal i1 As Integer, ByVal i2 As Integer, _
                                                        ByVal su As Integer) As Boolean
    Dim ix1 As Integer
    Dim ix2 As Integer
    Dim i1s As Integer
    Dim i2s As Integer
    chkSu = False
    
    For ix2 = 1 To side
        If ix2 <> i2 Then
            If SuAry(i1, ix2) = su Then
                chkSu = False
                Exit Function
            End If
        End If
    Next
    
    For ix1 = 1 To side
        If ix1 <> i1 Then
            If SuAry(ix1, i2) = su Then
                chkSu = False
                Exit Function
            End If
        End If
    Next
    

    i1s = (Int((i1 + 2) / 3) - 1) * 3 + 1 
    i2s = (Int((i2 + 2) / 3) - 1) * 3 + 1
    For ix1 = i1s To i1s + 2
        For ix2 = i2s To i2s + 2
            If SuAry(ix1, ix2) = su Then
                chkSu = False
                Exit Function
            End If
        Next
    Next
    chkSu = True
End Function















