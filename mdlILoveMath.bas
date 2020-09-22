Attribute VB_Name = "mdlILoveMath"
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public oTombol As Integer
Public oBatasan As Double
Public oPos As Integer


Public Operator As String

Public Function oPerhitungan(oCont1 As Label, oCont2 As Label, oCont3 As Label, oCont4 As Label) 'Random the question
    Dim oAcak As Integer
    Dim oVar1 As Double
    Dim oVar2 As Double
    Dim Hasil As Double
    Dim oTmp(3) As Double
           
    Randomize
    
    oVar1 = CInt(Rnd * 10)
    oVar2 = CInt(Rnd * 10)
    
    If oVar1 = 10 Or oVar2 = 10 Then
        oVar1 = Abs(oVar1 - CInt(Rnd * 10))
        oVar2 = Abs(oVar2 - CInt(Rnd * 10))
    End If
    
    Hasil = oVar1 * oVar2
    
    Do Until oAcak > 0 And oAcak < 5
        oAcak = CInt(Rnd * 10)
    Loop
    
    Select Case oAcak
        Case Is = 1
            oTmp(0) = Hasil
            
            oTmp(1) = Abs(Hasil - (CInt(Rnd * 10)))
            oTmp(2) = Abs(Hasil - (CInt(Rnd * 10)))
            oTmp(3) = Abs(Hasil - (CInt(Rnd * 10)))
        Case Is = 2
            oTmp(1) = Hasil
            
            oTmp(0) = Abs(Hasil - (CInt(Rnd * 10)))
            oTmp(2) = Abs(Hasil - (CInt(Rnd * 10)))
            oTmp(3) = Abs(Hasil - (CInt(Rnd * 10)))
        Case Is = 3
            oTmp(2) = Hasil
            
            oTmp(0) = Abs(Hasil - (CInt(Rnd * 10)))
            oTmp(1) = Abs(Hasil - (CInt(Rnd * 10)))
            oTmp(3) = Abs(Hasil - (CInt(Rnd * 10)))
        Case Is = 4
            oTmp(3) = Hasil
            
            oTmp(0) = Abs(Hasil - (CInt(Rnd * 10)))
            oTmp(1) = Abs(Hasil - (CInt(Rnd * 10)))
            oTmp(2) = Abs(Hasil - (CInt(Rnd * 10)))
    End Select
            
    For Q = 0 To 3
        For Z = 0 To 3
            If Not Q = Z Then
                If oTmp(Q) = oTmp(Z) Then
                    oTmp(Q) = oTmp(Q) + CInt(Rnd * 10)
                End If
            End If
        Next Z
    Next Q
                        
    oPos = oAcak
    oCont1 = oTmp(0)
    oCont2 = oTmp(1)
    oCont3 = oTmp(2)
    oCont4 = oTmp(3)
    oPerhitungan = Str(oVar1) & " " & "X" & Str(oVar2)
    
End Function



