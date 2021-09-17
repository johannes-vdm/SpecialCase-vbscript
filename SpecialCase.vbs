Function SpecialCase(sText)
 
    Dim a, iLen, bSpace, tmpX, tmpFull,SrchText
    SrchText = " -'/\&"
 
    iLen = Len(sText)
    For a = 1 To iLen
    If a <> 1 Then 'just to make sure 1st character is upper and the rest lower'
        If bSpace = True Then
            tmpX = UCase(mid(sText,a,1))
            bSpace = False
        Else
        tmpX=LCase(mid(sText,a,1))
            If Instr(1,SrchText,tmpX) > 0 Then
                       ' If tmpX = " " Or tmpX = "'" Or tmpx = "&" Then 
              bSpace = True
            End If  
        End if
        
    Else
        tmpX = UCase(mid(sText,a,1))
    End if
    tmpFull = tmpFull & tmpX
    Next
    SpecialCase = tmpFull
End Function


MsgBox SpecialCase("JOHN O'CONNEr Smith-JONES")
