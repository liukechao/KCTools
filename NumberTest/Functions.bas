Attribute VB_Name = "Functions"
Function sum(a As String, b As String) As String
    Dim i As Integer
    Dim len_a As Integer, len_b As Integer, len_s As Integer
    Dim array_a() As Integer, array_b() As Integer, array_s() As Integer
    
    len_a = Len(a)
    len_b = Len(b)
    len_s = IIf(len_a > len_b, len_a, len_b) + 1
    
    ReDim array_a(len_s) As Integer, array_b(len_s) As Integer, array_s(len_s) As Integer
    
    For i = 1 To len_a
        array_a(i) = Val(Mid(StrReverse(a), i, 1))
    Next i
    
    For i = 1 To len_b
        array_b(i) = Val(Mid(StrReverse(b), i, 1))
    Next i
    
    For i = 1 To len_s
        array_s(i) = array_a(i) + array_b(i)
    Next i
    
    For i = 1 To len_s
        If array_s(i) >= 10 Then
            array_s(i + 1) = array_s(i + 1) + array_s(i) \ 10
            array_s(i) = array_s(i) Mod 10
        End If
    Next i
    
    For i = len_s To 1 Step -1
        sum = sum & IIf(array_s(i) = 0, " ", array_s(i))
    Next i
    
    sum = Replace(LTrim(sum), " ", 0)
    
    If sum = "" Then
        sum = "0"
    End If
    
End Function

