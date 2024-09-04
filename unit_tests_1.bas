Attribute VB_Name = "unit_test"
Function test_PriceBond()
    y = 0.03
    face = 2000000
    couponRate = 0.04
    m = 10
    ppy = 2

    ' Week 1    
    x0 = PriceBond(y, face, couponRate, m)
    'x1 = PriceBond(y, face, couponRate, m, 1)
    'x2 = PriceBond(y, face, couponRate, m, ppy)
    
'    If Round(x2, 0) = 2171686 And Round(x1) = 2170604 And Round(x0, 0) = 2170604 Then
    If Round(x0, 0) = 2170604 Then
        result = "PASS"
    Else
        result = "FAIL"
    End If
        
    test_PriceBond = result
End Function

Function test_getBondPrice()
    y = 0.03
    face = 2000000
    couponRate = 0.04
    m = 10
    ppy = 2

    If Round(getBondPrice(y, face, couponRate, m, 1), 0) = 2170604 And Round(getBondPrice(y, face, couponRate, m, 2), 0) = 2171686 Then
        result = "PASS"
    Else
        result = "FAIL"
    End If
        
    test_getBondPrice = result
End Function

Function test_getBondDuration()
    y = 0.03
    face = 2000000
    couponRate = 0.04
    m = 10

    If Round(getBondDuration(y, face, couponRate, m), 2) = 8.51 Then
        result = "PASS"
    Else
        result = "FAIL"
    End If
        
    test_getBondDuration = result
End Function

Function test_BondDuration(Optional showMsg As Boolean = True)
    y = 0.03
    face = 2000000
    couponRate = 0.04
    m = 10

    If Round(BondDuration(y, face, couponRate, m), 2) = 8.51 Then
        result = "PASS"
    Else
        result = "FAIL"
    End If
    
    If showMsg Then
        MsgBox (result)
    End If
    
    test_BondDuration = result
End Function

Function test_MyMatMult()
    'Rnd is a random draw from 0 to 1
    m = Round(Rnd * 3, 0) + 1
    n = Round(Rnd * 3, 0) + 1
    
    m = 2
    n = 2
    ReDim mat(m, n) As Integer
    ReDim vec(m) As Integer
    
    For i = 0 To m
        For j = 0 To n
            mat(i, j) = i * 3 + j + 1
        Next j
    Next i
    
    For i = 0 To m
        vec(i) = i + 1
    Next i
    
    x = MyMatMult(vec, mat)
    
    nRand = Application.Min(Round(Rnd * 3, 0) + 1, n)
    
    Dim randAns As Double
    For i = 0 To m
        randAns = randAns + vec(i) * mat(i, nRand)
    Next i
    
    If x(nRand) = randAns Then
        result = "PASS"
    Else
        result = "FAIL"
    End If
        
    test_MyMatMult = result
End Function

Function test_MyTripDataObj()
    test_MyTripDataObj = "MANUAL REVIEW"
End Function

Function test_getPortfolioDuration()

    Dim myObj(1 To 3) As Variant
    
    Dim namevec(1 To 3)  As String
    namevec(1) = "Alice"
    namevec(2) = "Bob"
    namevec(3) = "Chuck"
    
    Dim valuevec(1 To 3) As Integer
    Dim durationvec(1 To 3) As Integer
    
    For i = 1 To 3
        valuevec(i) = i * 100
        durationvec(i) = i + 4
    Next i
    
    myObj(1) = valuevec
    myObj(2) = durationvec
    myObj(3) = namevec
    
    x1 = getPortfolioDuration(myObj)
    
    For i = 1 To 3
        valuevec(i) = i * 2
        durationvec(i) = i + 1
    Next i
    myObj(1) = valuevec
    myObj(2) = durationvec
    
    x2 = getPortfolioDuration(myObj)
    
    If Round(x1, 1) = 6.3 And Round(x2, 1) = 3.3 Then
        result = "PASS"
    Else
        result = "FAIL"
    End If
    
    test_getPortfolioDuration = result
End Function

Function test_getNamedRange()
    x = getNamedRange("xVector")
    If x(1, 1) = 10 Then
        result = "PASS"
    Else
        result = "FAIL"
    End If
    
    test_getNamedRange = result
End Function

Function test_FizzBuzz()
    Dim n As Integer
    n = Int(Round(Rnd * 3 + 1) * 15 + 1)
    x = FizzBuzz(n, n + 14)
    
    If IsEmpty(x) Then
        If showMsg Then
            MsgBox ("Your function did not return anything." & vbCrLf & "Assign your output value to the function name.")
        End If
        result = "FAIL"
        Exit Function
    End If
    
    Dim base As Integer
    base = LBound(x, 1)
    
    For i = base To (base + 14)
        If x(i) = "Fizz" Then x(i) = "fizz"
        If x(i) = "Buzz" Then x(i) = "buzz"
        If x(i) = "FizzBuzz" Then x(i) = "fizzbuzz"
        If x(i) = "Fizzbuzz" Then x(i) = "fizzbuzz"
    Next i

    If x(base + 2) = "fizz" And x(base + 14) = "fizzbuzz" Then
        result = "PASS"
    Else
        result = "FAIL"
    End If
    
    test_FizzBuzz = result
End Function

Function test_filter()
    Sheets("qOffice1_Filter").Range("G1").FormulaR1C1 = "=SUBTOTAL(9,C[-4])"
    x = Sheets("qOffice1_Filter").Range("G1").Value
    Sheets("qOffice1_Filter").Range("G1").ClearContents
    
    If Round(x, 0) = 4873 Then
        result = "PASS"
    Else
        result = "FAIL"
    End If
    
    selftest_filter = result
    
End Function

