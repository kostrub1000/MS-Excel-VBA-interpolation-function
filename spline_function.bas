Attribute VB_Name = "spline_function"


Public Function spline(x As Double, Xrange As Range, Yrange As Range)

Dim intervals As New Collection

Dim a As Integer, b As Integer, c As Integer, d As Integer

a = Xrange.Item(1).Row
b = Xrange.Item(Xrange.Count).Row
c = Xrange.Item(1).Column
d = Yrange.Item(1).Column

Set Xrange = Range(Cells(a, c), Cells(b, c))
Set Yrange = Range(Cells(a, d), Cells(b, d))



Dim n As Integer
n = Xrange.Count


For i = 1 To n - 1
    If x >= Xrange.Item(i) And x < Xrange.Item(i + 1) Then
        intervals.Add i
    ElseIf x <= Xrange.Item(i) And x > Xrange.Item(i + 1) Then
        intervals.Add i
    End If
Next i
If x = Xrange.Item(n) Then
    intervals.Add n - 1
End If

Dim t_coll As Collection
Dim y_coll As New Collection
Dim t As Double
Dim y As Double

If n >= 3 Then
    control_points_massive = Get_control_points(Xrange, Yrange)
    
    For i = 1 To intervals.Count
        Set t_coll = get_roots_of_cubic_polinom(x, Xrange.Item(intervals(i)), (control_points_massive(intervals(i) - 1, 0)), (control_points_massive(intervals(i) - 1, 2)), Xrange.Item(intervals(i) + 1))
        For j = 1 To t_coll.Count
            t = t_coll(j)
            y = (1 - t) ^ 3 * Yrange.Item(intervals(i)) + 3 * (1 - t) ^ 2 * t * (control_points_massive(intervals(i) - 1, 1)) + 3 * (1 - t) * (t ^ (2)) * (control_points_massive(intervals(i) - 1, 3)) + (t ^ (3)) * Yrange.Item(intervals(i) + 1)
            y_coll.Add (y)
        Next j
    Next i
Else
    Debug.Print "Function works with minimum three points"
End If

Dim results() As Double
ReDim results(y_coll.Count - 1)
For i = 1 To y_coll.Count
    results(i - 1) = y_coll(i)
Next i

spline = results

End Function



Private Function GetFirstControlPoints(rhs() As Double) As Double()

    Dim n As Integer
    n = UBound(rhs)
    Dim x() As Double
    Dim tmp() As Double
    ReDim x(n)
    ReDim tmp(n)
    
    
    Dim b As Double
    b = 2
    
    x(0) = rhs(0) / b
    
    For i = 1 To n
        tmp(i) = 1 / b
        If i < n Then
            b = 4 - tmp(i)
        Else
            b = 3.5 - tmp(i)
        End If
        x(i) = (rhs(i) - x(i - 1)) / b
    Next i
    For i = 0 To n - 1
        x(n - i - 1) = x(n - i - 1) - tmp(n - i) * x(n - i)
    Next i
    GetFirstControlPoints = x
    
    

End Function



Private Function Get_control_points(Xrange As Range, Yrange As Range) As Double()
    
    
    
    Dim n As Integer
    n = Xrange.Count - 2
    
    Dim rhs() As Double
    Dim x_1() As Double
    Dim y_1() As Double
    Dim x_2() As Double
    Dim y_2() As Double
    ReDim rhs(n)
    ReDim x_1(n)
    ReDim y_1(n)
    ReDim x_2(n)
    ReDim y_2(n)
    
    If n = 0 Then
        x_1(0) = ((2 * Xrange.Item(1).Value + Xrange.Item(2).Value) / 3)
        y_1(0) = ((2 * Yrange.Item(1).Value + Yrange.Item(2).Value) / 3)
        x_2(0) = (2 * x_1(0) - Xrange.Item(1).Value)
        y_2(0) = (2 * y_1(0) - Yrange.Item(1).Value)
    Else
        
        rhs(0) = Xrange.Item(1).Value + 2 * Xrange.Item(2).Value
        For i = 1 To n - 1
            rhs(i) = 4 * Xrange.Item(i + 1).Value + 2 * Xrange.Item(i + 2).Value
        Next i
        rhs(n) = (8 * Xrange.Item(n + 1).Value + Xrange.Item(n + 2).Value) / 2
        
        x_1 = GetFirstControlPoints(rhs)
        
        
        rhs(0) = Yrange.Item(1).Value + 2 * Yrange.Item(2).Value
        For i = 1 To n - 1
            rhs(i) = 4 * Yrange.Item(i + 1).Value + 2 * Yrange.Item(i + 2).Value
        Next i
        rhs(n) = (8 * Yrange.Item(n + 1).Value + Yrange.Item(n + 2).Value) / 2
        y_1 = GetFirstControlPoints(rhs)
        
        For i = 0 To n
            If i < n Then
                x_2(i) = 2 * Xrange.Item(i + 2) - x_1(i + 1)
                y_2(i) = 2 * Yrange.Item(i + 2) - y_1(i + 1)
            Else
                x_2(i) = (Xrange.Item(Xrange.Count) + x_1(n)) / 2
                y_2(i) = (Yrange.Item(Yrange.Count) + y_1(n)) / 2
            End If
        Next i
        
    End If
    
    Dim control_points_massive() As Double
    ReDim control_points_massive(n, 3)
    For i = 0 To UBound(x_1)
        control_points_massive(i, 0) = x_1(i)
        control_points_massive(i, 1) = y_1(i)
        control_points_massive(i, 2) = x_2(i)
        control_points_massive(i, 3) = y_2(i)
    Next i
    
    



    Get_control_points = control_points_massive

End Function

Private Function get_roots_of_cubic_polinom(x As Double, x0 As Double, x1 As Double, x2 As Double, x3 As Double) As Collection

Dim roots As New Collection

If x = x0 Then
    roots.Add 0
ElseIf x = x3 Then
    roots.Add 1
Else
    Dim pi As Double
    pi = 3.1415926535
    
    
    Dim a As Double
    Dim b As Double
    Dim c As Double
    Dim d As Double
    
    a = x3 - 3 * x2 + 3 * x1 - x0
    b = 3 * x2 - 6 * x1 + 3 * x0
    c = 3 * x1 - 3 * x0
    d = x0 - x
    
    If a <> 0 Then
        Dim p As Double
        Dim qm As Double
        
        
        p = (3 * a * c - b * b) / (3 * a * a)
        qm = (2 * b * b * b - 9 * a * b * c + 27 * a * a * d) / (27 * a * a * a)
        
        Dim Q As Double
        Q = (p / 3) ^ 3 + (qm / 2) ^ 2
        
        
        Dim y1 As Double, y2 As Double, y3 As Double
        
        Dim phi As Double
    
        If Q < 0 Then
            If qm < 0 Then
                phi = Atn((-Q) ^ 0.5 / (-qm / 2))
            ElseIf qm > 0 Then
                phi = Atn((-Q) ^ 0.5 / (-qm / 2)) + pi
            Else
                phi = pi / 2
            End If
            
            y1 = 2 * (-p / 3) ^ 0.5 * Cos(phi / 3)
            y2 = 2 * (-p / 3) ^ 0.5 * Cos(phi / 3 + 2 * pi / 3)
            y3 = 2 * (-p / 3) ^ 0.5 * Cos(phi / 3 + 4 * pi / 3)
            
            
            If (y1 - b / (3 * a)) >= 0 And (y1 - b / (3 * a)) <= 1 Then
                roots.Add y1 - b / (3 * a)
            End If
            
            
            If (y2 - b / (3 * a)) >= 0 And (y2 - b / (3 * a)) <= 1 Then
                roots.Add y2 - b / (3 * a)
            End If
            
            
            If (y3 - b / (3 * a)) >= 0 And (y3 - b / (3 * a)) <= 1 Then
                roots.Add y3 - b / (3 * a)
            End If
        
            
        ElseIf Q = 0 Then
            y1 = 2 * (-Q / 2) ^ (1 / 3)
            y2 = -(-Q / 2) ^ (1 / 3)
            
            If (y1 - b / (3 * a)) >= 0 And (y1 - b / (3 * a)) <= 1 Then
                roots.Add y1 - b / (3 * a)
            End If
            
            If (y2 - b / (3 * a)) >= 0 And (y2 - b / (3 * a)) <= 1 Then
                roots.Add y2 - b / (3 * a)
            End If
            
        Else
            If (-qm / 2 + (Q) ^ 0.5) >= 0 And (-qm / 2 - (Q) ^ 0.5) >= 0 Then
                roots.Add (-qm / 2 + (Q) ^ 0.5) ^ (1 / 3) + (-qm / 2 - (Q) ^ 0.5) ^ (1 / 3) - b / (3 * a)
            ElseIf (-qm / 2 + (Q) ^ 0.5) >= 0 And (-qm / 2 - (Q) ^ 0.5) < 0 Then
                roots.Add (-qm / 2 + (Q) ^ 0.5) ^ (1 / 3) - Abs(-qm / 2 - (Q) ^ 0.5) ^ (1 / 3) - b / (3 * a)
            ElseIf (-qm / 2 + (Q) ^ 0.5) < 0 And (-qm / 2 - (Q) ^ 0.5) >= 0 Then
                roots.Add -Abs(-qm / 2 + (Q) ^ 0.5) ^ (1 / 3) + (-qm / 2 - (Q) ^ 0.5) ^ (1 / 3) - b / (3 * a)
            Else
                roots.Add -Abs(-qm / 2 + (Q) ^ 0.5) ^ (1 / 3) - Abs(-qm / 2 - (Q) ^ 0.5) ^ (1 / 3) - b / (3 * a)
            End If
        End If
    ElseIf a = 0 And b = 0 And c <> 0 Then
        roots.Add (-d / c)
    ElseIf a = 0 And b <> 0 Then
        Dim Descr As Double
        Descr = c ^ 2 - 4 * b * d
        Dim root_1 As Double
        Dim root_2 As Double
        root_1 = (-c - (Descr) ^ 0.5) / 2 * b
        root_2 = (-c + (Descr) ^ 0.5) / 2 * b
        If root_1 >= 0 Or root_1 <= 1 Then
            roots.Add root_1
        End If
        
        If root_2 >= 0 Or root_2 <= 1 Then
            roots.Add root_2
        End If
        
    End If
End If

Set get_roots_of_cubic_polinom = roots


End Function



