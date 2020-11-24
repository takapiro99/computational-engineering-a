Sub calculate()
    Worksheets("Sheet1").Activate
    E = Cells(7, 2)
    r = Cells(9, 2)
    Nodes = 5
    Main = create_matrix(Nodes * 2, Nodes * 2 + 1, 0)
    For i = 0 To Nodes * 2 - 1
        Main(i)(UBound(Main(0))) = Cells(2 + i, 11)
    Next i
    ' Call print_array(Main, "prim")
    Dim cordinates(4)
    For i = 0 To 4
        cordinates(i) = Array(Cells(2 + i, 2), Cells(2 + i, 3))
    Next i
    sections = Array(Array(1, 2), Array(3, 4), Array(2, 4), Array(3, 2), Array(2, 5), Array(4, 5)) ' sinの正負
    Dim k_(5) As Variant
    For i = 0 To 5                               ' create k
        k_(i) = k_factory(cordinates(sections(i)(0) - 1), cordinates(sections(i)(1) - 1), E, r)
        ' Call print_array(k_(i), i & "番目")
    Next i
    
    ' k 代入していく、今回はuとv分かれてなくてu1v1u2v2って並んでるから詰まった
    yoko = 1
    step_ = 1
    For i = 0 To 5
        For j = 0 To 1
            For k = 0 To 1
                row = (sections(i)(j) - 1) * 2
                col = (sections(i)(k) - 1) * 2
                ' MsgBox (row & "   " & col)
                Main(row)(col) = Main(row)(col) + k_(i)(2 * j)(2 * k)
                Main(row + yoko)(col) = Main(row + yoko)(col) + k_(i)(2 * j + step_)(2 * k)
                Main(row)(col + yoko) = Main(row)(col + yoko) + k_(i)(2 * j)(2 * k + step_)
                Main(row + yoko)(col + yoko) = Main(row + yoko)(col + yoko) + k_(i)(2 * j + step_)(2 * k + step_)
            Next k
        Next j
    Next i
    ' はきだす
    big = shrink_array(Main)                     ' just deep copy
    f = Array(Cells(5, 9) - 1, Cells(4, 9) - 1, Cells(3, 9) - 1, Cells(2, 9) - 1)
    'Call print_array(Main, "掃き出前")
    Main = shrink_array(shrink_array(shrink_array(shrink_array(Main, f(0), f(0)), f(1), f(1)), f(2), f(2)), f(3), f(3))
    'Call print_array(Main, "ぎゅっとした後掃き出前")
    Call forward_elimination(Main)
    Call backward_substitution(Main)
    'Call print_array(Main, "掃き出し後")
    ' F_output出す
    Cells(2, 12) = 0
    Cells(3, 12) = 0
    Cells(4, 12) = Main(0)(UBound(Main(0)))
    Cells(5, 12) = Main(1)(UBound(Main(0)))
    Cells(6, 12) = 0
    Cells(7, 12) = 0
    Cells(8, 12) = Main(2)(UBound(Main(0)))
    Cells(9, 12) = Main(3)(UBound(Main(0)))
    Cells(10, 12) = Main(4)(UBound(Main(0)))
    Cells(11, 12) = Main(5)(UBound(Main(0)))
    For i = 0 To 9
      num = 0
      For j = 0 To 9
        num = num + big(i)(j) * Cells(2 + j, 12)
      Next j
      Cells(2 + i, 13) = Round(num, 3)
    Next i
End Sub

Function k_factory(first, second, E, r)      ' takes 2 cordinates
    diagonal = Sqr((second(0) - first(0)) ^ 2 + (second(1) - first(1)) ^ 2)
    cos_ = (second(0) - first(0)) / diagonal
    sin_ = (second(1) - first(1)) / diagonal
    ' MsgBox (Round(sin_, 5) & "    " & Round(cos_, 5))
    new_k = k_template_factory(sin_, cos_)
    Pi = 4 * Atn(1)
    new_k = multiply_each(new_k, E * Pi * r ^ 2 / diagonal) ' multiply by ea/l
    'print_array (new_k)
    k_factory = new_k
End Function

Function k_template_factory(sin_, cos_)
    k = create_matrix(4, 4, 0)
    k(0)(0) = cos_ ^ 2
    k(0)(1) = cos_ * sin_
    k(0)(2) = -1 * cos_ ^ 2
    k(0)(3) = -1 * cos_ * sin_
    k(1)(0) = cos_ * sin_
    k(1)(1) = sin_ ^ 2
    k(1)(2) = -1 * cos_ * sin_
    k(1)(3) = -1 * sin_ ^ 2
    k(2)(0) = -1 * cos_ ^ 2
    k(2)(1) = -1 * cos_ * sin_
    k(2)(2) = cos_ ^ 2
    k(2)(3) = cos_ * sin_
    k(3)(0) = -1 * cos_ * sin_
    k(3)(1) = -1 * sin_ ^ 2
    k(3)(2) = cos_ * sin_
    k(3)(3) = sin_ ^ 2
    k_template_factory = k
End Function

Function shrink_array(arr_, Optional row = -1, Optional col = -1)
    arr = arr_
    new_array_row = UBound(arr)
    new_array_col = UBound(arr(0))
    ' row詰める
    If row >= 0 Then
        new_array_row = new_array_row - 1
        For i = 0 To UBound(arr(0))
            For j = row To UBound(arr) - 1
                arr(j)(i) = arr(j + 1)(i)
            Next j
        Next i
    End If
    ' col 詰める
    If col >= 0 Then
        new_array_col = new_array_col - 1
        For i = 0 To UBound(arr)
            For j = col To UBound(arr(0)) - 1
                arr(i)(j) = arr(i)(j + 1)
            Next j
        Next i
    End If
    new_array = create_matrix(new_array_row + 1, new_array_col + 1)
    For i = 0 To new_array_row
        For j = 0 To new_array_col
            new_array(i)(j) = arr(i)(j)
        Next j
    Next i
    shrink_array = new_array
End Function

Function multiply_each(jug_array, number)
    new_array = create_matrix(UBound(jug_array) + 1, UBound(jug_array(0)) + 1)
    For i = 0 To UBound(jug_array)
        For j = 0 To UBound(jug_array(0))
            new_array(i)(j) = jug_array(i)(j) * number
        Next j
    Next i
    multiply_each = new_array
End Function

' Create a matrix of arbitrary size (jug array)
Function create_matrix(row_size, col_size, Optional initial = "")
    Dim ans, row As Variant
    ans = Array()
    ReDim ans(row_size - 1)
    For i = 0 To row_size - 1
        row = Array()                            ' new array Instance
        ReDim row(col_size - 1)
        For j = 0 To col_size - 1
            row(j) = initial
        Next j
        ans(i) = row
    Next i
    create_matrix = ans
End Function

Sub print_array(arr, Optional msg As String)
    Debug.Print ("--")
    For i = 0 To UBound(arr)
        tmp = ""
        For j = 0 To UBound(arr(0))
            'tmp = tmp & arr(i)(j) & "  "
            tmp = tmp & Round(arr(i)(j), 4) & "  "
        Next j
        Debug.Print (tmp)
    Next i
    Debug.Print ("--- " & Now & " " & msg & " ---")
End Sub

Sub forward_elimination(arr)
    For i = 0 To UBound(arr)
        Key = arr(i)(i)
        For j = 0 To UBound(arr(0)) - i
            ' divide each element with key
            arr(i)(j + i) = arr(i)(j + i) / Key
        Next j
        For k = i + 1 To UBound(arr)
            num = arr(k)(i)
            For L = i To UBound(arr(0))
                arr(k)(L) = arr(k)(L) - arr(i)(L) * num
            Next L
        Next k
    Next i
End Sub

Sub backward_substitution(arr)
    For i = UBound(arr) To 1 Step -1
        For j = i To 1 Step -1
            Key = arr(j - 1)(i)
            For k = UBound(arr(0)) To i Step -1
                arr(j - 1)(k) = arr(j - 1)(k) - arr(i)(k) * Key
            Next k
        Next j
    Next i
End Sub

' reset area
Sub reset_area()
    Dim SheetName As String
    SheetName = "Sheet1"
    Sheets(SheetName).Range("L4", "L5").ClearContents
    Sheets(SheetName).Range("L8", "L12").ClearContents
    Sheets(SheetName).Range("M2", "M12").ClearContents
End Sub

