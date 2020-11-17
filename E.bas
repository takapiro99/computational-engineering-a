' reset area
Sub reset_area()
    Dim SheetName As String
    SheetName = "Sheet1"
    Sheets(SheetName).Range("L2", "L13").ClearContents
    Sheets(SheetName).Range("M2", "M13").ClearContents
End Sub

Sub calculate()
    Worksheets("Sheet1").Activate
    E = Cells(7, 2)
    nu = Cells(8, 2)
    t = Cells(9, 2)
    Nodes = 6
    one = Array(Cells(2, 2), Cells(2, 3))
    three = Array(Cells(3, 2), Cells(3, 3))
    two = Array((one(0) + three(0)) / 2, (one(1) + three(1)) / 2)
    four = Array(Cells(4, 2), Cells(4, 3))
    six = Array(Cells(5, 2), Cells(5, 3))
    five = Array((four(0) + six(0)) / 2, (four(1) + six(1)) / 2)
    cordinates = Array(one, two, three, four, five, six) ' 1~6の座標
    sections = Array(Array(1, 2, 5), Array(2, 3, 6), Array(1, 5, 4), Array(2, 6, 5))
    'Call print_array(sections, "sections")
    'Call print_array(cordinates, "c")
    coefficient = E * t / 4 / (1 - nu * nu)
    nu_stress = Array(Array(1, nu, 0), Array(nu, 1, 0), Array(0, 0, (1 - nu) / 2))
    ' 全体は12×13ですね
    main = create_matrix(12, 13, 0)
    For i = 0 To UBound(main)
        main(i)(UBound(main(0))) = Cells(2 + i, 11)
    Next i
        
    Dim k_() As Variant
    ReDim k_(4)                                  ' magic number
    ' それぞれのkを作ります
    ' それぞれのA(面積)
    ' 1/2A はBの2回分
    ' e/1-v ^2はDから。
    ' そこから Bt D B をそれぞれ用意。βとか使って
    For i = 0 To 3                               ' k作るために4周
        Dim tmp()
        x1 = cordinates(sections(i)(0) - 1)(0)
        y1 = cordinates(sections(i)(0) - 1)(1)
        x2 = cordinates(sections(i)(1) - 1)(0)
        y2 = cordinates(sections(i)(1) - 1)(1)
        x3 = cordinates(sections(i)(2) - 1)(0)
        y3 = cordinates(sections(i)(2) - 1)(1)
        A = (x1 * y2 + x2 * y3 + x3 * y1 - x1 * y3 - x2 * y1 - x3 * y2) / 2
        B1 = Array(y2 - y3, y3 - y1, y1 - y2, 0, 0, 0)
        B2 = Array(0, 0, 0, x3 - x2, x1 - x3, x2 - x1)
        B3 = Array(x3 - x2, x1 - x3, x2 - x1, y2 - y3, y3 - y1, y1 - y2)
        B = Array(B1, B2, B3)
        Bt = matrix_t(B)
        tmp = matrix_cross(matrix_cross(Bt, nu_stress), B)
        tmp = multiply_each(tmp, 1 / A)
        tmp = multiply_each(tmp, coefficient)
        'Call print_array(tmp, "k_" & i)
        k_(i) = tmp
    Next i
    'Call print_array(main, "hello")
    ' それを全体係数行列にいれます
    For i = 0 To 3                               'ubound(k)
        For j = 0 To 2                           ' 3行 * 分
            For k = 0 To 2                       ' 一つずつ × 2
                'Debug.Print (row & col & " / " & row + 6 & col + 6)
                row = sections(i)(j) - 1
                col = sections(i)(k) - 1
                main(row)(col) = main(row)(col) + k_(i)(j)(k)
                main(row + 6)(col) = main(row + 6)(col) + k_(i)(j + 3)(k)
                main(row)(col + 6) = main(row)(col + 6) + k_(i)(j)(k + 3)
                main(row + 6)(col + 6) = main(row + 6)(col + 6) + k_(i)(j + 3)(k + 3)
            Next k
        Next j
    Next i
    ' debug
    For i = 0 To UBound(main)
        For j = 0 To UBound(main(0))
            'Cells(j + 20)(i + 20) = main(i)(j)
        Next j
    Next i

    Call print_array(main, "hello")
    big = shrink_array(main)
    ' 掃きだします
    main = shrink_array(shrink_array(shrink_array(main, 6, 6), 3, 3), 0, 0)
    Call forward_elimination(main)
    Call backward_substitution(main)
    Call print_array(main, "last")
    For i = 0 To UBound(main)
        For j = 0 To UBound(main(0))
            'Cells(j + 20)(i + 34) = main(i)(j)
        Next j
    Next i
    Cells(2, 12) = 0
    Cells(3, 12) = main(0)(9)
    Cells(4, 12) = main(1)(9)
    Cells(5, 12) = 0
    Cells(6, 12) = main(2)(9)
    Cells(7, 12) = main(3)(9)
    Cells(8, 12) = 0
    Cells(9, 12) = main(4)(9)
    Cells(10, 12) = main(5)(9)
    Cells(11, 12) = main(6)(9)
    Cells(12, 12) = main(7)(9)
    Cells(13, 12) = main(8)(9)
    For i = 0 To 11
      Dim num
      num = 0
      For j = 0 To 11
        num = num + big(i)(j) * Cells(2 + j, 12)
      Next j
      Debug.Print (Round(num, 3))
      Cells(2 + i, 13) = Round(num, 3)
    Next i
    
End Sub

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

' multiply matrixs
Function matrix_cross(m1, m2)
    ans = create_matrix(UBound(m1) + 1, UBound(m2(0)) + 1)
    For i = 0 To UBound(ans)
        For j = 0 To UBound(ans(0))
            sum_ = 0
            For k = 0 To UBound(m1(0))
                sum_ = sum_ + m1(i)(k) * m2(k)(j)
            Next k
            ans(i)(j) = sum_
        Next j
    Next i
    matrix_cross = ans
End Function

' transpose matrix
Function matrix_t(m)
    ans = create_matrix(UBound(m(0)) + 1, UBound(m) + 1)
    For i = 0 To UBound(m(0))
        For j = 0 To UBound(m)
            ans(i)(j) = m(j)(i)
        Next j
    Next i
    matrix_t = ans
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

Sub print_array(arr, Optional msg As String)
    Debug.Print ("--")
    For i = 0 To UBound(arr)
        tmp = ""
        For j = 0 To UBound(arr(0))
            'tmp = tmp & arr(i)(j) & "  "
            tmp = tmp & Round(arr(i)(j), 6) & "  "
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

Function shrink_array(arr_, Optional row As Integer = -1, Optional col As Integer = -1)
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

