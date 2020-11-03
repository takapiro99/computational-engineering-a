Sub calculate()
    Worksheets("Sheet1").Activate
    Nodes = 6
    E = Cells(1, 2)
    H0 = Cells(2, 2)
    H1 = Cells(3, 2)
    B = Cells(4, 2)
    L = Cells(5, 2)
    L_ = L / (Nodes - 1)                         ' 各要素の幅
    
    ' k
    base_k_ = Array(Array(12, 6 * L_, -12, 6 * L_), Array(6 * L_, 4 * L_ ^ 2, -6 * L_, 2 * L_ ^ 2), Array(-12, -6 * L_, 12, -6 * L_), Array(6 * L_, 2 * L_ ^ 2, -6 * L_, 4 * L_ ^ 2))
    base_k = multiply_each(base_k_, E / L_ ^ 3)
    'print_array (base_k_)
    Dim k_ As Variant
    k_ = Array()
    ReDim k_(Nodes - 1)
    ' kを一つずつ準備
    For i = 0 To Nodes - 1 - 1
        h_i = H0 - ((L_ * i) / L + (L_ * (i + 1) / L)) / 2 * (H0 - H1)
        new_arr = multiply_each(base_k, B * h_i ^ 3 / 12) ' multiply by Iz
        k_(i) = new_arr
    Next i
    ' mainにいれる
    main = create_matrix(Nodes * 2, Nodes * 2 + 1, 0) ' 全部入り係数行列
    For i = 0 To Nodes - 1 - 1
        For j = 0 To 3
            For k = 0 To 3
                r = j + i * 2
                c = k + i * 2
                main(r)(c) = main(r)(c) + k_(i)(j)(k)
            Next k
        Next j
    Next i
    ' F inputとかも入れる
    For i = 0 To 11
        main(i)(UBound(main(0))) = Cells(2 + i, 10)
    Next i
    'print_array (main)
    '縮小しておいてみる
    shrinked_main = shrink_array(shrink_array(main, 0, 0), 0, 0)
    ' はきだす
    'print_array (shrinked_main)
    Call forward_elimination(shrinked_main)
    Call backward_substitution(shrinked_main)
    'print_array (shrinked_main)
    
    ' セルにいれる
    Cells(2, 11) = 0
    Cells(3, 11) = 0
    For i = 0 To 9
      Cells(4 + i, 11) = shrinked_main(i)(UBound(shrinked_main(0)))
    Next i
    ' TODO: F_outputも計算してセルに入れる
End Sub

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

' print array for debug
Sub print_array(arr, Optional msg As String)
    Debug.Print ("--")
    For i = 0 To UBound(arr)
        tmp = ""
        For j = 0 To UBound(arr(0))
            tmp = tmp & arr(i)(j) & "  "
            'tmp = tmp & Round(arr(i)(j), 1) & "  "
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

Sub reset_area()
    Sheets("Sheet1").Range("K2", "K13").ClearContents
    Sheets("Sheet1").Range("L2", "L13").ClearContents
End Sub

