Sub calculate()
    Worksheets("Sheet1").Activate
    Nodes = Cells(1, 2)
    E = Cells(2, 2)
    H0 = Cells(3, 2)
    H1 = Cells(4, 2)
    B = Cells(5, 2)
    L = Cells(6, 2)
    L_ = L / (Nodes - 1)                         ' 各要素の幅
    p = Cells(7, 2)
    
    k_ = create_matrix(2, Nodes - 1)             ' kを集めたもの
    For i = 0 To UBound(k_(0))
        ' それぞれのkに代入
        k_(0)(i) = (((H0 - (L_ * i / L * (H0 - H1))) + (H0 - ((L_ * (i + 1)) / L * (H0 - H1)))) / 2) * B * E / L_
    Next i
    
    main = create_matrix(6, 7)                   ' 全部入り係数行列
    ' ゼロ埋め
    For i = 0 To UBound(main)
        For j = 0 To UBound(main(0))
            main(i)(j) = 0
        Next j
    Next i
    main(5)(6) = p
    ' いろいろ代入
    print_array (main)
    For i = 0 To UBound(k_(0))
        main(i)(i) = main(i)(i) + k_(0)(i)
        main(i)(i + 1) = main(i)(i + 1) - k_(0)(i)
        main(i + 1)(i) = main(i + 1)(i) - k_(0)(i)
        main(i + 1)(i + 1) = main(i + 1)(i + 1) + k_(0)(i)
    Next i
    Call print_array(main, "けいすう")
    compressed = shrink_array(main, 0, 0)
    Call print_array(compressed, "compressed")
    'Exit Sub
    Call forward_elimination(compressed)
    Call backward_substitution(compressed)
    print_array (compressed)
    
    Dim ans As Variant
    ans = Array()
    ReDim ans(UBound(compressed))
    For i = 0 To UBound(compressed)
        ans(i) = Round(compressed(i)(UBound(compressed(0))), 6)
    Next i
    For i = 0 To UBound(ans)
        Cells(3 + i, 7) = ans(i)
    Next i
    ' ずるいかも
    Cells(7, 6) = p
    Cells(2, 6) = -p
End Sub

Sub reset_area()
    Sheets("Sheet1").Range("F2", "F2").ClearContents
    Sheets("Sheet1").Range("F7", "F7").ClearContents
    Sheets("Sheet1").Range("G3", "G7").ClearContents
End Sub

Sub forward_elimination(arr)
    For i = 0 To UBound(arr)
        Dim key
        key = arr(i)(i)
        'print_array (arr)
        For j = 0 To UBound(arr(0)) - i
            ' divide each element with key
            arr(i)(j + i) = arr(i)(j + i) / key
        Next j
        
        ' substitute downward each row
        For k = i + 1 To UBound(arr)
            Dim num
            num = arr(k)(i)
            For L = i To UBound(arr(0))
                arr(k)(L) = arr(k)(L) - arr(i)(L) * num
            Next L
        Next k
        
    Next i
End Sub

Sub backward_substitution(arr)
    For i = UBound(arr) To 1 Step -1
        ' substitute upward each row
        For j = i To 1 Step -1
            Dim key
            key = arr(j - 1)(i)
            For k = UBound(arr(0)) To i Step -1
                arr(j - 1)(k) = arr(j - 1)(k) - arr(i)(k) * key
            Next k
        Next j
    Next i
End Sub

' Create a matrix of arbitrary size (jug array)
Function create_matrix(row_size, col_size)
    Dim ans, row As Variant
    ans = Array()
    ReDim ans(row_size - 1)
    For i = 0 To row_size - 1
        row = Array()                            ' new array Instance
        ReDim row(col_size - 1)
        ans(i) = row
    Next i
    create_matrix = ans
End Function

' print array for debug
Sub print_array(arr, Optional msg As String)
    Debug.Print ("--")
    For i = 0 To UBound(arr)
        Dim tmp
        tmp = ""
        For j = 0 To UBound(arr(0))
            tmp = tmp & arr(i)(j) & "  "
        Next j
        Debug.Print (tmp)
    Next i
    Debug.Print ("--- " & Now & " " & msg & " ---")
End Sub

Function shrink_array(arr, Optional row As Integer = -1, Optional col As Integer = -1)
    new_array_row = UBound(arr)
    new_array_col = UBound(arr(0))
    'row詰める
    If row >= 0 Then
        new_array_row = new_array_row - 1
        For i = 0 To UBound(arr(0))
            For j = row To UBound(arr) - 1
                arr(j)(i) = arr(j + 1)(i)
            Next j
        Next i
        For i = 0 To UBound(arr(0))
            arr(UBound(arr))(i) = ""
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
        For i = 0 To UBound(arr)
            arr(i)(UBound(arr(0))) = ""
        Next i
    End If
    'shrink_array = arr
    new_array = create_matrix(new_array_row + 1, new_array_col + 1)
    For i = 0 To new_array_row
        For j = 0 To new_array_col
            new_array(i)(j) = arr(i)(j)
        Next j
    Next i
    shrink_array = new_array
End Function

