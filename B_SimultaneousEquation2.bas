Sub reset_area()
    Sheets("Sheet1").Range("H3", "H4").ClearContents
    Sheets("Sheet1").Range("H6", "H7").ClearContents
    Sheets("Sheet1").Range("G2", "G2").ClearContents
    Sheets("Sheet1").Range("G5", "G5").ClearContents
End Sub

Sub calculate()
    Dim A
    A = create_matrix(4, 5)
    A(0)(0) = Cells(3, 2)
    A(0)(1) = Cells(3, 3)
    A(0)(2) = Cells(3, 5)
    A(0)(3) = Cells(3, 6)
    A(1)(0) = Cells(4, 2)
    A(1)(1) = Cells(4, 3)
    A(1)(2) = Cells(4, 5)
    A(1)(3) = Cells(4, 6)
    A(2)(0) = Cells(6, 2)
    A(2)(1) = Cells(6, 3)
    A(2)(2) = Cells(6, 5)
    A(2)(3) = Cells(6, 6)
    A(3)(0) = Cells(7, 2)
    A(3)(1) = Cells(7, 3)
    A(3)(2) = Cells(7, 5)
    A(3)(3) = Cells(7, 6)
    A(0)(4) = Cells(3, 7)
    A(1)(4) = Cells(4, 7)
    A(2)(4) = Cells(6, 7)
    A(3)(4) = Cells(7, 7)
    
    Call forward_elimination(A)
    Call backward_substitution(A)
    'Call print_array(A, "done")
    Cells(3, 8) = Round(A(0)(4), 3)
    Cells(4, 8) = Round(A(1)(4), 3)
    Cells(6, 8) = Round(A(2)(4), 3)
    Cells(7, 8) = Round(A(3)(4), 3)
    
    Cells(2, 7) = Cells(2, 2) * Cells(3, 8) + Cells(2, 3) * Cells(4, 8) + Cells(2, 5) * Cells(6, 8) + Cells(2, 6) * Cells(7, 8)
    Cells(5, 7) = Cells(5, 2) * Cells(3, 8) + Cells(5, 3) * Cells(4, 8) + Cells(5, 5) * Cells(6, 8) + Cells(5, 6) * Cells(7, 8)
End Sub


Sub forward_elimination(arr)
    For i = 0 To UBound(arr)
        Dim key
        key = arr(i)(i)
        For j = 0 To UBound(arr(0)) - i
            ' divide each element with key
            arr(i)(j + i) = arr(i)(j + i) / key
        Next j

        ' substitute downward each row
        For k = i + 1 To UBound(arr)
            Dim num
            num = arr(k)(i)
            For l = i To UBound(arr(0))
                arr(k)(l) = arr(k)(l) - arr(i)(l) * num
            Next l
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


' get values from sheet as jag array
Function read_matrix_from_sheet(row_origin, col_origin, row_size, col_size)
    Dim arr
    arr = create_matrix(row_size, col_size)
    For i = 0 To row_size - 1
        For j = 0 To col_size - 1
            'Debug.Print (Cells(row_origin + i, col_origin + j))
            arr(i)(j) = Cells(row_origin + i, col_origin + j)
        Next j
    Next i
    read_matrix_from_sheet = arr
End Function

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
