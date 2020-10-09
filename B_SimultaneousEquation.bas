Sub reset_area()
    Sheets("Sheet1").Range("H2", "H7").ClearContents
End Sub

' main
Sub solve_simultaneous_equation()
    Dim A
    A = read_matrix_from_sheet(2, 1, 6, 7)
    Call forward_elimination(A)
    Call backward_substitution(A)
    'Call print_array(A, "done")
    ' get answer
    Dim ans As Variant
    ans = Array()
    ReDim ans(UBound(A))
    For i = 0 To UBound(A)
        'ans(i) = A(i)(UBound(A(0)))
        ans(i) = Round(A(i)(UBound(A(0))), 3)
    Next i

    Worksheets("Sheet1").Activate
    For i = 0 To UBound(ans)
        Sheets("Sheet1").Cells(2 + i, 8) = ans(i)
    Next i
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

