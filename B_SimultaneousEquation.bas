Sub reset_area()
    MsgBox ("reset!")
End Sub

' main
Sub solve_simultaneous_equation()
    Dim A
    A = read_matrix_from_sheet(2, 1, 6, 7)
    
    ' each row
    For i = 0 To UBound(A) - 1
        Dim key
        key = A(i)(i)
        For j = 0 To UBound(A(0)) - i
            ' divide each element with key
            A(i)(j + i) = A(i)(j + i) / key
        Next j
        'Call print_array(A)
        ' substitute downwards each row
        For k = i + 1 To UBound(A)
            Dim number1
            number1 = 0
            number1 = A(k)(i)
            'MsgBox (number1 & " " & k & " " & i)
            ' substitute each elements
            For l = i To UBound(A(0))
                A(k)(l) = A(k)(l) - A(i)(l) * number1
            Next l
        Next k
        'MsgBox (UBound(A))
    Next i
    For i = 0 To 5
        'Debug.Print (A(5)(i))
    Next i
    Call print_array(A)
End Sub

Sub print_array(arr)
    Debug.Print ("==========")
    Debug.Print (Now)
    Debug.Print ("--")
    For i = 0 To UBound(arr)
        Dim tmp
        tmp = ""
        For j = 0 To UBound(arr(0))
            tmp = tmp & arr(i)(j) & "  "
        Next j
        Debug.Print (tmp)
    Next i
    Debug.Print ("==========")
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

