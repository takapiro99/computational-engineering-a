' reset area
Sub reset_area()
    Dim SheetName As String
    SheetName = "Sheet1"
    Sheets(SheetName).Range("E6", "J11").ClearContents
    Sheets(SheetName).Range("B6", "D11").ClearContents
End Sub

' main
' 2020/10/6
' transpose matrix and multiply
Sub transpose_matrix_and_multiply()
    Dim SheetName
    SheetName = "Sheet1"
    Dim B, D, Bt, BtDB As Variant
    B = read_matrix_from_sheet(2, 2, 3, 6)
    Bt = matrix_t(B)
    
    a = write_matrix_to_sheet(SheetName, Bt, 6, 2)

    D = read_matrix_from_sheet(2, 8, 3, 3)
    BtDB = matrix_cross(matrix_cross(Bt, D), B)

    a = write_matrix_to_sheet(SheetName, BtDB, 6, 5)
End Sub

' get values from sheet as jag array
Function read_matrix_from_sheet(row_origin, col_origin, row_size, col_size)
    Dim arr
    arr = create_matrix(row_size, col_size)
    For i = 0 To row_size - 1
        For j = 0 To col_size - 1
            Debug.Print (Cells(row_origin + i, col_origin + j))
            arr(i)(j) = Cells(row_origin + i, col_origin + j)
        Next j
    Next i
    read_matrix_from_sheet = arr
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

' write into sheet
Function write_matrix_to_sheet(SheetName, matrix, row_origin, col_origin)
    Dim i
    For i = 0 To UBound(matrix)
        Sheets(SheetName).Range(Cells(row_origin + i, col_origin), Cells(row_origin + i, col_origin + UBound(matrix(0)))) = matrix(i)
    Next i
End Function

