Sub reset_area()
    Sheets("Sheet3").Range("A2", "I11").ClearContents
End Sub

' main
' 2020/10/6
' Multiply the matrix of Sheet1 and Sheet2 and write to Sheet3
Sub multiply_matrix()
    Dim A, B, AB
    A = read_matrix_from_sheet("Sheet1")
    B = read_matrix_from_sheet("Sheet2")
    AB = matrix_cross(A, B)
    A = write_matrix_to_sheet("Sheet3", AB, 2, 1)
End Sub

' get values from sheet as jag array
Function read_matrix_from_sheet(SheetName)
    Worksheets(SheetName).Activate
    Dim arr, row, col
    row = Cells(1, 2)
    col = Cells(1, 3)
    arr = create_matrix(row, col)
    For i = 0 To row - 1
        For j = 0 To col - 1
            Debug.Print (Cells(2 + i, 1 + j))
            arr(i)(j) = Cells(2 + i, 1 + j)
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
    Worksheets(SheetName).Activate
    Dim i
    For i = 0 To UBound(matrix)
        Sheets(SheetName).Range(Cells(row_origin + i, col_origin), Cells(row_origin + i, col_origin + UBound(matrix(0)))) = matrix(i)
    Next i
End Function

