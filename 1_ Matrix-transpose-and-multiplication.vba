' エリアのリセット
Sub 四角形角を丸くする2_Click()
Dim SheetName As String
SheetName = "Sheet1"
Sheets(SheetName).Range("F11", "K16").ClearContents
Sheets(SheetName).Range("B11", "D16").ClearContents
End Sub

' 転置と掛け算を計算するくん
Sub 四角形角を丸くする3_Click()
Dim B, Bt
B = read_matrix_from_sheet(6, 2, 3, 6)
Bt = matrix_t(B)
Dim x As Integer
For x = 11 To 16
Sheets("Sheet1").Range("B" & x, "D" & x) = Bt(x - 11)
Next x
Dim D, BtDB As Variant
D = read_matrix_from_sheet(6, 9, 3, 3)
BtDB = matrix_cross(matrix_cross(Bt, D), B)
Dim y As Integer
For y = 11 To 16
Sheets("Sheet1").Range("F" & y, "K" & y) = BtDB(y - 11)
Next y
End Sub

' シートからジャグ配列として値を取得
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

' 行列の転置
Function matrix_t(m)
  ans = create_matrix(UBound(m(0)) + 1, UBound(m) + 1)
  For i = 0 To UBound(m(0))
    For j = 0 To UBound(m)
      ans(i)(j) = m(j)(i)
    Next j
  Next i
  matrix_t = ans
End Function

' 任意サイズの行列(ジャグ配列)を作成する
Function create_matrix(row_size, col_size)
    Dim ans, row As Variant
    ans = Array()
    ReDim ans(row_size - 1)
    For i = 0 To row_size - 1
        row = Array() ' 新しいオブジェクトのインスタンスが代入される
        ReDim row(col_size - 1)
        ans(i) = row
    Next i
    create_matrix = ans
End Function
' 行列同士の掛け算
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
