## Excel VBAで数値計算


#### 1. 行列の転置と掛け算

##### Function

- `reset_area(void) => void`
  - 直書きされたエリアをクリアする
-  `read_matrix_from_sheet(row_origin: int, col_origin: int, row_size: int, col_size: int) => jagArray`
- `matrix_t(m: jagArray) => jagArray`
  - 転置
- `create_matrix(row_size: int, col_size: int) => jagArray`
  - ジャグ行列の作成
- `matrix_cross(m1: jagArray, m2: jagArray) => jagArray`
  - 配列の掛け算
- `write_matrix_to_sheet(SheetName: string, matrix: jagArray, row_origin: int, col_origin: int) => null`
  - 配列をシートに記入


#### 1A. 行列の掛け算

##### Function

-  `read_matrix_from_sheet(SheetName: string) => jagArray`
   -  シートから値を読み込む
- `create_matrix(row_size: int, col_size: int) => jagArray`
  - ジャグ行列の作成
- `matrix_cross(m1: jagArray, m2: jagArray) => jagArray`
  - 配列の掛け算
- `write_matrix_to_sheet(SheetName: string, matrix: jagArray, row_origin: int, col_origin: int) => null`
  - 配列をシートに記入