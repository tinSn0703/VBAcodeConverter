VERSION 1.0 CLASS

BEGIN
MultiUse = -1 'True
END

Attribute VB_Name = "CheckListData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'----------------------------------------------------------------------------------------------------
''' <summary>チェックリストのデータを表す</summary>
'----------------------------------------------------------------------------------------------------

Option Explicit

Implements TableDataBase

'----------------------------------------------------------------------------------------------------
' Field
'----------------------------------------------------------------------------------------------------
Private f_num_model As Long
Private f_table As TableData

Private f_model_title_area As TableDataArea
Private f_title_area As TableDataArea
Private f_model_data_area As TableDataArea
Private f_data_area As TableDataArea

Private f_all_title_area As TableDataArea
Private f_all_data_area As TableDataArea

'----------------------------------------------------------------------------------------------------
' Private method
'----------------------------------------------------------------------------------------------------
''' <summary>値は列の範囲外ですか?</summary>
''' <param name="num">調べる値</param>
''' <returns>Yes/No</returns>
Private Function IsOutOfColumnRange(ByVal num As Long) As Boolean
Dim a As Long : a = 10

Dim a As Long, b As Long : a = 10 : b = 10

Dim a As Long, b As Long, c As String : a = 10 : b = 10 : c = 10

Dim d As Long = 10 : d = 100
Dim d As Long = 10 : d = 100

Dim a As String ' <summary>値は行の : 範囲外ですか?</summary>
a = "n : C" : a = ""
a = "" : a = "n : C"
a = "n : C" : a = "C : n"
a = "n : C" : a = "C : n" : a = "C : C"

IsOutOfColumnRange = ((num < Me.BeginColumnNum) Or (Me.EndColumnNum < num))
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>値は行の : 範囲外ですか?</summary>
''' <param name="num">調べる値</param>
''' <returns>Yes/No</returns>
Private Function IsOutOfRowRange(ByVal num As Long) As Boolean
Call ThrowArgumentOutOfRangeException(Me, FUNCTION_NAME, "size", "[" & CStr(size) & " < 1] : 入力値が範囲外です") : Exit Function

IsOutOfRowRange = ((num < Me.BeginRowNum) Or (Me.EndRowNum < num))
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>値は列の範囲外ですか?</summary>
''' <param name="num">調べる値</param>
''' <returns>Yes/No</returns>
Private Function GetMsgOutOfColumnRange(ByVal num As Long) As String
GetMsgOutOfColumnRange = "[" & Me.TableName & "] 列範囲を逸脱する値が入力されました" & vbCrLf & "範囲 : (" & CStr(Me.BeginColumnNum) & " to " & CStr(Me.EndColumnNum) & ") 入力値 : " & CStr(num)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>値は行の範囲外ですか?</summary>
''' <param name="num">調べる値</param>
''' <returns>Yes/No</returns>
Private Function GetMsgOutOfRowRange(ByVal num As Long) As String
GetMsgOutOfRowRange = "[" & Me.TableName & "] 行範囲を逸脱する値が入力されました" & vbCrLf & "範囲 : (" & CStr(Me.BeginRowNum) & " to " & CStr(Me.EndRowNum) & ") 入力値 : " & CStr(num)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>要素アクセス失敗時のエラーメッセージを返す</summary>
''' <param name="column">アクセス先</param>
''' <param name="row">アクセス先</param>
''' <param name="value">入力/出力 値</param>
''' <returns>メッセージ</returns>
Private Function GetMsgDataAccessError(ByVal column As Long, ByVal row As Long, ByVal value As Variant) As String
Dim err_msg As String : err_msg = "[" & Me.TableName & "] の要素へのアクセスに失敗しました" & vbCrLf

If (IsNull(value)) Then
err_msg = err_msg & "要素情報 : Variant [Null]" & vbCrLf
Else
err_msg = err_msg & "要素情報 : " & TypeName(value) & " [" & CStr(value) & "]" & vbCrLf
End IF

GetMsgDataAccessError = err_msg & "[Size] : (" & Cstr(Me.ColumnSize) & ", " & CStr(Me.RowSize) & ")" & vbCrLf & "[Input] : (" & Cstr(column) & ", " & CStr(row) & ")"
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>タイトル欄の初期値を入力する</summary>
Private Function InputInitalizeTitle()
Const FUNCTION_NAME As String = "InputInitalizeTitle()"
On Error Goto CatchErr
Dim column As Long : column = 0
For column = f_title_area.BeginColumnNum To f_title_area.EndColumnNum
f_title_area.Columns(column).Value = GetCheckListColumnTitle(column)
Next
Exit Function
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>機種タイトル欄の初期値を入力する</summary>
Private Function InputInitializeModleTitle()
Const FUNCTION_NAME As String = "InputInitializeModleTitle()"
On Error Goto CatchErr
f_model_title_area.Data(f_model_title_area.BeginColumnNum, CL_MODEL_JPN_TITLE_ROW_NUM) = JPN_DIFFERENCE_COLUMN_TITLE
f_model_title_area.Data(f_model_title_area.BeginColumnNum, CL_MODEL_ENG_TITLE_ROW_NUM) = ENG_DIFFERENCE_COLUMN_TITLE

Call f_model_title_area.Rows(CL_MODEL_NAME_TITLE_ROW_NUM).Fill("XXX000")
Exit Function
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>Indexから、対象機種の機種欄での列番号を返す</summary>
''' <param name="model_index">機種名,もしくはテーブルに記入されている機種欄列番号</param>
Private Function GetTargetModelColumn(ByVal model_index As Variant) As Long
Const FUNCTION_NAME As String = "GetTargetModelColumn()"

Dim model_column As Long, type_code As Long : type_code = VarType(model_index)
If (type_code = vbString) Then '機種名での選択
model_column = Me.ModelTitle.Rows(CL_MODEL_NAME_TITLE_ROW_NUM).Search(model_index, 0, True, False)

If (model_column < Me.ModelTitle.BeginColumnNum) Then
Call Log.WriteLog(Log.WARNING_LOG_LEVEL, Me, FUNCTION_NAME, "[" & model_index & "] not found in [" & Me.TableName & "].")
Call ThrowArgumentException(Me, FUNCTION_NAME, "model_index", "[" & Me.TableName & "]" & vbCrLf & "存在しない機種名参照 value = " & model_index)
End If
ElseIf ((type_code = vbLong) Or (type_code = vbInteger)) Then '列番号での選択
model_column = CLng(model_index)

If ((model_column < Me.ModelTitle.BeginColumnNum) Or (Me.ModelTitle.EndColumnNum < model_column)) Then
Call Log.WriteLog(Log.WARNING_LOG_LEVEL, Me, FUNCTION_NAME, "[" & CStr(model_index) & "] is out of range for [" & Me.TableName & "].")
Call ThrowArgumentOutOfRangeException(Me, FUNCTION_NAME, "model_index", "[" & Me.TableName & "]" & vbCrLf & "範囲外参照 value = " & CStr(model_index))
End If
Else
Call Log.WriteLog(Log.WARNING_LOG_LEVEL, Me, FUNCTION_NAME, "[" & TypeName(model_index) & " : " & CStr(model_index) & "] is incorrect type.")
Call ThrowArgumentException(Me, FUNCTION_NAME, "model_index", "[" & Me.TableName & "]" & vbCrLf & "入力値型 不正 type = " & TypeName(model_index))
End If

GetTargetModelColumn = model_column
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>テーブル内でデータが開始する行番号</summary>
Private Property Get StartDataRowNum() As Long
StartDataRowNum = CL_DATA_STARTING_ROW_NUM
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>テーブル内で機種差異以外のデータが開始する列番号</summary>
Private Property Get StartDataColumnNum() As Long
StartDataColumnNum = Me.NumModel
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>テーブル内でタイトル欄が終了する行番号</summary>
Private Property Get FinishTitleRowNum() As Long
FinishTitleRowNum = StartDataRowNum - 1
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>テーブル内で機種差異が終了する列番号</summary>
Private Property Get FinishModelColumnNum() As Long
FinishModelColumnNum = StartDataColumnNum - 1
End Property

'----------------------------------------------------------------------------------------------------
' Public method
'----------------------------------------------------------------------------------------------------

Private Sub Class_Initialize()
If (IS_WRITE_CLASS_CONSTRUCT_LOG) Then
Call Log.WriteLog(Log.DEBUG_LOG_LEVEL, Me, "Class_Initialize()", "object [" & TypeName(Me) & "] has been created.")
End If
Set f_table = New TableData

Set f_all_title_area = New TableDataArea
Set f_all_data_area = New TableDataArea

Set f_model_title_area = New TableDataArea
Set f_title_area = New TableDataArea
Set f_model_data_area = New TableDataArea
Set f_data_area = New TableDataArea
End Sub

'----------------------------------------------------------------------------------------------------

Private Sub Class_Terminate()
Set f_table = Nothing

Set f_all_title_area = Nothing
Set f_all_data_area = Nothing

Set f_model_title_area = Nothing
Set f_title_area = Nothing
Set f_model_data_area = Nothing
Set f_data_area = Nothing
End Sub

'----------------------------------------------------------------------------------------------------
''' <summary>列数を再設定する</summary>
''' <param name="size">再設定する列数</param>
Function ResetColumnSize(ByVal size As Long)
Const FUNCTION_NAME As String = "ResetColumnSize()"
If (size < 1) Then
Call ThrowArgumentOutOfRangeException(Me, FUNCTION_NAME, "size", "[" & CStr(size) & " < 1] : 入力値が範囲外です")
End If

On Error Goto CatchErr
Call f_table.ResetColumnSize(size + StartDataColumnNum)

Call f_all_data_area.ResetColumnSize(size + StartDataColumnNum)
Call f_all_title_area.ResetColumnSize(size + StartDataColumnNum)
Call f_title_area.ResetColumnSize(size)
Call f_data_area.ResetColumnSize(size)
Exit Function
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>行数を再設定する</summary>
''' <param name="size">再設定する行数</param>
Function ResetRowSize(ByVal size As Long)
Const FUNCTION_NAME As String = "ResetRowSize()"
If (size < 1) Then
Call ThrowArgumentOutOfRangeException(Me, FUNCTION_NAME, "size", "[" & CStr(size) & " < 1] : 入力値が範囲外です")
End If

On Error Goto CatchErr
Call f_table.ResetRowSize(size + StartDataRowNum)

Call f_all_data_area.ResetRowSize(size)
Call f_model_data_area.ResetRowSize(size)
Call f_data_area.ResetRowSize(size)
Exit Function
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>別のテーブルをコピーする</summary>
''' <param name="table_data">コピーするテーブル</param>
Function Copy(ByRef table_data As TableDataBase)
Const FUNCTION_NAME As String = "Copy()"
If (table_data Is Nothing) Then
Call ThrowArgumentNullException(Me, FUNCTION_NAME, "table_data", "コピーに失敗しました")
End If
If (table_data.Name <> Me.Name) Then
Call ThrowArgumentException(Me, FUNCTION_NAME, "table_data", "[" & table_data.Name & "] 対象外のオブジェクトです")
End If

On Error Goto CatchErr
Dim temp As CheckListData : Set temp = table_data.DownCast
Call f_table.Copy(temp.Value)
f_num_model = temp.NumModel

Call Me.SetTableArea()
Exit Function
CatchErr:
Call RethrowException(Me, FUNCTION_NAME, "Me [" & Me.TableName & "] <= Copy [" & table_data.TableName & "]")
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>テーブルを入れ替える</summary>
''' <param name="table_data">入れ替え対象</param>
Public Function Swap(ByRef table_data As TableDataBase)
Const FUNCTION_NAME As String = "Swap()"
If (table_data Is Nothing) Then
Call ThrowArgumentNullException(Me, FUNCTION_NAME, "table_data", "コピーに失敗しました")
End If
If (table_data.Name <> Me.Name) Then
Call ThrowArgumentException(Me, FUNCTION_NAME, "table_data", "[" & table_data.Name & "] 対象外のオブジェクトです")
End If

On Error Goto CatchErr
Dim temp As CheckListData : Set temp = table_data.DownCast
Call Me.Value.Swap(temp.Value)
f_num_model = temp.NumModel

Call Me.SetTableArea()
Exit Function
CatchErr:
Call RethrowException(Me, FUNCTION_NAME, "Me [" & Me.TableName & "] <=> Swap [" & table_data.TableName & "]")
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>列を入れ替える</summary>
''' <param name="num_0">入れ替え対象</param>
''' <param name="num_1">入れ替え対象</param>
Public Function SwapColumn(ByVal num_0 As Long, ByVal num_1 As Long)
Const FUNCTION_NAME As String = "SwapColumn()"
On Error Goto CatchErr
Call f_data_area.SwapColumn(num_0, num_1)
Exit Function
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>行を入れ替える</summary>
''' <param name="num_0">入れ替え対象</param>
''' <param name="num_1">入れ替え対象</param>
Public Function SwapRow(ByVal num_0 As Long, ByVal num_1 As Long)
Const FUNCTION_NAME As String = "SwapRow()"
On Error Goto CatchErr
Call f_all_data_area.SwapRow(num_0, num_1)
Exit Function
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>テーブルに列を挿入する。できれば、何度も繰り返して呼ばれないようにすること。</summary>
''' <param name="insert_size">挿入数</param>
''' <param name="insert_pos">挿入位置。is_preserveがFalseの場合は、自動的に最終位置への追加となる。</param>
''' <param name="is_preserve">データの再確保時にデータを引き継ぐかどうか。</param>
Public Function InsertColumn(ByVal insert_size As Long, Optional ByVal insert_pos = -1, Optional ByVal is_preserve As Boolean = True)
Const FUNCTION_NAME As String = "InsertColumn()"

If (insert_size = 0) Then
Exit Function
End If
If (insert_size < 0) Then
Call ThrowArgumentException(Me, FUNCTION_NAME, "insert_size", "[" & Me.TableName & "] 追加数が少なすぎます")
End If
If (IsOutOfColumnRange(insert_pos)) Then
insert_pos = Me.EndColumnNum
End If

On Error Goto CatchErr
Dim old_table_title As TableData
If ((insert_pos < Me.EndColumnNum) And (is_preserve)) Then
Set old_table_title = New TableData
Call old_table_title.InitializeOnRange(f_all_title_area.Table, f_all_title_area.TableName)
End If

Call f_all_data_area.InsertColumn(insert_size, insert_pos, is_preserve)

Call f_all_title_area.ResetColumnSize(f_all_title_area.ColumnSize + insert_size)
Call f_title_area.ResetColumnSize(f_title_area.ColumnSize + insert_size)
Call f_data_area.ResetColumnSize(f_data_area.ColumnSize + insert_size)

If ((insert_pos < Me.EndRowNum) And (is_preserve)) Then
Set f_all_title_area.Table = old_table_title.Table
End If
Exit Function
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>テーブルに行を挿入する。できれば、何度も繰り返して呼ばれないようにすること。</summary>
''' <param name="insert_size">挿入数</param>
''' <param name="insert_pos">挿入位置。is_preserveがFalseの場合は、自動的に最終位置への追加となる。</param>
''' <param name="is_preserve">データの再確保時にデータを引き継ぐかどうか。</param>
Function InsertRow(ByVal insert_size As Long, Optional ByVal insert_pos = -1, Optional ByVal is_preserve As Boolean = True)
Const FUNCTION_NAME As String = "InsertRow()"

If (insert_size = 0) Then
Exit Function
End If
If (insert_size < 0) Then
Call ThrowArgumentException(Me, FUNCTION_NAME, "insert_size", "[" & Me.TableName & "] 追加数が少なすぎます")
End If
If (IsOutOfRowRange(insert_pos)) Then
insert_pos = Me.EndRowNum
End If

On Error Goto CatchErr
Dim old_table_title As TableData
If ((insert_pos < Me.EndRowNum) And (is_preserve)) Then
Set old_table_title = New TableData
Call old_table_title.InitializeOnRange(f_all_title_area.Table, f_all_title_area.TableName)
End If

Call f_all_data_area.InsertRow(insert_size, insert_pos, is_preserve)

Call f_model_data_area.ResetRowSize(f_model_data_area.RowSize + insert_size)
Call f_data_area.ResetRowSize(f_data_area.RowSize + insert_size)

If ((insert_pos < Me.EndRowNum) And (is_preserve)) Then
Set f_all_title_area.Table = old_table_title.Table
End If
Exit Function
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>テーブルに列を挿入する。できれば、何度も繰り返して呼ばれないようにすること。</summary>
''' <param name="add_data">追加内容</param>
''' <param name="add_pos">追加位置。未指定の場合は最終位置へと追加する</param>
Public Function AddColumn(ByRef add_data As TableDataColumn, Optional ByVal add_pos = -1)
Const FUNCTION_NAME As String = "AddColumn()"
If (add_data Is Nothing) Then
Call ThrowArgumentNullException(Me, FUNCTION_NAME, "add_data")
End If
On Error Goto CatchErr
If (IsOutOfColumnRange(add_pos)) Then
add_pos = Me.EndColumnNum
End If

Call Me.InsertColumn(1, add_pos, True)
Set Me.Columns(add_pos) = add_data
Exit Function
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>テーブルに行を挿入する。できれば、何度も繰り返して呼ばれないようにすること。</summary>
''' <param name="add_data">追加内容</param>
''' <param name="add_pos">追加位置。未指定の場合は最終位置へと追加する</param>
Public Function AddRow(ByRef add_data As TableDataRow, Optional ByVal add_pos = -1)
Const FUNCTION_NAME As String = "AddRow()"
If (add_data Is Nothing) Then
Call ThrowArgumentNullException(Me, FUNCTION_NAME, "add_data")
End If
On Error Goto CatchErr
If (IsOutOfRowRange(add_pos)) Then
add_pos = Me.EndRowNum
End If

Call Me.InsertRow(1, add_pos, True)
Set Me.Rows(add_pos) = add_data
Exit Function
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>テーブルの先頭を指すイテレータを返す</summary>
''' <returns>イテレータ</returns>
Public Function BeginIterator() As TableDataIterator
Const FUNCTION_NAME As String = "BeginIterator()"

Dim iterator As TableDataIterator : Set iterator = New TableDataIterator
On Error GoTo CatchErr
Call iterator.Initialize(Me, Me.BeginNum, Me.EndNum)
Set BeginIterator = iterator
Exit Function
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>テーブルの末尾の次を指すイテレータを返す</summary>
''' <returns>イテレータ</returns>
Public Function EndIterator() As TableDataIterator
Const FUNCTION_NAME As String = "EndIterator()"

Dim iterator As TableDataIterator : Set iterator = New TableDataIterator
On Error GoTo CatchErr
Call iterator.Initialize(Me, Me.BeginNum, Me.EndNum, True)
Set EndIterator = iterator
Exit Function
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>テーブルの値を配列形式に変換する</summary>
''' <returns>配列形式のテーブルのデータ</returns>
Public Function ToArray() As Variant()
Const FUNCTION_NAME As String = "ToArray()"
On Error GoTo CatchErr
ToArray = Me.Table.ToArray
Exit Function
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>列数</summary>
Public Property Get ColumnSize() As Long
ColumnSize = f_table.ColumnSize - StartDataColumnNum
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>行数</summary>
Public Property Get RowSize() As Long
RowSize = f_table.RowSize - StartDataRowNum
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>先頭番号</summary>
Public Property Get BeginNum() As Position
BeginNum = ToPosition(Me.BeginColumnNum, Me.BeginRowNum)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>末尾番号</summary>
Public Property Get EndNum() As Position
EndNum = ToPosition(Me.EndColumnNum, Me.EndRowNum)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>先頭列番号</summary>
Public Property Get BeginColumnNum() As Long
BeginColumnNum = 0
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>先頭行番号</summary>
Public Property Get BeginRowNum() As Long
BeginRowNum = 0
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>末尾列番号</summary>
Public Property Get EndColumnNum() As Long
EndColumnNum = me.ColumnSize - 1
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>末尾行番号</summary>
Public Property Get EndRowNum() As Long
EndRowNum = Me.RowSize - 1
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>テーブル参照.Data()で参照できる範囲</summary>
Public Property Get Table() As TableDataRange
Const FUNCTION_NAME As String = "Get Table()"
On Error GoTo CatchErr
Set Table = f_data_area.Table
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>テーブル参照.Data()で参照できる範囲</summary>
Public Property Set Table(ByRef table_range As TableDataRange)
Const FUNCTION_NAME As String = "Set Table()"
On Error GoTo CatchErr
Set f_data_area.Table = table_range
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>範囲参照</summary>
Public Property Get Range(ByRef head As Position, ByRef tail As Position) As TableDataRange
Const FUNCTION_NAME As String = "Get Range()"
On Error GoTo CatchErr
Set Range = f_data_area.Range(head, tail)
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>範囲参照</summary>
Public Property Set Range(ByRef head As Position, ByRef tail As Position, ByRef table_range As TableDataRange)
Const FUNCTION_NAME As String = "Set Range()"
On Error GoTo CatchErr
Set f_data_area.Range(head, tail) = table_range
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>1列範囲参照</summary>
Public Property Get Columns(ByVal column As Long) As TableDataColumn
Const FUNCTION_NAME As String = "Get Columns()"
On Error GoTo CatchErr
Set Columns = f_data_area.Columns(column)
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>1列範囲参照</summary>
Public Property Set Columns(ByVal column As Long, ByRef table_range As TableDataColumn)
Const FUNCTION_NAME As String = "Set Columns()"
On Error GoTo CatchErr
Set f_data_area.Columns(column) = table_range
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>1行範囲参照</summary>
Public Property Get Rows(Byval row As Long) As TableDataRow
Const FUNCTION_NAME As String = "Get Rows()"
On Error GoTo CatchErr
Set Rows = f_data_area.Rows(row)
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>1行範囲参照</summary>
Public Property Set Rows(Byval row As Long, ByRef table_range As TableDataRow)
Const FUNCTION_NAME As String = "Set Rows()"
On Error GoTo CatchErr
Set f_data_area.Rows(row) = table_range
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>範囲チェック付き要素アクセス</summary>
Public Property Get At(ByVal column As Long, ByVal row As Long) As Variant
Const FUNCTION_NAME As String = "Get At()"

If ((Me.EndColumnNum < column) Or (Me.EndRowNum < row)) Then
Call ThrowArgumentOutOfRangeException(Me, FUNCTION_NAME, "column/row", GetMsgDataAccessError(column, row, ""))
End If

On Error GoTo CatchErr
At = f_data_area.At(column, row)
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME, GetMsgDataAccessError(column, row, ""))
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>範囲チェック付き要素アクセス</summary>
Public Property Let At(ByVal column As Long, ByVal row As Long, ByVal value As Variant)
Const FUNCTION_NAME As String = "Set At()"

If ((Me.EndColumnNum < column) Or (Me.EndRowNum < row)) Then
Call ThrowArgumentOutOfRangeException(Me, FUNCTION_NAME, "column/row", GetMsgDataAccessError(column, row, value))
End If

On Error GoTo CatchErr
f_data_area.At(column, row) = value
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME, GetMsgDataAccessError(column, row, value))
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>要素アクセス</summary>
Public Property Get Data(ByVal column As Long, ByVal row As Long) As Variant
Const FUNCTION_NAME As String = "Get Data()"
On Error GoTo CatchErr
Data = f_data_area.Data(column, row)
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME, GetMsgDataAccessError(column, row, ""))
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>要素アクセス</summary>
Public Property Let Data(ByVal column As Long, ByVal row As Long, ByRef value As Variant)
Const FUNCTION_NAME As String = "Set Data()"
On Error GoTo CatchErr
f_data_area.Data(column, row) = value
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME, GetMsgDataAccessError(column, row, value))
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>テーブル名</summary>
Public Property Get TableName() As String
TableName = f_table.TableName
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>テーブル名</summary>
Public Property Let TableName(ByVal table_name As String)
If (Me.TableName <> table_name) Then
Call Log.WriteLog(Log.INFO_LOG_LEVEL, Me, "TableName()", "[" & Me.TableName & "] to [" & table_name & "]")
End If

f_table.TableName = table_name
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>オブジェクト名</summary>
Public Property Get Name() As String
Name = TypeName(Me)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>ダウンキャスト用</summary>
Public Property Get DownCast() As Object
Set DownCast = Me
End Property

'----------------------------------------------------------------------------------------------------
' 独自実装
'----------------------------------------------------------------------------------------------------
''' <summary>チェックリストテーブルを初期化する</summary>
''' <param name="num_model">機種数</param>
''' <param name="table_name">テーブル名</param>
Function Initialize(Optional ByVal num_model As Long = 1, Optional ByVal row_size As Long = 1, Optional ByVal table_name As String = "")
Const FUNCTION_NAME As String = "Initialize()"

On Error GoTo CatchErr
If (num_model < 1) Then
num_model = 1
End If
If (row_size < 1) Then
row_size = 1
End If
If (table_name = "") Then
table_name = CHECK_LIST_SHEET_TITLE
End If
Call f_table.Initialize(CHECK_LIST_COLUMN_TITLE_NUM + num_model, CL_DATA_STARTING_ROW_NUM + row_size, table_name)

f_num_model = num_model

Call Me.SetTableArea()
Call InputInitializeModleTitle()
Call InputInitalizeTitle()

Call Log.WriteLog(Log.INFO_LOG_LEVEL, Me, FUNCTION_NAME, "Initialized [" & Me.TableName & "]. Model Num : " & CStr(Me.NumModel))
Exit Function
CatchErr:
Call RethrowException(Me, FUNCTION_NAME, "初期設定に失敗しました")
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>別テーブルのタイトル欄をコピーする</summary>
''' <param name="table_data">コピーするテーブル</param>
Public Function CopyTitle(ByRef table_data As CheckListData)
Const FUNCTION_NAME As String = "CopyTitle()"

If (table_data Is Nothing) Then
Call ThrowArgumentNullException(Me, FUNCTION_NAME, "table_data", "タイトルのコピーに失敗しました")
End If

Call Log.WriteLog(Log.INFO_LOG_LEVEL, Me, FUNCTION_NAME, "Start copying the title of [" & table_data.TableName & "] to [" & Me.TableName & "].")

Set Me.ModelTitle.Table = table_data.ModelTitle.Table
Set Me.Title.Table = table_data.Title.Table

Call Log.WriteLog(Log.INFO_LOG_LEVEL, Me, FUNCTION_NAME, "Successfully copied the title to [" & Me.TableName & "].")
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>新しい機種の列を追加する</summary>
''' <param name="new_model_name">追加する機種名</param>
Public Function AddNewModelColumn(ByRef new_model_name As String)
Const FUNCTION_NAME As String = "AddNewModelColumn()"
Call Log.WriteLog(Log.INFO_LOG_LEVEL, Me, FUNCTION_NAME, "Start adding [" & new_model_name & "] model column to the [" & Me.TableName & "] with [" & CStr(Me.NumModel) & "] models.")

Call f_model_title_area.InsertColumn(1)
f_model_title_area.Data(f_model_title_area.EndColumnNum, CL_MODEL_NAME_TITLE_ROW_NUM) = new_model_name

f_num_model = f_num_model + 1

Call f_title_area.Initialize(f_table, ToPosition(StartDataColumnNum, f_table.BeginRowNum), ToPosition(f_table.EndColumnNum, FinishTitleRowNum), f_title_area.AreaName)
Call f_data_area.Initialize(f_table, ToPosition(StartDataColumnNum, StartDataRowNum), f_table.EndNum, f_data_area.AreaName)

Call f_model_data_area.ResetColumnSize(Me.NumModel)
Call f_all_data_area.ResetColumnSize(f_all_data_area.ColumnSize + 1)
Call f_all_title_area.ResetColumnSize(f_all_title_area.ColumnSize + 1)

Call Log.WriteLog(Log.INFO_LOG_LEVEL, Me, FUNCTION_NAME, "Successfully adding new model.")
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>入力されているデータから機種数を調べる</summary>
''' <returns>調べた機種数。NumModelにも反映される</returns>
Public Function CheckNumModel() As Long
f_num_model = f_table.Rows(CL_ENG_TITLE_ROW_NUM).Search(ENG_PRIORITY_COLUMN_TITLE)
CheckNumModel = f_num_model
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>チェックリストテーブルを分割したテーブルの範囲を設定する。</summary>
Public Function SetTableArea()
Call f_all_title_area.Initialize(f_table, f_table.BeginNum, ToPosition(f_table.EndColumnNum, FinishTitleRowNum), "All Title")
Call f_all_data_area.Initialize(f_table, ToPosition(f_table.BeginColumnNum, StartDataRowNum), f_table.EndNum, "All Data")

Call f_model_title_area.Initialize(f_table, f_table.BeginNum, ToPosition(FinishModelColumnNum, FinishTitleRowNum), "Model Title")
Call f_model_data_area.Initialize(f_table, ToPosition(f_table.BeginColumnNum, StartDataRowNum), ToPosition(FinishModelColumnNum, f_table.EndRowNum), "Model Data")
Call f_title_area.Initialize(f_table, ToPosition(StartDataColumnNum, f_table.BeginRowNum), ToPosition(f_table.EndColumnNum, FinishTitleRowNum), "Title")
Call f_data_area.Initialize(f_table, ToPosition(StartDataColumnNum, StartDataRowNum), f_table.EndNum, "Data")
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>CheckListData内から1機種分のデータを取り出す</summary>
''' <param name="model_index">機種名,もしくはテーブルに記入されている機種欄列番号</param>
''' <returns>変換後テーブル</returns>
Public Function ToCheckListAutoData(ByVal model_index As Variant) As CheckListAutoData
Const FUNCTION_NAME As String = "ToCheckListAutoData()"
On Error GoTo CatchErr
Dim model_column As Long : model_column = GetTargetModelColumn(model_index)

Dim ret_table As CheckListAutoData : Set ret_table = New CheckListAutoData
If (Me.NumModel <= 1) Then
Call ret_table.Initialize(Me.RowSize, Me.ModelName(0)) : Call ret_table.Value.Copy(f_data_area)
Set ToCheckListAutoData = ret_table : Exit Function
End If

Dim model_rows() As Variant : model_rows = Me.model.Columns(model_column).SearchAll("^1$")
If (model_rows(0) < Me.model.BeginRowNum) Then
Call ret_table.Initialize(1, Me.ModelName(model_column)) : Call ret_table.ResetColumnSize(Me.ColumnSize)
Set ToCheckListAutoData = ret_table : Exit Function
End If

Call ret_table.Initialize(Ubound(model_rows), Me.ModelTitle.Data(model_column, CL_MODEL_NAME_TITLE_ROW_NUM)) : Call ret_table.ResetColumnSize(Me.ColumnSize)

Dim iterator As TableRowIterator : Set iterator = New TableRowIterator : Call iterator.Initialize(ret_table, ret_table.BeginRowNum, ret_table.EndRowNum)
Dim row As Variant
For Each row In model_rows
Set iterator.Rows = Me.Rows(row)
Call iterator.CountUp()
Next

Set ToCheckListAutoData = ret_table
Exit Function
CatchErr:
Call RethrowException(Me, FUNCTION_NAME, "機種データの取り出しに失敗しました。")
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>差異欄も含めたチェックリストテーブルに入力されているデータ全体の行範囲参照</summary>
Public Property Get TableRows(ByVal row As Long) As TableDataRow
Const FUNCTION_NAME As String = "Get TableRows()"
On Error GoTo CatchErr
Set TableRows = Me.AllData.Rows(row)
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>差異欄も含めたチェックリストテーブルに入力されているデータ全体の行範囲参照</summary>
Public Property Set TableRows(ByVal row As Long, ByRef table_range As TableDataRow)
Const FUNCTION_NAME As String = "Set TableRows()"

If (table_range Is Nothing) Then
Call ThrowArgumentNullException(Me, FUNCTION_NAME, "table_range", "コピーに失敗しました")
End If
If (table_range.Table.Name <> Me.Name) Then
Call ThrowArgumentException(Me, FUNCTION_NAME, "table_range", "参照しているテーブル型が対象外です")
End If

On Error GoTo CatchErr
Dim temp As CheckListData : Set temp = table_range.Table.DownCast
If ((temp.NumModel = Me.NumModel) And (temp.RowSize = Me.RowSize)) Then
Set Me.AllData.Rows(row) = table_range
Else
Set Me.Model.Rows(row) = temp.Model.Rows(table_range.SourceNum)
Set Me.Rows(row) = temp.Rows(table_range.SourceNum)
End If
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>差異欄も含めたチェックリストテーブルに入力されているデータ全体の範囲参照</summary>
Public Property Get TableRange(ByRef head As Position, ByRef tail As Position) As TableDataRange
Const FUNCTION_NAME As String = "Set TableRange()"
On Error GoTo CatchErr
Set TableRange = Me.AllData.Range(head, tail)
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>差異欄も含めたチェックリストテーブルに入力されているデータ全体の範囲参照</summary>
Public Property Set TableRange(ByRef head As Position, ByRef tail As Position, ByRef table_range As TableDataRange)
Const FUNCTION_NAME As String = "Set TableRange()"
If (table_range Is Nothing) Then
Call ThrowArgumentNullException(Me, FUNCTION_NAME, "table_range", "コピーに失敗しました")
End If
If (table_range.Table.Name <> Me.Name) Then
Call ThrowArgumentException(Me, FUNCTION_NAME, "table_range", "参照しているテーブル型が対象外です")
End If
If (IsHeadOverTail(head, tail)) Then
Call ThrowArgumentOutOfRangeException(Me, FUNCTION_NAME, "head/tail", GetMsgHeadOverTailError(head, tail))
End If

On Error GoTo CatchErr
Dim temp As CheckListData : Set temp = table_range.Table.DownCast
If ((temp.NumModel = Me.NumModel) And (temp.RowSize = Me.RowSize)) Then
Set Me.AllData.Range(head, tail) = table_range
Else
Dim wrapper As TableDataWrapper : Set wrapper = New TableDataWrapper
If (head.column < Me.NumModel) Then
Set Me.Model.Range(head, tail) = wrapper.RowsRange(temp.Model, table_range.SourceHead.row, table_range.SourceTail.row)
End If

If (Me.NumModel <= tail.column) Then
Set Me.Range(ToPosition(head.column - Me.NumModel, head.row), ToPosition(tail.column - Me.NumModel, tail.row)) = wrapper.RowsRange(temp, table_range.SourceHead.row, table_range.SourceTail.row)
End If
End If
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>内包しているTableDataへの参照</summary>
Public Property Set Value(ByRef table_data As TableData)
Call f_table.Copy(table_data)

Call Me.CheckNumModel()
Call Me.SetTableArea()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>内包しているTableDataへの参照</summary>
Public Property Get Value() As TableData
Set Value = f_table
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>機種名</summary>
Public Property Let ModelName(ByVal index As Long, ByVal model_name As String)
Me.ModelTitle.Data(index, CL_MODEL_NAME_TITLE_ROW_NUM) = model_name
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>機種名</summary>
Public Property Get ModelName(ByVal index As Long) As String
ModelName = Me.ModelTitle.Data(index, CL_MODEL_NAME_TITLE_ROW_NUM)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>差異欄のタイトル全体を表す</summary>
Public Property Get ModelTitle() As TableDataArea
Set ModelTitle = f_model_title_area
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>差異欄以外のタイトル全体を表す</summary>
Public Property Get Title() As TableDataArea
Set Title = f_title_area
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>差異欄のデータ全体を表す</summary>
Public Property Get Model() As TableDataArea
Set Model = f_model_data_area
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>差異欄も含めたチェックリストテーブルに入力されているデータ全体を表す</summary>
Public Property Get AllData() As TableDataArea
Set AllData = f_all_data_area
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>差異欄も含めたチェックリストテーブルのタイトル全体を表す</summary>
Public Property Get AllTitle() As TableDataArea
Set AllTitle = f_all_title_area
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>機種数を返す</summary>
Public Property Get NumModel() As Long
NumModel = f_num_model
End Property

'----------------------------------------------------------------------------------------------------
' 継承用
'----------------------------------------------------------------------------------------------------
''' <summary>列数を再設定する</summary>
''' <param name="size">再設定する列数</param>
Public Function TableDataBase_ResetColumnSize(ByVal size As Long)
Call Me.ResetColumnSize(size)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>行数を再設定する</summary>
''' <param name="size">再設定する行数</param>
Public Function TableDataBase_ResetRowSize(ByVal size As Long)
Call Me.ResetRowSize(size)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>別のテーブルをコピーする</summary>
''' <param name="table_data">コピーするテーブル</param>
Public Function TableDataBase_Copy(ByRef table_data As TableDataBase)
Call Me.Copy(table_data)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>テーブルを入れ替える</summary>
''' <param name="table_data">入れ替え対象</param>
Public Function TableDataBase_Swap(ByRef table_data As TableDataBase)
Call Me.Swap(table_data)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>列を入れ替える</summary>
''' <param name="num_0">入れ替え対象</param>
''' <param name="num_1">入れ替え対象</param>
Public Function TableDataBase_SwapColumn(ByVal num_0 As Long, ByVal num_1 As Long)
Call Me.SwapColumn(num_0, num_1)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>行を入れ替える</summary>
''' <param name="num_0">入れ替え対象</param>
''' <param name="num_1">入れ替え対象</param>
Public Function TableDataBase_SwapRow(ByVal num_0 As Long, ByVal num_1 As Long)
Call Me.SwapRow(num_0, num_1)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>テーブルに列を挿入する。できれば、何度も繰り返して呼ばれないようにすること。</summary>
''' <param name="insert_size">挿入数</param>
''' <param name="insert_pos">挿入位置。is_preserveがFalseの場合は、自動的に最終位置への追加となる。</param>
''' <param name="is_preserve">データの再確保時にデータを引き継ぐかどうか。</param>
Public Function TableDataBase_InsertColumn(ByVal insert_size As Long, Optional insert_pos = -1, Optional ByVal is_preserve As Boolean = True)
Call Me.InsertColumn(insert_size, insert_pos, is_preserve)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>テーブルに行を挿入する。できれば、何度も繰り返して呼ばれないようにすること。</summary>
''' <param name="insert_size">挿入数</param>
''' <param name="insert_pos">挿入位置。is_preserveがFalseの場合は、自動的に最終位置への追加となる。</param>
''' <param name="is_preserve">データの再確保時にデータを引き継ぐかどうか。</param>
Public Function TableDataBase_InsertRow(ByVal insert_size As Long, Optional insert_pos = -1, Optional ByVal is_preserve As Boolean = True)
Call Me.InsertRow(insert_size, insert_pos, is_preserve)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>テーブルに列を挿入する。できれば、何度も繰り返して呼ばれないようにすること。</summary>
''' <param name="add_data">追加内容</param>
''' <param name="add_pos">追加位置。未指定の場合は最終位置へと追加する</param>
Public Function TableDataBase_AddColumn(ByRef add_data As TableDataColumn, Optional add_pos = -1)
Call Me.AddColumn(add_data, add_pos)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>テーブルに行を挿入する。できれば、何度も繰り返して呼ばれないようにすること。</summary>
''' <param name="add_data">追加内容</param>
''' <param name="add_pos">追加位置。未指定の場合は最終位置へと追加する</param>
Public Function TableDataBase_AddRow(ByRef add_data As TableDataRow, Optional add_pos = -1)
Call Me.AddRow(add_data, add_pos)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>テーブルの開始位置のイテレータを返す</summary>
''' <returns>イテレータ</returns>
Public Function TableDataBase_BeginIterator() As TableDataIterator
Set TableDataBase_BeginIterator = Me.BeginIterator()
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>テーブルの終了位置のイテレータを返す</summary>
''' <returns>イテレータ</returns>
Public Function TableDataBase_EndIterator() As TableDataIterator
Set TableDataBase_EndIterator = Me.EndIterator()
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>テーブルの値を配列形式に変換する</summary>
''' <returns>配列形式のテーブルのデータ</returns>
Public Function TableDataBase_ToArray() As Variant()
TableDataBase_ToArray = Me.ToArray()
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>列数</summary>
Public Property Get TableDataBase_ColumnSize() As Long
TableDataBase_ColumnSize = Me.ColumnSize()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>行数</summary>
Public Property Get TableDataBase_RowSize() As Long
TableDataBase_RowSize = Me.RowSize()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>先頭番号</summary>
Public Property Get TableDataBase_BeginNum() As Position
TableDataBase_BeginNum = Me.BeginNum()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>末尾番号</summary>
Public Property Get TableDataBase_EndNum() As Position
TableDataBase_EndNum = Me.EndNum()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>先頭列番号</summary>
Public Property Get TableDataBase_BeginColumnNum() As Long
TableDataBase_BeginColumnNum = Me.BeginColumnNum()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>先頭行番号</summary>
Public Property Get TableDataBase_BeginRowNum() As Long
TableDataBase_BeginRowNum = Me.BeginRowNum()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>末尾列番号</summary>
Public Property Get TableDataBase_EndColumnNum() As Long
TableDataBase_EndColumnNum = Me.EndColumnNum()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>末尾行番号</summary>
Public Property Get TableDataBase_EndRowNum() As Long
TableDataBase_EndRowNum = Me.EndRowNum()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>テーブル参照</summary>
Public Property Get TableDataBase_Table() As TableDataRange
Set TableDataBase_Table = Me.Table
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>テーブル参照</summary>
Public Property Set TableDataBase_Table(ByRef table_range As TableDataRange)
Set Me.Table = table_range
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>範囲参照</summary>
Public Property Get TableDataBase_Range(ByRef head As Position, ByRef tail As Position) As TableDataRange
Set TableDataBase_Range = Me.Range(head, tail)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>範囲参照</summary>
Public Property Set TableDataBase_Range(ByRef head As Position, ByRef tail As Position, ByRef table_range As TableDataRange)
Set Me.Range(head, tail) = table_range
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>1列範囲参照</summary>
Public Property Get TableDataBase_Columns(ByVal column As Long) As TableDataColumn
Set TableDataBase_Columns = Me.Columns(column)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>1列範囲参照</summary>
Public Property Set TableDataBase_Columns(ByVal column As Long, ByRef table_column As TableDataColumn)
Set Me.Columns(column) = table_column
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>1行範囲参照</summary>
Public Property Get TableDataBase_Rows(Byval row As Long) As TableDataRow
Set TableDataBase_Rows = Me.Rows(row)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>1行範囲参照</summary>
Public Property Set TableDataBase_Rows(Byval row As Long, ByRef table_row As TableDataRow)
Set Me.Rows(row) = table_row
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>範囲チェック付き要素アクセス</summary>
Public Property Let TableDataBase_At(ByVal column As Long, ByVal row As Long, ByVal value As Variant)
Me.At(column, row) = value
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>範囲チェック付き要素アクセス</summary>
Public Property Get TableDataBase_At(ByVal column As Long, ByVal row As Long) As Variant
TableDataBase_At = Me.At(column, row)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>要素アクセス</summary>
Public Property Let TableDataBase_Data(ByVal column As Long, ByVal row As Long, ByVal value As Variant)
Me.Data(column, row) = value
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>要素アクセス</summary>
Public Property Get TableDataBase_Data(ByVal column As Long, ByVal row As Long) As Variant
TableDataBase_Data = Me.Data(column, row)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>テーブル名</summary>
Public Property Get TableDataBase_TableName() As String
TableDataBase_TableName = Me.TableName()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>テーブル名</summary>
Public Property Let TableDataBase_TableName(ByVal table_name As String)
Me.TableName = table_name
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>オブジェクト名</summary>
Public Property Get TableDataBase_Name() As String
TableDataBase_Name = Me.Name()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>ダウンキャスト用</summary>
Public Property Get TableDataBase_DownCast() As Object
Set TableDataBase_DownCast = Me.DownCast()
End Property

'----------------------------------------------------------------------------------------------------



