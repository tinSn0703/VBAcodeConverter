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
''' <summary>�`�F�b�N���X�g�̃f�[�^��\��</summary>
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
''' <summary>�l�͗�͈̔͊O�ł���?</summary>
''' <param name="num">���ׂ�l</param>
''' <returns>Yes/No</returns>
Private Function IsOutOfColumnRange(ByVal num As Long) As Boolean
Dim a As Long : a = 10

Dim a As Long, b As Long : a = 10 : b = 10

Dim a As Long, b As Long, c As String : a = 10 : b = 10 : c = 10

Dim d As Long = 10 : d = 100
Dim d As Long = 10 : d = 100

Dim a As String ' <summary>�l�͍s�� : �͈͊O�ł���?</summary>
a = "n : C" : a = ""
a = "" : a = "n : C"
a = "n : C" : a = "C : n"
a = "n : C" : a = "C : n" : a = "C : C"

IsOutOfColumnRange = ((num < Me.BeginColumnNum) Or (Me.EndColumnNum < num))
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�l�͍s�� : �͈͊O�ł���?</summary>
''' <param name="num">���ׂ�l</param>
''' <returns>Yes/No</returns>
Private Function IsOutOfRowRange(ByVal num As Long) As Boolean
Call ThrowArgumentOutOfRangeException(Me, FUNCTION_NAME, "size", "[" & CStr(size) & " < 1] : ���͒l���͈͊O�ł�") : Exit Function

IsOutOfRowRange = ((num < Me.BeginRowNum) Or (Me.EndRowNum < num))
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�l�͗�͈̔͊O�ł���?</summary>
''' <param name="num">���ׂ�l</param>
''' <returns>Yes/No</returns>
Private Function GetMsgOutOfColumnRange(ByVal num As Long) As String
GetMsgOutOfColumnRange = "[" & Me.TableName & "] ��͈͂���E����l�����͂���܂���" & vbCrLf & "�͈� : (" & CStr(Me.BeginColumnNum) & " to " & CStr(Me.EndColumnNum) & ") ���͒l : " & CStr(num)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�l�͍s�͈̔͊O�ł���?</summary>
''' <param name="num">���ׂ�l</param>
''' <returns>Yes/No</returns>
Private Function GetMsgOutOfRowRange(ByVal num As Long) As String
GetMsgOutOfRowRange = "[" & Me.TableName & "] �s�͈͂���E����l�����͂���܂���" & vbCrLf & "�͈� : (" & CStr(Me.BeginRowNum) & " to " & CStr(Me.EndRowNum) & ") ���͒l : " & CStr(num)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�v�f�A�N�Z�X���s���̃G���[���b�Z�[�W��Ԃ�</summary>
''' <param name="column">�A�N�Z�X��</param>
''' <param name="row">�A�N�Z�X��</param>
''' <param name="value">����/�o�� �l</param>
''' <returns>���b�Z�[�W</returns>
Private Function GetMsgDataAccessError(ByVal column As Long, ByVal row As Long, ByVal value As Variant) As String
Dim err_msg As String : err_msg = "[" & Me.TableName & "] �̗v�f�ւ̃A�N�Z�X�Ɏ��s���܂���" & vbCrLf

If (IsNull(value)) Then
err_msg = err_msg & "�v�f��� : Variant [Null]" & vbCrLf
Else
err_msg = err_msg & "�v�f��� : " & TypeName(value) & " [" & CStr(value) & "]" & vbCrLf
End IF

GetMsgDataAccessError = err_msg & "[Size] : (" & Cstr(Me.ColumnSize) & ", " & CStr(Me.RowSize) & ")" & vbCrLf & "[Input] : (" & Cstr(column) & ", " & CStr(row) & ")"
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�^�C�g�����̏����l����͂���</summary>
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
''' <summary>�@��^�C�g�����̏����l����͂���</summary>
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
''' <summary>Index����A�Ώۋ@��̋@�헓�ł̗�ԍ���Ԃ�</summary>
''' <param name="model_index">�@�햼,�������̓e�[�u���ɋL������Ă���@�헓��ԍ�</param>
Private Function GetTargetModelColumn(ByVal model_index As Variant) As Long
Const FUNCTION_NAME As String = "GetTargetModelColumn()"

Dim model_column As Long, type_code As Long : type_code = VarType(model_index)
If (type_code = vbString) Then '�@�햼�ł̑I��
model_column = Me.ModelTitle.Rows(CL_MODEL_NAME_TITLE_ROW_NUM).Search(model_index, 0, True, False)

If (model_column < Me.ModelTitle.BeginColumnNum) Then
Call Log.WriteLog(Log.WARNING_LOG_LEVEL, Me, FUNCTION_NAME, "[" & model_index & "] not found in [" & Me.TableName & "].")
Call ThrowArgumentException(Me, FUNCTION_NAME, "model_index", "[" & Me.TableName & "]" & vbCrLf & "���݂��Ȃ��@�햼�Q�� value = " & model_index)
End If
ElseIf ((type_code = vbLong) Or (type_code = vbInteger)) Then '��ԍ��ł̑I��
model_column = CLng(model_index)

If ((model_column < Me.ModelTitle.BeginColumnNum) Or (Me.ModelTitle.EndColumnNum < model_column)) Then
Call Log.WriteLog(Log.WARNING_LOG_LEVEL, Me, FUNCTION_NAME, "[" & CStr(model_index) & "] is out of range for [" & Me.TableName & "].")
Call ThrowArgumentOutOfRangeException(Me, FUNCTION_NAME, "model_index", "[" & Me.TableName & "]" & vbCrLf & "�͈͊O�Q�� value = " & CStr(model_index))
End If
Else
Call Log.WriteLog(Log.WARNING_LOG_LEVEL, Me, FUNCTION_NAME, "[" & TypeName(model_index) & " : " & CStr(model_index) & "] is incorrect type.")
Call ThrowArgumentException(Me, FUNCTION_NAME, "model_index", "[" & Me.TableName & "]" & vbCrLf & "���͒l�^ �s�� type = " & TypeName(model_index))
End If

GetTargetModelColumn = model_column
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u�����Ńf�[�^���J�n����s�ԍ�</summary>
Private Property Get StartDataRowNum() As Long
StartDataRowNum = CL_DATA_STARTING_ROW_NUM
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u�����ŋ@�퍷�وȊO�̃f�[�^���J�n�����ԍ�</summary>
Private Property Get StartDataColumnNum() As Long
StartDataColumnNum = Me.NumModel
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u�����Ń^�C�g�������I������s�ԍ�</summary>
Private Property Get FinishTitleRowNum() As Long
FinishTitleRowNum = StartDataRowNum - 1
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u�����ŋ@�퍷�ق��I�������ԍ�</summary>
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
''' <summary>�񐔂��Đݒ肷��</summary>
''' <param name="size">�Đݒ肷���</param>
Function ResetColumnSize(ByVal size As Long)
Const FUNCTION_NAME As String = "ResetColumnSize()"
If (size < 1) Then
Call ThrowArgumentOutOfRangeException(Me, FUNCTION_NAME, "size", "[" & CStr(size) & " < 1] : ���͒l���͈͊O�ł�")
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
''' <summary>�s�����Đݒ肷��</summary>
''' <param name="size">�Đݒ肷��s��</param>
Function ResetRowSize(ByVal size As Long)
Const FUNCTION_NAME As String = "ResetRowSize()"
If (size < 1) Then
Call ThrowArgumentOutOfRangeException(Me, FUNCTION_NAME, "size", "[" & CStr(size) & " < 1] : ���͒l���͈͊O�ł�")
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
''' <summary>�ʂ̃e�[�u�����R�s�[����</summary>
''' <param name="table_data">�R�s�[����e�[�u��</param>
Function Copy(ByRef table_data As TableDataBase)
Const FUNCTION_NAME As String = "Copy()"
If (table_data Is Nothing) Then
Call ThrowArgumentNullException(Me, FUNCTION_NAME, "table_data", "�R�s�[�Ɏ��s���܂���")
End If
If (table_data.Name <> Me.Name) Then
Call ThrowArgumentException(Me, FUNCTION_NAME, "table_data", "[" & table_data.Name & "] �ΏۊO�̃I�u�W�F�N�g�ł�")
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
''' <summary>�e�[�u�������ւ���</summary>
''' <param name="table_data">����ւ��Ώ�</param>
Public Function Swap(ByRef table_data As TableDataBase)
Const FUNCTION_NAME As String = "Swap()"
If (table_data Is Nothing) Then
Call ThrowArgumentNullException(Me, FUNCTION_NAME, "table_data", "�R�s�[�Ɏ��s���܂���")
End If
If (table_data.Name <> Me.Name) Then
Call ThrowArgumentException(Me, FUNCTION_NAME, "table_data", "[" & table_data.Name & "] �ΏۊO�̃I�u�W�F�N�g�ł�")
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
''' <summary>������ւ���</summary>
''' <param name="num_0">����ւ��Ώ�</param>
''' <param name="num_1">����ւ��Ώ�</param>
Public Function SwapColumn(ByVal num_0 As Long, ByVal num_1 As Long)
Const FUNCTION_NAME As String = "SwapColumn()"
On Error Goto CatchErr
Call f_data_area.SwapColumn(num_0, num_1)
Exit Function
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�s�����ւ���</summary>
''' <param name="num_0">����ւ��Ώ�</param>
''' <param name="num_1">����ւ��Ώ�</param>
Public Function SwapRow(ByVal num_0 As Long, ByVal num_1 As Long)
Const FUNCTION_NAME As String = "SwapRow()"
On Error Goto CatchErr
Call f_all_data_area.SwapRow(num_0, num_1)
Exit Function
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u���ɗ��}������B�ł���΁A���x���J��Ԃ��ČĂ΂�Ȃ��悤�ɂ��邱�ƁB</summary>
''' <param name="insert_size">�}����</param>
''' <param name="insert_pos">�}���ʒu�Bis_preserve��False�̏ꍇ�́A�����I�ɍŏI�ʒu�ւ̒ǉ��ƂȂ�B</param>
''' <param name="is_preserve">�f�[�^�̍Ċm�ێ��Ƀf�[�^�������p�����ǂ����B</param>
Public Function InsertColumn(ByVal insert_size As Long, Optional ByVal insert_pos = -1, Optional ByVal is_preserve As Boolean = True)
Const FUNCTION_NAME As String = "InsertColumn()"

If (insert_size = 0) Then
Exit Function
End If
If (insert_size < 0) Then
Call ThrowArgumentException(Me, FUNCTION_NAME, "insert_size", "[" & Me.TableName & "] �ǉ��������Ȃ����܂�")
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
''' <summary>�e�[�u���ɍs��}������B�ł���΁A���x���J��Ԃ��ČĂ΂�Ȃ��悤�ɂ��邱�ƁB</summary>
''' <param name="insert_size">�}����</param>
''' <param name="insert_pos">�}���ʒu�Bis_preserve��False�̏ꍇ�́A�����I�ɍŏI�ʒu�ւ̒ǉ��ƂȂ�B</param>
''' <param name="is_preserve">�f�[�^�̍Ċm�ێ��Ƀf�[�^�������p�����ǂ����B</param>
Function InsertRow(ByVal insert_size As Long, Optional ByVal insert_pos = -1, Optional ByVal is_preserve As Boolean = True)
Const FUNCTION_NAME As String = "InsertRow()"

If (insert_size = 0) Then
Exit Function
End If
If (insert_size < 0) Then
Call ThrowArgumentException(Me, FUNCTION_NAME, "insert_size", "[" & Me.TableName & "] �ǉ��������Ȃ����܂�")
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
''' <summary>�e�[�u���ɗ��}������B�ł���΁A���x���J��Ԃ��ČĂ΂�Ȃ��悤�ɂ��邱�ƁB</summary>
''' <param name="add_data">�ǉ����e</param>
''' <param name="add_pos">�ǉ��ʒu�B���w��̏ꍇ�͍ŏI�ʒu�ւƒǉ�����</param>
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
''' <summary>�e�[�u���ɍs��}������B�ł���΁A���x���J��Ԃ��ČĂ΂�Ȃ��悤�ɂ��邱�ƁB</summary>
''' <param name="add_data">�ǉ����e</param>
''' <param name="add_pos">�ǉ��ʒu�B���w��̏ꍇ�͍ŏI�ʒu�ւƒǉ�����</param>
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
''' <summary>�e�[�u���̐擪���w���C�e���[�^��Ԃ�</summary>
''' <returns>�C�e���[�^</returns>
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
''' <summary>�e�[�u���̖����̎����w���C�e���[�^��Ԃ�</summary>
''' <returns>�C�e���[�^</returns>
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
''' <summary>�e�[�u���̒l��z��`���ɕϊ�����</summary>
''' <returns>�z��`���̃e�[�u���̃f�[�^</returns>
Public Function ToArray() As Variant()
Const FUNCTION_NAME As String = "ToArray()"
On Error GoTo CatchErr
ToArray = Me.Table.ToArray
Exit Function
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>��</summary>
Public Property Get ColumnSize() As Long
ColumnSize = f_table.ColumnSize - StartDataColumnNum
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�s��</summary>
Public Property Get RowSize() As Long
RowSize = f_table.RowSize - StartDataRowNum
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�擪�ԍ�</summary>
Public Property Get BeginNum() As Position
BeginNum = ToPosition(Me.BeginColumnNum, Me.BeginRowNum)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�����ԍ�</summary>
Public Property Get EndNum() As Position
EndNum = ToPosition(Me.EndColumnNum, Me.EndRowNum)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�擪��ԍ�</summary>
Public Property Get BeginColumnNum() As Long
BeginColumnNum = 0
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�擪�s�ԍ�</summary>
Public Property Get BeginRowNum() As Long
BeginRowNum = 0
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>������ԍ�</summary>
Public Property Get EndColumnNum() As Long
EndColumnNum = me.ColumnSize - 1
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�����s�ԍ�</summary>
Public Property Get EndRowNum() As Long
EndRowNum = Me.RowSize - 1
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u���Q��.Data()�ŎQ�Ƃł���͈�</summary>
Public Property Get Table() As TableDataRange
Const FUNCTION_NAME As String = "Get Table()"
On Error GoTo CatchErr
Set Table = f_data_area.Table
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u���Q��.Data()�ŎQ�Ƃł���͈�</summary>
Public Property Set Table(ByRef table_range As TableDataRange)
Const FUNCTION_NAME As String = "Set Table()"
On Error GoTo CatchErr
Set f_data_area.Table = table_range
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�͈͎Q��</summary>
Public Property Get Range(ByRef head As Position, ByRef tail As Position) As TableDataRange
Const FUNCTION_NAME As String = "Get Range()"
On Error GoTo CatchErr
Set Range = f_data_area.Range(head, tail)
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�͈͎Q��</summary>
Public Property Set Range(ByRef head As Position, ByRef tail As Position, ByRef table_range As TableDataRange)
Const FUNCTION_NAME As String = "Set Range()"
On Error GoTo CatchErr
Set f_data_area.Range(head, tail) = table_range
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>1��͈͎Q��</summary>
Public Property Get Columns(ByVal column As Long) As TableDataColumn
Const FUNCTION_NAME As String = "Get Columns()"
On Error GoTo CatchErr
Set Columns = f_data_area.Columns(column)
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>1��͈͎Q��</summary>
Public Property Set Columns(ByVal column As Long, ByRef table_range As TableDataColumn)
Const FUNCTION_NAME As String = "Set Columns()"
On Error GoTo CatchErr
Set f_data_area.Columns(column) = table_range
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>1�s�͈͎Q��</summary>
Public Property Get Rows(Byval row As Long) As TableDataRow
Const FUNCTION_NAME As String = "Get Rows()"
On Error GoTo CatchErr
Set Rows = f_data_area.Rows(row)
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>1�s�͈͎Q��</summary>
Public Property Set Rows(Byval row As Long, ByRef table_range As TableDataRow)
Const FUNCTION_NAME As String = "Set Rows()"
On Error GoTo CatchErr
Set f_data_area.Rows(row) = table_range
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�͈̓`�F�b�N�t���v�f�A�N�Z�X</summary>
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
''' <summary>�͈̓`�F�b�N�t���v�f�A�N�Z�X</summary>
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
''' <summary>�v�f�A�N�Z�X</summary>
Public Property Get Data(ByVal column As Long, ByVal row As Long) As Variant
Const FUNCTION_NAME As String = "Get Data()"
On Error GoTo CatchErr
Data = f_data_area.Data(column, row)
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME, GetMsgDataAccessError(column, row, ""))
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�v�f�A�N�Z�X</summary>
Public Property Let Data(ByVal column As Long, ByVal row As Long, ByRef value As Variant)
Const FUNCTION_NAME As String = "Set Data()"
On Error GoTo CatchErr
f_data_area.Data(column, row) = value
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME, GetMsgDataAccessError(column, row, value))
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u����</summary>
Public Property Get TableName() As String
TableName = f_table.TableName
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u����</summary>
Public Property Let TableName(ByVal table_name As String)
If (Me.TableName <> table_name) Then
Call Log.WriteLog(Log.INFO_LOG_LEVEL, Me, "TableName()", "[" & Me.TableName & "] to [" & table_name & "]")
End If

f_table.TableName = table_name
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�I�u�W�F�N�g��</summary>
Public Property Get Name() As String
Name = TypeName(Me)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�_�E���L���X�g�p</summary>
Public Property Get DownCast() As Object
Set DownCast = Me
End Property

'----------------------------------------------------------------------------------------------------
' �Ǝ�����
'----------------------------------------------------------------------------------------------------
''' <summary>�`�F�b�N���X�g�e�[�u��������������</summary>
''' <param name="num_model">�@�퐔</param>
''' <param name="table_name">�e�[�u����</param>
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
Call RethrowException(Me, FUNCTION_NAME, "�����ݒ�Ɏ��s���܂���")
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�ʃe�[�u���̃^�C�g�������R�s�[����</summary>
''' <param name="table_data">�R�s�[����e�[�u��</param>
Public Function CopyTitle(ByRef table_data As CheckListData)
Const FUNCTION_NAME As String = "CopyTitle()"

If (table_data Is Nothing) Then
Call ThrowArgumentNullException(Me, FUNCTION_NAME, "table_data", "�^�C�g���̃R�s�[�Ɏ��s���܂���")
End If

Call Log.WriteLog(Log.INFO_LOG_LEVEL, Me, FUNCTION_NAME, "Start copying the title of [" & table_data.TableName & "] to [" & Me.TableName & "].")

Set Me.ModelTitle.Table = table_data.ModelTitle.Table
Set Me.Title.Table = table_data.Title.Table

Call Log.WriteLog(Log.INFO_LOG_LEVEL, Me, FUNCTION_NAME, "Successfully copied the title to [" & Me.TableName & "].")
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�V�����@��̗��ǉ�����</summary>
''' <param name="new_model_name">�ǉ�����@�햼</param>
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
''' <summary>���͂���Ă���f�[�^����@�퐔�𒲂ׂ�</summary>
''' <returns>���ׂ��@�퐔�BNumModel�ɂ����f�����</returns>
Public Function CheckNumModel() As Long
f_num_model = f_table.Rows(CL_ENG_TITLE_ROW_NUM).Search(ENG_PRIORITY_COLUMN_TITLE)
CheckNumModel = f_num_model
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�`�F�b�N���X�g�e�[�u���𕪊������e�[�u���͈̔͂�ݒ肷��B</summary>
Public Function SetTableArea()
Call f_all_title_area.Initialize(f_table, f_table.BeginNum, ToPosition(f_table.EndColumnNum, FinishTitleRowNum), "All Title")
Call f_all_data_area.Initialize(f_table, ToPosition(f_table.BeginColumnNum, StartDataRowNum), f_table.EndNum, "All Data")

Call f_model_title_area.Initialize(f_table, f_table.BeginNum, ToPosition(FinishModelColumnNum, FinishTitleRowNum), "Model Title")
Call f_model_data_area.Initialize(f_table, ToPosition(f_table.BeginColumnNum, StartDataRowNum), ToPosition(FinishModelColumnNum, f_table.EndRowNum), "Model Data")
Call f_title_area.Initialize(f_table, ToPosition(StartDataColumnNum, f_table.BeginRowNum), ToPosition(f_table.EndColumnNum, FinishTitleRowNum), "Title")
Call f_data_area.Initialize(f_table, ToPosition(StartDataColumnNum, StartDataRowNum), f_table.EndNum, "Data")
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>CheckListData������1�@�핪�̃f�[�^�����o��</summary>
''' <param name="model_index">�@�햼,�������̓e�[�u���ɋL������Ă���@�헓��ԍ�</param>
''' <returns>�ϊ���e�[�u��</returns>
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
Call RethrowException(Me, FUNCTION_NAME, "�@��f�[�^�̎��o���Ɏ��s���܂����B")
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>���ٗ����܂߂��`�F�b�N���X�g�e�[�u���ɓ��͂���Ă���f�[�^�S�̂̍s�͈͎Q��</summary>
Public Property Get TableRows(ByVal row As Long) As TableDataRow
Const FUNCTION_NAME As String = "Get TableRows()"
On Error GoTo CatchErr
Set TableRows = Me.AllData.Rows(row)
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>���ٗ����܂߂��`�F�b�N���X�g�e�[�u���ɓ��͂���Ă���f�[�^�S�̂̍s�͈͎Q��</summary>
Public Property Set TableRows(ByVal row As Long, ByRef table_range As TableDataRow)
Const FUNCTION_NAME As String = "Set TableRows()"

If (table_range Is Nothing) Then
Call ThrowArgumentNullException(Me, FUNCTION_NAME, "table_range", "�R�s�[�Ɏ��s���܂���")
End If
If (table_range.Table.Name <> Me.Name) Then
Call ThrowArgumentException(Me, FUNCTION_NAME, "table_range", "�Q�Ƃ��Ă���e�[�u���^���ΏۊO�ł�")
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
''' <summary>���ٗ����܂߂��`�F�b�N���X�g�e�[�u���ɓ��͂���Ă���f�[�^�S�͈͎̂̔Q��</summary>
Public Property Get TableRange(ByRef head As Position, ByRef tail As Position) As TableDataRange
Const FUNCTION_NAME As String = "Set TableRange()"
On Error GoTo CatchErr
Set TableRange = Me.AllData.Range(head, tail)
Exit Property
CatchErr:
Call RethrowException(Me, FUNCTION_NAME)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>���ٗ����܂߂��`�F�b�N���X�g�e�[�u���ɓ��͂���Ă���f�[�^�S�͈͎̂̔Q��</summary>
Public Property Set TableRange(ByRef head As Position, ByRef tail As Position, ByRef table_range As TableDataRange)
Const FUNCTION_NAME As String = "Set TableRange()"
If (table_range Is Nothing) Then
Call ThrowArgumentNullException(Me, FUNCTION_NAME, "table_range", "�R�s�[�Ɏ��s���܂���")
End If
If (table_range.Table.Name <> Me.Name) Then
Call ThrowArgumentException(Me, FUNCTION_NAME, "table_range", "�Q�Ƃ��Ă���e�[�u���^���ΏۊO�ł�")
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
''' <summary>����Ă���TableData�ւ̎Q��</summary>
Public Property Set Value(ByRef table_data As TableData)
Call f_table.Copy(table_data)

Call Me.CheckNumModel()
Call Me.SetTableArea()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>����Ă���TableData�ւ̎Q��</summary>
Public Property Get Value() As TableData
Set Value = f_table
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�@�햼</summary>
Public Property Let ModelName(ByVal index As Long, ByVal model_name As String)
Me.ModelTitle.Data(index, CL_MODEL_NAME_TITLE_ROW_NUM) = model_name
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�@�햼</summary>
Public Property Get ModelName(ByVal index As Long) As String
ModelName = Me.ModelTitle.Data(index, CL_MODEL_NAME_TITLE_ROW_NUM)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>���ٗ��̃^�C�g���S�̂�\��</summary>
Public Property Get ModelTitle() As TableDataArea
Set ModelTitle = f_model_title_area
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>���ٗ��ȊO�̃^�C�g���S�̂�\��</summary>
Public Property Get Title() As TableDataArea
Set Title = f_title_area
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>���ٗ��̃f�[�^�S�̂�\��</summary>
Public Property Get Model() As TableDataArea
Set Model = f_model_data_area
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>���ٗ����܂߂��`�F�b�N���X�g�e�[�u���ɓ��͂���Ă���f�[�^�S�̂�\��</summary>
Public Property Get AllData() As TableDataArea
Set AllData = f_all_data_area
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>���ٗ����܂߂��`�F�b�N���X�g�e�[�u���̃^�C�g���S�̂�\��</summary>
Public Property Get AllTitle() As TableDataArea
Set AllTitle = f_all_title_area
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�@�퐔��Ԃ�</summary>
Public Property Get NumModel() As Long
NumModel = f_num_model
End Property

'----------------------------------------------------------------------------------------------------
' �p���p
'----------------------------------------------------------------------------------------------------
''' <summary>�񐔂��Đݒ肷��</summary>
''' <param name="size">�Đݒ肷���</param>
Public Function TableDataBase_ResetColumnSize(ByVal size As Long)
Call Me.ResetColumnSize(size)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�s�����Đݒ肷��</summary>
''' <param name="size">�Đݒ肷��s��</param>
Public Function TableDataBase_ResetRowSize(ByVal size As Long)
Call Me.ResetRowSize(size)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�ʂ̃e�[�u�����R�s�[����</summary>
''' <param name="table_data">�R�s�[����e�[�u��</param>
Public Function TableDataBase_Copy(ByRef table_data As TableDataBase)
Call Me.Copy(table_data)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u�������ւ���</summary>
''' <param name="table_data">����ւ��Ώ�</param>
Public Function TableDataBase_Swap(ByRef table_data As TableDataBase)
Call Me.Swap(table_data)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>������ւ���</summary>
''' <param name="num_0">����ւ��Ώ�</param>
''' <param name="num_1">����ւ��Ώ�</param>
Public Function TableDataBase_SwapColumn(ByVal num_0 As Long, ByVal num_1 As Long)
Call Me.SwapColumn(num_0, num_1)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�s�����ւ���</summary>
''' <param name="num_0">����ւ��Ώ�</param>
''' <param name="num_1">����ւ��Ώ�</param>
Public Function TableDataBase_SwapRow(ByVal num_0 As Long, ByVal num_1 As Long)
Call Me.SwapRow(num_0, num_1)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u���ɗ��}������B�ł���΁A���x���J��Ԃ��ČĂ΂�Ȃ��悤�ɂ��邱�ƁB</summary>
''' <param name="insert_size">�}����</param>
''' <param name="insert_pos">�}���ʒu�Bis_preserve��False�̏ꍇ�́A�����I�ɍŏI�ʒu�ւ̒ǉ��ƂȂ�B</param>
''' <param name="is_preserve">�f�[�^�̍Ċm�ێ��Ƀf�[�^�������p�����ǂ����B</param>
Public Function TableDataBase_InsertColumn(ByVal insert_size As Long, Optional insert_pos = -1, Optional ByVal is_preserve As Boolean = True)
Call Me.InsertColumn(insert_size, insert_pos, is_preserve)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u���ɍs��}������B�ł���΁A���x���J��Ԃ��ČĂ΂�Ȃ��悤�ɂ��邱�ƁB</summary>
''' <param name="insert_size">�}����</param>
''' <param name="insert_pos">�}���ʒu�Bis_preserve��False�̏ꍇ�́A�����I�ɍŏI�ʒu�ւ̒ǉ��ƂȂ�B</param>
''' <param name="is_preserve">�f�[�^�̍Ċm�ێ��Ƀf�[�^�������p�����ǂ����B</param>
Public Function TableDataBase_InsertRow(ByVal insert_size As Long, Optional insert_pos = -1, Optional ByVal is_preserve As Boolean = True)
Call Me.InsertRow(insert_size, insert_pos, is_preserve)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u���ɗ��}������B�ł���΁A���x���J��Ԃ��ČĂ΂�Ȃ��悤�ɂ��邱�ƁB</summary>
''' <param name="add_data">�ǉ����e</param>
''' <param name="add_pos">�ǉ��ʒu�B���w��̏ꍇ�͍ŏI�ʒu�ւƒǉ�����</param>
Public Function TableDataBase_AddColumn(ByRef add_data As TableDataColumn, Optional add_pos = -1)
Call Me.AddColumn(add_data, add_pos)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u���ɍs��}������B�ł���΁A���x���J��Ԃ��ČĂ΂�Ȃ��悤�ɂ��邱�ƁB</summary>
''' <param name="add_data">�ǉ����e</param>
''' <param name="add_pos">�ǉ��ʒu�B���w��̏ꍇ�͍ŏI�ʒu�ւƒǉ�����</param>
Public Function TableDataBase_AddRow(ByRef add_data As TableDataRow, Optional add_pos = -1)
Call Me.AddRow(add_data, add_pos)
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u���̊J�n�ʒu�̃C�e���[�^��Ԃ�</summary>
''' <returns>�C�e���[�^</returns>
Public Function TableDataBase_BeginIterator() As TableDataIterator
Set TableDataBase_BeginIterator = Me.BeginIterator()
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u���̏I���ʒu�̃C�e���[�^��Ԃ�</summary>
''' <returns>�C�e���[�^</returns>
Public Function TableDataBase_EndIterator() As TableDataIterator
Set TableDataBase_EndIterator = Me.EndIterator()
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u���̒l��z��`���ɕϊ�����</summary>
''' <returns>�z��`���̃e�[�u���̃f�[�^</returns>
Public Function TableDataBase_ToArray() As Variant()
TableDataBase_ToArray = Me.ToArray()
End Function

'----------------------------------------------------------------------------------------------------
''' <summary>��</summary>
Public Property Get TableDataBase_ColumnSize() As Long
TableDataBase_ColumnSize = Me.ColumnSize()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�s��</summary>
Public Property Get TableDataBase_RowSize() As Long
TableDataBase_RowSize = Me.RowSize()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�擪�ԍ�</summary>
Public Property Get TableDataBase_BeginNum() As Position
TableDataBase_BeginNum = Me.BeginNum()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�����ԍ�</summary>
Public Property Get TableDataBase_EndNum() As Position
TableDataBase_EndNum = Me.EndNum()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�擪��ԍ�</summary>
Public Property Get TableDataBase_BeginColumnNum() As Long
TableDataBase_BeginColumnNum = Me.BeginColumnNum()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�擪�s�ԍ�</summary>
Public Property Get TableDataBase_BeginRowNum() As Long
TableDataBase_BeginRowNum = Me.BeginRowNum()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>������ԍ�</summary>
Public Property Get TableDataBase_EndColumnNum() As Long
TableDataBase_EndColumnNum = Me.EndColumnNum()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�����s�ԍ�</summary>
Public Property Get TableDataBase_EndRowNum() As Long
TableDataBase_EndRowNum = Me.EndRowNum()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u���Q��</summary>
Public Property Get TableDataBase_Table() As TableDataRange
Set TableDataBase_Table = Me.Table
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u���Q��</summary>
Public Property Set TableDataBase_Table(ByRef table_range As TableDataRange)
Set Me.Table = table_range
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�͈͎Q��</summary>
Public Property Get TableDataBase_Range(ByRef head As Position, ByRef tail As Position) As TableDataRange
Set TableDataBase_Range = Me.Range(head, tail)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�͈͎Q��</summary>
Public Property Set TableDataBase_Range(ByRef head As Position, ByRef tail As Position, ByRef table_range As TableDataRange)
Set Me.Range(head, tail) = table_range
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>1��͈͎Q��</summary>
Public Property Get TableDataBase_Columns(ByVal column As Long) As TableDataColumn
Set TableDataBase_Columns = Me.Columns(column)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>1��͈͎Q��</summary>
Public Property Set TableDataBase_Columns(ByVal column As Long, ByRef table_column As TableDataColumn)
Set Me.Columns(column) = table_column
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>1�s�͈͎Q��</summary>
Public Property Get TableDataBase_Rows(Byval row As Long) As TableDataRow
Set TableDataBase_Rows = Me.Rows(row)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>1�s�͈͎Q��</summary>
Public Property Set TableDataBase_Rows(Byval row As Long, ByRef table_row As TableDataRow)
Set Me.Rows(row) = table_row
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�͈̓`�F�b�N�t���v�f�A�N�Z�X</summary>
Public Property Let TableDataBase_At(ByVal column As Long, ByVal row As Long, ByVal value As Variant)
Me.At(column, row) = value
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�͈̓`�F�b�N�t���v�f�A�N�Z�X</summary>
Public Property Get TableDataBase_At(ByVal column As Long, ByVal row As Long) As Variant
TableDataBase_At = Me.At(column, row)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�v�f�A�N�Z�X</summary>
Public Property Let TableDataBase_Data(ByVal column As Long, ByVal row As Long, ByVal value As Variant)
Me.Data(column, row) = value
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�v�f�A�N�Z�X</summary>
Public Property Get TableDataBase_Data(ByVal column As Long, ByVal row As Long) As Variant
TableDataBase_Data = Me.Data(column, row)
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u����</summary>
Public Property Get TableDataBase_TableName() As String
TableDataBase_TableName = Me.TableName()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�e�[�u����</summary>
Public Property Let TableDataBase_TableName(ByVal table_name As String)
Me.TableName = table_name
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�I�u�W�F�N�g��</summary>
Public Property Get TableDataBase_Name() As String
TableDataBase_Name = Me.Name()
End Property

'----------------------------------------------------------------------------------------------------
''' <summary>�_�E���L���X�g�p</summary>
Public Property Get TableDataBase_DownCast() As Object
Set TableDataBase_DownCast = Me.DownCast()
End Property

'----------------------------------------------------------------------------------------------------



