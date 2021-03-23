
#include <iostream>
#include <string>
#include <list>
#include <regex>
#include <map>

/**
 * 解析上無視する項目
 *  デフォルトプロパティ
 * -> よくわからん。無視。
 *  Attribute
 * -> 後でつけろ
 * 解析の流れ
 * -> ファイルを読み込む(一体で)
 * -> 拡張子、Attributeからオブジェクト名の取り込み
 * -> 行分割、結合を戻す(行ごとに分割したリストにする)
 * -> クラス(名前,継承関係のみ)、構造体、列挙型を取り込む
 * -> グローバル変数(定数)、
 *    コメント、
 *    #If,#Const、
 *    関数(関数名、戻り値、引数)を取り込む(ファイル内の順番も保持する)
 * -> メソッド、プロパティ、フィールドを取り込む
 * -> 関数、メソッド内の取り込み
 *  -> 変数宣言(VB6の変数宣言はどこに記入しても関数の先頭で行う扱い)
 *  -> 関数の上からブロックを取り込む
 * -> C#への置き換え
 * 分割要素
 * -> 関数外
 *  -> 列挙型
 *   -> 列挙子
 *  -> 構造体
 *   -> 変数
 * 
 * 配列 扱い
 変数保持オブジェクトに配列を表すフィールドを用意する?
 型名から推測する?

型名 扱い
 認識した型名と、型オブジェクトをリストに登録する。
 ライブラリの型はファイルに記載し、実行時に読み取り、リストに登録

コメント 扱い
 行の扱い。以下の3つで持つ。
  インデント 正直どうでいい気がする。変換時に新しくネストすればよいとも思う。
  処理
  コメント 文字列内に ‘ がある場合どのように分割するか? コメントより文字列が優先されるため、” の数を確認しながら分割する。

 ソース改行記号の後にコメントは記入出来ない(例 以下のようにはできない)
Call MsgBox(“test ” & _ ‘NG
 	“B”
 コメントの末尾にソース改行記号がある場合、改行先もコメントとなる。
 A = 1 ‘ コメント _
   この行もコメント

 ファイルの読み取り時に行の分割、結合を解除で問題ない

 コードの読み取り方。
 1. 行途中での改行、複数行の1行へのまとめを解除
 2. ファイル名から、クラスモジュール、標準モジュールを分別。クラス名だけ、読み取っておく。
 3. 標準モジュールのファイルを上から順に、「型定義、関数定義、変数宣言、コメント等」で分別して読み取り、スタックしていく。
    読み込んだ、型、グローバル変数(定数)は別でリストを作成しておく。
 4. クラスモジュールのファイルを上から順に、「メンバー、メソッド、コメント等」で分別して読み取り、スタックしていく。
 * 
# If文などの読み取り方。
 まず、関数ごとに読み込んだブロック毎に読む。
 If文などが出現すると、Blockオブジェクトを作成し、終了地点が出現するまで、行の取り込みを開始する。
 読み込み途中で、上の処理を再帰的に呼び出す。
 If文などが出現しないまま、終了地点が来たら、ループから抜け、上のループに戻す。
 これ、本当に再帰で数千行はあるようなお排泄物プロシージャを取り込んだら、メモリがやばいことになるかも
**/

/// <summary>アクセス修飾子</summary>
enum class AccessAdornment
{
	LOCAL_ACCESS,
	PUBLIC_ACCESS,
	PRIVATE_ACCESS,
	FRIEND_ACCESS,
};

std::string ToAceessString(AccessAdornment _Access)
{
	switch (_Access)
	{
		case AccessAdornment::LOCAL_ACCESS:		return "";
		case AccessAdornment::PUBLIC_ACCESS:	return "public";
		case AccessAdornment::PRIVATE_ACCESS:	return "private";
		case AccessAdornment::FRIEND_ACCESS:	return "friend";
		default: break;
	}
}

/// <summary>演算子</summary>
enum class Operator
{
	MULTIPLY_OPERATOR,	//乗算
	RAISING_OPERATOR,	//累乗
	DIVIDE_OPERATOR,	//除算
	SURPLUS_OPERATOR,	//余り
	PLUS_OPERATOR,		//加算
	MINUS_OPERATOR,		//減算
	EQUAL_OPERATOR,		//等価
	IS_OPERATOR,		//型互換確認
	LIKE_OPRATOR,		//文字列比較(正規表現で再現する)
	AND_OPERATOR,		// A & B
	OR_OPERATOR,		// A | B
	NOT_OPRATOR,		// A ! B
	XOR_OPERATOR,		// A ^ B
	EQV_OPRATOR,		// !(A ^ B)
	IMP_OPRATOR,		// (!A) | B
};

/// <summary></summary>
/// <param name="_Operator"></param>
/// <param name="_RightValue"></param>
/// <param name="_LeftValue"></param>
/// <returns></returns>
std::string ToString(Operator _Operator, std::string _RightValue, std::string _LeftValue)
{
	switch (_Operator)
	{
		case Operator::MULTIPLY_OPERATOR:	return _RightValue + " * " + _LeftValue;
		case Operator::RAISING_OPERATOR:	return "Math.Pow(" + _RightValue + ", " + _LeftValue + ")";
		case Operator::DIVIDE_OPERATOR:		return _RightValue + " / " + _LeftValue;
		case Operator::SURPLUS_OPERATOR:	return _RightValue + " % " + _LeftValue;
		case Operator::PLUS_OPERATOR:		return _RightValue + " + " + _LeftValue;
		case Operator::MINUS_OPERATOR:		return _RightValue + " - " + _LeftValue;
		case Operator::EQUAL_OPERATOR:		return _RightValue + " = " + _LeftValue;
		case Operator::IS_OPERATOR:			return _RightValue + " is " + _LeftValue;
		case Operator::LIKE_OPRATOR:		return _RightValue + "" + _LeftValue;
		case Operator::NOT_OPRATOR:			return " !" + _RightValue;
		case Operator::AND_OPERATOR:		return _RightValue + " && " + _LeftValue;
		case Operator::OR_OPERATOR:			return _RightValue + " || " + _LeftValue;
		case Operator::XOR_OPERATOR:		return _RightValue + " ^ " + _LeftValue;
		case Operator::EQV_OPRATOR:			return "!(" + _RightValue + " ^ " + _LeftValue + ")";
		case Operator::IMP_OPRATOR:			return "(!(" + _RightValue + ") | " + _LeftValue + ")";
	}
}

class VariableElement;

/// <summary>グローバル変数のリスト</summary>
std::map<std::string, VariableElement> _GlobalVariableList; //Name, Variable

/// <summary>変数要素</summary>
class VariableElement
{
	AccessAdornment _Access;

	std::string _TypeName;
	std::string _Name;

public:
	VariableElement() {}

	VariableElement(const std::string _TypeName, const std::string _Name, const AccessAdornment _Access = AccessAdornment::LOCAL_ACCESS)
		: _TypeName (_TypeName), _Name(_Name), _Access(_Access)
	{}

	constexpr std::string& GetTypeName() { return _TypeName; }
	constexpr std::string& GetName() { return _Name; }
	constexpr AccessAdornment GetAccess() { return _Access; }
};

class ConstantElement : public VariableElement
{
	std::string _Value;

public:
	ConstantElement() {}

	ConstantElement(const VariableElement _Variable, const std::string _Value)
		: VariableElement(_Variable), _Value(_Value)
	{}

	constexpr std::string& Value() { return _Value; }
};

/// <summary>行の要素</summary>
class LineElement
{
public:
	std::string ToCShapeCode();
};

/// <summary>生値</summary>
class ValueElement : public LineElement
{
	std::string _Value;

public:
	ValueElement();

	std::string ToCShapeCode() { return _Value; }
};

/// <summary>代入</summary>
class AssignElement : public LineElement
{
private:
	std::shared_ptr<LineElement> _RightSideValue;
	std::shared_ptr<LineElement> _LeftSideValue;
};

/// <summary>演算子要素</summary>
class OperatorElement : public LineElement
{
};

/// <summary>計算要素(一つの括弧で区切れる分)</summary>
class FormulaElement : public LineElement
{
	std::vector<LineElement> _Formula;

};

/// <summary>変数呼び出し</summary>
class CallVariableElement : public LineElement
{
private:
	std::shared_ptr<VariableElement> _Variable;
	
public:
	CallVariableElement();
};

/// <summary>変数宣言</summary>
class DeclarationElement : public LineElement
{
private:
	std::shared_ptr<VariableElement> _Variable;

public:
	DeclarationElement();
};

class FunctionBlock;
/// <summary>関数呼び出し</summary>
class CallFUnctionElement : public LineElement
{
private:
	std::shared_ptr<FunctionBlock> _CallFunction;
	std::map<std::string, LineElement> _ArugmentAssign;
public:
	CallFUnctionElement();
};
 
/// <summary>コメントの部分</summary>
class CommentElement : public LineElement
{
	std::string _Comment;
};

class CodeBlock
{
};

class LineBlock : public CodeBlock
{
	std::vector<LineElement> _Element;
};

/// <summary>If文</summary>
class IfElseBlock : public CodeBlock
{

};

class SwitchBlock : public CodeBlock
{
};

class WhileBlock : public CodeBlock
{
};

class ForBlock : public CodeBlock
{
	std::string _Begin;
	std::string _End;
};

class ArgumentCode : public VariableElement
{
	std::string _BaseValue;

	bool _IsReference;
	bool _IsParmArray;
	bool _IsOptional;
public:
	ArgumentCode() {}
};

const std::string _ArgumentMatchCode = "(ByRef|ByVal|Optional|ParmArray)?[ ]?";

/// <summary>関数のブロック</summary>
class FunctionBlock : public CodeBlock
{
	bool _IsStatic;
	AccessAdornment _Access;
	std::string _FunctionName;
	std::string _ReturnType;
	std::list<ArgumentCode> _Args;
	std::vector<CodeBlock> _Code;

public:
	FunctionBlock();

	void SplitElement(std::smatch _Code)
	{
		if ((_Code[1] == true) && (_Code[1].str() == "Private"))		this->_Access = AccessAdornment::PRIVATE_ACCESS;
		else if ((_Code[1] == true) && (_Code[1].str() == "Friend"))	this->_Access = AccessAdornment::FRIEND_ACCESS;
		else															this->_Access = AccessAdornment::PUBLIC_ACCESS;

		this->_FunctionName = _Code[3].str();
		this->_ReturnType = _Code[5].str();

		//this->_Code = _Code[7].str();
	}
};

class TypeBlock;
/// <summary>プロジェクト内で定義されている型の一覧</summary>
std::map<std::string, TypeBlock> _ProjectTypeList;

/// <summary>型の要素を表す</summary>
class TypeBlock : public CodeBlock
{
protected:
	std::string _Name;
public:
	TypeBlock(const std::string& _Name) : _Name(_Name) {}

	constexpr std::string& GetName() { return TypeBlock::_Name; }
};

/// <summary>列挙型</summary>
class EnumBlock : public TypeBlock
{
	std::list<ConstantElement> _ElementList;
	
public:

	EnumBlock(const std::string& _Name) : TypeBlock(_Name) {}
};

/// <summary>構造体</summary>
class StructBlock : public TypeBlock
{
	std::list<VariableElement> _MemberList;

public:
};

/// <summary>クラス</summary>
class ClassBlock : public TypeBlock
{
	std::string _BaseTypeName; //継承元

	std::list< std::shared_ptr<VariableElement>> _MemberList; //メンバー一覧
	std::list<std::shared_ptr<FunctionBlock>> _MethodList; //メソッド一覧

	std::vector<CodeBlock> _ClassCode; //クラスのコード全体。
public:

};

class FileBlock
{

};

std::list<FunctionBlock> SplitFunctionElement(std::string _Code)
{
	/*
Match
[1] アクセス修飾子
[2] Static
[3] Function|Sub //VBのどうでもいい区分
[4] 関数名
[5] 引数部分
[6] 戻り値の部分
[7] 戻り値 型名
[8] 関数内コード
[9] Function|Sub //どうでもいい
*/
	const std::string _AccessMatchCode = R"((Public|Private|Friend)??\s*)";
	const std::string _StaticMatchCode = R"((Static)??\s*)";
	const std::string _FunctionNameMatchCode = R"((Function|Sub)\s*(\w+?)\s*)";
	const std::string _ArgumentMatchCode = R"([(]([\w=(),\s]*?)[)])";
	const std::string _ReturnMatchCode = R"((\s*As\s*(\w+?))??\n)";
	const std::string _FunctionCodeMatchCode = R"(([\w\W\s]*?)End (Function|Sub))";
	const std::string _FunctionMatchCode = _AccessMatchCode + _StaticMatchCode + _FunctionNameMatchCode + _ArgumentMatchCode + _ReturnMatchCode + _FunctionCodeMatchCode;

	std::regex _Regex(_FunctionMatchCode, std::regex_constants::icase);
	std::list<std::smatch> _List;

	for (std::sregex_iterator _Itr(std::begin(_Code), std::end(_Code), _Regex), _End; _Itr != _End; _Itr ++)
	{
		_List.push_back(*_Itr);
	}

	return std::list<FunctionBlock>();
}

int OperateVB6codeLine(int argc, char* argv[], char* envp[]);

int main(int argc, char* argv[], char* envp[])
{
	return OperateVB6codeLine(argc, argv, envp);
	/*
	SplitFunctionElement(
R"('----------------------------------------------------------------------------------------------------

Function A_FUNC()
	If (vec.size > Ubound(vec.data)) Then	ReDim Preserve vec.data(0 To vec.size * 1.5)
	
	vec.data(vec.size) = value
	vec.size = vec.size + 1
End Function

'----------------------------------------------------------------------------------------------------

Function B_FUNC() As Boolean
	If (vec.size > Ubound(vec.data)) Then	ReDim Preserve vec.data(0 To vec.size * 1.5)
	
	vec.data(vec.size) = value
	vec.size = vec.size + 1
End Function

'----------------------------------------------------------------------------------------------------

Function C_FUNC(ByVal str As String)
	If (vec.size > Ubound(vec.data)) Then	ReDim Preserve vec.data(0 To vec.size * 1.5)
	
	vec.data(vec.size) = value
	vec.size = vec.size + 1
End Function

'----------------------------------------------------------------------------------------------------
'数値が始まる位置を調べる
'str 	: 調べる文字列
'head	: 調べ始める位置(最小値は1)
'return	: 数値が始まる位置
Function D_FUNC(ByVal str As String, Optional head As Long = 1) As Long
	For head = head To Len(str)
		If (Mid(str, head, 1) Like "#") Then	Exit For
	Next
	
	FindValuePosition = head
End Function)");
*/
}
