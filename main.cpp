
#include <iostream>
#include <string>
#include <list>
#include <regex>

enum class AccessAdornment
{
	PUBLIC_ACCESS,
	PRIVATE_ACCESS,
	PROTECTED_ACCESS,
	FRIEND_ACCESS
};

struct ArgumentElement
{
	std::string _Type;
	std::string _ArgName;
	bool _IsParms;
	bool _IsBaseValue;
	std::string _BaseValue;

	void SplitElement(std::smatch _Code)
	{

	}
};

const std::string _ArgumentMatchCode = 
"(ByRef|ByVal|Optional|ParmArray)?[ ]?";

struct FunctionElement
{
	AccessAdornment _Access;
	std::string _FunctionName;
	std::string _ReturnType;
	std::list<ArgumentElement> _Args;
	std::string _Code;

	void SplitElement(std::smatch _Code)
	{
		if ((_Code[1] == true) && (_Code[1].str() == "Private"))		this->_Access = AccessAdornment::PRIVATE_ACCESS;
		else if ((_Code[1] == true) && (_Code[1].str() == "Friend"))	this->_Access = AccessAdornment::FRIEND_ACCESS;
		else															this->_Access = AccessAdornment::PUBLIC_ACCESS;

		this->_FunctionName = _Code[3].str();
		this->_ReturnType = _Code[5].str();

		this->_Code = _Code[7].str();
	}
};

/*
Match
[1] アクセス修飾子
[2] Function|Sub //どうでもいい
[3] 関数名
[4] 引数部分
[5] 戻り値の部分
[6] 戻り値 型名
[7] 関数内コード
[8] Function|Sub //どうでもいい
*/
const std::string _FunctionMatchCode =
R"((Public|Private|Friend)?[ ]?(Function|Sub) (\w+?)[(]([a-zA-Z_0-9=(), ]*?)[)]( As (\w+?))?\n([\w\W\s]*?)End (Function|Sub))";

std::list<FunctionElement> SplitFunctionElement(std::string _Code)
{
	std::regex _Regex(_FunctionMatchCode, std::regex_constants::icase);
	std::list<std::smatch> _List;

	for (std::sregex_iterator _Itr(std::begin(_Code), std::end(_Code), _Regex), _End; _Itr != _End; _Itr ++)
	{
		_List.push_back(*_Itr);
	}

	return std::list<FunctionElement>();
}

int main()
{
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
}
