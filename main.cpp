
#include <iostream>
#include <string>
#include <list>
#include <regex>
#include <map>

#include <boost/version.hpp>
#include <boost/algorithm/string.hpp>

/**
 * ��͏㖳�����鍀��
 *  �f�t�H���g�v���p�e�B
 * -> �悭�킩���B�����B
 *  Attribute
 * -> ��ł���
 * ��̗͂���
 * -> �t�@�C����ǂݍ���(��̂�)
 * -> �g���q�AAttribute����I�u�W�F�N�g���̎�荞��
 * -> �s�����A������߂�(�s���Ƃɕ����������X�g�ɂ���)
 * -> �N���X(���O,�p���֌W�̂�)�A�\���́A�񋓌^����荞��
 * -> �O���[�o���ϐ�(�萔)�A
 *    �R�����g�A
 *    #If,#Const�A
 *    �֐�(�֐����A�߂�l�A����)����荞��(�t�@�C�����̏��Ԃ��ێ�����)
 * -> ���\�b�h�A�v���p�e�B�A�t�B�[���h����荞��
 * -> �֐��A���\�b�h���̎�荞��
 *  -> �ϐ��錾(VB6�̕ϐ��錾�͂ǂ��ɋL�����Ă��֐��̐擪�ōs������)
 *  -> �֐��̏ォ��u���b�N����荞��
 * -> C#�ւ̒u������
 * �����v�f
 * -> �֐��O
 *  -> �񋓌^
 *   -> �񋓎q
 *  -> �\����
 *   -> �ϐ�
 * 
 * �z�� ����
 �ϐ��ێ��I�u�W�F�N�g�ɔz���\���t�B�[���h��p�ӂ���?
 �^�����琄������?

�^�� ����
 �F�������^���ƁA�^�I�u�W�F�N�g�����X�g�ɓo�^����B
 ���C�u�����̌^�̓t�@�C���ɋL�ڂ��A���s���ɓǂݎ��A���X�g�ɓo�^

�R�����g ����
 �s�̈����B�ȉ���3�Ŏ��B
  �C���f���g �����ǂ��ł����C������B�ϊ����ɐV�����l�X�g����΂悢�Ƃ��v���B
  ����
  �R�����g ��������� �e ������ꍇ�ǂ̂悤�ɕ������邩? �R�����g��蕶���񂪗D�悳��邽�߁A�h �̐����m�F���Ȃ��番������B

 �\�[�X���s�L���̌�ɃR�����g�͋L���o���Ȃ�(�� �ȉ��̂悤�ɂ͂ł��Ȃ�)
Call MsgBox(�gtest �h & _ �eNG
 	�gB�h
 �R�����g�̖����Ƀ\�[�X���s�L��������ꍇ�A���s����R�����g�ƂȂ�B
 A = 1 �e �R�����g _
   ���̍s���R�����g

 �t�@�C���̓ǂݎ�莞�ɍs�̕����A�����������Ŗ��Ȃ�

 �R�[�h�̓ǂݎ����B
 1. �s�r���ł̉��s�A�����s��1�s�ւ̂܂Ƃ߂�����
 2. �t�@�C��������A�N���X���W���[���A�W�����W���[���𕪕ʁB�N���X�������A�ǂݎ���Ă����B
 3. �W�����W���[���̃t�@�C�����ォ�珇�ɁA�u�^��`�A�֐���`�A�ϐ��錾�A�R�����g���v�ŕ��ʂ��ēǂݎ��A�X�^�b�N���Ă����B
    �ǂݍ��񂾁A�^�A�O���[�o���ϐ�(�萔)�͕ʂŃ��X�g���쐬���Ă����B
 4. �N���X���W���[���̃t�@�C�����ォ�珇�ɁA�u�����o�[�A���\�b�h�A�R�����g���v�ŕ��ʂ��ēǂݎ��A�X�^�b�N���Ă����B
 * 
# If���Ȃǂ̓ǂݎ����B
 �܂��A�֐����Ƃɓǂݍ��񂾃u���b�N���ɓǂށB
 If���Ȃǂ��o������ƁABlock�I�u�W�F�N�g���쐬���A�I���n�_���o������܂ŁA�s�̎�荞�݂��J�n����B
 �ǂݍ��ݓr���ŁA��̏������ċA�I�ɌĂяo���B
 If���Ȃǂ��o�����Ȃ��܂܁A�I���n�_��������A���[�v���甲���A��̃��[�v�ɖ߂��B
 ����A�{���ɍċA�Ő���s�͂���悤�Ȃ��r�����v���V�[�W������荞�񂾂�A����������΂����ƂɂȂ邩��
**/

#include "Atteribute.h"


/*
class VariableElement;

/// <summary>�O���[�o���ϐ��̃��X�g</summary>
std::map<std::string, VariableElement> _GlobalVariableList; //Name, Variable

/// <summary>�ϐ��v�f</summary>
class VariableElement
{
	AccessAdornment _Access;

	std::string _TypeName;
	std::string _Name;

public:
	VariableElement() : _Access(AccessAdornment::PUBLIC_ACCESS) {}

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

/// <summary>�s�̗v�f</summary>
class LineElement
{
public:
	std::string ToCShapeCode();
};

/// <summary>���l</summary>
class ValueElement : public LineElement
{
	std::string _Value;

public:
	ValueElement();

	std::string ToCShapeCode() { return _Value; }
};

/// <summary>���</summary>
class AssignElement : public LineElement
{
private:
	std::shared_ptr<LineElement> _RightSideValue;
	std::shared_ptr<LineElement> _LeftSideValue;
};

/// <summary>���Z�q�v�f</summary>
class OperatorElement : public LineElement
{
};

/// <summary>�v�Z�v�f(��̊��ʂŋ�؂�镪)</summary>
class FormulaElement : public LineElement
{
	std::vector<LineElement> _Formula;

};

/// <summary>�ϐ��Ăяo��</summary>
class CallVariableElement : public LineElement
{
private:
	std::shared_ptr<VariableElement> _Variable;
	
public:
	CallVariableElement();
};

/// <summary>�ϐ��錾</summary>
class DeclarationElement : public LineElement
{
private:
	std::shared_ptr<VariableElement> _Variable;

public:
	DeclarationElement();
};

class FunctionBlock;
/// <summary>�֐��Ăяo��</summary>
class CallFUnctionElement : public LineElement
{
private:
	std::shared_ptr<FunctionBlock> _CallFunction;
	std::map<std::string, LineElement> _ArugmentAssign;
public:
	CallFUnctionElement();
};
 
/// <summary>�R�����g�̕���</summary>
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

/// <summary>If��</summary>
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
	ArgumentCode() : _IsReference(true), _IsParmArray(false), _IsOptional(false) {}
};

const std::string _ArgumentMatchCode = "(ByRef|ByVal|Optional|ParmArray)?[ ]?";

class TypeBlock;
/// <summary>�v���W�F�N�g���Œ�`����Ă���^�̈ꗗ</summary>
std::map<std::string, TypeBlock> _ProjectTypeList;

/// <summary>�^�̗v�f��\��</summary>
class TypeBlock : public CodeBlock
{
protected:
	std::string _Name;
public:
	TypeBlock(const std::string& _Name) : _Name(_Name) {}

	constexpr std::string& GetName() { return TypeBlock::_Name; }
};

/// <summary>�񋓌^</summary>
class EnumBlock : public TypeBlock
{
	std::list<ConstantElement> _ElementList;
	
public:

	EnumBlock(const std::string& _Name) : TypeBlock(_Name) {}
};

/// <summary>�\����</summary>
class StructBlock : public TypeBlock
{
	std::list<VariableElement> _MemberList;

public:
	StructBlock(const std::string& _Name) : TypeBlock(_Name) {}
};

/// <summary>�N���X</summary>
class ClassBlock : public TypeBlock
{
	std::string _BaseTypeName; //�p����

	std::list <std::shared_ptr<TypeBlock>> _TypeList;
	std::list<std::shared_ptr<VariableElement>> _MemberList; //�����o�[�ꗗ
	std::list<std::shared_ptr<FunctionBlock>> _MethodList; //���\�b�h�ꗗ

	std::vector<CodeBlock> _ClassCode; //�N���X�̃R�[�h�S�́B
public:

};

class FileBlock
{

};

std::list<FunctionBlock> SplitFunctionElement(std::string _Code)
{
	/*
Match
[1] �A�N�Z�X�C���q
[2] Static
[3] Function|Sub //VB�̂ǂ��ł������敪
[4] �֐���
[5] ��������
[6] �߂�l�̕���
[7] �߂�l �^��
[8] �֐����R�[�h
[9] Function|Sub //�ǂ��ł�����

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

*/

int OperateVB6codeLine(int argc, char* argv[], char* envp[]);
int SplitVB6codeLine(int argc, char* argv[], char* envp[]);

int main(int argc, char* argv[], char* envp[])
{
	//std::cout << std::to_string(1 % 2);
	
	
	std::string _code = "Call Log.WriteLog(Log.DEBUG_LOG_LEVEL, Me, \"Class_Initialize()\", \"object[\" & TypeName(Me) & \"] has been created \"\" aa \"\" .\")";
	std::list<std::string> _code_list;
	boost::split(_code_list, _code, boost::is_any_of("\""));

	std::cout << _code << std::endl;
	for (auto _itr = std::begin(_code_list), _end = std::end(_code_list); _itr != _end; _itr++)
	{
		std::cout << (*_itr) << std::endl;
	}

	/*
	const auto _begin = std::begin(_code);
	for (auto _itr = std::begin(_code), _end = std::end(_code); _itr != _end; _itr++)
	{
		if ((*_itr) == 'a')
		{
			std::cout << (size_t)(_itr - _begin) << std::endl;
		}
	}
	*/


	//return OperateVB6codeLine(argc, argv, envp);
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
'���l���n�܂�ʒu�𒲂ׂ�
'str 	: ���ׂ镶����
'head	: ���׎n�߂�ʒu(�ŏ��l��1)
'return	: ���l���n�܂�ʒu
Function D_FUNC(ByVal str As String, Optional head As Long = 1) As Long
	For head = head To Len(str)
		If (Mid(str, head, 1) Like "#") Then	Exit For
	Next
	
	FindValuePosition = head
End Function)");
*/
}
