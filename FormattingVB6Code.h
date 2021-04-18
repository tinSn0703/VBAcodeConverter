#pragma once

#include <list>
#include <string>

/// <summary>�t�@�C�������荞��VB6�R�[�h�𐮌`����</summary>
class FormattingVB6Code
{
	/// <summary>�R�[�h���̗]���ȃX�y�[�X���폜����B</summary>
	/// <param name="_code">�ΏۃR�[�h</param>
	void remove_code_space(std::string& _code);
	
	/// <summary>��sIf���� If () Then ~ End If�`���ɂ���</summary>
	/// <param name="_code">�ΏۃR�[�h</param>
	void replace_one_if_line(std::string& _code);

	/// <summary>1�R�[�h�s�𕪊�����</summary>
	/// <param name="_code">�ΏۃR�[�h</param>
	void replace_one_code_line(std::string& _code);
public:

	/// <summary>VB6�R�[�h��ǂݎ��₷���`���ɕϊ�����B</summary>
	/// <param name="_vb6_code_list"></param>
	std::string Format(std::list<std::string>& _vb6_code_list);
};

