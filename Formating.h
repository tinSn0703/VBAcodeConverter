#pragma once

#include <list>
#include <string>

class Formating
{
public:
	/// <summary>VB6�R�[�h��ǂݎ��₷���`���ɕϊ�����B</summary>
	/// <param name="_vb6_code_list">VB6�R�[�h�B�ϊ������</param>
	/// <returns>�ϊ����VB6�R�[�h���ꊇ���������́B\n��1�s���̋�؂�ɂ�����</returns>
	virtual std::string Format(std::list<std::string>& _vb6_code_list) = 0;
};

