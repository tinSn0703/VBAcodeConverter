#pragma once

#include "Formating.h"

/// <summary>
/// <para>�t�@�C�������荞��VB6�R�[�h�𐮌`����</para>
/// <para># �s�̕����A������߂������̎菇</para>
/// <para>1. �t�@�C����S�ēǂݍ���</para>
/// <para>2. ���s�ȊO�̋󔒂�S�āA���p�X�y�[�X�ɕς���</para>
/// <para>3. �C���f���g�A�s�����̃X�y�[�X��S�č폜</para>
/// <para>4. �����̔��p�X�y�[�X�͑S�ĒP���̔��p�X�y�[�X�ɕς���</para>
/// <para>5. �������ꂽ�s��1�s�ɖ߂�</para>
/// <para>6. �ȗ��\�L���ꂽIf����߂��B(���̏������̓ǂݍ��݂��s���₷�����邽��)</para>
/// <para>7. ��������ꂽ�s��߂�</para>
/// <para></para>
/// </summary>
class FormattingVB6Code : public Formating
{
	/// <summary>�s�������s�ɓn�����R�[�h�ł��邩?</summary>
	/// <param name="_code">�ΏۃR�[�h</param>
	/// <returns>Yes/No</returns>
	bool is_continuous_line(const std::string& _code);

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
	/// <param name="_vb6_code_list">VB6�R�[�h�B�ϊ������</param>
	/// <returns>�ϊ����VB6�R�[�h���ꊇ���������́B\n��1�s���̋�؂�ɂ�����</returns>
	std::string Format(std::list<std::string>& _vb6_code_list);
};

