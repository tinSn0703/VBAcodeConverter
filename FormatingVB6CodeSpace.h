#pragma once

#include <vector>

#include "Formating.h"

/// <summary>
/// ���Z�q�Ȃǂ̗��[�̃X�y�[�X���Ȃ��ꍇ�A�ǉ�����
/// </summary>
class FormatingVB6CodeSpace : public Formating
{
private:
	/// <summary>["'(��ڂ̂�)]�̏o���ʒu�𒲂ׂĕԂ�</summary>
	/// <param name="_code">��������R�[�h</param>
	/// <returns>�������ʁBvector�ŕԂ��B���݂��Ȃ��ꍇ�͋�ŕԂ�</returns>
	std::vector<size_t> find_all_quotation(std::string& _code);

	/// <summary>�w�肵���ꏊ���A������(�܂��̓R�����g)���ł���?</summary>
	/// <param name="quotation_positions"></param>
	/// <param name="position"></param>
	/// <returns></returns>
	bool is_this_in_string(const std::vector<size_t>& quotation_positions, size_t position);
public:

	/// <summary>VB6�R�[�h��ǂݎ��₷���`���ɕϊ�����B</summary>
	/// <param name="_vb6_code_list">VB6�R�[�h�B�ϊ������</param>
	/// <returns>�ϊ����VB6�R�[�h���ꊇ���������́B\n��1�s���̋�؂�ɂ�����</returns>
	std::string Format(std::list<std::string>& _vb6_code_list);
};

