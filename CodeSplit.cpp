#include <string>
#include <list>
#include <regex>

#include <boost/version.hpp>
#include <boost/algorithm/string.hpp>

#include "CodeSplit.h"

const std::string LINE_SPLIT_PATTERN = R"(^([ \w&.,+-/*<>()]+?)(=??)$)";

/*
�R�[�h�����菇
 �u"�v�̗L�����m�F
 �u"�v������ꍇ�A�ŏ��́u"�v����u"�v�܂ł𕪊�����
�@��������������̓��A�擪���u"�v�łȂ����̂����L�̕ϊ����s��
 �u*&/-+=�v�̗��[���X�y�[�X�ƂȂ�悤�ɕϊ�����B
 �u<>�v�̗��[���X�y�[�X�ƂȂ�悤�ɕϊ�����B�񓙉����ł͂Ȃ�����
 �u,�v�̈�オ�X�y�[�X�ƂȂ�悤�ɕϊ�����B
  
  
*/

std::vector<std::string> split_quotation(std::string& _code)
{
	std::vector<std::string> _return_code(1);
	//["']�̍ŏ��̏o���ʒu������
	//["]�𔭌������ꍇ�A����["]�̏o���ʒu�������B�ŏ���["]�܂łƎ���["]�܂ł𕪊����čŏ��̏����ɂ��ǂ�
	//[']�𔭌������ꍇ[']�̂ЂƂO�ƈȍ~�ŕ�������

	return _return_code;
}

void format_operater_space(std::string& _code)
{
}

int SplitVB6codeLine(int argc, char* argv[], char* envp[])
{
	std::string _code = "";
	std::vector<std::string> _split_code = split_quotation(_code); //������A�R�����g���Ƃ��̑��ŕ�������

	for (auto _itr = std::begin(_split_code), _end = std::end(_split_code); _itr != _end; _itr++)
	{
		if ((*_itr)[0] == '\'') break;		//�u'�v����n�܂���̂̓R�����g�Ȃ̂ŏI��
		if ((*_itr)[0] == '\"') continue;	//�u"�v����n�܂���͕̂�����Ȃ̂ŃX�L�b�v

		format_operater_space(*_itr);
	}

	return 0;
}
