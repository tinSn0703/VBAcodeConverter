
#include <iostream>
#include <fstream>
#include <string>
#include <list>
#include <regex>
#include <filesystem>

#include <boost/version.hpp>
#include <boost/algorithm/string.hpp>

#include "FormattingVB6Code.h"
#include "StringOperate.h"

//-------------------------------------------------------------------------

void FormattingVB6Code::remove_code_space(std::string& _code)
{
	if (_code.size() < 1) return;

	std::replace(std::begin(_code), std::end(_code), '\t', ' '); //�^�u���X�y�[�X�ɕϊ��B����������Z�߂Ă���Ă��܂��B
	boost::trim(_code); //�C���f���g�A�s�����X�y�[�X�̍폜

	trim_consecutive_char(_code, ' '); //2�ȏ�̍s�X�y�[�X��P�X�y�[�X�ɒu��������B��������̃X�y�[�X���u��������B
}

//-------------------------------------------------------------------------

void FormattingVB6Code::replace_one_if_line(std::string& _code)
{
	std::smatch _match;
	if (std::regex_match(_code, _match, std::regex(R"((.*?)If[ ]?([^\n]+?)[ ]?Then ([^'\n].+?))", std::regex_constants::icase)))
	{
		std::string temp = _match[3].str();
		replace_one_if_line(temp); //�ċA�I�Ɋ֐����Ăяo�����ƂŁA������if����1�s�Ɍ�������Ă����Ԃ�If () Then ~ End If�`���̃l�X�g�ɂ���B

		_code = _match[1].str() + "If " + _match[2].str() + " Then\n" + temp + "\nEnd If";
	}
}

//-------------------------------------------------------------------------

void FormattingVB6Code::replace_one_code_line(std::string& _code)
{
	std::uint16_t _count = 0;
	bool _is_remove = false;

	std::smatch _match;
	for (auto _itr = std::begin(_code), _end = std::end(_code); _itr != _end; _itr++)
	{
		if ((*_itr) == '\'') break; //�R�����g���ɓ˓��B����ȏ�͖���
		else if ((*_itr) == '\"') _count++; //�u"�v�̏o���񐔂��J�E���g
		else if ((*_itr) == ':')//�R�[�h�̘A���L���̏o��
		{
			//����(���s�O)�Ɂu:�v������ꍇ�AGoto���Ȃ̂ŃX�L�b�v. �u"�v�̏o�����������܂���0 == ������O �Ɣ��f
			if (((_itr + 1) < _end) && ((*(_itr + 1)) != '\n') && ((_count % 2) == 0))
			{
				(*_itr) = '\n';
				_is_remove = true;
			}
		}
	}

	if (_is_remove) _code = std::regex_replace(_code, std::regex(R"([ ]?\n[ ]?)"), "\n");
}

//-------------------------------------------------------------------------

std::string FormattingVB6Code::Format(std::list<std::string>& _vb6_code_list)
{
	std::string _vb6_code = "";
	// �s�����̃X�y�[�X�y�уC���f���g�̍폜�B�^�u�͑S�ăX�y�[�X�ɒu��������B���K�\���ō폜���s���ƁA�Ƃ�ł��Ȃ����Ԃ������邽�߁A������̌����ƍĊm�ۂōs���B
	for (auto _itr = std::begin(_vb6_code_list), _end = std::end(_vb6_code_list); _itr != _end; )
	{
		auto _temp_code = (*_itr);
		remove_code_space(_temp_code);

		// �����s�ւ̕�����P�s�ɖ߂��B
		if ((2 < _temp_code.size()) && (_temp_code.back() == '_') && (_temp_code[_temp_code.size() - 2] == ' '))
		{
			++_itr;								//���̍s�փC�e���[�^��i�߂�B
			_temp_code.pop_back();				//�u_�v���폜
			_temp_code += (*_itr);				//���s��A��
			_itr = _vb6_code_list.erase(_itr);	//���s���폜
			--_itr;								//�C�e���[�^��߂�
			(*_itr) = _temp_code;				//�ēx�A���s�̏������s��
		}
		else
		{
			replace_one_if_line(_temp_code); //��sIf���𕡐��s�ɂ���
			replace_one_code_line(_temp_code); //��s�R�[�h�𕡐��s�ɂ���

			_vb6_code += _temp_code + "\n"; //�R�[�h����̕ϐ��ɓZ�߂�
			++_itr;
		}
	}

	_vb6_code_list.clear();
	boost::split(_vb6_code_list, _vb6_code, boost::is_any_of("\n"));

	return _vb6_code;
}

//-------------------------------------------------------------------------
