
#include <iostream>
#include <fstream>
#include <string>
#include <list>
#include <regex>
#include <filesystem>

#include <boost/version.hpp>
#include <boost/algorithm/string.hpp>

#include "StringOperate.h"

/// <summary>
/// <para>�s�̕����A������߂������̎菇</para>
/// <para>�t�@�C����S�ēǂݍ���</para>
/// <para>���s�ȊO�̋󔒂�S�āA���p�X�y�[�X�ɕς���</para>
/// <para>�C���f���g�A�s�����̃X�y�[�X��S�č폜</para>
/// <para>�����̔��p�X�y�[�X�͑S�ĒP���̔��p�X�y�[�X�ɕς���</para>
/// <para>�������ꂽ�s��1�s�ɖ߂�</para>
/// <para>�ȗ��\�L���ꂽIf����߂��B(���̏������̓ǂݍ��݂��s���₷�����邽��)</para>
/// <para>��������ꂽ�s��߂�</para>
/// <para></para>
/// </summary>

/// <summary>
///  <para>�C���f���g�A�s�����X�y�[�X���폜����B</para>
///  <para>20�����炢�X�y�[�X����͂���Ȃ�ĒN�������낤���A���K�\���ł�����������Ȃ�(�y�ϓI�ϑ�)</para>
///  <para>boost�ɂ���܂���...</para>
/// </summary>
/// <param name="_code"></param>
/// <returns></returns>
std::string& remove_head_and_tail_space(std::string &_code)
{
	if (_code.size() < 1) return _code;

	if (_code[0] == ' ')
	{
		auto i = _code.find_first_not_of(' '); //�C���f���g����R�[�h�ւ̕ω��_����������
		if (i == std::string::npos)	_code = ""; //
		else						_code = _code.substr(i, _code.size() - 1); //�C���f���g���폜
	}
	
	if ((0 < _code.size()) && _code[_code.size() - 1] == ' ')
	{
		auto i = _code.find_last_not_of(' '); //�s�����X�y�[�X����R�[�h�ւ̕ω��_����������
		if (i == std::string::npos)	_code = "";
		else						_code = _code.substr(0, i + 1); //�s�����X�y�[�X���폜
	}

	return _code;
}

/// <summary>�R�[�h���̗]���ȃX�y�[�X���폜����B</summary>
/// <param name="_code"></param>
/// <returns></returns>
std::string& remove_code_space(std::string& _code)
{
	if (_code.size() < 1) return _code;

	std::replace(std::begin(_code), std::end(_code), '\t', ' '); //�^�u���X�y�[�X�ɕϊ��B����������Z�߂Ă���Ă��܂��B
	boost::trim(_code); //�C���f���g�A�s�����X�y�[�X�̍폜

	trim_consecutive_char(_code, ' '); //2�ȏ�̍s�X�y�[�X��P�X�y�[�X�ɒu��������B��������̃X�y�[�X���u��������B

	return _code;
}

/// <summary>��sIf���� If () Then ~ End If�`���ɂ���</summary>
/// <returns>�ϊ������������Ԃ��B�ϊ�����Ȃ������ꍇ�͌��̕����񂪕Ԃ�</returns>
std::string& replace_one_if_line(std::string& _code)
{
	std::smatch _match;
	if (std::regex_match(_code, _match, std::regex(R"((.*?)If[ ]?([^\n]+?)[ ]?Then ([^'\n].+?))", std::regex_constants::icase)))
	{
		std::string temp = _match[3].str();
		replace_one_if_line(temp); //�ċA�I�Ɋ֐����Ăяo�����ƂŁA������if����1�s�Ɍ�������Ă����Ԃ�If () Then ~ End If�`���̃l�X�g�ɂ���B

		_code = _match[1].str() + "If " + _match[2].str() + " Then\n" + temp + "\nEnd If";
	}

	return _code;
}

/// <summary>1�R�[�h�s�𕪊�����</summary>
/// <returns>�ϊ������������Ԃ��B�ϊ�����Ȃ������ꍇ�͌��̕����񂪕Ԃ�</returns>
std::string& replace_one_code_line(std::string& _code)
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
	return _code;
}

#define DEBUG_MODE 0

/// <summary>VB6�́A�r�����s�A��s�����ꂽ�s�����̏�Ԃɖ߂��B�R�[�h�̃e�X�g</summary>
/// <param name="argc"></param>
/// <param name="argv"></param>
/// <param name="envp"></param>
/// <returns></returns>
int OperateVB6codeLine(int argc, char* argv[], char* envp[])
{
#if DEBUG_MODE != 0
	std::string _test_code = "If (IsNull(value)) Then err_msg = err_msg";
		//"  ";

	std::cout << replace_one_if_line(_test_code) << std::endl;

#else
	std::ifstream _vb6_file(R"(D:\Project\VBAcodeConverter\test_vb6_base.bas)");
	if (_vb6_file.fail()) throw std::exception("");
	
	std::list<std::string> _vb6_code_list;
	std::string _vb6_code;
	while (!_vb6_file.eof()) { std::getline(_vb6_file, _vb6_code); _vb6_code_list.push_back(_vb6_code); }
	_vb6_file.close();
	_vb6_code = "";
	
	// �s�����̃X�y�[�X�y�уC���f���g�̍폜�B�^�u�͑S�ăX�y�[�X�ɒu��������B���K�\���ō폜���s���ƁA�Ƃ�ł��Ȃ����Ԃ������邽�߁A������̌����ƍĊm�ۂōs���B
	for (auto _itr = std::begin(_vb6_code_list), _end = std::end(_vb6_code_list); _itr != _end; )
	{
		auto _temp_code = remove_code_space(*_itr);
		
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
			_temp_code = replace_one_if_line(_temp_code); //��sIf���𕡐��s�ɂ���
			_temp_code = replace_one_code_line(_temp_code); //��s�R�[�h�𕡐��s�ɂ���
			
			_vb6_code += _temp_code + "\n"; //�R�[�h����̕ϐ��ɓZ�߂�
			++_itr;
		}
	}
	
	_vb6_code_list.clear();
	boost::split(_vb6_code_list, _vb6_code, boost::is_any_of("\n"));

	std::ofstream _output_file(R"(D:\Project\VBAcodeConverter\test_vb6.bas)");
	_output_file << _vb6_code;
	_output_file.close();
#endif
	// ����I��
	return(0);
}

//LINE_SPLIT_PATTERN = L"\\n([ ]+?)([ \\w=&.,+-/*<>()]+?|([ \\w=&.,+-/*<>()]+?+?\"[^\"]*?\"[ \\w=&.,+-/*<>()]*?){1,})[ ]*:[ ]*([^�f\\n].+?)\\n)"
