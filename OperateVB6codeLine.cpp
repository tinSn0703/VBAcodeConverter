
#include <iostream>
#include <fstream>
#include <string>
#include <list>
#include <regex>
#include <filesystem>

#include <boost/version.hpp>
#include <boost/algorithm/string.hpp>

/// <summary>
/// <para>�s�̕����A������߂������̎菇</para>
/// <para>�t�@�C����S�ēǂݍ���</para>
/// <para>���s�ȊO�̋󔒂�S�āA���p�X�y�[�X�ɕς���</para>
/// <para>�����̔��p�X�y�[�X�͑S�ĒP���̔��p�X�y�[�X�ɕς���</para>
/// <para>�s�����̃X�y�[�X��S�č폜</para>
/// <para>��s�v�f��0�̕�����Ƃ���</para>
/// <para>�������ꂽ�s��1�s�ɖ߂�</para>
/// <para>�ȗ��\�L���ꂽIf����߂��B(���̏������̓ǂݍ��݂��s���₷�����邽��)</para>
/// <para>��������ꂽ�s��߂��^�C�~���O�́A�s�v�f�̔F�����ɍs���B</para>
/// <para></para>
/// </summary>

class ArgumentElement
{
	bool _IsReference;

	std::string _Type;
	std::string _Name;
	std::string _InitValue;
};

class FunctionBlock
{
	std::string _Name;
	std::string _ReturnType;
	std::vector<ArgumentElement> _ArgumentList;
};

std::wstring multi_to_wide_capi(std::string const& src)
{
	std::size_t converted{};
	std::vector<wchar_t> dest(src.size(), L'\0');
	if (::_mbstowcs_s_l(&converted, dest.data(), dest.size(), src.data(), _TRUNCATE, ::_create_locale(LC_ALL, "jpn")) != 0) {
		throw std::system_error{ errno, std::system_category() };
	}
	dest.resize(std::char_traits<wchar_t>::length(dest.data()));
	dest.shrink_to_fit();
	return std::wstring(dest.begin(), dest.end());
}

std::string wide_to_multi_capi(std::wstring const& src)
{
	std::size_t converted{};
	std::vector<char> dest(src.size() * sizeof(wchar_t) + 1, '\0');
	if (::_wcstombs_s_l(&converted, dest.data(), dest.size(), src.data(), _TRUNCATE, ::_create_locale(LC_ALL, "jpn")) != 0) {
		throw std::system_error{ errno, std::system_category() };
	}
	dest.resize(std::char_traits<char>::length(dest.data()));
	dest.shrink_to_fit();
	return std::string(dest.begin(), dest.end());
}

std::uint16_t count_num_word_in_str(const std::string& str, const std::string& word, std::string::size_type pos = 0, std::string::size_type end = std::string::npos)
{
	std::uint16_t count = 0;
	if (end == std::string::npos) end = str.size();
	for (; ((pos = str.find(word, pos)) != std::string::npos) && (pos < end); pos += word.size())	++count;
	return count;
}

std::uint16_t rcount_num_word_in_str(const std::string& str, const std::string& word, std::string::size_type pos = std::string::npos, std::string::size_type end = -1)
{
	std::uint16_t count = 0;
	if (pos == std::string::npos) pos = str.size() - 1;
	for (; ((pos = str.find(word, pos)) != std::string::npos) && (end < pos); pos -= word.size())	++count;
	return count;
}

/// <summary></summary>
/// <param name="_code"></param>
/// <returns></returns>
std::string& remove_code_space(std::string& _code)
{
	//2�ȏ�̍s�X�y�[�X��P�X�y�[�X�ɒu��������B��������̃X�y�[�X���u��������B
	for (auto _space_pos = _code.find(' '); _space_pos != std::string::npos; )
	{
		const auto _end_space_pos = _code.find_first_not_of(' ', _space_pos);

		if (1 < (_end_space_pos - _space_pos))
		{
			_code = _code.substr(0, _space_pos) + " " + _code.substr(_end_space_pos, _code.size() - 1);
			_space_pos = _code.find(' ', _space_pos + 1);
		}
		else
		{
			_space_pos = _code.find(' ', _end_space_pos);
		}
	}

	return _code;
}

/// <summary></summary>
/// <param name="_code"></param>
/// <returns></returns>
std::string& remove_excess_code(std::string &_code)
{
	if (_code.size() < 1) return _code;

	std::replace(std::begin(_code), std::end(_code), '\t', ' ');
	
	if (_code[0] == ' ')
	{
		auto i = _code.find_first_not_of(' '); //�C���f���g����R�[�h�ւ̕ω��_����������
		if (i == std::string::npos)	_code = "";
		else						_code = _code.substr(i, _code.size() - 1); //�C���f���g���폜
	}
	
	if ((0 < _code.size()) && _code[_code.size() - 1] == ' ')
	{
		auto i = _code.find_last_not_of(' '); //�s�����X�y�[�X����R�[�h�ւ̕ω��_����������
		if (i == std::string::npos)	_code = "";
		else						_code = _code.substr(0, i + 1); //�s�����X�y�[�X���폜
	}

	return remove_code_space(_code);
}

/// <summary>��sIf���� If () Then ~ End If�`���ɂ���</summary>
/// <returns>�ϊ������������Ԃ��B�ϊ�����Ȃ������ꍇ�͌��̕����񂪕Ԃ�</returns>
std::string& replace_one_if_line(std::string& _str)
{
	std::smatch _match;
	if (std::regex_match(_str, _match, std::regex(R"(If[ ]?([^\n]+?)[ ]?Then ([^'\n].+?))")))
	{
		_str = "If " + _match[1].str() + " Then\n" + _match[2].str() + "\nEnd If";
	}

	return _str;
}

/// <summary>��sIf���� If () Then ~ End If�`���ɂ���</summary>
/// <returns>�ϊ������������Ԃ��B�ϊ�����Ȃ������ꍇ�͌��̕����񂪕Ԃ�</returns>
std::string& replace_one_code_line(std::string& _str)
{
	std::smatch _match;
	for (auto _itr = std::begin(_str), _end = std::end(_str); _itr != _end; _itr++)
	{

	}

	return _str;
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
	std::string _vb6_code; //= std::string(std::istreambuf_iterator<char>(_vb6_file), std::istreambuf_iterator<char>());
	while (!_vb6_file.eof()) { std::getline(_vb6_file, _vb6_code); _vb6_code_list.push_back(_vb6_code); }
	_vb6_file.close();
	_vb6_code = "";
	
	// �s�����̃X�y�[�X�y�уC���f���g�̍폜�B�^�u�͑S�ăX�y�[�X�ɒu��������B���K�\���ō폜���s���ƂƂ�ł��Ȃ����Ԃ������邽�߁A������̌����ƍĊm�ۂōs���B
	for (auto _itr = std::begin(_vb6_code_list), _end = std::end(_vb6_code_list); _itr != _end; )
	{
		auto _temp_code = remove_excess_code(*_itr);
		
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
			
			_vb6_code += _temp_code + "\n"; //�R�[�h����̕ϐ��ɓZ�߂�
			++_itr;
		}
	}
	
	//_vb6_code_list.clear();
	//boost::split(_vb6_code_list, _vb6_code, boost::is_any_of("\n"));

	std::ofstream _output_file(R"(D:\Project\VBAcodeConverter\test_vb6.bas)");
	_output_file << _vb6_code;
	_output_file.close();
#endif
	// ����I��
	return(0);
}

//LINE_SPLIT_PATTERN = L"\\n([ ]+?)([ \\w=&.,+-/*<>()]+?|([ \\w=&.,+-/*<>()]+?+?\"[^\"]*?\"[ \\w=&.,+-/*<>()]*?){1,})[ ]*:[ ]*([^�f\\n].+?)\\n)"
