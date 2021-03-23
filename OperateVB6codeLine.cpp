
#include <iostream>
#include <fstream>
#include <string>
#include <list>
#include <regex>
#include <filesystem>

#include <boost/version.hpp>
#include <boost/algorithm/string.hpp>

/// <summary>
/// <para>行の分割、複合を戻す処理の手順</para>
/// <para>ファイルを全て読み込む</para>
/// <para>改行以外の空白を全て、半角スペースに変える</para>
/// <para>複数の半角スペースは全て単数の半角スペースに変える</para>
/// <para>行末尾のスペースを全て削除</para>
/// <para>空行要素数0の文字列とする</para>
/// <para>分割された行を1行に戻す</para>
/// <para>省略表記されたIf文を戻す。(次の処理時の読み込みを行いやすくするため)</para>
/// <para>結合されれた行を戻すタイミングは、行要素の認識時に行う。</para>
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
	//2つ以上の行スペースを単スペースに置き換える。文字列内のスペースも置き換える。
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
		auto i = _code.find_first_not_of(' '); //インデントからコードへの変化点を検索する
		if (i == std::string::npos)	_code = "";
		else						_code = _code.substr(i, _code.size() - 1); //インデントを削除
	}
	
	if ((0 < _code.size()) && _code[_code.size() - 1] == ' ')
	{
		auto i = _code.find_last_not_of(' '); //行末尾スペースからコードへの変化点を検索する
		if (i == std::string::npos)	_code = "";
		else						_code = _code.substr(0, i + 1); //行末尾スペースを削除
	}

	return remove_code_space(_code);
}

/// <summary>一行If文を If () Then ~ End If形式にする</summary>
/// <returns>変換した文字列を返す。変換されなかった場合は元の文字列が返る</returns>
std::string& replace_one_if_line(std::string& _str)
{
	std::smatch _match;
	if (std::regex_match(_str, _match, std::regex(R"(If[ ]?([^\n]+?)[ ]?Then ([^'\n].+?))")))
	{
		_str = "If " + _match[1].str() + " Then\n" + _match[2].str() + "\nEnd If";
	}

	return _str;
}

/// <summary>一行If文を If () Then ~ End If形式にする</summary>
/// <returns>変換した文字列を返す。変換されなかった場合は元の文字列が返る</returns>
std::string& replace_one_code_line(std::string& _str)
{
	std::smatch _match;
	for (auto _itr = std::begin(_str), _end = std::end(_str); _itr != _end; _itr++)
	{

	}

	return _str;
}

#define DEBUG_MODE 0

/// <summary>VB6の、途中改行、一行化された行を元の状態に戻す。コードのテスト</summary>
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
	
	// 行末尾のスペース及びインデントの削除。タブは全てスペースに置き換える。正規表現で削除を行うととんでもない時間がかかるため、文字列の検索と再確保で行う。
	for (auto _itr = std::begin(_vb6_code_list), _end = std::end(_vb6_code_list); _itr != _end; )
	{
		auto _temp_code = remove_excess_code(*_itr);
		
		// 複数行への分割を単行に戻す。
		if ((2 < _temp_code.size()) && (_temp_code.back() == '_') && (_temp_code[_temp_code.size() - 2] == ' '))
		{
			++_itr;								//次の行へイテレータを進める。
			_temp_code.pop_back();				//「_」を削除
			_temp_code += (*_itr);				//次行を連結
			_itr = _vb6_code_list.erase(_itr);	//次行を削除
			--_itr;								//イテレータを戻す
			(*_itr) = _temp_code;				//再度、今行の処理を行う
		}
		else
		{
			_temp_code = replace_one_if_line(_temp_code); //一行If文を複数行にする
			
			_vb6_code += _temp_code + "\n"; //コードを一つの変数に纏める
			++_itr;
		}
	}
	
	//_vb6_code_list.clear();
	//boost::split(_vb6_code_list, _vb6_code, boost::is_any_of("\n"));

	std::ofstream _output_file(R"(D:\Project\VBAcodeConverter\test_vb6.bas)");
	_output_file << _vb6_code;
	_output_file.close();
#endif
	// 正常終了
	return(0);
}

//LINE_SPLIT_PATTERN = L"\\n([ ]+?)([ \\w=&.,+-/*<>()]+?|([ \\w=&.,+-/*<>()]+?+?\"[^\"]*?\"[ \\w=&.,+-/*<>()]*?){1,})[ ]*:[ ]*([^’\\n].+?)\\n)"
