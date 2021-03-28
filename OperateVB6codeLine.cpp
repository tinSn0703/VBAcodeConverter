
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
/// <para>行の分割、複合を戻す処理の手順</para>
/// <para>ファイルを全て読み込む</para>
/// <para>改行以外の空白を全て、半角スペースに変える</para>
/// <para>インデント、行末尾のスペースを全て削除</para>
/// <para>複数の半角スペースは全て単数の半角スペースに変える</para>
/// <para>分割された行を1行に戻す</para>
/// <para>省略表記されたIf文を戻す。(次の処理時の読み込みを行いやすくするため)</para>
/// <para>結合されれた行を戻す</para>
/// <para></para>
/// </summary>

/// <summary>
///  <para>インデント、行末尾スペースを削除する。</para>
///  <para>20万ぐらいスペースを入力するなんて誰もやらんやろうし、正規表現でいいかもしれない(楽観的観測)</para>
///  <para>boostにありました...</para>
/// </summary>
/// <param name="_code"></param>
/// <returns></returns>
std::string& remove_head_and_tail_space(std::string &_code)
{
	if (_code.size() < 1) return _code;

	if (_code[0] == ' ')
	{
		auto i = _code.find_first_not_of(' '); //インデントからコードへの変化点を検索する
		if (i == std::string::npos)	_code = ""; //
		else						_code = _code.substr(i, _code.size() - 1); //インデントを削除
	}
	
	if ((0 < _code.size()) && _code[_code.size() - 1] == ' ')
	{
		auto i = _code.find_last_not_of(' '); //行末尾スペースからコードへの変化点を検索する
		if (i == std::string::npos)	_code = "";
		else						_code = _code.substr(0, i + 1); //行末尾スペースを削除
	}

	return _code;
}

/// <summary>コード内の余分なスペースを削除する。</summary>
/// <param name="_code"></param>
/// <returns></returns>
std::string& remove_code_space(std::string& _code)
{
	if (_code.size() < 1) return _code;

	std::replace(std::begin(_code), std::end(_code), '\t', ' '); //タブをスペースに変換。文字列内も纏めてやってしまう。
	boost::trim(_code); //インデント、行末尾スペースの削除

	trim_consecutive_char(_code, ' '); //2つ以上の行スペースを単スペースに置き換える。文字列内のスペースも置き換える。

	return _code;
}

/// <summary>一行If文を If () Then ~ End If形式にする</summary>
/// <returns>変換した文字列を返す。変換されなかった場合は元の文字列が返る</returns>
std::string& replace_one_if_line(std::string& _code)
{
	std::smatch _match;
	if (std::regex_match(_code, _match, std::regex(R"((.*?)If[ ]?([^\n]+?)[ ]?Then ([^'\n].+?))", std::regex_constants::icase)))
	{
		std::string temp = _match[3].str();
		replace_one_if_line(temp); //再帰的に関数を呼び出すことで、複数のif文が1行に結合されている状態をIf () Then ~ End If形式のネストにする。

		_code = _match[1].str() + "If " + _match[2].str() + " Then\n" + temp + "\nEnd If";
	}

	return _code;
}

/// <summary>1コード行を分割する</summary>
/// <returns>変換した文字列を返す。変換されなかった場合は元の文字列が返る</returns>
std::string& replace_one_code_line(std::string& _code)
{
	std::uint16_t _count = 0;
	bool _is_remove = false;

	std::smatch _match;
	for (auto _itr = std::begin(_code), _end = std::end(_code); _itr != _end; _itr++)
	{
		if ((*_itr) == '\'') break; //コメント文に突入。これ以上は無視
		else if ((*_itr) == '\"') _count++; //「"」の出現回数をカウント
		else if ((*_itr) == ':')//コードの連結記号の出現
		{
			//末尾(改行前)に「:」がある場合、Goto文なのでスキップ. 「"」の出現数が偶数または0 == 文字列外 と判断
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
	std::string _vb6_code;
	while (!_vb6_file.eof()) { std::getline(_vb6_file, _vb6_code); _vb6_code_list.push_back(_vb6_code); }
	_vb6_file.close();
	_vb6_code = "";
	
	// 行末尾のスペース及びインデントの削除。タブは全てスペースに置き換える。正規表現で削除を行うと、とんでもない時間がかかるため、文字列の検索と再確保で行う。
	for (auto _itr = std::begin(_vb6_code_list), _end = std::end(_vb6_code_list); _itr != _end; )
	{
		auto _temp_code = remove_code_space(*_itr);
		
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
			_temp_code = replace_one_code_line(_temp_code); //一行コードを複数行にする
			
			_vb6_code += _temp_code + "\n"; //コードを一つの変数に纏める
			++_itr;
		}
	}
	
	_vb6_code_list.clear();
	boost::split(_vb6_code_list, _vb6_code, boost::is_any_of("\n"));

	std::ofstream _output_file(R"(D:\Project\VBAcodeConverter\test_vb6.bas)");
	_output_file << _vb6_code;
	_output_file.close();
#endif
	// 正常終了
	return(0);
}

//LINE_SPLIT_PATTERN = L"\\n([ ]+?)([ \\w=&.,+-/*<>()]+?|([ \\w=&.,+-/*<>()]+?+?\"[^\"]*?\"[ \\w=&.,+-/*<>()]*?){1,})[ ]*:[ ]*([^’\\n].+?)\\n)"
