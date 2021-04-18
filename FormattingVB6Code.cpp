
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

	std::replace(std::begin(_code), std::end(_code), '\t', ' '); //タブをスペースに変換。文字列内も纏めてやってしまう。
	boost::trim(_code); //インデント、行末尾スペースの削除

	trim_consecutive_char(_code, ' '); //2つ以上の行スペースを単スペースに置き換える。文字列内のスペースも置き換える。
}

//-------------------------------------------------------------------------

void FormattingVB6Code::replace_one_if_line(std::string& _code)
{
	std::smatch _match;
	if (std::regex_match(_code, _match, std::regex(R"((.*?)If[ ]?([^\n]+?)[ ]?Then ([^'\n].+?))", std::regex_constants::icase)))
	{
		std::string temp = _match[3].str();
		replace_one_if_line(temp); //再帰的に関数を呼び出すことで、複数のif文が1行に結合されている状態をIf () Then ~ End If形式のネストにする。

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
}

//-------------------------------------------------------------------------

std::string FormattingVB6Code::Format(std::list<std::string>& _vb6_code_list)
{
	std::string _vb6_code = "";
	// 行末尾のスペース及びインデントの削除。タブは全てスペースに置き換える。正規表現で削除を行うと、とんでもない時間がかかるため、文字列の検索と再確保で行う。
	for (auto _itr = std::begin(_vb6_code_list), _end = std::end(_vb6_code_list); _itr != _end; )
	{
		auto _temp_code = (*_itr);
		remove_code_space(_temp_code);

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
			replace_one_if_line(_temp_code); //一行If文を複数行にする
			replace_one_code_line(_temp_code); //一行コードを複数行にする

			_vb6_code += _temp_code + "\n"; //コードを一つの変数に纏める
			++_itr;
		}
	}

	_vb6_code_list.clear();
	boost::split(_vb6_code_list, _vb6_code, boost::is_any_of("\n"));

	return _vb6_code;
}

//-------------------------------------------------------------------------
