#include <string>
#include <list>
#include <regex>

#include <boost/version.hpp>
#include <boost/algorithm/string.hpp>

#include "CodeSplit.h"

const std::string LINE_SPLIT_PATTERN = R"(^([ \w&.,+-/*<>()]+?)(=??)$)";

/*
コード分割手順
 「"」の有無を確認
 「"」がある場合、最初の「"」から「"」までを分割する
　分割した文字列の内、先頭が「"」でないものを下記の変換を行う
 「*&/-+=」の両端がスペースとなるように変換する。
 「<>」の両端がスペースとなるように変換する。非等価時ではないこと
 「,」の一つ後がスペースとなるように変換する。
  
  
*/

std::vector<std::string> split_quotation(std::string& _code)
{
	std::vector<std::string> _return_code(1);
	//["']の最初の出現位置を検索
	//["]を発見した場合、次の["]の出現位置を検索。最初の["]までと次の["]までを分割して最初の処理にもどる
	//[']を発見した場合[']のひとつ前と以降で分割する

	return _return_code;
}

void format_operater_space(std::string& _code)
{
}

int SplitVB6codeLine(int argc, char* argv[], char* envp[])
{
	std::string _code = "";
	std::vector<std::string> _split_code = split_quotation(_code); //文字列、コメント部とその他で分割する

	for (auto _itr = std::begin(_split_code), _end = std::end(_split_code); _itr != _end; _itr++)
	{
		if ((*_itr)[0] == '\'') break;		//「'」から始まるものはコメントなので終了
		if ((*_itr)[0] == '\"') continue;	//「"」から始まるものは文字列なのでスキップ

		format_operater_space(*_itr);
	}

	return 0;
}
