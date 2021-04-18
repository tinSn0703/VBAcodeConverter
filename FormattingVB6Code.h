#pragma once

#include <list>
#include <string>

/// <summary>ファイルから取り込んだVB6コードを整形する</summary>
class FormattingVB6Code
{
	/// <summary>コード内の余分なスペースを削除する。</summary>
	/// <param name="_code">対象コード</param>
	void remove_code_space(std::string& _code);
	
	/// <summary>一行If文を If () Then ~ End If形式にする</summary>
	/// <param name="_code">対象コード</param>
	void replace_one_if_line(std::string& _code);

	/// <summary>1コード行を分割する</summary>
	/// <param name="_code">対象コード</param>
	void replace_one_code_line(std::string& _code);
public:

	/// <summary>VB6コードを読み取りやすい形式に変換する。</summary>
	/// <param name="_vb6_code_list"></param>
	std::string Format(std::list<std::string>& _vb6_code_list);
};

