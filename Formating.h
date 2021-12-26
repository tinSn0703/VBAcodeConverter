#pragma once

#include <list>
#include <string>

class Formating
{
public:
	/// <summary>VB6コードを読み取りやすい形式に変換する。</summary>
	/// <param name="_vb6_code_list">VB6コード。変換される</param>
	/// <returns>変換後のVB6コードを一括化したもの。\nが1行毎の区切りにあたる</returns>
	virtual std::string Format(std::list<std::string>& _vb6_code_list) = 0;
};

