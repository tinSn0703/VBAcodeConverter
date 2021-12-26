#pragma once

#include <vector>

#include "Formating.h"

/// <summary>
/// 演算子などの両端のスペースがない場合、追加する
/// </summary>
class FormatingVB6CodeSpace : public Formating
{
private:
	/// <summary>["'(一つ目のみ)]の出現位置を調べて返す</summary>
	/// <param name="_code">検索するコード</param>
	/// <returns>検索結果。vectorで返す。存在しない場合は空で返す</returns>
	std::vector<size_t> find_all_quotation(std::string& _code);

	/// <summary>指定した場所が、文字列(またはコメント)内ですか?</summary>
	/// <param name="quotation_positions"></param>
	/// <param name="position"></param>
	/// <returns></returns>
	bool is_this_in_string(const std::vector<size_t>& quotation_positions, size_t position);
public:

	/// <summary>VB6コードを読み取りやすい形式に変換する。</summary>
	/// <param name="_vb6_code_list">VB6コード。変換される</param>
	/// <returns>変換後のVB6コードを一括化したもの。\nが1行毎の区切りにあたる</returns>
	std::string Format(std::list<std::string>& _vb6_code_list);
};

