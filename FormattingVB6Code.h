#pragma once

#include "Formating.h"

/// <summary>
/// <para>ファイルから取り込んだVB6コードを整形する</para>
/// <para># 行の分割、複合を戻す処理の手順</para>
/// <para>1. ファイルを全て読み込む</para>
/// <para>2. 改行以外の空白を全て、半角スペースに変える</para>
/// <para>3. インデント、行末尾のスペースを全て削除</para>
/// <para>4. 複数の半角スペースは全て単数の半角スペースに変える</para>
/// <para>5. 分割された行を1行に戻す</para>
/// <para>6. 省略表記されたIf文を戻す。(次の処理時の読み込みを行いやすくするため)</para>
/// <para>7. 結合されれた行を戻す</para>
/// <para></para>
/// </summary>
class FormattingVB6Code : public Formating
{
	/// <summary>行が複数行に渡ったコードであるか?</summary>
	/// <param name="_code">対象コード</param>
	/// <returns>Yes/No</returns>
	bool is_continuous_line(const std::string& _code);

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
	/// <param name="_vb6_code_list">VB6コード。変換される</param>
	/// <returns>変換後のVB6コードを一括化したもの。\nが1行毎の区切りにあたる</returns>
	std::string Format(std::list<std::string>& _vb6_code_list);
};

