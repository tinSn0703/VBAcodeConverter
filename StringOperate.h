#pragma once

#include <string>

std::wstring multi_to_wide_capi(std::string const& src);

std::string wide_to_multi_capi(std::wstring const& src);

std::uint16_t count_num_word_in_str(const std::string& str, const std::string& word, std::string::size_type pos = 0, std::string::size_type end = std::string::npos);

std::uint16_t rcount_num_word_in_str(const std::string& str, const std::string& word, std::string::size_type pos = std::string::npos, std::string::size_type end = -1);

/// <summary>��������̘A�����ďo�����镶����P���ɒu��������</summary>
/// <param name="_code"></param>
/// <param name="_char"></param>
/// <returns></returns>
std::string& trim_consecutive_char(std::string& _code, const char _char);