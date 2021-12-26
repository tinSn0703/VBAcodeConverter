
#include <vector>
#include <boost/algorithm/string.hpp>

#include "StringOperate.h"

//-------------------------------------------------------------------------

std::wstring multi_to_wide_capi(std::string const& src)
{
	std::size_t converted{};
	std::vector<wchar_t> dest(src.size(), L'\0');
	if (::_mbstowcs_s_l(&converted, dest.data(), dest.size(), src.data(), _TRUNCATE, ::_create_locale(LC_ALL, "jpn")) != 0)
	{
		throw std::system_error{ errno, std::system_category() };
	}
	dest.resize(std::char_traits<wchar_t>::length(dest.data()));
	dest.shrink_to_fit();
	return std::wstring(dest.begin(), dest.end());
}

//-------------------------------------------------------------------------

std::string wide_to_multi_capi(std::wstring const& src)
{
	std::size_t converted{};
	std::vector<char> dest(src.size() * sizeof(wchar_t) + 1, '\0');
	if (::_wcstombs_s_l(&converted, dest.data(), dest.size(), src.data(), _TRUNCATE, ::_create_locale(LC_ALL, "jpn")) != 0)
	{
		throw std::system_error{ errno, std::system_category() };
	}
	dest.resize(std::char_traits<char>::length(dest.data()));
	dest.shrink_to_fit();
	return std::string(dest.begin(), dest.end());
}

//-------------------------------------------------------------------------

std::uint16_t count_num_word_in_str(const std::string& str, const std::string& word, std::string::size_type pos, std::string::size_type end)
{
	std::uint16_t count = 0;
	if (end == std::string::npos) end = str.size();
	for (; ((pos = str.find(word, pos)) != std::string::npos) && (pos < end); pos += word.size())	++count;
	return count;
}

//-------------------------------------------------------------------------

std::uint16_t rcount_num_word_in_str(const std::string& str, const std::string& word, std::string::size_type pos, std::string::size_type end)
{
	std::uint16_t count = 0;
	if (pos == std::string::npos) pos = str.size() - 1;
	for (; ((pos = str.find(word, pos)) != std::string::npos) && (end < pos); pos -= word.size())	++count;
	return count;
}

//-------------------------------------------------------------------------

std::string& trim_consecutive_char(std::string& _str, const char _char)
{
	for (auto _pos = _str.find(_char); _pos != std::string::npos; )
	{
		const auto _end_pos = _str.find_first_not_of(' ', _pos);

		if (1 < (_end_pos - _pos))
		{
			_str = _str.substr(0, _pos) + _char + _str.substr(_end_pos, _str.size() - 1);
			_pos = _str.find(_char, _pos + 1);
		}
		else
		{
			_pos = _str.find(_char, _end_pos);
		}
	}

	return _str;
}