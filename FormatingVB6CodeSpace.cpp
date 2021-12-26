
#include <regex>

#include "FormatingVB6CodeSpace.h"

//-------------------------------------------------------------------------

std::vector<size_t> FormatingVB6CodeSpace::find_all_quotation(std::string& _code)
{
	std::vector<size_t> _ret;

	const auto _begin = std::begin(_code);
	for (auto _itr = std::begin(_code), _end = std::end(_code); _itr != _end; _itr++)
	{
		if (((*_itr) == '\"') || ((*_itr) == '\''))	_ret.push_back((size_t)(_itr - _begin));
	}
	
	return _ret;
}

//-------------------------------------------------------------------------

bool FormatingVB6CodeSpace::is_this_in_string(const std::vector<size_t>& quotation_positions, size_t position)
{
	if (quotation_positions.size() < 1) return false;

	std::uint16_t i = 0;
	for (auto quotation_position : quotation_positions)
	{
		if (position < quotation_position) break;
		i++;
	}

	return ((i % 2) != 0);
}

//-------------------------------------------------------------------------

std::string FormatingVB6CodeSpace::Format(std::list<std::string>& _vb6_code_list)
{
	// æ“ª["]––”ö‚Å•ªŠ„‚·‚é
	// •ªŠ„Œã‚ÌÅŒã‚Ì‚İAÅ‰‚ÉoŒ»‚·‚é[']‚Å•ªŠ„‚·‚é

	return std::string();
}

//-------------------------------------------------------------------------