#pragma once

#include <string>

/// <summary>ÉAÉNÉZÉXèCè¸éq</summary>
enum class AccessAdornment
{
	LOCAL_ACCESS,
	PUBLIC_ACCESS,
	PRIVATE_ACCESS,
	FRIEND_ACCESS,
};

inline std::string ToAceessString(const AccessAdornment _Access)
{
	switch (_Access)
	{
		case AccessAdornment::LOCAL_ACCESS:		return "";
		case AccessAdornment::PUBLIC_ACCESS:	return "public";
		case AccessAdornment::PRIVATE_ACCESS:	return "private";
		case AccessAdornment::FRIEND_ACCESS:	return "friend";
		default: break;
	}
}

/// <summary>ââéZéq</summary>
enum class Operator
{
	MULTIPLY_OPERATOR,	//èÊéZ
	RAISING_OPERATOR,	//ó›èÊ
	DIVIDE_OPERATOR,	//èúéZ
	SURPLUS_OPERATOR,	//ó]ÇË
	PLUS_OPERATOR,		//â¡éZ
	MINUS_OPERATOR,		//å∏éZ
	EQUAL_OPERATOR,		//ìôâø
	IS_OPERATOR,		//å^å›ä∑ämîF
	LIKE_OPRATOR,		//ï∂éöóÒî‰är(ê≥ãKï\åªÇ≈çƒåªÇ∑ÇÈ)
	AND_OPERATOR,		// A & B
	OR_OPERATOR,		// A | B
	NOT_OPRATOR,		// A ! B
	XOR_OPERATOR,		// A ^ B
	EQV_OPRATOR,		// !(A ^ B)
	IMP_OPRATOR,		// (!A) | B
};

/// <summary></summary>
/// <param name="_Operator"></param>
/// <param name="_RightValue"></param>
/// <param name="_LeftValue"></param>
/// <returns></returns>
inline std::string ToString(const Operator _Operator, const std::string& _RightValue, const std::string& _LeftValue)
{
	switch (_Operator)
	{
		case Operator::MULTIPLY_OPERATOR:	return _RightValue + " * " + _LeftValue;
		case Operator::RAISING_OPERATOR:	return "Math.Pow(" + _RightValue + ", " + _LeftValue + ")";
		case Operator::DIVIDE_OPERATOR:		return _RightValue + " / " + _LeftValue;
		case Operator::SURPLUS_OPERATOR:	return _RightValue + " % " + _LeftValue;
		case Operator::PLUS_OPERATOR:		return _RightValue + " + " + _LeftValue;
		case Operator::MINUS_OPERATOR:		return _RightValue + " - " + _LeftValue;
		case Operator::EQUAL_OPERATOR:		return _RightValue + " = " + _LeftValue;
		case Operator::IS_OPERATOR:			return _RightValue + " is " + _LeftValue;
		case Operator::LIKE_OPRATOR:		return _RightValue + "" + _LeftValue;
		case Operator::NOT_OPRATOR:			return " !" + _RightValue;
		case Operator::AND_OPERATOR:		return _RightValue + " && " + _LeftValue;
		case Operator::OR_OPERATOR:			return _RightValue + " || " + _LeftValue;
		case Operator::XOR_OPERATOR:		return _RightValue + " ^ " + _LeftValue;
		case Operator::EQV_OPRATOR:			return "!(" + _RightValue + " ^ " + _LeftValue + ")";
		case Operator::IMP_OPRATOR:			return "(!(" + _RightValue + ") | " + _LeftValue + ")";
	}
}
