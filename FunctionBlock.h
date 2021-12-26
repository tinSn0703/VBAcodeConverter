#pragma once

#include <string>
#include <vector>
#include <list>

#include "Atteribute.h"

class CodeBlock {};
class ArgumentCode : public CodeBlock {};

/// <summary>ä÷êîÇÃÉuÉçÉbÉN</summary>
class FunctionBlock : public CodeBlock
{
	bool _IsStatic;
	AccessAdornment _Access;
	std::string _FunctionName;
	std::string _ReturnType;
	//std::list<ArgumentCode> _Args;
	//std::vector<CodeBlock> _Code;

public:
	FunctionBlock();
	//FunctionBlock(std::string _ReturnType, std::string _FunctionName, std::list<ArgumentCode> _Args, std::vector<CodeBlock> _Code, AccessAdornment _Access = AccessAdornment::PUBLIC_ACCESS, bool _IsStatic = false);

	//void Initialize(std::string _ReturnType, std::string _FunctionName, std::list<ArgumentCode> _Args, std::vector<CodeBlock> _Code, AccessAdornment _Access = AccessAdornment::PUBLIC_ACCESS, bool _IsStatic = false);




};

