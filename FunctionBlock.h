#pragma once

/// <summary>ŠÖ”‚ÌƒuƒƒbƒN</summary>
class FunctionBlock : public CodeBlock
{
	bool _IsStatic;
	AccessAdornment _Access;
	std::string _FunctionName;
	std::string _ReturnType;
	std::list<ArgumentCode> _Args;
	std::vector<CodeBlock> _Code;

public:
	FunctionBlock();

	/*
	void SplitElement(std::smatch _Code)
	{
		if ((_Code[1] == true) && (_Code[1].str() == "Private"))		this->_Access = AccessAdornment::PRIVATE_ACCESS;
		else if ((_Code[1] == true) && (_Code[1].str() == "Friend"))	this->_Access = AccessAdornment::FRIEND_ACCESS;
		else															this->_Access = AccessAdornment::PUBLIC_ACCESS;

		this->_FunctionName = _Code[3].str();
		this->_ReturnType = _Code[5].str();

		//this->_Code = _Code[7].str();
	}
	*/
};

