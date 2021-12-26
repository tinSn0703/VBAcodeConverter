
#include <iostream>
#include <fstream>
#include <string>
#include <list>
#include <regex>
#include <filesystem>

#include <boost/version.hpp>
#include <boost/algorithm/string.hpp>

#include "FormattingVB6Code.h"
#include "StringOperate.h"

#define DEBUG_MODE 0

/// <summary>VB6の、途中改行、一行化された行を元の状態に戻す。コードのテスト</summary>
/// <param name="argc"></param>
/// <param name="argv"></param>
/// <param name="envp"></param>
/// <returns></returns>
int OperateVB6codeLine(int argc, char* argv[], char* envp[])
{
	std::ifstream _vb6_file(R"(D:\Project\VBAcodeConverter\test_vb6_base.bas)");
	if (_vb6_file.fail()) throw std::exception("");
	
	std::list<std::string> _vb6_code_list;
	std::string _vb6_code;
	while (!_vb6_file.eof()) { std::getline(_vb6_file, _vb6_code); _vb6_code_list.push_back(_vb6_code); }
	_vb6_file.close();
	
	FormattingVB6Code _format;
	_vb6_code = _format.Format(_vb6_code_list);
	
	std::ofstream _output_file(R"(D:\Project\VBAcodeConverter\test_vb6.bas)");
	_output_file << _vb6_code;
	_output_file.close();
	
	// 正常終了
	return(0);
}
