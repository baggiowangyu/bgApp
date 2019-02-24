// ExcelApp.cpp : 定义控制台应用程序的入口点。
//

#include "stdafx.h"
#include <iostream>
#include <string>
#include <vector>

#include "GxxGmExcelHandler.h"


void TestWrite()
{
	std::vector<std::wstring> header;
	header.push_back(L"设备ID");
	header.push_back(L"设备名称");
	header.push_back(L"设备厂商");
	header.push_back(L"设备门类");
	header.push_back(L"连接信息");
	header.push_back(L"配置版本");
	header.push_back(L"用户名称");
	header.push_back(L"用户密码");
	header.push_back(L"设备编号");
	header.push_back(L"扩展信息");
	header.push_back(L"国标编码");
	header.push_back(L"缩写名称");
	header.push_back(L"设备版本");
	header.push_back(L"归属网关");

	GxxGmExcelHandler excel;
	excel.InitializeWrite("设备信息表.xls");
	excel.WriteHeader(L"设备表", header);

	// 写入内容
	std::vector<std::wstring> content;
	content.push_back(L"100001");
	content.push_back(L"王煜的执法仪");
	content.push_back(L"高新兴");
	content.push_back(L"执法仪");
	content.push_back(L"[][][]");
	content.push_back(L"1");
	content.push_back(L"admin");
	content.push_back(L"admin");
	content.push_back(L"123456");
	content.push_back(L"");
	content.push_back(L"44000000132000000001");
	content.push_back(L"1");
	content.push_back(L"v1.0");
	content.push_back(L"600001");
	excel.WriteLine(content);

	excel.FinishWrite();
}

void TestRead()
{
	GxxGmExcelHandler excel;
	excel.InitializeRead(L"设备信息表.xls", L"设备表");

	std::vector<std::wstring> content;
	excel.ReadLine(L"A1:N1", content);
	excel.FinishRead();
}

int _tmain(int argc, _TCHAR* argv[])
{
	//TestWrite();
	TestRead();

	return 0;
}

