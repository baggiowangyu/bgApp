// ExcelApp.cpp : �������̨Ӧ�ó������ڵ㡣
//

#include "stdafx.h"
#include <iostream>
#include <string>
#include <vector>

#include "GxxGmExcelHandler.h"


void TestWrite()
{
	std::vector<std::wstring> header;
	header.push_back(L"�豸ID");
	header.push_back(L"�豸����");
	header.push_back(L"�豸����");
	header.push_back(L"�豸����");
	header.push_back(L"������Ϣ");
	header.push_back(L"���ð汾");
	header.push_back(L"�û�����");
	header.push_back(L"�û�����");
	header.push_back(L"�豸���");
	header.push_back(L"��չ��Ϣ");
	header.push_back(L"�������");
	header.push_back(L"��д����");
	header.push_back(L"�豸�汾");
	header.push_back(L"��������");

	GxxGmExcelHandler excel;
	excel.InitializeWrite("�豸��Ϣ��.xls");
	excel.WriteHeader(L"�豸��", header);

	// д������
	std::vector<std::wstring> content;
	content.push_back(L"100001");
	content.push_back(L"���ϵ�ִ����");
	content.push_back(L"������");
	content.push_back(L"ִ����");
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
	excel.InitializeRead(L"�豸��Ϣ��.xls", L"�豸��");

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

