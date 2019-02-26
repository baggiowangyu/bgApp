#ifndef _bgMfcExcelModule_H_
#define _bgMfcExcelModule_H_

#include <vector>
#include <string>

#include "CWorkbook.h"		// ������������
#include "CWorkbooks.h"		// ͳ�����еĹ�����
#include "CApplication.h"	// ExcelӦ�ó�����,�������Ǵ򿪵�������ExcelӦ��
#include "CRange.h"			// �����࣬��EXcel�Ĵ󲿷ֲ�����Ҫ��������к�
#include "CWorksheet.h"		// �������еĵ���������
#include "CWorksheets.h"	// ͳ�ܵ�ǰ�������е����й�����

class bgMfcExcelModule
{
public:
	bgMfcExcelModule();
	~bgMfcExcelModule();

public:
	int WriteInitialize(const wchar_t *xls_path, const wchar_t *sheet_name);
	int WriteHeader(std::vector<std::wstring> headers);
	int WriteLine(std::vector<std::wstring> rows);
	void WriteFinish();

public:
	int ReadInitialize(const wchar_t *xls_path, const wchar_t *sheet_name);
	int ReadHeader(std::vector<std::wstring> &headers);
	int ReadNextLine(std::vector<std::wstring> &rows);
	void ReadFinish();

private:
	std::wstring xls_path_;
	std::wstring sheet_name_;

private:
	CApplication app1;
	CWorkbooks books;
	CWorkbook book;
	CWorksheets sheets;
	CWorksheet sheet;
	CRange range;
	CRange iCell;
	LPDISPATCH lpDisp;

	long rows_;
	long cols_;
	long current_row_;
};

#endif//_bgMfcExcelModule_H_
