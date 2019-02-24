#ifndef _GxxGmExcelHandler_H_
#define _GxxGmExcelHandler_H_

#include <string>
#include <vector>

#pragma region Import the type libraries

#import "libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52" \
	rename("RGB", "MSORGB") \
	rename("DocumentProperties", "MSODocumentProperties")
// 或者
//#import "C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE12\\MSO.DLL" \
//    rename("RGB", "MSORGB") \
//    rename("DocumentProperties", "MSODocumentProperties")

using namespace Office;

#import "libid:0002E157-0000-0000-C000-000000000046"
// 或者
//#import "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA6\\VBE6EXT.OLB"

using namespace VBIDE;

#import "libid:00020813-0000-0000-C000-000000000046" \
	rename("DialogBox", "ExcelDialogBox") \
	rename("RGB", "ExcelRGB") \
	rename("CopyFile", "ExcelCopyFile") \
	rename("ReplaceText", "ExcelReplaceText") \
	no_auto_exclude
// 或者
//#import "C:\\Program Files\\Microsoft Office\\Office12\\EXCEL.EXE" \
//    rename("DialogBox", "ExcelDialogBox") \
//    rename("RGB", "ExcelRGB") \
//    rename("CopyFile", "ExcelCopyFile") \
//    rename("ReplaceText", "ExcelReplaceText") \
//    no_auto_exclude

#pragma endregion

class GxxGmExcelHandler
{
public:
	GxxGmExcelHandler();
	~GxxGmExcelHandler();

public:
	//////////////////////////////////////////////////////////////////////////
	//
	// 创建EXCEL
	//
	//////////////////////////////////////////////////////////////////////////

public:
	//////////////////////////////////////////////////////////////////////////
	//
	// 写入EXCEL
	//
	//////////////////////////////////////////////////////////////////////////
	int InitializeWrite(const char * xls_path);
	int WriteHeader(const wchar_t *sheet_name, std::vector<std::wstring> &headers);
	int WriteLine(std::vector<std::wstring> item);
	void FinishWrite();

public:
	//////////////////////////////////////////////////////////////////////////
	//
	// 读取EXCEL
	//
	//////////////////////////////////////////////////////////////////////////
	int InitializeRead(const wchar_t * xls_path, const wchar_t *sheet_name);
	int ReadLine(const wchar_t *range, std::vector<std::wstring> &item);
	void FinishRead();

private:
	int InitializeExcel();
	void CleanupExcel();

	//HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp,	LPOLESTR ptName, int cArgs, ...);
	//HRESULT SafeArrayPutName(SAFEARRAY* psa, long index, int args, ...);
	HRESULT SafeArrayPutName(SAFEARRAY* psa, long index, std::vector<std::wstring> args);

private:
	Excel::_ApplicationPtr spXlApp;
	Excel::WorkbooksPtr spXlBooks;
	Excel::_WorkbookPtr spXlBook;
	Excel::_WorksheetPtr spXlSheet;
	std::string file_path;

	long current_content_line;
};

#endif//_GxxGmExcelHandler_H_
