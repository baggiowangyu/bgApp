#include "stdafx.h"
#include "bgMfcExcelModule.h"

// 参考页面：https://blog.csdn.net/cai_niaocainiao/article/details/81806928
// 参考页面：https://www.cnblogs.com/tianya2543/p/4165997.html

bgMfcExcelModule::bgMfcExcelModule()
: rows_(0)
, cols_(0)
, current_row_(0)
{
	HRESULT hr = CoInitialize(NULL);
	if (FAILED(hr))
	{
		// 初始化失败
	}
}

bgMfcExcelModule::~bgMfcExcelModule()
{
	// 
	CoUninitialize();
}

int bgMfcExcelModule::WriteInitialize(const wchar_t *xls_path, const wchar_t *sheet_name)
{
	int errCode = 0;

	xls_path_ = xls_path;
	std::wstring xls_tmp_path = xls_path_ + L".tmp";
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	BOOL bret = app1.CreateDispatch(_T("Excel.Application"), NULL);
	if (!bret)
	{
		return -1;
	}

	books.AttachDispatch(app1.get_Workbooks());
	lpDisp = books.Open(xls_tmp_path.c_str(), covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional);
	book.AttachDispatch(lpDisp); 

	//得到Worksheets   
	sheets.AttachDispatch(book.get_Worksheets());
	//sheet = sheets.get_Item(COleVariant((short)1));
	//得到当前活跃sheet 
	//如果有单元格正处于编辑状态中，此操作不能返回，会一直等待 
	lpDisp = book.get_ActiveSheet();
	sheet.AttachDispatch(lpDisp);

	// 设置sheet名称？
	sheet.put_Name(sheet_name);
	current_row_ = 0;

	return errCode;
}

int bgMfcExcelModule::WriteHeader(std::vector<std::wstring> headers)
{
	int errCode = 0;

	++current_row_;
	COleVariant vResult;
	for (long col_index = 1; col_index <= headers.size(); ++col_index)
	{
		range.AttachDispatch(sheet.set_Cells());
		range.AttachDispatch(range.get_Item(COleVariant(current_row_), COleVariant(col_index)).pdispVal);   //第一变量是行，第二个变量是列，即第二行第一列
		vResult = range.get_Value2();

		CString str;
		if (vResult.vt == VT_BSTR) //字符串  
		{
			str = vResult.bstrVal;
		}
		else if (vResult.vt == VT_R8) //8字节的数字  
		{
			str.Format(_T("%f"), vResult.dblVal);
		}
		else if (vResult.vt == VT_NULL)
		{
			str = L""
		}

		headers.push_back(str.GetString());
	}

	return errCode;
}

int bgMfcExcelModule::WriteLine(std::vector<std::wstring> rows)
{
	int errCode = 0;

	return errCode;
}

void bgMfcExcelModule::WriteFinish()
{
	//
}

int bgMfcExcelModule::ReadInitialize(const wchar_t *xls_path, const wchar_t *sheet_name)
{
	int errCode = 0;

	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	BOOL bret = app1.CreateDispatch(_T("Excel.Application"), NULL);
	if (!bret)
	{
		return -1;
	}

	books.AttachDispatch(app1.get_Workbooks());
	lpDisp = books.Open(xls_path, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional, covOptional, 
		covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional);

	//得到Workbook    
	book.AttachDispatch(lpDisp);
	//得到Worksheets   
	sheets.AttachDispatch(book.get_Worksheets());
	//sheet = sheets.get_Item(COleVariant((short)1));
	//得到当前活跃sheet 
	//如果有单元格正处于编辑状态中，此操作不能返回，会一直等待 
	lpDisp = book.get_ActiveSheet();
	sheet.AttachDispatch(lpDisp);

	// 得到已使用的行列数
	CRange usedRange;
	usedRange.AttachDispatch(sheet.get_UsedRange());

	//已经使用的行数
	range.AttachDispatch(usedRange.get_Rows());
	rows_ = range.get_Count();                   

	//已经使用的列数
	range.AttachDispatch(usedRange.get_Columns());
	cols_ = range.get_Count(); 

	current_row_ = 0;

	return errCode;
}

int bgMfcExcelModule::ReadHeader(std::vector<std::wstring> &headers)
{
	int errCode = 0;

	//读取第一个单元格的值 
	COleVariant vResult;
	current_row_ = 1;
	for (long col_index = 1; col_index <= cols_; ++col_index)
	{
		range.AttachDispatch(sheet.get_Cells());
		range.AttachDispatch(range.get_Item(COleVariant(current_row_), COleVariant(col_index)).pdispVal);   //第一变量是行，第二个变量是列，即第二行第一列
		vResult = range.get_Value2();

		CString str;
		if (vResult.vt == VT_BSTR) //字符串  
		{
			str = vResult.bstrVal;
		}
		else if (vResult.vt == VT_R8) //8字节的数字  
		{
			str.Format(_T("%f"), vResult.dblVal);
		}
		else if (vResult.vt == VT_NULL)
		{
			str = L""
		}

		headers.push_back(str.GetString());
	}

	return errCode;
}

int bgMfcExcelModule::ReadNextLine(std::vector<std::wstring> &rows)
{
	int errCode = 0;

	++current_row_;
	COleVariant vResult;
	for (long col_index = 1; col_index <= cols_; ++col_index)
	{
		range.AttachDispatch(sheet.get_Cells());
		range.AttachDispatch(range.get_Item(COleVariant(current_row_), COleVariant(col_index)).pdispVal);   //第一变量是行，第二个变量是列，即第二行第一列
		vResult = range.get_Value2();

		CString str;
		if (vResult.vt == VT_BSTR) //字符串  
		{
			str = vResult.bstrVal;
		}
		else if (vResult.vt == VT_R8) //8字节的数字  
		{
			str.Format(_T("%f"), vResult.dblVal);
		}
		else if (vResult.vt == VT_NULL)
		{
			str = L""
		}

		rows.push_back(str.GetString());
	}

	return errCode;
}

void bgMfcExcelModule::ReadFinish()
{
	//
	books.Close();
	app1.Quit(); 
	//释放对象      
	range.ReleaseDispatch();
	sheet.ReleaseDispatch();
	sheets.ReleaseDispatch();
	book.ReleaseDispatch();
	books.ReleaseDispatch();
	app1.ReleaseDispatch();
}