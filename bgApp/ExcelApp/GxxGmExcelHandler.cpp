#include "GxxGmExcelHandler.h"

#include "Poco/Path.h"

#include <odbcinst.h>
#include <afxdb.h>
#include <atlstr.h>

GxxGmExcelHandler::GxxGmExcelHandler()
: spXlApp(NULL)
, current_content_line(1)
{
	// ��ʼ��COM����
	CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
}

GxxGmExcelHandler::~GxxGmExcelHandler()
{
	// Uninitialize COM for this thread.
	CoUninitialize();
}

int GxxGmExcelHandler::InitializeWrite(const char * xls_path)
{
	file_path = xls_path;
	int errCode = InitializeExcel();

	// ����һ���µĹ�����
	spXlBook = spXlBooks->Add();

	return errCode;
}

int GxxGmExcelHandler::WriteHeader(const wchar_t *sheet_name, std::vector<std::wstring> &headers)
{
	int errCode = 0;

	int element_count = headers.size();

	// Get the active Worksheet and set its name.
	// ��ȡ��Ĺ�����������������
	spXlSheet = spXlBook->ActiveSheet;
	spXlSheet->Name = _bstr_t(sheet_name);

	// ����һ��1��x�ľ���
	VARIANT saNames;
	saNames.vt = VT_ARRAY | VT_VARIANT;

	SAFEARRAYBOUND sab[2];
	sab[0].lLbound = 1; sab[0].cElements = 1;
	sab[1].lLbound = 1; sab[1].cElements = element_count;
	saNames.parray = SafeArrayCreate(VT_VARIANT, 2, sab);

	HRESULT ret = SafeArrayPutName(saNames.parray, current_content_line, headers);
	if (FAILED(ret))
		return ret;

	wchar_t range_str[4096] = {0};
	wchar_t letter_A = L'A';
	wchar_t letter_last = letter_A + element_count - 1;
	swprintf(range_str, L"A%d:%c%d", current_content_line, letter_last, current_content_line);

	++current_content_line;

	VARIANT param;
	param.vt = VT_BSTR;
	param.bstrVal = SysAllocString(range_str);
	Excel::RangePtr spXlRange = spXlSheet->Range[param];

	spXlRange->Value2 = saNames;
	VariantClear(&saNames);

	return errCode;
}

int GxxGmExcelHandler::WriteLine(std::vector<std::wstring> item)
{
	int errCode = 0;

	int element_count = item.size();

	// ����һ��1��x�ľ���
	VARIANT saNames;
	saNames.vt = VT_ARRAY | VT_VARIANT;

	SAFEARRAYBOUND sab[2];
	sab[0].lLbound = 1; sab[0].cElements = 1;
	sab[1].lLbound = 1; sab[1].cElements = element_count;
	saNames.parray = SafeArrayCreate(VT_VARIANT, 2, sab);

	HRESULT ret = SafeArrayPutName(saNames.parray, current_content_line, item);
	if (FAILED(ret))
		return ret;

	wchar_t range_str[4096] = {0};
	wchar_t letter_A = L'A';
	wchar_t letter_last = letter_A + element_count - 1;
	swprintf(range_str, L"A%d:%c%d", current_content_line, letter_last, current_content_line);

	++current_content_line;

	VARIANT param;
	param.vt = VT_BSTR;
	param.bstrVal = SysAllocString(range_str);
	Excel::RangePtr spXlRange = spXlSheet->Range[param];

	spXlRange->Value2 = saNames;
	VariantClear(&saNames);

	return errCode;
}

void GxxGmExcelHandler::FinishWrite()
{
	// �õ���ǰĿ¼
	Poco::Path xls_path(Poco::Path::current());
	xls_path.append(file_path);

	// Convert the NULL-terminated string to BSTR
	variant_t vtFileName(xls_path.toString().c_str());

	spXlBook->SaveAs(vtFileName, Excel::xlOpenXMLWorkbook, vtMissing, 
		vtMissing, vtMissing, vtMissing, Excel::xlNoChange);

	spXlBook->Close();

	// Quit the Excel application. (i.e. Application.Quit)
	spXlApp->Quit();
}

int GxxGmExcelHandler::InitializeRead(const wchar_t * xls_path, const wchar_t *sheet_name)
{
	CString strXlsFileName = xls_path;
	CString strSql;
	strSql.Format(L"DRIVER={MICROSOFT EXCEL DRIVER (*.xls)};DSN='';FIRSTROWHASNAMES=1;READONLY=FALSE;CREATE_DB=\"%s\";DBQ=%s", strXlsFileName, strXlsFileName);

	CDatabase *pCDatabase = new CDatabase;   
	pCDatabase->OpenEx(strSql, CDatabase::noOdbcDialog);

	CString strSheetName = sheet_name;
	CString strSql;
	strSql.Format(L"SELECT * FROM [%s$A1:IV65536]", strSheetName);   

	CRecordset *pCRecordset = new CRecordset(pCDatabase);
	pCRecordset->Open(CRecordset::forwardOnly, strSql, CRecordset::readOnly);

	CString strField;
	pCRecordset->GetFieldValue((short)0, strField);//��ȡ��1�е�1�еĵ�Ԫ������
	pCRecordset->GetFieldValue((short)3, strField);//��1�е�4��
	pCRecordset->MoveNext();//�α�ָ����һ�У�����GetFieldValue�Ϳ��Զ�ȡ��һ�е�����

	return 0;
}

int GxxGmExcelHandler::ReadLine(const wchar_t *range, std::vector<std::wstring> &item)
{
	int errCode = 0;

	_variant_t v_rang(range);
	Excel::RangePtr spXlRange = spXlSheet->GetRange(v_rang);

	return errCode;
}

void GxxGmExcelHandler::FinishRead()
{

}

int GxxGmExcelHandler::InitializeExcel()
{
	int errCode = 0;

	HRESULT hr = spXlApp.CreateInstance(__uuidof(Excel::Application));
	if (FAILED(hr))
	{
		// ����EXCELʵ��ʧ�ܣ�
		return hr;
	}

	// ����Excelʵ���ķ���״̬
	spXlApp->Visible[0] = VARIANT_FALSE;
	spXlBooks = spXlApp->Workbooks;
	
	return errCode;
}

void GxxGmExcelHandler::CleanupExcel()
{
	// �����ͷŵ�EXCEL
	spXlApp->Quit();
}


HRESULT GxxGmExcelHandler::SafeArrayPutName(SAFEARRAY* psa, long index, std::vector<std::wstring> args)
{
	int errCode = 0;

	// Extract arguments...
	long indices[2] = {0, 0};
	int i = 0;
	std::vector<std::wstring>::iterator iter;
	for(iter = args.begin(); iter != args.end(); ++iter) 
	{
		int element_item_index = i + 1;

		indices[0] = index;
		indices[1] = element_item_index;

		VARIANT element;
		element.vt = VT_BSTR;
		element.bstrVal = SysAllocString(iter->c_str());

		// Copies the VARIANT into the SafeArray
		HRESULT hr = SafeArrayPutElement(psa, indices, (void*)&element);
		if (FAILED(hr))
		{
			errCode = hr;
			break;
		}

		++i;
	}

	return errCode;
}