
// bgFileDecrypt.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CbgFileDecryptApp:
// �йش����ʵ�֣������ bgFileDecrypt.cpp
//

class CbgFileDecryptApp : public CWinAppEx
{
public:
	CbgFileDecryptApp();

// ��д
	public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CbgFileDecryptApp theApp;