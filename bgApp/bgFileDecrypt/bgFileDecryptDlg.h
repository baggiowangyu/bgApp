
// bgFileDecryptDlg.h : 头文件
//

#pragma once
#include "afxwin.h"
#include "afxcmn.h"

#include "Poco/Thread.h"

// CbgFileDecryptDlg 对话框
class CbgFileDecryptDlg : public CDialog
{
// 构造
public:
	CbgFileDecryptDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_BGFILEDECRYPT_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	CEdit m_cPassword;
	CListCtrl m_cFileList;
	CProgressCtrl m_cWorkingProgress;

public:
	CString password;
	Poco::Thread decrypt_thread_;
	static void DecryptThread(void* param);

public:
	afx_msg void OnDropFiles(HDROP hDropInfo);
	afx_msg void OnBnClickedBtnStart();
	afx_msg void OnLvnKeydownListFile(NMHDR *pNMHDR, LRESULT *pResult);
};
