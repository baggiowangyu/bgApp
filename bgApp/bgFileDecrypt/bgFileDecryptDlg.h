
// bgFileDecryptDlg.h : ͷ�ļ�
//

#pragma once
#include "afxwin.h"
#include "afxcmn.h"

#include "Poco/Thread.h"

// CbgFileDecryptDlg �Ի���
class CbgFileDecryptDlg : public CDialog
{
// ����
public:
	CbgFileDecryptDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_BGFILEDECRYPT_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
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
