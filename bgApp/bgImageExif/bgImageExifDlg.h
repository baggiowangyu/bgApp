
// bgImageExifDlg.h : ͷ�ļ�
//

#pragma once


// CbgImageExifDlg �Ի���
class CbgImageExifDlg : public CDialog
{
// ����
public:
	CbgImageExifDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_BGIMAGEEXIF_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
};
