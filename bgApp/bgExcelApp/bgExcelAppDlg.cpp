
// bgExcelAppDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "bgExcelApp.h"
#include "bgExcelAppDlg.h"

#include "bgMfcExcelModule.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CbgExcelAppDlg �Ի���




CbgExcelAppDlg::CbgExcelAppDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CbgExcelAppDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CbgExcelAppDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CbgExcelAppDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BUTTON1, &CbgExcelAppDlg::OnBnClickedButton1)
END_MESSAGE_MAP()


// CbgExcelAppDlg ��Ϣ�������

BOOL CbgExcelAppDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CbgExcelAppDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CbgExcelAppDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CbgExcelAppDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CbgExcelAppDlg::OnBnClickedButton1()
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

	std::vector<std::wstring> item;
	item.push_back(L"100001");
	item.push_back(L"������4Gִ����¼��");
	item.push_back(L"�����˹���");
	item.push_back(L"ִ����¼��");
	item.push_back(L"conn:[][][]");
	item.push_back(L"1");
	item.push_back(L"admin");
	item.push_back(L"admin");
	item.push_back(L"100001");
	item.push_back(L"��չA");
	item.push_back(L"44000000001320000001");
	item.push_back(L"aaa");
	item.push_back(L"V1");
	item.push_back(L"600001");

	bgMfcExcelModule excel;

	excel.WriteInitialize(L"D:\\�豸��Ϣ��.xlsx", L"�豸��");
	excel.WriteHeader(header);
	excel.WriteLine(item);
	excel.WriteFinish();

	std::vector<std::wstring> read_header;
	excel.ReadInitialize(L"D:\\�豸��Ϣ��.xlsx", L"�豸��");
	excel.ReadHeader(read_header);

	while (true)
	{
		std::vector<std::wstring> read_item;
		int errCode = excel.ReadNextLine(read_item);
		if (errCode != 0)
		{
			break;
		}
	}

	excel.ReadFinish();
	
}
