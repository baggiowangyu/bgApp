
// bgExcelAppDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "bgExcelApp.h"
#include "bgExcelAppDlg.h"

#include "bgMfcExcelModule.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
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


// CbgExcelAppDlg 对话框




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


// CbgExcelAppDlg 消息处理程序

BOOL CbgExcelAppDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
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

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
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

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CbgExcelAppDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CbgExcelAppDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CbgExcelAppDlg::OnBnClickedButton1()
{
	std::vector<std::wstring> header;
	header.push_back(L"设备ID");
	header.push_back(L"设备名称");
	header.push_back(L"设备厂商");
	header.push_back(L"设备门类");
	header.push_back(L"连接信息");
	header.push_back(L"配置版本");
	header.push_back(L"用户名称");
	header.push_back(L"用户密码");
	header.push_back(L"设备编号");
	header.push_back(L"扩展信息");
	header.push_back(L"国标编码");
	header.push_back(L"缩写名称");
	header.push_back(L"设备版本");
	header.push_back(L"所属网关");

	std::vector<std::wstring> item;
	item.push_back(L"100001");
	item.push_back(L"高新兴4G执法记录仪");
	item.push_back(L"高新兴国迈");
	item.push_back(L"执法记录仪");
	item.push_back(L"conn:[][][]");
	item.push_back(L"1");
	item.push_back(L"admin");
	item.push_back(L"admin");
	item.push_back(L"100001");
	item.push_back(L"扩展A");
	item.push_back(L"44000000001320000001");
	item.push_back(L"aaa");
	item.push_back(L"V1");
	item.push_back(L"600001");

	bgMfcExcelModule excel;

	excel.WriteInitialize(L"D:\\设备信息表.xlsx", L"设备表");
	excel.WriteHeader(header);
	excel.WriteLine(item);
	excel.WriteFinish();

	std::vector<std::wstring> read_header;
	excel.ReadInitialize(L"D:\\设备信息表.xlsx", L"设备表");
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
