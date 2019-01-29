
// bgFileEncryptDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "bgFileEncrypt.h"
#include "bgFileEncryptDlg.h"


#include "openssl/bio.h"
#include "openssl/ssl.h"
#include "openssl/err.h"
#include "openssl/evp.h"
#include "openssl/ossl_typ.h"
#include "openssl/evp_locl.h"

#include "Poco/Base64Encoder.h"
#include <sstream>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

//// 内置密钥，用于加密口令
//const char *internal_password = "12345678qwertyui";


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


// CbgFileEncryptDlg 对话框



#ifndef WM_COPYGLOBALDATA
#define WM_COPYGLOBALDATA 0x0049
#endif

typedef WINUSERAPI BOOL WINAPI CHANGEWINDOWMESSAGEFILTER(UINT message, DWORD dwFlag);


//修复OnDropFile在win7管理员权限下接收不到的问题
void DropFileFix()
{
	HINSTANCE hDllInst = LoadLibraryW(_T("user32.dll"));
	if (hDllInst)
	{
		CHANGEWINDOWMESSAGEFILTER *pAddMessageFilterFunc = (CHANGEWINDOWMESSAGEFILTER *)GetProcAddress(hDllInst, "ChangeWindowMessageFilter");
		if (pAddMessageFilterFunc)
		{
			pAddMessageFilterFunc(WM_DROPFILES, MSGFLT_ADD);
			pAddMessageFilterFunc(WM_COPYDATA, MSGFLT_ADD);
			pAddMessageFilterFunc(0x0049, MSGFLT_ADD);
		}
		FreeLibrary(hDllInst);
	}
}


CbgFileEncryptDlg::CbgFileEncryptDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CbgFileEncryptDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CbgFileEncryptDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EDIT_PASSWORD, m_cPassword);
	DDX_Control(pDX, IDC_LIST_FILES, m_cFileList);
	DDX_Control(pDX, IDC_PROGRESS_WORKING, m_cWorkingProgress);
}

BEGIN_MESSAGE_MAP(CbgFileEncryptDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_WM_DROPFILES()
	ON_BN_CLICKED(IDC_BTN_START, &CbgFileEncryptDlg::OnBnClickedBtnStart)
	ON_NOTIFY(LVN_KEYDOWN, IDC_LIST_FILES, &CbgFileEncryptDlg::OnLvnKeydownListFiles)
END_MESSAGE_MAP()


// CbgFileEncryptDlg 消息处理程序

BOOL CbgFileEncryptDlg::OnInitDialog()
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
	SetWindowText(_T("留留游园会 加密工具 V1.0"));
	OpenSSL_add_all_algorithms();

	DropFileFix();

	m_cFileList.InsertColumn(0, _T("文件名"), LVCFMT_LEFT, 450);
	m_cPassword.LimitText(16);

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CbgFileEncryptDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void CbgFileEncryptDlg::OnPaint()
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
HCURSOR CbgFileEncryptDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CbgFileEncryptDlg::OnDropFiles(HDROP hDropInfo)
{
	// 如果线程在工作中，则不处理文件拖动
	if (!encrypt_thread_.isRunning())
	{
		//取得被拖动文件的数目
		int total_drop_count = DragQueryFile(hDropInfo, -1, NULL, 0);

		for(int index = 0; index < total_drop_count; ++index)
		{
			//获得拖曳的第i个文件的文件名
			TCHAR wcStr[MAX_PATH] = {0};
			DragQueryFile(hDropInfo, index, wcStr, MAX_PATH);

			int count = m_cFileList.GetItemCount();
			m_cFileList.InsertItem(count, wcStr);
		} 

		//拖放结束后,释放内存
		DragFinish(hDropInfo);
	}

	CDialog::OnDropFiles(hDropInfo);
}

void CbgFileEncryptDlg::OnBnClickedBtnStart()
{
	// 先检查加密线程是否在工作中，如果在工作，那么就弹出告警
	if (!encrypt_thread_.isRunning())
	{
		// 读取加密密钥
		m_cPassword.GetWindowText(password);

		// 启动线程
		encrypt_thread_.start(CbgFileEncryptDlg::EncryptThread, this);
	}
}


void CbgFileEncryptDlg::EncryptThread(void* param)
{
	CbgFileEncryptDlg *dlg = (CbgFileEncryptDlg *)param;

	int errCode = 0;
	CWnd *pcwnd = dlg->GetDlgItem(IDC_STATIC_STATE);

	// 在这里处理这轮加密的相关参数
	EVP_CIPHER_CTX *ctx = EVP_CIPHER_CTX_new();

	// 处理密钥
	// 定义向量
	unsigned char iv[17] = "\xc2\x86\x69\x6d\x88\x7c\x9a\xa0\x61\x1b\xbb\x3e\x20\x25\xa4\x5a";

	// 定义密钥
	int count = dlg->m_cFileList.GetItemCount();
	for(int index = 0; index < count; ++index)
	{
		USES_CONVERSION;
		unsigned char key[17] = {0};
		memset(key, 0, 17);
		memcpy_s(key, 16, T2A(dlg->password.GetBuffer(0)), dlg->password.GetLength());
		errCode = EVP_EncryptInit_ex(ctx, EVP_aes_128_cbc(), NULL, key, iv);

		TCHAR path[4096] = {0};
		dlg->m_cFileList.GetItemText(index, 0, path, 4096);

		TCHAR state_msg[4096] = {0};
		_stprintf_s(state_msg, 4096, _T("[%d]正在加密%s..."), index + 1, path);
		pcwnd->SetWindowText(state_msg);

		// 打开源文件
		HANDLE file_handle = CreateFileW(path, GENERIC_ALL, FILE_SHARE_READ|FILE_SHARE_WRITE|FILE_SHARE_DELETE, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
		if (file_handle == INVALID_HANDLE_VALUE)
		{
			errCode = GetLastError();
			_stprintf_s(state_msg, 4096, _T("[%d]打开文件%s失败，错误码:%d"), index + 1, path, errCode);
			pcwnd->SetWindowText(state_msg);
			break ;
		}

		// 打开密文文件
		TCHAR new_path[4096] = {0};
		_stprintf_s(new_path, 4096, _T("%s.cipher"), path);
		HANDLE new_file_handle = CreateFileW(new_path, GENERIC_ALL, FILE_SHARE_READ|FILE_SHARE_WRITE|FILE_SHARE_DELETE, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
		if (file_handle == INVALID_HANDLE_VALUE)
		{
			errCode = GetLastError();
			_stprintf_s(state_msg, 4096, _T("[%d]打开文件%s失败，错误码:%d"), index + 1, new_path, errCode);
			pcwnd->SetWindowText(state_msg);

			CloseHandle(file_handle);
			break ;
		}

		// 先将口令Base64，然后记录Base64的长度，然后先写入4字节长度，再写入Base64
		std::ostringstream str;
		Poco::Base64Encoder encoder(str);
		encoder << key;
		
		std::string base64encode_key = str.str();
		encoder.close();
		DWORD base64_len = base64encode_key.size();

		DWORD written = 0;
		WriteFile(new_file_handle, &base64_len, sizeof(DWORD), &written, NULL);
		WriteFile(new_file_handle, base64encode_key.c_str(), base64_len, &written, NULL);

		// 获取文件大小
		DWORD file_size = GetFileSize(file_handle, NULL);
		dlg->m_cWorkingProgress.SetRange32(0, file_size);

		// 读取文件
		int progress = 0;
		dlg->m_cWorkingProgress.SetPos(progress);

		bool is_failed = false;
		DWORD readed = 0;
		unsigned char data_buffer[4096] = {0};
		while (ReadFile(file_handle, data_buffer, 4096, &readed, NULL))
		{
			if (readed == 0)
				break;
			
			// 更新加密
			int cipher_len = 8192;
			unsigned char cipher[8192] = {0};
			errCode = EVP_EncryptUpdate(ctx, cipher, &cipher_len, data_buffer, readed);
			if (errCode != 1)
			{
				is_failed = true;
				break;
			}

			if (is_failed)
				break;

			// 写入新文件
			DWORD written = 0;
			BOOL bret= WriteFile(new_file_handle, cipher, cipher_len, &written, NULL);
			if (!bret)
			{
				errCode = GetLastError();
				_stprintf_s(state_msg, 4096, _T("[%d]写入文件%s失败，错误码:%d"), index + 1, new_path, errCode);
				pcwnd->SetWindowText(state_msg);

				is_failed = true;
				break;
			}

			progress += cipher_len;
			dlg->m_cWorkingProgress.SetPos(progress);
		}

		// 停止加密
		int cipher_len = 8192;
		unsigned char cipher[8192] = {0};
		errCode = EVP_EncryptFinal_ex(ctx, cipher, &cipher_len);
		if (errCode == 1)
		{
			DWORD written = 0;
			BOOL bret= WriteFile(new_file_handle, cipher, cipher_len, &written, NULL);
			if (!bret)
			{
				errCode = GetLastError();
				_stprintf_s(state_msg, 4096, _T("[%d]写入文件%s失败，错误码:%d"), index + 1, new_path, errCode);
				pcwnd->SetWindowText(state_msg);

				is_failed = true;
			}

			dlg->m_cWorkingProgress.SetPos(file_size);
		}
		else
			is_failed = true;

		if (is_failed)
			break;

		dlg->m_cWorkingProgress.SetPos(file_size);

		// 关闭文件
		CloseHandle(file_handle);
		CloseHandle(new_file_handle);

		_stprintf_s(state_msg, 4096, _T("[%d]加密文件%s完成！"), index + 1, new_path);
		pcwnd->SetWindowText(state_msg);
	}

	EVP_CIPHER_CTX_free(ctx);

	return ;
}

void CbgFileEncryptDlg::OnLvnKeydownListFiles(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	
	// 找到删除键
	if (pLVKeyDow->wVKey == VK_DELETE)
	{
		// 找到当前选中的行
		POSITION pos = m_cFileList.GetFirstSelectedItemPosition();
		if (pos == NULL)
			return ;

		do 
		{
			int index = m_cFileList.GetNextSelectedItem(pos);
			if (index != -1)
				m_cFileList.DeleteItem(index);
			else
				break;

		} while (true);
	}

	*pResult = 0;
}
