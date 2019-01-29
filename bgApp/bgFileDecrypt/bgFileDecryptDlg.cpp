
// bgFileDecryptDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "bgFileDecrypt.h"
#include "bgFileDecryptDlg.h"

#include "openssl/bio.h"
#include "openssl/ssl.h"
#include "openssl/err.h"
#include "openssl/evp.h"
#include "openssl/ossl_typ.h"
#include "openssl/evp_locl.h"

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


// CbgFileDecryptDlg �Ի���


#ifndef WM_COPYGLOBALDATA
#define WM_COPYGLOBALDATA 0x0049
#endif

typedef WINUSERAPI BOOL WINAPI CHANGEWINDOWMESSAGEFILTER(UINT message, DWORD dwFlag);


//�޸�OnDropFile��win7����ԱȨ���½��ղ���������
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

CbgFileDecryptDlg::CbgFileDecryptDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CbgFileDecryptDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CbgFileDecryptDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EDIT_PASSWORD, m_cPassword);
	DDX_Control(pDX, IDC_LIST_FILE, m_cFileList);
	DDX_Control(pDX, IDC_PROGRESS1, m_cWorkingProgress);
}

BEGIN_MESSAGE_MAP(CbgFileDecryptDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_WM_DROPFILES()
	ON_BN_CLICKED(IDC_BTN_START, &CbgFileDecryptDlg::OnBnClickedBtnStart)
	ON_NOTIFY(LVN_KEYDOWN, IDC_LIST_FILE, &CbgFileDecryptDlg::OnLvnKeydownListFile)
END_MESSAGE_MAP()


// CbgFileDecryptDlg ��Ϣ�������

BOOL CbgFileDecryptDlg::OnInitDialog()
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
	SetWindowText(_T("������԰�� ���ܹ��� V1.0"));
	OpenSSL_add_all_algorithms();

	DropFileFix();

	m_cFileList.InsertColumn(0, _T("�ļ���"), LVCFMT_LEFT, 450);
	m_cPassword.LimitText(16);

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CbgFileDecryptDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void CbgFileDecryptDlg::OnPaint()
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
HCURSOR CbgFileDecryptDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CbgFileDecryptDlg::OnDropFiles(HDROP hDropInfo)
{
	// TODO: �ڴ������Ϣ�����������/�����Ĭ��ֵ
	// ����߳��ڹ����У��򲻴����ļ��϶�
	if (!decrypt_thread_.isRunning())
	{
		m_cFileList.DeleteAllItems();

		// ȡ�ñ��϶��ļ�����Ŀ
		int total_drop_count = DragQueryFile(hDropInfo, -1, NULL, 0);

		for(int index = 0; index < total_drop_count; ++index)
		{
			// �����ҷ�ĵ�i���ļ����ļ���
			TCHAR wcStr[MAX_PATH] = {0};
			DragQueryFile(hDropInfo, index, wcStr, MAX_PATH);

			int count = m_cFileList.GetItemCount();
			m_cFileList.InsertItem(count, wcStr);
		} 

		// �ϷŽ�����,�ͷ��ڴ�
		DragFinish(hDropInfo);
	}

	CDialog::OnDropFiles(hDropInfo);
}

void CbgFileDecryptDlg::OnBnClickedBtnStart()
{
	// �ȼ������߳��Ƿ��ڹ����У�����ڹ�������ô�͵����澯
	if (!decrypt_thread_.isRunning())
	{
		// ��ȡ������Կ
		m_cPassword.GetWindowText(password);

		// �����߳�
		decrypt_thread_.start(CbgFileDecryptDlg::DecryptThread, this);
	}
}

void CbgFileDecryptDlg::DecryptThread(void* param)
{
	CbgFileDecryptDlg *dlg = (CbgFileDecryptDlg *)param;

	int errCode = 0;
	CWnd *pcwnd = dlg->GetDlgItem(IDC_STATIC_STATE);

	// �����ﴦ�����ּ��ܵ���ز���
	EVP_CIPHER_CTX *ctx = EVP_CIPHER_CTX_new();

	// ������Կ
	// ��������
	unsigned char iv[17] = "\xc2\x86\x69\x6d\x88\x7c\x9a\xa0\x61\x1b\xbb\x3e\x20\x25\xa4\x5a";

	// ������Կ
	int count = dlg->m_cFileList.GetItemCount();
	for(int index = 0; index < count; ++index)
	{
		USES_CONVERSION;
		unsigned char key[16] = {0};
		memset(key, 0, 16);
		memcpy_s(key, 16, T2A(dlg->password.GetBuffer(0)), dlg->password.GetLength());
		errCode = EVP_DecryptInit_ex(ctx, EVP_aes_128_cbc(), NULL, key, iv);

		TCHAR path[4096] = {0};
		dlg->m_cFileList.GetItemText(index, 0, path, 4096);

		TCHAR state_msg[4096] = {0};
		_stprintf_s(state_msg, 4096, _T("[%d]���ڽ���%s..."), index + 1, path);
		pcwnd->SetWindowText(state_msg);

		// ��Դ�ļ�
		HANDLE file_handle = CreateFileW(path, GENERIC_ALL, FILE_SHARE_READ|FILE_SHARE_WRITE|FILE_SHARE_DELETE, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
		if (file_handle == INVALID_HANDLE_VALUE)
		{
			errCode = GetLastError();
			_stprintf_s(state_msg, 4096, _T("[%d]���ļ�%sʧ�ܣ�������:%d"), index + 1, path, errCode);
			pcwnd->SetWindowText(state_msg);
			break ;
		}

		// �������ļ�
		TCHAR new_path[4096] = {0};
		_stprintf_s(new_path, 4096, _T("%s.plain"), path);
		HANDLE new_file_handle = CreateFileW(new_path, GENERIC_ALL, FILE_SHARE_READ|FILE_SHARE_WRITE|FILE_SHARE_DELETE, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
		if (file_handle == INVALID_HANDLE_VALUE)
		{
			errCode = GetLastError();
			_stprintf_s(state_msg, 4096, _T("[%d]���ļ�%sʧ�ܣ�������:%d"), index + 1, new_path, errCode);
			pcwnd->SetWindowText(state_msg);

			CloseHandle(file_handle);
			break ;
		}

		

		// ��ȡ�ļ���С
		DWORD file_size = GetFileSize(file_handle, NULL);
		dlg->m_cWorkingProgress.SetRange32(0, file_size);

		// �ȶ�ȡ���ȣ�����ת
		DWORD readed = 0;
		DWORD base64_len = 0;
		ReadFile(file_handle, &base64_len, sizeof(DWORD), &readed, NULL);

		SetFilePointer(file_handle, sizeof(DWORD) + base64_len, NULL, FILE_BEGIN);

		// ��ȡ�ļ�
		int progress = 0;
		dlg->m_cWorkingProgress.SetPos(progress);

		bool is_failed = false;
		//DWORD readed = 0;
		unsigned char data_buffer[16384] = {0};
		while (ReadFile(file_handle, data_buffer, 16384, &readed, NULL))
		{
			if (readed == 0)
				break;

			// ���¼���
			int plain_len = 32768;
			unsigned char plain[32768] = {0};
			errCode = EVP_DecryptUpdate(ctx, plain, &plain_len, data_buffer, readed);
			if (errCode != 1)
			{
				is_failed = true;
				break;
			}

			if (is_failed)
				break;

			// д�����ļ�
			DWORD written = 0;
			BOOL bret= WriteFile(new_file_handle, plain, plain_len, &written, NULL);
			if (!bret)
			{
				errCode = GetLastError();
				_stprintf_s(state_msg, 4096, _T("[%d]д���ļ�%sʧ�ܣ�������:%d"), index + 1, new_path, errCode);
				pcwnd->SetWindowText(state_msg);

				is_failed = true;
				break;
			}

			progress += plain_len;
			dlg->m_cWorkingProgress.SetPos(progress);
		}

		// ֹͣ����
		int plain_len = 4096;
		unsigned char plain[4096] = {0};
		errCode = EVP_DecryptFinal_ex(ctx, plain, &plain_len);
		if (errCode == 1)
		{
			DWORD written = 0;
			BOOL bret= WriteFile(new_file_handle, plain, plain_len, &written, NULL);
			if (!bret)
			{
				errCode = GetLastError();
				_stprintf_s(state_msg, 4096, _T("[%d]д���ļ�%sʧ�ܣ�������:%d"), index + 1, new_path, errCode);
				pcwnd->SetWindowText(state_msg);

				is_failed = true;
			}

			progress += plain_len;
			dlg->m_cWorkingProgress.SetPos(progress);
		}
		else
			is_failed = true;

		if (is_failed)
			break;

		_stprintf_s(state_msg, 4096, _T("[%d]�����ļ�%s��ɣ�"), index + 1, new_path);
		pcwnd->SetWindowText(state_msg);

		// �ر��ļ�
		CloseHandle(file_handle);
		CloseHandle(new_file_handle);
	}

	EVP_CIPHER_CTX_free(ctx);
}

void CbgFileDecryptDlg::OnLvnKeydownListFile(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);

	// �ҵ�ɾ����
	if (pLVKeyDow->wVKey == VK_DELETE)
	{
		// �ҵ���ǰѡ�е���
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
