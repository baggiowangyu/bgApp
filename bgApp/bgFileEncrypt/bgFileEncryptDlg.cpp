
// bgFileEncryptDlg.cpp : ʵ���ļ�
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

//// ������Կ�����ڼ��ܿ���
//const char *internal_password = "12345678qwertyui";


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


// CbgFileEncryptDlg �Ի���



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


// CbgFileEncryptDlg ��Ϣ�������

BOOL CbgFileEncryptDlg::OnInitDialog()
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CbgFileEncryptDlg::OnPaint()
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
HCURSOR CbgFileEncryptDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CbgFileEncryptDlg::OnDropFiles(HDROP hDropInfo)
{
	// ����߳��ڹ����У��򲻴����ļ��϶�
	if (!encrypt_thread_.isRunning())
	{
		//ȡ�ñ��϶��ļ�����Ŀ
		int total_drop_count = DragQueryFile(hDropInfo, -1, NULL, 0);

		for(int index = 0; index < total_drop_count; ++index)
		{
			//�����ҷ�ĵ�i���ļ����ļ���
			TCHAR wcStr[MAX_PATH] = {0};
			DragQueryFile(hDropInfo, index, wcStr, MAX_PATH);

			int count = m_cFileList.GetItemCount();
			m_cFileList.InsertItem(count, wcStr);
		} 

		//�ϷŽ�����,�ͷ��ڴ�
		DragFinish(hDropInfo);
	}

	CDialog::OnDropFiles(hDropInfo);
}

void CbgFileEncryptDlg::OnBnClickedBtnStart()
{
	// �ȼ������߳��Ƿ��ڹ����У�����ڹ�������ô�͵����澯
	if (!encrypt_thread_.isRunning())
	{
		// ��ȡ������Կ
		m_cPassword.GetWindowText(password);

		// �����߳�
		encrypt_thread_.start(CbgFileEncryptDlg::EncryptThread, this);
	}
}


void CbgFileEncryptDlg::EncryptThread(void* param)
{
	CbgFileEncryptDlg *dlg = (CbgFileEncryptDlg *)param;

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
		unsigned char key[17] = {0};
		memset(key, 0, 17);
		memcpy_s(key, 16, T2A(dlg->password.GetBuffer(0)), dlg->password.GetLength());
		errCode = EVP_EncryptInit_ex(ctx, EVP_aes_128_cbc(), NULL, key, iv);

		TCHAR path[4096] = {0};
		dlg->m_cFileList.GetItemText(index, 0, path, 4096);

		TCHAR state_msg[4096] = {0};
		_stprintf_s(state_msg, 4096, _T("[%d]���ڼ���%s..."), index + 1, path);
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
		_stprintf_s(new_path, 4096, _T("%s.cipher"), path);
		HANDLE new_file_handle = CreateFileW(new_path, GENERIC_ALL, FILE_SHARE_READ|FILE_SHARE_WRITE|FILE_SHARE_DELETE, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
		if (file_handle == INVALID_HANDLE_VALUE)
		{
			errCode = GetLastError();
			_stprintf_s(state_msg, 4096, _T("[%d]���ļ�%sʧ�ܣ�������:%d"), index + 1, new_path, errCode);
			pcwnd->SetWindowText(state_msg);

			CloseHandle(file_handle);
			break ;
		}

		// �Ƚ�����Base64��Ȼ���¼Base64�ĳ��ȣ�Ȼ����д��4�ֽڳ��ȣ���д��Base64
		std::ostringstream str;
		Poco::Base64Encoder encoder(str);
		encoder << key;
		
		std::string base64encode_key = str.str();
		encoder.close();
		DWORD base64_len = base64encode_key.size();

		DWORD written = 0;
		WriteFile(new_file_handle, &base64_len, sizeof(DWORD), &written, NULL);
		WriteFile(new_file_handle, base64encode_key.c_str(), base64_len, &written, NULL);

		// ��ȡ�ļ���С
		DWORD file_size = GetFileSize(file_handle, NULL);
		dlg->m_cWorkingProgress.SetRange32(0, file_size);

		// ��ȡ�ļ�
		int progress = 0;
		dlg->m_cWorkingProgress.SetPos(progress);

		bool is_failed = false;
		DWORD readed = 0;
		unsigned char data_buffer[4096] = {0};
		while (ReadFile(file_handle, data_buffer, 4096, &readed, NULL))
		{
			if (readed == 0)
				break;
			
			// ���¼���
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

			// д�����ļ�
			DWORD written = 0;
			BOOL bret= WriteFile(new_file_handle, cipher, cipher_len, &written, NULL);
			if (!bret)
			{
				errCode = GetLastError();
				_stprintf_s(state_msg, 4096, _T("[%d]д���ļ�%sʧ�ܣ�������:%d"), index + 1, new_path, errCode);
				pcwnd->SetWindowText(state_msg);

				is_failed = true;
				break;
			}

			progress += cipher_len;
			dlg->m_cWorkingProgress.SetPos(progress);
		}

		// ֹͣ����
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
				_stprintf_s(state_msg, 4096, _T("[%d]д���ļ�%sʧ�ܣ�������:%d"), index + 1, new_path, errCode);
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

		// �ر��ļ�
		CloseHandle(file_handle);
		CloseHandle(new_file_handle);

		_stprintf_s(state_msg, 4096, _T("[%d]�����ļ�%s��ɣ�"), index + 1, new_path);
		pcwnd->SetWindowText(state_msg);
	}

	EVP_CIPHER_CTX_free(ctx);

	return ;
}

void CbgFileEncryptDlg::OnLvnKeydownListFiles(NMHDR *pNMHDR, LRESULT *pResult)
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
