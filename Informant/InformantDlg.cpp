
// InformantDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "Informant.h"
#include "InformantDlg.h"
#include "afxdialogex.h"
#include "timeInfoDlg.h"
#include "dateInfo.h"
#include "fesInfo.h"
#include "diaryDlg.h"
#include "noteDlg.h"
#include "aboutDlg.h"
#include "remindDlg.h"
#include "LoginDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif
#define   WM_SYSTEMTRAY WM_USER+5


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

	// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

														// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
	ON_BN_CLICKED(IDOK, &CAboutDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// CInformantDlg �Ի���



CInformantDlg::CInformantDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_INFORMANT_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CInformantDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, infoMain);
	DDX_Control(pDX, IDC_LIST2, infoDate);
}

BEGIN_MESSAGE_MAP(CInformantDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_COMMAND(ID_32773, &CInformantDlg::On32773)
	ON_COMMAND(ID_32774, &CInformantDlg::On32774)
	ON_COMMAND(ID_32775, &CInformantDlg::On32775)
	ON_COMMAND(ID_32777, &CInformantDlg::On32777)
	ON_COMMAND(ID_32778, &CInformantDlg::On32778)
	ON_COMMAND(ID_32771, &CInformantDlg::On32771)
	ON_COMMAND(ID_32776, &CInformantDlg::On32776)
	ON_COMMAND(ID_32779, &CInformantDlg::On32779)
	ON_COMMAND(ID_32772, &CInformantDlg::On32772)
	ON_WM_INITMENUPOPUP()
	ON_MESSAGE(WM_SYSTEMTRAY, &CInformantDlg::OnSystemtray)
	ON_WM_TIMER()
END_MESSAGE_MAP()


// CInformantDlg ��Ϣ�������

BOOL CInformantDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	// ��������...���˵�����ӵ�ϵͳ�˵��С�
	state = FALSE;

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

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������
	SetTimer(1, 5000, NULL);
	SetTimer(2, 300000, NULL);
	CRect rect;
	infoMain.GetClientRect(&rect);
	infoMain.SetExtendedStyle(infoMain.GetExtendedStyle() | LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);
	infoMain.InsertColumn(0, L"ʱ��", LVCFMT_CENTER, rect.Width() * 5 / 32);
	infoMain.InsertColumn(1, L"����", LVCFMT_CENTER, rect.Width()  / 4);
	infoMain.InsertColumn(2, L"����", LVCFMT_CENTER, rect.Width() * 11 / 32);
	infoMain.InsertColumn(3, L"�����޸ģ�", LVCFMT_CENTER, rect.Width()  / 4);
	CHeaderCtrl*   pHeaderCtrl = (CHeaderCtrl*)infoMain.GetHeaderCtrl();
	pHeaderCtrl->EnableWindow(FALSE);


	infoDate.GetClientRect(&rect);
	infoDate.SetExtendedStyle(infoMain.GetExtendedStyle() | LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);
	infoDate.InsertColumn(0, L"ʱ��", LVCFMT_CENTER, rect.Width() * 5 / 32);
	infoDate.InsertColumn(1, L"����", LVCFMT_CENTER, rect.Width()  / 4);
	infoDate.InsertColumn(2, L"����", LVCFMT_CENTER, rect.Width() * 11 / 32);
	infoDate.InsertColumn(3, L"����", LVCFMT_CENTER, rect.Width()  / 4);
	pHeaderCtrl = (CHeaderCtrl*)infoDate.GetHeaderCtrl();
	pHeaderCtrl->EnableWindow(FALSE);

	//���ݿ�����
	CoInitialize(NULL);
	pCon.CreateInstance(__uuidof(Connection));
	_bstr_t strCon =
	"Provider = Microsoft.Jet.OLEDB.4.0; Data Source = .\\Informant.mdb";
	//"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\\Informant.mdb";
	try {
	pCon->Open(strCon, "", "", adModeUnknown);
	}
	catch (_com_error e) {
	AfxMessageBox(L"��ʼ�����Ӵ���\n");
	AfxMessageBox(e.ErrorMessage()+e.Description());
	}

	readAll();

	//����ϵͳ����
	NotifyIcon.cbSize = sizeof(NOTIFYICONDATA);
	NotifyIcon.hIcon = m_hIcon;  
	NotifyIcon.hWnd = m_hWnd;
	lstrcpy(NotifyIcon.szTip, _T("Informant�ճ̹���"));
	NotifyIcon.uCallbackMessage = WM_SYSTEMTRAY;
	NotifyIcon.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
	Shell_NotifyIcon(NIM_ADD, &NotifyIcon);   //���ϵͳ����

	
	CLoginDlg dlg;
	dlg.DoModal();

    remindDlg dlg1;
	dlg1.DoModal();

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CInformantDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CInformantDlg::OnPaint()
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
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CInformantDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CInformantDlg::On32773()
{
	// TODO: �ڴ���������������
	this->ShowWindow(SW_HIDE);
	timeInfoDlg* dlg = new timeInfoDlg;
	dlg->DoModal();
	this->ShowWindow(SW_SHOW);
}


void CInformantDlg::On32774()
{
	// TODO: �ڴ���������������
	this->ShowWindow(SW_HIDE);
	dateInfo dlg;
	dlg.DoModal();
	this->ShowWindow(SW_SHOW);
}


void CInformantDlg::On32775()
{
	// TODO: �ڴ���������������
	this->ShowWindow(SW_HIDE);
	fesInfo dlg;
	dlg.DoModal();
	this->ShowWindow(SW_SHOW);
}


void CInformantDlg::On32777()
{
	// TODO: �ڴ���������������
	this->ShowWindow(SW_HIDE);
	diaryDlg dlg;
	dlg.DoModal();
	this->ShowWindow(SW_SHOW);
}


void CInformantDlg::On32778()
{
	// TODO: �ڴ���������������
	this->ShowWindow(SW_HIDE);
	noteDlg dlg;
	dlg.DoModal();
	this->ShowWindow(SW_SHOW);
}


void CAboutDlg::OnBnClickedOk()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CDialogEx::OnOK();
}




//����������������ȡ��
void CInformantDlg::On32771()
{
	// TODO: �ڴ���������������
	HKEY hKey;
	CString strRegPath = _T("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run");
	LPCTSTR szModule;
	//�ҵ�ϵͳ��������  
	if (state == FALSE) {
		if (RegOpenKeyEx(HKEY_CURRENT_USER, strRegPath, 0, KEY_ALL_ACCESS, &hKey) == ERROR_SUCCESS) {//��������
			TCHAR szModule[MAX_PATH];
			GetModuleFileName(NULL, (LPWSTR)szModule, MAX_PATH);//�õ������������ȫ·��  
			RegSetValueEx(hKey, _T("Informant"), 0, REG_SZ, (LPBYTE)szModule, (lstrlen(szModule) + 1) * sizeof(TCHAR));
			//���һ����Key,������ֵ��"Informant"��Ӧ�ó������֣����Ӻ�׺.exe��  
			RegCloseKey(hKey); //�ر�ע���  
			state = !state;
			GetMenu()->GetSubMenu(0)->CheckMenuItem(ID_32771, MF_BYCOMMAND | MF_CHECKED);
			AfxMessageBox(L"�����������óɹ���");
		}
		else {
			AfxMessageBox(_T("ϵͳ��������,������ϵͳ����"));
		}
	}
	else if (state == TRUE) {
		if (RegOpenKeyEx(HKEY_CURRENT_USER, strRegPath, 0, KEY_ALL_ACCESS, &hKey) == ERROR_SUCCESS)
		{
			RegDeleteValue(hKey, _T("Informant"));
			RegCloseKey(hKey);
			state = !state;
			GetMenu()->GetSubMenu(0)->CheckMenuItem(ID_32771, MF_BYCOMMAND | MF_UNCHECKED);
			AfxMessageBox(L"��������ȡ���ɹ���");
		}
	}
}


void CInformantDlg::On32776()
{
	// TODO: �ڴ���������������
	CString StrMusic = L"�����ļ�(*.mp3)|*.mp3|�����ļ�(*.*)|*.*|";
	CFileDialog Dlg(TRUE, NULL, NULL, OFN_HIDEREADONLY | OFN_READONLY, StrMusic, NULL);
	if (Dlg.DoModal() == IDOK) {
		CString sql, musicPath = Dlg.GetPathName();
		pRec.CreateInstance(__uuidof(Recordset));
		try {       //����Open����ִ��SQL����      
			sql = L"DELETE * FROM tbl_music";
			pRec->Open(_variant_t(sql),
				pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);

			sql = L"INSERT INTO tbl_music VALUES ('" + musicPath + L"')";
			pRec->Open(_variant_t(sql),
				pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);

			AfxMessageBox(L"�������óɹ���");
		}
		catch (_com_error e) {
			AfxMessageBox(e.ErrorMessage() + e.Description());
		}
	}
}
//����ϵͳ
void CInformantDlg::On32779()
{
	// TODO: �ڴ���������������
	aboutDlg dlg;
	dlg.DoModal();
}
//�˳�
void CInformantDlg::On32772()
{
	// TODO: �ڴ���������������
	CDialogEx::OnOK();
}
//����������Ӧ
LRESULT CInformantDlg::OnSystemtray(WPARAM wParam, LPARAM lParam)
{
	//wParam���յ���ͼ���ID����lParam���յ���������Ϊ
	//  if(wParam!=IDR_MAINFRAME)     
	//      return    1;     
	switch (lParam)
	{
	case  WM_LBUTTONDBLCLK://�������
	{
		DestroyWindow();
	}
	break;
	case  WM_LBUTTONDOWN://�������   
	{
		ModifyStyleEx(0, WS_EX_TOPMOST);   //���Ըı䴰�ڵ���ʾ���
		ShowWindow(SW_SHOWNORMAL);
	}
	break;
	}
	return 0;
	return LRESULT();
}


//�رմ��ں�����д
void CInformantDlg::OnCancel()
{
	// TODO: �ڴ����ר�ô����/����û���
	this->ShowWindow(SW_HIDE);
	//CDialogEx::OnCancel();
}

//�ݻٴ��ڲ���������
BOOL CInformantDlg::DestroyWindow()
{
	// TODO: �ڴ����ר�ô����/����û���
	Shell_NotifyIcon(NIM_DELETE, &NotifyIcon);//��������ͼ��
	return CDialogEx::DestroyWindow();
}
//��������ȡ��
void CInformantDlg::readAll(){
	//��ѯ��ȡ�����ݵ�List
	CString sql = L"SELECT * FROM tbl_festival ORDER BY date";
	pRec.CreateInstance(__uuidof(Recordset));
	//pRec.CreateInstance("ADODB.Recordset");
	try {       //����Open����ִ��SQL�����ò�ѯ�����¼�� 
				//pRec->Open(_variant_t(sql),
				//	pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
		pRec = pCon->Execute((_bstr_t)sql, NULL, adCmdText);
	}
	catch (_com_error e) {
		AfxMessageBox(L"��ѯ���ݴ���1\n");
		AfxMessageBox(e.ErrorMessage() + e.Description());
	}
	while (!pRec->adoEOF) {
		infoMain.InsertItem(0, (CString)pRec->GetCollect(L"date"));
		infoMain.SetItemText(0, 1, (CString)pRec->GetCollect(L"name_"));
		infoMain.SetItemText(0, 2, (CString)pRec->GetCollect(L"remarks"));
		infoMain.SetItemText(0, 3, (CString)pRec->GetCollect(L"change"));

		pRec->MoveNext();
	}

	sql = L"SELECT * FROM tbl_time ORDER BY date_";
	pRec.CreateInstance(__uuidof(Recordset));
	try {       //����Open����ִ��SQL�����ò�ѯ�����¼�� 
		pRec->Open(_variant_t(sql),
			pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	}
	catch (_com_error e) {
		AfxMessageBox(L"��ѯ���ݴ���2\n");
		AfxMessageBox(e.ErrorMessage() + e.Description());
	}
	while (!pRec->adoEOF) {
		infoMain.InsertItem(0, (CString)pRec->GetCollect(L"date_"));
		infoMain.SetItemText(0, 1, (CString)pRec->GetCollect(L"title_"));
		infoMain.SetItemText(0, 2, (CString)pRec->GetCollect(L"detail_"));
		pRec->MoveNext();
	}

	sql = L"SELECT * FROM tbl_dateInfo ORDER BY begintime";
	pRec.CreateInstance(__uuidof(Recordset));
	try {       //����Open����ִ��SQL�����ò�ѯ�����¼�� 
		pRec->Open(_variant_t(sql),
			pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	}
	catch (_com_error e) {
		AfxMessageBox(L"��ѯ���ݴ���3\n");
		AfxMessageBox(e.ErrorMessage() + e.Description());
	}
	while (!pRec->adoEOF) {
		infoDate.InsertItem(0, (CString)pRec->GetCollect(L"begintime"));
		infoDate.SetItemText(0, 1, (CString)pRec->GetCollect(L"title"));
		infoDate.SetItemText(0, 2, (CString)pRec->GetCollect(L"detail"));
		infoDate.SetItemText(0, 3, (CString)pRec->GetCollect(L"period"));
		pRec->MoveNext();
	}
}


void CInformantDlg::OnTimer(UINT_PTR nIDEvent)
{
	// TODO: �ڴ������Ϣ�����������/�����Ĭ��ֵ
	switch (nIDEvent){
	default:
		break;
	case 1:
		infoMain.DeleteAllItems();
		infoDate.DeleteAllItems();
		readAll();
		break;
	case 2:
		remindDlg dlg;
		dlg.DoModal();
		break;
	}
	CDialogEx::OnTimer(nIDEvent);
}
