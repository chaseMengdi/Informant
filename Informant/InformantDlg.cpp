
// InformantDlg.cpp : 实现文件
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


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

	// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

														// 实现
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


// CInformantDlg 对话框



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


// CInformantDlg 消息处理程序

BOOL CInformantDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	// 将“关于...”菜单项添加到系统菜单中。
	state = FALSE;

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

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	SetTimer(1, 5000, NULL);
	SetTimer(2, 300000, NULL);
	CRect rect;
	infoMain.GetClientRect(&rect);
	infoMain.SetExtendedStyle(infoMain.GetExtendedStyle() | LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);
	infoMain.InsertColumn(0, L"时间", LVCFMT_CENTER, rect.Width() * 5 / 32);
	infoMain.InsertColumn(1, L"标题", LVCFMT_CENTER, rect.Width()  / 4);
	infoMain.InsertColumn(2, L"内容", LVCFMT_CENTER, rect.Width() * 11 / 32);
	infoMain.InsertColumn(3, L"允许修改？", LVCFMT_CENTER, rect.Width()  / 4);
	CHeaderCtrl*   pHeaderCtrl = (CHeaderCtrl*)infoMain.GetHeaderCtrl();
	pHeaderCtrl->EnableWindow(FALSE);


	infoDate.GetClientRect(&rect);
	infoDate.SetExtendedStyle(infoMain.GetExtendedStyle() | LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);
	infoDate.InsertColumn(0, L"时间", LVCFMT_CENTER, rect.Width() * 5 / 32);
	infoDate.InsertColumn(1, L"标题", LVCFMT_CENTER, rect.Width()  / 4);
	infoDate.InsertColumn(2, L"内容", LVCFMT_CENTER, rect.Width() * 11 / 32);
	infoDate.InsertColumn(3, L"周期", LVCFMT_CENTER, rect.Width()  / 4);
	pHeaderCtrl = (CHeaderCtrl*)infoDate.GetHeaderCtrl();
	pHeaderCtrl->EnableWindow(FALSE);

	//数据库连接
	CoInitialize(NULL);
	pCon.CreateInstance(__uuidof(Connection));
	_bstr_t strCon =
	"Provider = Microsoft.Jet.OLEDB.4.0; Data Source = .\\Informant.mdb";
	//"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\\Informant.mdb";
	try {
	pCon->Open(strCon, "", "", adModeUnknown);
	}
	catch (_com_error e) {
	AfxMessageBox(L"初始化连接错误\n");
	AfxMessageBox(e.ErrorMessage()+e.Description());
	}

	readAll();

	//设置系统托盘
	NotifyIcon.cbSize = sizeof(NOTIFYICONDATA);
	NotifyIcon.hIcon = m_hIcon;  
	NotifyIcon.hWnd = m_hWnd;
	lstrcpy(NotifyIcon.szTip, _T("Informant日程管理"));
	NotifyIcon.uCallbackMessage = WM_SYSTEMTRAY;
	NotifyIcon.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
	Shell_NotifyIcon(NIM_ADD, &NotifyIcon);   //添加系统托盘

	
	CLoginDlg dlg;
	dlg.DoModal();

    remindDlg dlg1;
	dlg1.DoModal();

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
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

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CInformantDlg::OnPaint()
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
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CInformantDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CInformantDlg::On32773()
{
	// TODO: 在此添加命令处理程序代码
	this->ShowWindow(SW_HIDE);
	timeInfoDlg* dlg = new timeInfoDlg;
	dlg->DoModal();
	this->ShowWindow(SW_SHOW);
}


void CInformantDlg::On32774()
{
	// TODO: 在此添加命令处理程序代码
	this->ShowWindow(SW_HIDE);
	dateInfo dlg;
	dlg.DoModal();
	this->ShowWindow(SW_SHOW);
}


void CInformantDlg::On32775()
{
	// TODO: 在此添加命令处理程序代码
	this->ShowWindow(SW_HIDE);
	fesInfo dlg;
	dlg.DoModal();
	this->ShowWindow(SW_SHOW);
}


void CInformantDlg::On32777()
{
	// TODO: 在此添加命令处理程序代码
	this->ShowWindow(SW_HIDE);
	diaryDlg dlg;
	dlg.DoModal();
	this->ShowWindow(SW_SHOW);
}


void CInformantDlg::On32778()
{
	// TODO: 在此添加命令处理程序代码
	this->ShowWindow(SW_HIDE);
	noteDlg dlg;
	dlg.DoModal();
	this->ShowWindow(SW_SHOW);
}


void CAboutDlg::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码
	CDialogEx::OnOK();
}




//开机启动的设置与取消
void CInformantDlg::On32771()
{
	// TODO: 在此添加命令处理程序代码
	HKEY hKey;
	CString strRegPath = _T("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run");
	LPCTSTR szModule;
	//找到系统的启动项  
	if (state == FALSE) {
		if (RegOpenKeyEx(HKEY_CURRENT_USER, strRegPath, 0, KEY_ALL_ACCESS, &hKey) == ERROR_SUCCESS) {//打开启动项
			TCHAR szModule[MAX_PATH];
			GetModuleFileName(NULL, (LPWSTR)szModule, MAX_PATH);//得到本程序自身的全路径  
			RegSetValueEx(hKey, _T("Informant"), 0, REG_SZ, (LPBYTE)szModule, (lstrlen(szModule) + 1) * sizeof(TCHAR));
			//添加一个子Key,并设置值，"Informant"是应用程序名字（不加后缀.exe）  
			RegCloseKey(hKey); //关闭注册表  
			state = !state;
			GetMenu()->GetSubMenu(0)->CheckMenuItem(ID_32771, MF_BYCOMMAND | MF_CHECKED);
			AfxMessageBox(L"开机启动设置成功！");
		}
		else {
			AfxMessageBox(_T("系统参数错误,不能随系统启动"));
		}
	}
	else if (state == TRUE) {
		if (RegOpenKeyEx(HKEY_CURRENT_USER, strRegPath, 0, KEY_ALL_ACCESS, &hKey) == ERROR_SUCCESS)
		{
			RegDeleteValue(hKey, _T("Informant"));
			RegCloseKey(hKey);
			state = !state;
			GetMenu()->GetSubMenu(0)->CheckMenuItem(ID_32771, MF_BYCOMMAND | MF_UNCHECKED);
			AfxMessageBox(L"开机启动取消成功！");
		}
	}
}


void CInformantDlg::On32776()
{
	// TODO: 在此添加命令处理程序代码
	CString StrMusic = L"音乐文件(*.mp3)|*.mp3|所有文件(*.*)|*.*|";
	CFileDialog Dlg(TRUE, NULL, NULL, OFN_HIDEREADONLY | OFN_READONLY, StrMusic, NULL);
	if (Dlg.DoModal() == IDOK) {
		CString sql, musicPath = Dlg.GetPathName();
		pRec.CreateInstance(__uuidof(Recordset));
		try {       //利用Open函数执行SQL命令      
			sql = L"DELETE * FROM tbl_music";
			pRec->Open(_variant_t(sql),
				pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);

			sql = L"INSERT INTO tbl_music VALUES ('" + musicPath + L"')";
			pRec->Open(_variant_t(sql),
				pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);

			AfxMessageBox(L"音乐设置成功！");
		}
		catch (_com_error e) {
			AfxMessageBox(e.ErrorMessage() + e.Description());
		}
	}
}
//关于系统
void CInformantDlg::On32779()
{
	// TODO: 在此添加命令处理程序代码
	aboutDlg dlg;
	dlg.DoModal();
}
//退出
void CInformantDlg::On32772()
{
	// TODO: 在此添加命令处理程序代码
	CDialogEx::OnOK();
}
//托盘命令响应
LRESULT CInformantDlg::OnSystemtray(WPARAM wParam, LPARAM lParam)
{
	//wParam接收的是图标的ID，而lParam接收的是鼠标的行为
	//  if(wParam!=IDR_MAINFRAME)     
	//      return    1;     
	switch (lParam)
	{
	case  WM_LBUTTONDBLCLK://左键单击
	{
		DestroyWindow();
	}
	break;
	case  WM_LBUTTONDOWN://左键单击   
	{
		ModifyStyleEx(0, WS_EX_TOPMOST);   //可以改变窗口的显示风格
		ShowWindow(SW_SHOWNORMAL);
	}
	break;
	}
	return 0;
	return LRESULT();
}


//关闭窗口函数重写
void CInformantDlg::OnCancel()
{
	// TODO: 在此添加专用代码和/或调用基类
	this->ShowWindow(SW_HIDE);
	//CDialogEx::OnCancel();
}

//摧毁窗口并消除托盘
BOOL CInformantDlg::DestroyWindow()
{
	// TODO: 在此添加专用代码和/或调用基类
	Shell_NotifyIcon(NIM_DELETE, &NotifyIcon);//消除托盘图标
	return CDialogEx::DestroyWindow();
}
//所有数据取出
void CInformantDlg::readAll(){
	//查询并取出数据到List
	CString sql = L"SELECT * FROM tbl_festival ORDER BY date";
	pRec.CreateInstance(__uuidof(Recordset));
	//pRec.CreateInstance("ADODB.Recordset");
	try {       //利用Open函数执行SQL命令，获得查询结果记录集 
				//pRec->Open(_variant_t(sql),
				//	pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
		pRec = pCon->Execute((_bstr_t)sql, NULL, adCmdText);
	}
	catch (_com_error e) {
		AfxMessageBox(L"查询数据错误1\n");
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
	try {       //利用Open函数执行SQL命令，获得查询结果记录集 
		pRec->Open(_variant_t(sql),
			pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	}
	catch (_com_error e) {
		AfxMessageBox(L"查询数据错误2\n");
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
	try {       //利用Open函数执行SQL命令，获得查询结果记录集 
		pRec->Open(_variant_t(sql),
			pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	}
	catch (_com_error e) {
		AfxMessageBox(L"查询数据错误3\n");
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
	// TODO: 在此添加消息处理程序代码和/或调用默认值
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
