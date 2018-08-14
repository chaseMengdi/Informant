// remindDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "Informant.h"
#include "remindDlg.h"
#include "afxdialogex.h"


// remindDlg 对话框

IMPLEMENT_DYNAMIC(remindDlg, CDialogEx)

remindDlg::remindDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_REMIND_DIALOG, pParent)
{

}

remindDlg::~remindDlg()
{
}

void remindDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_OCX1, myPlayer);
	DDX_Control(pDX, IDC_LIST1, infoTime);
	DDX_Control(pDX, IDC_LIST3, infoDate);
}


BEGIN_MESSAGE_MAP(remindDlg, CDialogEx)
END_MESSAGE_MAP()


// remindDlg 消息处理程序


BOOL remindDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  在此添加额外的初始化
	CRect rect;
	infoTime.GetClientRect(&rect);
	infoTime.SetExtendedStyle(infoTime.GetExtendedStyle() | LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);
	infoTime.InsertColumn(0, L"标题", LVCFMT_CENTER, rect.Width()/2);
	infoTime.InsertColumn(1, L"内容", LVCFMT_CENTER, rect.Width()/2);
	CHeaderCtrl* pHeaderCtrl = (CHeaderCtrl*)infoTime.GetHeaderCtrl();
	pHeaderCtrl->EnableWindow(FALSE);

	infoDate.GetClientRect(&rect);
	infoDate.SetExtendedStyle(infoDate.GetExtendedStyle() | LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);
	infoDate.InsertColumn(0, L"标题", LVCFMT_CENTER, rect.Width()/3);
	infoDate.InsertColumn(1, L"内容", LVCFMT_CENTER, rect.Width() / 3);
	infoDate.InsertColumn(2, L"周期", LVCFMT_CENTER, rect.Width()/3);
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
		AfxMessageBox(e.ErrorMessage() + e.Description());
	}
	//获取系统时间
	CTime now = CTime::GetCurrentTime();
	int nowY = now.GetYear(), nowM = now.GetMonth(), nowD = now.GetDay();
	COleDateTime now_1(nowY, nowM, nowD,0,0,0);
	CString nowS = now.Format("%Y/%m/%d");
	SetDlgItemText(IDC_STATIC1, L"今天是" + nowS + L"。");
	CString sql = L"SELECT * FROM tbl_time WHERE date_='"+nowS+L"'";
	pRec.CreateInstance(__uuidof(Recordset));
	try {       //利用Open函数执行SQL命令，获得查询结果记录集 
		pRec->Open(_variant_t(sql),
			pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	}
	catch (_com_error e) {
		AfxMessageBox(L"查询数据错误1\n");
		AfxMessageBox(e.ErrorMessage() + e.Description());
	}
	while (!pRec->adoEOF) {
		infoTime.InsertItem(0, (CString)pRec->GetCollect(L"title_"));
		infoTime.SetItemText(0, 1, (CString)pRec->GetCollect(L"detail_"));
		pRec->MoveNext();
	}
	//定期提醒数据获取
	BOOL state = FALSE;
	CString date_, year_, month_, day_, cycle_;
	int year, month, day,sub;
	sql = L"SELECT * FROM tbl_dateInfo";
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
		date_=(CString)pRec->GetCollect(L"begintime");
		cycle_=(CString)pRec->GetCollect(L"period");//获取数据库中的日期和周期
		year_ = date_.Left(4);
		month_ = date_.Right(5);
		month_ = month_.Left(2);
		day_ = date_.Right(2);
		year=_ttoi(year_);
		month=_ttoi(month_);
		day=_ttoi(day_);
		COleDateTime time_1(year,month,day,0,0,0);
		sub = abs((int)(time_1-now_1).m_span);
		if (cycle_ == L"每天")//判定是否到达周期，此处做了简化处理，
			state = TRUE;        //未考虑平年闰年、不同月份天数问题
		if (cycle_ == L"每周" && (sub % 7 == 0))
			state = TRUE;
		if (cycle_ == L"每月" && nowD==day && nowM!=month)
			state = TRUE;
		if (cycle_ == L"每年"&& nowD==day&&nowM==month&&nowY!=year)
			state = TRUE;
		if (state) {//如果到达周期则提醒
			infoDate.InsertItem(0,(CString)pRec->GetCollect(L"title"));
			infoDate.SetItemText(0, 1, (CString)pRec->GetCollect(L"detail"));
			infoDate.SetItemText(0, 2, (CString)pRec->GetCollect(L"period"));
		}
		pRec->MoveNext();

		//提醒音乐数据获取
		sql = L"SELECT * FROM tbl_music";
		pRec.CreateInstance(__uuidof(Recordset));
		try {       //利用Open函数执行SQL命令，获得查询结果记录集 
			pRec->Open(_variant_t(sql),
				pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
		}
		catch (_com_error e) {
			AfxMessageBox(L"查询音乐错误\n");
			AfxMessageBox(e.ErrorMessage() + e.Description());
		}
		while (!pRec->adoEOF) {
			myPlayer.put_URL((CString)pRec->GetCollect(L"music"));
			pRec->MoveNext();
		}
	}

	//播放音乐



	return TRUE;  // return TRUE unless you set the focus to a control
				  // 异常: OCX 属性页应返回 FALSE
}
