// fesInfo.cpp : 实现文件
//

#include "stdafx.h"
#include "Informant.h"
#include "fesInfo.h"
#include "afxdialogex.h"


// fesInfo 对话框

IMPLEMENT_DYNAMIC(fesInfo, CDialogEx)

fesInfo::fesInfo(CWnd* pParent /*=NULL*/)
	: CDialogEx(DD_FESINFO_DIALOG, pParent)
{
}

fesInfo::~fesInfo()
{
}

void fesInfo::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, fesList);
	DDX_Control(pDX, IDC_DATETIMEPICKER1, dateTimePicker);
}


BEGIN_MESSAGE_MAP(fesInfo, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON1, &fesInfo::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &fesInfo::OnBnClickedButton2)
	ON_BN_CLICKED(IDC_BUTTON4, &fesInfo::OnBnClickedButton4)
	ON_BN_CLICKED(IDC_BUTTON5, &fesInfo::OnBnClickedButton5)
	ON_BN_CLICKED(IDC_BUTTON3, &fesInfo::OnBnClickedButton3)
END_MESSAGE_MAP()

BOOL fesInfo::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  在此添加额外的初始化
	//fesList初始化
	dateTimePicker.SetFormat(L"yyyy/MM/dd");
	CRect rec;
	fesList.GetClientRect(&rec);
	fesList.SetExtendedStyle(fesList.GetExtendedStyle() | LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);
	fesList.InsertColumn(0, L"节日时间", LVCFMT_CENTER, rec.Width() * 5 / 32);
	fesList.InsertColumn(1, L"节日名称", LVCFMT_CENTER, rec.Width() / 4);
	fesList.InsertColumn(2, L"详细信息", LVCFMT_CENTER, rec.Width() * 11 / 32);
	fesList.InsertColumn(3, L"允许更改？", LVCFMT_CENTER, rec.Width() / 4);
	CHeaderCtrl* pHeaderCtrl = (CHeaderCtrl*)fesList.GetHeaderCtrl();
	pHeaderCtrl->EnableWindow(FALSE);

	//数据库连接
	//AfxOleInit();
	CoInitialize(NULL);
	pCon.CreateInstance(__uuidof(Connection));
	_bstr_t strCon =
		"Provider = Microsoft.Jet.OLEDB.4.0; Data Source = .\\Informant.mdb";
	//"DRIVER={Microsoft Access Driver (*.mdb)};DBQ=Informant.mdb;";
	try {
		pCon->Open(strCon, "", "", adModeUnknown);
	}
	catch (_com_error e) {
		AfxMessageBox(e.ErrorMessage() + e.Description());
	}

	//查询并取出数据到fesList
	CString sql = L"SELECT * FROM tbl_festival ORDER BY date";
	pRec.CreateInstance(__uuidof(Recordset));
	try {       //利用Open函数执行SQL命令，获得查询结果记录集         
				//pRec->Open(_variant_t(sql),
				//	pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
		pRec = pCon->Execute((_bstr_t)sql, NULL, adCmdText);
	}
	catch (_com_error e) {
		AfxMessageBox(e.ErrorMessage() + e.Description());
	}
	while (!pRec->adoEOF) {
		fesList.InsertItem(0, (CString)pRec->GetCollect(L"date"));
		fesList.SetItemText(0, 1, (CString)pRec->GetCollect(L"name_"));
		fesList.SetItemText(0, 2, (CString)pRec->GetCollect(L"remarks"));
		fesList.SetItemText(0, 3, (CString)pRec->GetCollect(L"change"));
		pRec->MoveNext();
	}

	return TRUE;  // return TRUE unless you set the focus to a control
				  // 异常: OCX 属性页应返回 FALSE
}
// fesInfo 消息处理程序

//添加新的节日提醒
void fesInfo::OnBnClickedButton1()
{
	// TODO: 在此添加控件通知处理程序代码
	CString date, name, detail, sql;

	GetDlgItem(IDC_DATETIMEPICKER1)->GetWindowText(date);
	GetDlgItem(IDC_EDIT2)->GetWindowTextW(name);
	GetDlgItem(IDC_EDIT3)->GetWindowTextW(detail);//获取节日相关信息
	if (name == L"" || detail == L"")
		AfxMessageBox(L"名称或详细不能为空！");
	else {//输入信息不为空时
		try {       //执行SQL命令      
			sql = L"SELECT * FROM tbl_festival WHERE date='" + date + L"' AND name_='" + name + L"' AND remarks='" + detail + L"' AND change='TRUE'";
			pRec.CreateInstance(__uuidof(Recordset));
			pRec->Open(_variant_t(sql),
				pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
		}
		catch (_com_error e) {
			AfxMessageBox(e.ErrorMessage() + e.Description());
			
		}
		if (pRec->adoEOF) {//如果未找到
			sql = L"INSERT INTO tbl_festival VALUES ('" + date + L"','" + name + L"','" + detail + L"', 'TRUE')";
			pRec.CreateInstance(__uuidof(Recordset));
			try {       //利用Open函数执行SQL命令      
				pRec->Open(_variant_t(sql),
					pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
				//AfxMessageBox(L"添加成功！");
				fesList.InsertItem(0, date);
				fesList.SetItemText(0, 1, name);
				fesList.SetItemText(0, 2, detail);
				fesList.SetItemText(0, 3, L"TRUE");
			}
			catch (_com_error e) {
				AfxMessageBox(e.ErrorMessage() + e.Description());
			}
		}
		else
			AfxMessageBox(L"不允许重复添加！");
	}
}

//更新节日提醒
void fesInfo::OnBnClickedButton3()
{
	// TODO: 在此添加控件通知处理程序代码
	CString date, name, detail, change, sql, date_n, name_n, detail_n;
	POSITION pos = fesList.GetFirstSelectedItemPosition();
	int nId;
	if (pos != NULL) {//当鼠标选中某行时
		nId = (int)fesList.GetNextSelectedItem(pos);
		date = fesList.GetItemText(nId, 0);
		name = fesList.GetItemText(nId, 1);
		detail = fesList.GetItemText(nId, 2);
		change = fesList.GetItemText(nId, 3);//获取选中的节日相关信息
		if (change == L"TRUE") {
			GetDlgItem(IDC_DATETIMEPICKER1)->GetWindowText(date_n);
			GetDlgItem(IDC_EDIT2)->GetWindowTextW(name_n);
			GetDlgItem(IDC_EDIT3)->GetWindowTextW(detail_n);//获取节日新的相关信息
															//sql = L"UPDATE tbl_festival  SET tbl_festival.date='" + date_n + L"' AND tbl_festival.name_='" + name_n + L"' AND tbl_festival.remarks='" + detail_n
															//	+ L"' WHERE tbl_festival.date='" + date + L"' AND tbl_festival.name_='" + name + L"' AND tbl_festival.remarks='" + detail + L"'";
															//GetDlgItem(IDC_EDIT2)->SetWindowTextW(sql);
			pRec.CreateInstance(__uuidof(Recordset));
			try {       //利用Open函数执行SQL命令    
				sql = L"DELETE FROM tbl_festival WHERE date='" + date + L"' AND name_='" + name + L"' AND remarks='" + detail + L"' AND change='TRUE'";
				pRec->Open(_variant_t(sql),
					pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
				sql = L"INSERT INTO tbl_festival VALUES ('" + date_n + L"','" + name_n + L"','" + detail_n + L"','TRUE')";
				pRec->Open(_variant_t(sql),
					pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
				//AfxMessageBox(L"更新成功！");
				fesList.SetItemText(nId, 0, date_n);
				fesList.SetItemText(nId, 1, name_n);
				fesList.SetItemText(nId, 2, detail_n);
			}
			catch (_com_error e) {
				AfxMessageBox(e.ErrorMessage() + e.Description());
			}
		}
		else
			AfxMessageBox(L"系统节日不允许更改！");
	}
	else//鼠标未选中数据
		AfxMessageBox(L"未选中节日！");
}

//清空节日提醒
void fesInfo::OnBnClickedButton4()
{
	// TODO: 在此添加控件通知处理程序代码
	CString sql = L"DELETE  FROM tbl_festival WHERE change='TRUE'";
	pRec.CreateInstance(__uuidof(Recordset));
	try {       //利用Open函数执行SQL命令      
		pRec->Open(_variant_t(sql),
			pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
		//AfxMessageBox(L"清空成功！");
		fesList.DeleteAllItems();//清空当前fesList
								 //再次初始化fesList
		sql = L"SELECT * FROM tbl_festival ORDER BY date";
		pRec.CreateInstance(__uuidof(Recordset));
		try {       //利用Open函数执行SQL命令，获得查询结果记录集         
					//pRec->Open(_variant_t(sql),
					//	pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
			pRec = pCon->Execute((_bstr_t)sql, NULL, adCmdText);
		}
		catch (_com_error e) {
			AfxMessageBox(e.ErrorMessage() + e.Description());
		}
		while (!pRec->adoEOF) {
			fesList.InsertItem(0, (CString)pRec->GetCollect(L"date"));
			fesList.SetItemText(0, 1, (CString)pRec->GetCollect(L"name_"));
			fesList.SetItemText(0, 2, (CString)pRec->GetCollect(L"remarks"));
			fesList.SetItemText(0, 3, (CString)pRec->GetCollect(L"change"));
			pRec->MoveNext();
		}

	}
	catch (_com_error e) {
		AfxMessageBox(e.ErrorMessage() + e.Description());
	}
}

//退出节日提醒并关闭数据库连接
void fesInfo::OnBnClickedButton5()
{
	// TODO: 在此添加控件通知处理程序代码
	CDialogEx::OnOK();
	CoUninitialize();
	/*if (pCon->State){
	pCon->Close();
	pRec->Close();
	}*/
}

//删除选中的节日提醒
void fesInfo::OnBnClickedButton2()
{
	// TODO: 在此添加控件通知处理程序代码
	CString date, name, detail, sql, change;
	POSITION pos = fesList.GetFirstSelectedItemPosition();
	int nId;
	if (pos != NULL) {//当鼠标选中某行时
		nId = (int)fesList.GetNextSelectedItem(pos);
		date = fesList.GetItemText(nId, 0);
		name = fesList.GetItemText(nId, 1);
		detail = fesList.GetItemText(nId, 2);
		change = fesList.GetItemText(nId, 3);//获取选中的节日相关信息
		if (change == L"TRUE") {//用户添加节日可删除
			sql = L"DELETE FROM tbl_festival WHERE date='" + date + L"' AND name_='" + name + L"' AND remarks='" + detail + L"'";
			pRec.CreateInstance(__uuidof(Recordset));
			try {       //利用Open函数执行SQL命令      
				pRec->Open(_variant_t(sql),
					pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
				//AfxMessageBox(L"删除成功！");
				fesList.DeleteItem(nId);
			}
			catch (_com_error e) {
				AfxMessageBox(e.ErrorMessage() + e.Description());
			}
		}
		else
			AfxMessageBox(L"系统节日不允许删除！");
	}
	else//鼠标未选中数据
		AfxMessageBox(L"未选中节日！");
}