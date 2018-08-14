// dateInfo.cpp : 实现文件
//

#include "stdafx.h"
#include "Informant.h"
#include "dateInfo.h"
#include "afxdialogex.h"
#include "InformantDlg.h"




// dateInfo 对话框

IMPLEMENT_DYNAMIC(dateInfo, CDialogEx)

dateInfo::dateInfo(CWnd* pParent /*=NULL*/)
	: CDialogEx(DD_DATEINFO_DIALOG, pParent)
{

	EnableAutomation();
	DA.OnInitADOConn();
	m_pRecordset = DA.m_pRecordset;
	m_pConnection = DA.m_pConnection;
	
}

dateInfo::~dateInfo()
{
	DA.ExitConnect();

	
}

void dateInfo::OnFinalRelease()
{
	// 释放了对自动化对象的最后一个引用后，将调用
	// OnFinalRelease。  基类将自动
	// 删除该对象。  在调用该基类之前，请添加您的
	// 对象所需的附加清理代码。

	CDialogEx::OnFinalRelease();
}

void dateInfo::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, m_list);
	DDX_Control(pDX, IDC_DATETIMEPICKER1, m_dateCtrl);
}


BEGIN_MESSAGE_MAP(dateInfo, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON1, &dateInfo::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &dateInfo::OnBnClickedButton2)
	ON_BN_CLICKED(IDC_BUTTON3, &dateInfo::OnBnClickedButton3)
	ON_BN_CLICKED(IDC_BUTTON4, &dateInfo::OnBnClickedButton4)
	ON_BN_CLICKED(IDC_BUTTON5, &dateInfo::OnBnClickedButton5)

	ON_NOTIFY(NM_CLICK, IDC_LIST1, &dateInfo::OnNMClickList1)
	ON_WM_TIMER()
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(dateInfo, CDialogEx)
END_DISPATCH_MAP()

// 注意: 我们添加 IID_IdateInfo 支持
//  以支持来自 VBA 的类型安全绑定。  此 IID 必须同附加到 .IDL 文件中的
//  调度接口的 GUID 匹配。

// {A4FF1B1E-608A-4A40-A25A-A1B5A164C827}
static const IID IID_IdateInfo =
{ 0xA4FF1B1E, 0x608A, 0x4A40, { 0xA2, 0x5A, 0xA1, 0xB5, 0xA1, 0x64, 0xC8, 0x27 } };

BEGIN_INTERFACE_MAP(dateInfo, CDialogEx)
	INTERFACE_PART(dateInfo, IID_IdateInfo, Dispatch)
END_INTERFACE_MAP()


// dateInfo 消息处理程序

//添加
void dateInfo::OnBnClickedButton1()
{
	// TODO: 在此添加控件通知处理程序代码
	CString str;
	GetDlgItem(IDC_EDIT2)->GetWindowText(str);
	CString str1;
	GetDlgItem(IDC_EDIT3)->GetWindowText(str1);
	if (str == _T("") || str1 == _T(""))
		AfxMessageBox(_T("标题和内容不能为空"));
	else
	{
		if (Find())AfxMessageBox(_T("不要重复添加"));
		else{
			_bstr_t sql;
			sql = "select * from tbl_dateInfo";
			m_pRecordset.CreateInstance(__uuidof(Recordset));
			m_pRecordset->Open(sql, m_pConnection.GetInterfacePtr(),
				adOpenDynamic, adLockOptimistic, adCmdText);
			m_pRecordset->AddNew();

			//m_pRecordset->PutCollect("username", _T("132"));

			m_pRecordset->PutCollect("title", (_bstr_t)str);

			m_pRecordset->PutCollect("detail", (_bstr_t)str1);

			CString str3;
			CTime m_date;
			m_dateCtrl.GetTime(m_date);
			m_pRecordset->PutCollect("begintime", (_bstr_t)m_date.Format("%Y/%m/%d"));
			CString text;
			GetDlgItemText(IDC_COMBO1, text);
			m_pRecordset->PutCollect("period", (_bstr_t)text);
			m_pRecordset->Update();
			readAll();
		}
	}
}

//删除
void dateInfo::OnBnClickedButton2()
{
	// TODO:  在此添加控件通知处理程序代码
	//CString sql;
	/*CString str;
	GetDlgItem(IDC_EDIT2)->GetWindowText(str);*/
	CString date, name, detail, sql, period;
	POSITION pos = m_list.GetFirstSelectedItemPosition();
	int nId;
	if (pos != NULL) 
	{//当鼠标选中某行时
		nId = (int)m_list.GetNextSelectedItem(pos);
		date = m_list.GetItemText(nId, 0);
		name = m_list.GetItemText(nId, 1);
		detail = m_list.GetItemText(nId, 2);//获取选中的节日相关信息
		period = m_list.GetItemText(nId, 3);
		sql.Format(_T("DELETE from tbl_dateInfo where begintime='%s' and title='%s'and detail='%s'and period='%s'"), date, name, detail,period);
		m_pConnection->Execute((_bstr_t)sql, NULL, adCmdText);
		//AfxMessageBox(L"删除成功！");
		readAll();
			
		}
	else//鼠标未选中数据
		AfxMessageBox(L"未选中信息！");
	
}

//修改
void dateInfo::OnBnClickedButton3()
{
	// TODO:  在此添加控件通知处理程序代码
	CString date, name, detail,period;
	POSITION pos = m_list.GetFirstSelectedItemPosition();
	int nId;
	if (pos != NULL) 
	{//当鼠标选中某行时
		nId = (int)m_list.GetNextSelectedItem(pos);
		date = m_list.GetItemText(nId, 0);
		name = m_list.GetItemText(nId, 1);
		detail = m_list.GetItemText(nId, 2);
		period = m_list.GetItemText(nId, 3);
		CString str1;
		GetDlgItem(IDC_EDIT3)->GetWindowText(str1);
		CString str2;
		GetDlgItem(IDC_EDIT2)->GetWindowText(str2);
		CTime m_date;
		m_dateCtrl.GetTime(m_date);
		CString text;
		GetDlgItemText(IDC_COMBO1, text);
		CString sql;
		sql.Format(_T("UPDATE tbl_dateInfo set begintime='%s',title='%s',detail='%s',period='%s' where begintime='%s' and title='%s'and detail='%s'and period='%s'"), //,title='%s',detail='%s',period='%s'
			m_date.Format("%Y/%m/%d"), str2, str1, text, date, name, detail, period);//, name, detail, period
		m_pConnection->Execute((_bstr_t)sql, NULL, adCmdText);
		m_list.DeleteItem(nId);
		readAll();
		//AfxMessageBox(L"修改成功！");
	}
	else//鼠标未选中数据
		AfxMessageBox(_T("未选中信息！"));
	
}

//清空
void dateInfo::OnBnClickedButton4()
{
	// TODO:  在此添加控件通知处理程序代码
	CString sql;
	sql.Format(_T("DELETE * from tbl_dateInfo"));
	m_pConnection->Execute((_bstr_t)sql, NULL, adCmdText);
	readAll();
}

//退出
void dateInfo::OnBnClickedButton5()
{
	// TODO:  在此添加控件通知处理程序代码
	CDialogEx::OnOK();
}


BOOL dateInfo::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  在此添加额外的初始化
	m_list.SetExtendedStyle(LVS_EX_FLATSB | LVS_EX_FULLROWSELECT |
		LVS_EX_HEADERDRAGDROP | LVS_EX_ONECLICKACTIVATE | LVS_EX_GRIDLINES);//
	CRect rect;
	m_list.GetClientRect(&rect);
	m_list.InsertColumn(0, _T("时间"), LVCFMT_LEFT, rect.Width() / 4, 0);
	m_list.InsertColumn(1, _T("标题"), LVCFMT_LEFT, rect.Width() / 4, 1);
	m_list.InsertColumn(2, _T("内容"), LVCFMT_LEFT, rect.Width() / 4, 2);
	m_list.InsertColumn(3, _T("周期"), LVCFMT_LEFT, rect.Width() / 4, 3);
	readAll();
	CHeaderCtrl* pHeaderCtrl = (CHeaderCtrl*)m_list.GetHeaderCtrl();
	pHeaderCtrl->EnableWindow(FALSE);

	CTime m_date;
	m_dateCtrl.GetTime(m_date);
	m_dateCtrl.SetFormat(_T("dddd yyyy/MM/dd"));

	SetDlgItemText(IDC_COMBO1, _T("每天"));
	//CString str;
	//str.Format(_T("星期%d"), m_date.GetDayOfWeek()-1);
	//GetDlgItem(IDC_EDIT4)->SetWindowText(str);


	
	SetTimer(1, 1000, NULL);
	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常:  OCX 属性页应返回 FALSE
}


void dateInfo::readAll()
{
	m_list.DeleteAllItems();
	//_bstr_t bstrSQL = "select * from tbl_dateInfo order by ID desc";

	//m_pRecordset.CreateInstance(_uuidof(Recordset));
	//m_pRecordset->Open(bstrSQL, m_pConnection.GetInterfacePtr(), adOpenDynamic,
	//	adLockOptimistic, adCmdText);

	CString sql;
	sql.Format(_T("select * from tbl_dateInfo order by begintime asc"));
	m_pRecordset = m_pConnection->Execute((_bstr_t)sql, NULL, adCmdText);
	while(!m_pRecordset->adoEOF)
	{
		m_list.InsertItem(0, _T(""));
		m_list.SetItemText(0, 0, (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("begintime"));
		m_list.SetItemText(0, 1, (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("title"));
		m_list.SetItemText(0, 2, (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("detail"));
		m_list.SetItemText(0, 3, (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("period"));
		m_pRecordset->MoveNext();
		//m_pRecordset->MovePrevious();
	}
}

BOOL dateInfo::Find()
{
	CString sql;
	CString str;
	GetDlgItem(IDC_EDIT2)->GetWindowText(str);
	CString str1;
	GetDlgItem(IDC_EDIT3)->GetWindowText(str1);
	CTime m_date;
	m_dateCtrl.GetTime(m_date);
	CString text;
	GetDlgItemText(IDC_COMBO1, text);
	sql.Format(_T("select * from tbl_dateInfo where begintime='%s'and title='%s'and detail='%s'and period='%s'"), 
		m_date.Format("%Y/%m/%d"), str,str1,text);
	m_pRecordset = m_pConnection->Execute((_bstr_t)sql, NULL, adCmdText);
	if (!m_pRecordset->adoEOF)
		return TRUE;
	return FALSE;
}


void dateInfo::OnNMClickList1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO:  在此添加控件通知处理程序代码
	int nItem = -1;
	if (pNMItemActivate != NULL)
	{
		nItem = pNMItemActivate->iItem;
	}
	//设置当前行号的内容
	m_list.GetItemText(nItem, 1);
	m_list.GetItemText(nItem, 2);
	SetDlgItemText(IDC_STATIC8, _T("时间：") + m_list.GetItemText(nItem, 0)
		+"\n"+"标题："+ m_list.GetItemText(nItem, 1) 
		+ "\n" + "内容：" + m_list.GetItemText(nItem, 2)
		+ "\n" + "周期：" + m_list.GetItemText(nItem, 3));
	*pResult = 0;
}


//UINT ThreadFun(LPVOID pParam){  //线程要调用的函数
//
//	MessageBox(NULL, _T("i am called by a thread."), _T("thread func"), MB_OK);
//
//}

void dateInfo::OnTimer(UINT_PTR nIDEvent)
{
	// TODO:  在此添加消息处理程序代码和/或调用默认值
	CDialogEx::OnTimer(nIDEvent);
	if (nIDEvent == 1)
	{
		CTime tm;
		tm = CTime::GetCurrentTime();
		_bstr_t bstrSQL = "select * from tbl_dateInfo order by begintime asc";
		m_pRecordset.CreateInstance(_uuidof(Recordset));
		m_pRecordset->Open(bstrSQL, m_pConnection.GetInterfacePtr(), adOpenDynamic,
			adLockOptimistic, adCmdText);

		while(!m_pRecordset->adoEOF)
		{
			CString time = m_pRecordset->GetCollect("begintime");
			CString period = m_pRecordset->GetCollect("period");
			if (time == tm.Format("%Y/%m/%d"))
			{

				//CWinThread* pThread=::AfxBeginThread(ThreadFun, NULL);

			}
			m_pRecordset->MoveNext();
		}
	}

}




