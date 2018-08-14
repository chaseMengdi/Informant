// timeInfoDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "Informant.h"
#include "timeInfoDlg.h"
#include "afxdialogex.h"

// timeInfoDlg 对话框
CString sql = L"SELECT *FROM tbl_time", Str1, Str2, Str3;
IMPLEMENT_DYNAMIC(timeInfoDlg, CDialogEx)

CFont font;
timeInfoDlg::timeInfoDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(DD_TIMEINFO_DIALOG, pParent)
{
	//m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

timeInfoDlg::~timeInfoDlg()
{
}

void timeInfoDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, timeInfo);
	DDX_Control(pDX, IDC_DATETIMEPICKER1, timeC1);
}


BEGIN_MESSAGE_MAP(timeInfoDlg, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON1, &timeInfoDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &timeInfoDlg::OnBnClickedButton2)
	ON_BN_CLICKED(IDC_BUTTON3, &timeInfoDlg::OnBnClickedButton3)
	ON_BN_CLICKED(IDC_BUTTON4, &timeInfoDlg::OnBnClickedButton4)
	ON_BN_CLICKED(IDC_BUTTON5, &timeInfoDlg::OnBnClickedButton5)
	ON_NOTIFY(LVN_ITEMCHANGED, IDC_LIST1, &timeInfoDlg::OnLvnItemchangedList1)
END_MESSAGE_MAP()
//更新数据表
void timeInfoDlg::Update(CString sql)
{
	//打开数据表
	//m_pRecordset.CreateInstance(__uuidof(Recordset));
	try
	{
		m_pRecordset->Open(_variant_t(sql), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	}
	catch (_com_error e)
	{
		CString errormessage;
		errormessage.Format(L"查询失败:%s", e.ErrorMessage());
		AfxMessageBox(errormessage);//显示错误信息
	}
	while (!m_pRecordset->adoEOF)
	{
		timeInfo.InsertItem(0, _T(""));
		timeInfo.SetItemText(0, 0, (CString)m_pRecordset->GetCollect(_T("date_")));
		timeInfo.SetItemText(0, 1, (CString)m_pRecordset->GetCollect(_T("title_")));
		timeInfo.SetItemText(0, 2, (CString)m_pRecordset->GetCollect(_T("detail_")));
		m_pRecordset->MoveNext();
	}
	m_pRecordset->Close();
	UpdateData();
}
// timeInfoDlg 消息处理程序
BOOL timeInfoDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	timeC1.SetFormat(L"yyyy/MM/dd");
	//连接数据库
	try
	{
		CoInitialize(NULL);
		m_pConnection.CreateInstance(__uuidof(Connection));//初始化指针
		m_pConnection->Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Informant.mdb", "", "", adModeUnknown);
	}
	catch (_com_error e)
	{
		CString errormessage;
		errormessage.Format(L"连接数据库失败!\r\n错误信息:%s", e.ErrorMessage());
		AfxMessageBox(errormessage);//显示错误信息
	}
	CRect rect;
	timeInfo.GetClientRect(&rect);
	timeInfo.SetExtendedStyle(timeInfo.GetExtendedStyle() | LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);
	timeInfo.InsertColumn(0, L"时间", LVCFMT_CENTER, rect.Width() * 2 / 7);
	timeInfo.InsertColumn(1, L"标题", LVCFMT_CENTER, rect.Width() * 2 / 7);
	timeInfo.InsertColumn(2, L"内容", LVCFMT_CENTER, rect.Width() * 3 / 7);
	CHeaderCtrl* pHeaderCtrl = (CHeaderCtrl*)timeInfo.GetHeaderCtrl();
	pHeaderCtrl->EnableWindow(FALSE);
	//打开数据表
	m_pRecordset.CreateInstance(__uuidof(Recordset));
	Update(sql);

	return TRUE;
}
//添加
void timeInfoDlg::OnBnClickedButton1()
{
	int i = 0;
	CString str1, str2, str3;
	GetDlgItemText(IDC_DATETIMEPICKER1, str1);
	GetDlgItemText(IDC_EDIT3, str2);
	GetDlgItemText(IDC_EDIT4, str3);
	if (str1 != ""&&str2 != ""&&str3 != "")
	{
		try
		{
			m_pRecordset.CreateInstance(__uuidof(Recordset));
			m_pRecordset->Open(_variant_t(sql), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
		}
		catch (_com_error e)
		{
			AfxMessageBox(_T("打开数据表失败") + e.Description());//显示错误信息
		}
		while (!m_pRecordset->adoEOF)
		{
			if (str1 == (CString)m_pRecordset->GetCollect(_T("date_"))
				&& str2 == (CString)m_pRecordset->GetCollect(_T("title_"))
				&& str3 == (CString)m_pRecordset->GetCollect(_T("detail_")))
			{
				i++;
				AfxMessageBox(_T("消息已经存在！"));
				break;
			}
			else m_pRecordset->MoveNext();
		}
		if (i == 0)
		{
			CString	sq = _T("INSERT INTO tbl_time VALUES ('") + str1 + _T("','") + str2 + _T("','") + str3 + _T("')");
			m_pRecordset.CreateInstance(__uuidof(Recordset));
			try {       //利用Open函数执行SQL命令      
				m_pRecordset->Open(_variant_t(sq),
					m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
				timeInfo.InsertItem(0, str1);
				timeInfo.SetItemText(0, 1, str2);
				timeInfo.SetItemText(0, 2, str3);
			}
			catch (_com_error e) {
				AfxMessageBox(_T("添加失败") + e.Description());
			}

		}
	}
	else if (str1 == "" || str2 == "" || str3 == "")
		AfxMessageBox(_T("请输入有效信息"));
}
//删除
void timeInfoDlg::OnBnClickedButton2()
{
	POSITION pos = timeInfo.GetFirstSelectedItemPosition();
	if (Str1 != ""&&Str2 != ""&&Str3 != ""&&pos != NULL)
	{
		CString sq = _T("DELETE FROM tbl_time WHERE date_='") + Str1 + _T("' AND title_='") + Str2 + _T("' AND detail_='") + Str3 + _T("'");
		m_pRecordset.CreateInstance(__uuidof(Recordset));
		try {       //利用Open函数执行SQL命令      
			m_pRecordset->Open(_variant_t(sq),
				m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
			//AfxMessageBox(_T("删除成功！"));
		}
		catch (_com_error e)
		{
			AfxMessageBox(_T("删除失败") + e.Description());
		}
		timeInfo.DeleteItem((int)timeInfo.GetNextSelectedItem(pos));
	}
	else
		AfxMessageBox(_T("请选择数据！"));
}
//修改
void timeInfoDlg::OnBnClickedButton3()
{
	int i = 0;
	CString str1, str2, str3;
	GetDlgItemText(IDC_DATETIMEPICKER1, str1);
	GetDlgItemText(IDC_EDIT3, str2);
	GetDlgItemText(IDC_EDIT4, str3);
	POSITION pos = timeInfo.GetFirstSelectedItemPosition();
	if (str1 == "" || str2 == "" || str3 == "" || pos == NULL)
		AfxMessageBox(_T("请输入有效信息"));
	else if (str1 != ""&&str2 != ""&&str3 != "")
	{
		{
			try
			{
				m_pRecordset.CreateInstance(__uuidof(Recordset));
				m_pRecordset->Open(_variant_t(sql),
					m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
				m_pRecordset->MoveFirst();
				while (!m_pRecordset->adoEOF)
				{
					if (str1 == (CString)m_pRecordset->GetCollect(_T("date_"))
						&& str2 == (CString)m_pRecordset->GetCollect(_T("title_"))
						&& str3 == (CString)m_pRecordset->GetCollect(_T("detail_")))
					{
						i++;
						AfxMessageBox(_T("消息已经存在！"));
						break;
					}
					else m_pRecordset->MoveNext();
				}
			}
			catch (_com_error e)
			{
				AfxMessageBox(e.Description());
			}

		}
		if (i == 0)
		{
			try
			{
				CString sq = _T("DELETE FROM tbl_time WHERE date_='") + Str1 + _T("' AND title_='") + Str2 + _T("' AND detail_='") + Str3 + _T("'");
				m_pRecordset.CreateInstance(__uuidof(Recordset));
				m_pRecordset->Open(_variant_t(sq),
					m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
				sq = _T("INSERT INTO tbl_time VALUES ('") + str1 + _T("','") + str2 + _T("','") + str3 + _T("')");
				m_pRecordset.CreateInstance(__uuidof(Recordset));
				m_pRecordset->Open(_variant_t(sq),
					m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
				//AfxMessageBox(_T("更新成功！"));
			}
			catch (_com_error e)
			{
				AfxMessageBox(_T("更新失败！") + e.Description());
			}
			timeInfo.DeleteAllItems();
			Update(sql);
		}
	}
}
//清空
void timeInfoDlg::OnBnClickedButton4()
{
	try
	{
		m_pRecordset->Open(_variant_t(sql), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	}
	catch (_com_error e)
	{
		CString errormessage;
		errormessage.Format(L"打开数据表失败:%s", e.ErrorMessage());
		AfxMessageBox(errormessage);//显示错误信息
	}
	if (m_pRecordset->adoEOF)
		AfxMessageBox(_T("表已清空！"));
	else if (IDYES == AfxMessageBox(L"是否清空数据？", MB_YESNO))
	{
		m_pRecordset->Close();

		CString sq = _T("DELETE * FROM tbl_time");
		m_pRecordset.CreateInstance(__uuidof(Recordset));
		try {       //利用Open函数执行SQL命令      
			m_pRecordset->Open(_variant_t(sq),
				m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
			//AfxMessageBox(_T("清空成功！"));
			timeInfo.DeleteAllItems();
		}
		catch (_com_error e)
		{
			AfxMessageBox(_T("清空失败"));
		}
	}
}
//退出
void timeInfoDlg::OnBnClickedButton5()
{
	//关闭数据库
	::CoUninitialize(); //关闭COM环境
	CDialogEx::OnOK();//退出
}
//显示数据
void timeInfoDlg::OnLvnItemchangedList1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	NMLISTVIEW *pnml = (NMLISTVIEW*)pNMLV;
	if (-1 != pnml->iItem)
	{
		Str1 = timeInfo.GetItemText(pnml->iItem, 0);
		Str2 = timeInfo.GetItemText(pnml->iItem, 1);
		Str3 = timeInfo.GetItemText(pnml->iItem, 2);
		SetDlgItemText(IDC_STATIC1, _T("时间:  ") + Str1
			+ "\n" + "标题： " + Str2
			+ "\n" + "内容： " + Str3);
	}
	*pResult = 0;
}