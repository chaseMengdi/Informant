// dateInfo.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "Informant.h"
#include "dateInfo.h"
#include "afxdialogex.h"
#include "InformantDlg.h"




// dateInfo �Ի���

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
	// �ͷ��˶��Զ�����������һ�����ú󣬽�����
	// OnFinalRelease��  ���ཫ�Զ�
	// ɾ���ö���  �ڵ��øû���֮ǰ�����������
	// ��������ĸ���������롣

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

// ע��: ������� IID_IdateInfo ֧��
//  ��֧������ VBA �����Ͱ�ȫ�󶨡�  �� IID ����ͬ���ӵ� .IDL �ļ��е�
//  ���Ƚӿڵ� GUID ƥ�䡣

// {A4FF1B1E-608A-4A40-A25A-A1B5A164C827}
static const IID IID_IdateInfo =
{ 0xA4FF1B1E, 0x608A, 0x4A40, { 0xA2, 0x5A, 0xA1, 0xB5, 0xA1, 0x64, 0xC8, 0x27 } };

BEGIN_INTERFACE_MAP(dateInfo, CDialogEx)
	INTERFACE_PART(dateInfo, IID_IdateInfo, Dispatch)
END_INTERFACE_MAP()


// dateInfo ��Ϣ�������

//���
void dateInfo::OnBnClickedButton1()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CString str;
	GetDlgItem(IDC_EDIT2)->GetWindowText(str);
	CString str1;
	GetDlgItem(IDC_EDIT3)->GetWindowText(str1);
	if (str == _T("") || str1 == _T(""))
		AfxMessageBox(_T("��������ݲ���Ϊ��"));
	else
	{
		if (Find())AfxMessageBox(_T("��Ҫ�ظ����"));
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

//ɾ��
void dateInfo::OnBnClickedButton2()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	//CString sql;
	/*CString str;
	GetDlgItem(IDC_EDIT2)->GetWindowText(str);*/
	CString date, name, detail, sql, period;
	POSITION pos = m_list.GetFirstSelectedItemPosition();
	int nId;
	if (pos != NULL) 
	{//�����ѡ��ĳ��ʱ
		nId = (int)m_list.GetNextSelectedItem(pos);
		date = m_list.GetItemText(nId, 0);
		name = m_list.GetItemText(nId, 1);
		detail = m_list.GetItemText(nId, 2);//��ȡѡ�еĽ��������Ϣ
		period = m_list.GetItemText(nId, 3);
		sql.Format(_T("DELETE from tbl_dateInfo where begintime='%s' and title='%s'and detail='%s'and period='%s'"), date, name, detail,period);
		m_pConnection->Execute((_bstr_t)sql, NULL, adCmdText);
		//AfxMessageBox(L"ɾ���ɹ���");
		readAll();
			
		}
	else//���δѡ������
		AfxMessageBox(L"δѡ����Ϣ��");
	
}

//�޸�
void dateInfo::OnBnClickedButton3()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	CString date, name, detail,period;
	POSITION pos = m_list.GetFirstSelectedItemPosition();
	int nId;
	if (pos != NULL) 
	{//�����ѡ��ĳ��ʱ
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
		//AfxMessageBox(L"�޸ĳɹ���");
	}
	else//���δѡ������
		AfxMessageBox(_T("δѡ����Ϣ��"));
	
}

//���
void dateInfo::OnBnClickedButton4()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	CString sql;
	sql.Format(_T("DELETE * from tbl_dateInfo"));
	m_pConnection->Execute((_bstr_t)sql, NULL, adCmdText);
	readAll();
}

//�˳�
void dateInfo::OnBnClickedButton5()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	CDialogEx::OnOK();
}


BOOL dateInfo::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  �ڴ���Ӷ���ĳ�ʼ��
	m_list.SetExtendedStyle(LVS_EX_FLATSB | LVS_EX_FULLROWSELECT |
		LVS_EX_HEADERDRAGDROP | LVS_EX_ONECLICKACTIVATE | LVS_EX_GRIDLINES);//
	CRect rect;
	m_list.GetClientRect(&rect);
	m_list.InsertColumn(0, _T("ʱ��"), LVCFMT_LEFT, rect.Width() / 4, 0);
	m_list.InsertColumn(1, _T("����"), LVCFMT_LEFT, rect.Width() / 4, 1);
	m_list.InsertColumn(2, _T("����"), LVCFMT_LEFT, rect.Width() / 4, 2);
	m_list.InsertColumn(3, _T("����"), LVCFMT_LEFT, rect.Width() / 4, 3);
	readAll();
	CHeaderCtrl* pHeaderCtrl = (CHeaderCtrl*)m_list.GetHeaderCtrl();
	pHeaderCtrl->EnableWindow(FALSE);

	CTime m_date;
	m_dateCtrl.GetTime(m_date);
	m_dateCtrl.SetFormat(_T("dddd yyyy/MM/dd"));

	SetDlgItemText(IDC_COMBO1, _T("ÿ��"));
	//CString str;
	//str.Format(_T("����%d"), m_date.GetDayOfWeek()-1);
	//GetDlgItem(IDC_EDIT4)->SetWindowText(str);


	
	SetTimer(1, 1000, NULL);
	return TRUE;  // return TRUE unless you set the focus to a control
	// �쳣:  OCX ����ҳӦ���� FALSE
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
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	int nItem = -1;
	if (pNMItemActivate != NULL)
	{
		nItem = pNMItemActivate->iItem;
	}
	//���õ�ǰ�кŵ�����
	m_list.GetItemText(nItem, 1);
	m_list.GetItemText(nItem, 2);
	SetDlgItemText(IDC_STATIC8, _T("ʱ�䣺") + m_list.GetItemText(nItem, 0)
		+"\n"+"���⣺"+ m_list.GetItemText(nItem, 1) 
		+ "\n" + "���ݣ�" + m_list.GetItemText(nItem, 2)
		+ "\n" + "���ڣ�" + m_list.GetItemText(nItem, 3));
	*pResult = 0;
}


//UINT ThreadFun(LPVOID pParam){  //�߳�Ҫ���õĺ���
//
//	MessageBox(NULL, _T("i am called by a thread."), _T("thread func"), MB_OK);
//
//}

void dateInfo::OnTimer(UINT_PTR nIDEvent)
{
	// TODO:  �ڴ������Ϣ�����������/�����Ĭ��ֵ
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




