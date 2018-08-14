// diaryDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "Informant.h"
#include "diaryDlg.h"
#include "InformantDlg.h"
#include "afxdialogex.h"
#include "Database.h"



// diaryDlg �Ի���

IMPLEMENT_DYNAMIC(diaryDlg, CDialogEx)

diaryDlg::diaryDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(DD_DIARY_DIALOG, pParent)
	, m_datetime(COleDateTime::GetCurrentTime())
	, m_detail(_T(""))
	, m_title(_T(""))
	, m_weather(_T(""))
{
	EnableAutomation();
	Informant.OnInitADOConn();
	m_pRecordset = Informant.m_pRecordset;
	m_pConnection = Informant.m_pConnection;
}

diaryDlg::~diaryDlg()
{
	Informant.ExitConnect();
}

void diaryDlg::OnFinalRelease()
{
	// �ͷ��˶��Զ�����������һ�����ú󣬽�����
	// OnFinalRelease��  ���ཫ�Զ�
	// ɾ���ö���  �ڵ��øû���֮ǰ�����������
	// ��������ĸ���������롣

	CDialogEx::OnFinalRelease();
}

void diaryDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_Add, m_add);
	DDX_Control(pDX, IDC_Change, m_change);
	//  DDX_Control(pDX, IDC_DATETIMEPICKER1, m_datetime);
	DDX_DateTimeCtrl(pDX, IDC_DATEMEPICKER1, m_datetime);
	DDX_Control(pDX, IDC_Delete, m_delete);
	DDX_Text(pDX, IDC_DetailEdit, m_detail);
	DDX_Control(pDX, IDC_Empty, m_empty);
	DDX_Control(pDX, IDC_Exit, m_exit);
	DDX_Control(pDX, IDC_OutputList, m_output);
	DDX_Text(pDX, IDC_TitleEdit, m_title);
	DDX_Text(pDX, IDC_WeatherEdit, m_weather);
	DDX_Control(pDX, IDC_DATEMEPICKER1, dateTimePicker1);
}


BEGIN_MESSAGE_MAP(diaryDlg, CDialogEx)

	//ON_BN_CLICKED(IDC_Add, &diaryDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_Add, &diaryDlg::OnBnClickedAdd)
	ON_NOTIFY(LVN_ITEMCHANGED, IDC_OutputList, &diaryDlg::OnLvnItemchangedOutputlist)
	//ON_BN_CLICKED(IDC_Change, &diaryDlg::OnBnClickedChange)
	ON_BN_CLICKED(IDC_Delete, &diaryDlg::OnBnClickedDelete)
	ON_BN_CLICKED(IDC_Exit, &diaryDlg::OnBnClickedExit)
	ON_BN_CLICKED(IDC_Change, &diaryDlg::OnBnClickedChange)
	ON_BN_CLICKED(IDC_Empty, &diaryDlg::OnBnClickedEmpty)
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(diaryDlg, CDialogEx)
END_DISPATCH_MAP()

// ע��: ������� IID_IdiaryDlg ֧��
//  ��֧������ VBA �����Ͱ�ȫ�󶨡�  �� IID ����ͬ���ӵ� .IDL �ļ��е�
//  ���Ƚӿڵ� GUID ƥ�䡣

// {A000407D-CB48-4208-ADCE-C0B9BE55A9AF}
static const IID IID_IdiaryDlg =
{ 0xA000407D, 0xCB48, 0x4208,{ 0xAD, 0xCE, 0xC0, 0xB9, 0xBE, 0x55, 0xA9, 0xAF } };

BEGIN_INTERFACE_MAP(diaryDlg, CDialogEx)
	INTERFACE_PART(diaryDlg, IID_IdiaryDlg, Dispatch)
END_INTERFACE_MAP()


// diaryDlg ��Ϣ�������






void diaryDlg::OnBnClickedAdd()
{
	UpdateData(TRUE);
	/*int year = m_datetime.GetYear();
	int mounth = m_datetime.GetMonth();
	int day = m_datetime.GetDay();
	CString yearTime;
	CString mounthTime;
	CString dayTime;
	yearTime.Format(L"%d", year);
	mounthTime.Format(L"%d", mounth);
	dayTime.Format(L"%d", day);*/

	CString timeFinal;
	GetDlgItem(IDC_DATEMEPICKER1)->GetWindowText(timeFinal);
	//timeFinal.Format(L"%s-%s-%s", yearTime, mounthTime, dayTime);


	Database Informant;
	Informant.OnInitADOConn();
	_bstr_t sql;
	sql = L"select * from tbl_diary";

	Informant.m_pRecordset.CreateInstance(__uuidof(Recordset));
	Informant.m_pRecordset->Open(sql, Informant.m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	try
	{
		Informant.m_pRecordset->AddNew(); //�������
		Informant.m_pRecordset->PutCollect("detail", (_bstr_t)m_detail);
		Informant.m_pRecordset->PutCollect("title", (_bstr_t)m_title);
		Informant.m_pRecordset->PutCollect("weather", (_bstr_t)m_weather);
		Informant.m_pRecordset->PutCollect("date", (_bstr_t)(timeFinal));
		Informant.m_pRecordset->Update(); //�������ݱ�
		Informant.ExitConnect();
		readAll();
	}
	catch (...)
	{
		MessageBox(L"����ʧ��");
		return;
	}
	//MessageBox(L"��ӳɹ�");
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
}

void diaryDlg::OnBnClickedDelete()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	CString sql, date, title, weather, detail;
	POSITION pos = m_output.GetFirstSelectedItemPosition();
	int nID;
	if (pos != NULL)
	{
		nID = (int)m_output.GetNextSelectedItem(pos);
		date = m_output.GetItemText(nID, 0);
		title = m_output.GetItemText(nID, 2);
		weather = m_output.GetItemText(nID, 1);
		detail = m_output.GetItemText(nID, 3);
		//sql = _T("DELETE��FROM tbl_diary WHERE date='") + date + _T("'AND weather='") + weather + _T("AND title='") + title + _T("AND detail='") + detail + _T("'");

		//sql = "DELETE FROM tbl_diary WHERE date='" + date + "' AND title='" + title + "'AND weather= '" + weather + "' AND detail='" + detail + "'";
		//AfxMessageBox(sql);
		sql.Format(L"DELETE FROM tbl_diary WHERE date='%s' AND weather= '%s' AND title='%s' AND detail='%s'", date, weather, title, detail);
		//GetDlgItem(IDC_TitleEdit)->SetWindowTextA(sql);
		//m_pRecordset.CreateInstance(__uuidof(Recordset));
		try
		{       //����Open����ִ��SQL����  
			m_pRecordset.CreateInstance(__uuidof(Recordset));
			m_pRecordset->Open(_variant_t(sql), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
			//m_pConnection->Execute((_bstr_t)sql, NULL, adCmdText);
			//AfxMessageBox(L"ɾ���ɹ���");
			readAll();
			//m_output.DeleteItem(nID); 

		}
		catch (_com_error e)
		{
			AfxMessageBox(e.ErrorMessage() + e.Description());
		}
	}
	else	AfxMessageBox(L"δѡ�����ݣ�");	 //���δѡ������
}

void diaryDlg::OnBnClickedChange()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	CString sql;
	CString str1;
	GetDlgItem(IDC_DetailEdit)->GetWindowText(str1);
	CString str2;
	GetDlgItem(IDC_TitleEdit)->GetWindowText(str2);
	if (Find())
	{
		sql.Format((L"UPDATE tbl_diary set detail='%s' where title='%s'"), str1, str2);
		m_pConnection->Execute((_bstr_t)sql, NULL, adCmdText);
		readAll();
	}
	else 
		AfxMessageBox(L"û�иñ�������");
}


void diaryDlg::OnBnClickedEmpty()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	CString sql;
	sql.Format(L"DELETE * from tbl_diary");
	m_pConnection->Execute((_bstr_t)sql, NULL, adCmdText);
	readAll();
}


void diaryDlg::OnBnClickedExit()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	CDialogEx::OnOK();
}

BOOL diaryDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  �ڴ���Ӷ���ĳ�ʼ��
	dateTimePicker1.SetFormat(L"yyyy/MM/dd");
	m_output.SetExtendedStyle(LVS_EX_FLATSB | LVS_EX_FULLROWSELECT |
		LVS_EX_HEADERDRAGDROP | LVS_EX_ONECLICKACTIVATE | LVS_EX_GRIDLINES);
	CRect rect;
	m_output.GetClientRect(&rect);
	m_output.InsertColumn(0, _T("            ʱ��"), LVCFMT_CENTER, rect.Width() / 4, 0);
	m_output.InsertColumn(1, _T("����"), LVCFMT_CENTER, rect.Width() / 4, 1);
	m_output.InsertColumn(2, _T("����"), LVCFMT_CENTER, rect.Width() / 4, 2);
	m_output.InsertColumn(3, _T("����"), LVCFMT_CENTER, rect.Width() / 4, 3);
	CHeaderCtrl* pHeaderCtrl = (CHeaderCtrl*)m_output.GetHeaderCtrl();
	pHeaderCtrl->EnableWindow(FALSE);

	// �쳣:  OCX ����ҳӦ���� FALSE
	readAll();
	return TRUE;  // return TRUE unless you set the focus to a control
}
void diaryDlg::readAll()
{
	Database Informant;
	Informant.OnInitADOConn();
	m_output.DeleteAllItems();
	//_bstr_t bstrSQL = "select * from tbl_dateInfo order by ID desc";

	//m_pRecordset.CreateInstance(_uuidof(Recordset));
	//m_pConnection.CreateInstance(__uuidof(Connection));
	m_pConnection = Informant.m_pConnection;
	m_pRecordset = Informant.m_pRecordset;


	CString sql;
	sql.Format(_T("select * from tbl_diary"));
	m_pRecordset = m_pConnection->Execute((_bstr_t)sql, NULL, adCmdText);
	//m_pRecordset->Open((bstr_t)sql, m_pConnection.GetInterfacePtr(), adOpenDynamic,
	//adLockOptimistic, adCmdText);

	while (!m_pRecordset->adoEOF)
	{
		m_output.InsertItem(0, _T(""));
		m_output.SetItemText(0, 0, (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("date"));
		m_output.SetItemText(0, 1, (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("weather"));
		m_output.SetItemText(0, 2, (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("title"));
		m_output.SetItemText(0, 3, (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("detail"));
		m_pRecordset->MoveNext();
		//m_pRecordset->MovePrevious();
	}
	//Informant.ExitConnect();
}
BOOL diaryDlg::Find()
{
	CString sql;
	CString str;
	GetDlgItem(IDC_TitleEdit)->GetWindowText(str);
	sql.Format(L"select * from tbl_diary where title='%s'", str);
	m_pRecordset = m_pConnection->Execute((_bstr_t)sql, NULL, adCmdText);
	if (!m_pRecordset->adoEOF)
		return TRUE;
	return FALSE;
}


void diaryDlg::OnLvnItemchangedOutputlist(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	int nItem = -1;
	if (pNMLV != NULL)
	{
		nItem = pNMLV->iItem;
	}
	//���õ�ǰ�кŵ�����
	m_output.GetItemText(nItem, 1);
	m_output.GetItemText(nItem, 2);
	SetDlgItemText(IDC_OutputList, _T("ʱ�䣺") + m_output.GetItemText(nItem, 0)
		+ "\n" + "������" + m_output.GetItemText(nItem, 1)
		+ "\n" + "���⣺" + m_output.GetItemText(nItem, 2)
		+ "\n" + "���ݣ�" + m_output.GetItemText(nItem, 3));
	getData();
	*pResult = 0;
}
void diaryDlg::getData() {
	CString date, title, weather, detail;
	POSITION pos = m_output.GetFirstSelectedItemPosition();
	int nID;
	if (pos != NULL)
	{
		nID = (int)m_output.GetNextSelectedItem(pos);
		date = m_output.GetItemText(nID, 0);
		weather = m_output.GetItemText(nID, 1);
		title = m_output.GetItemText(nID, 2);
		detail = m_output.GetItemText(nID, 3);
	}
	SetDlgItemText(IDC_WeatherEdit, weather);
	SetDlgItemText(IDC_TitleEdit, title);
	SetDlgItemText(IDC_DetailEdit, detail);
}