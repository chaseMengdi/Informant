// timeInfoDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "Informant.h"
#include "timeInfoDlg.h"
#include "afxdialogex.h"

// timeInfoDlg �Ի���
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
//�������ݱ�
void timeInfoDlg::Update(CString sql)
{
	//�����ݱ�
	//m_pRecordset.CreateInstance(__uuidof(Recordset));
	try
	{
		m_pRecordset->Open(_variant_t(sql), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	}
	catch (_com_error e)
	{
		CString errormessage;
		errormessage.Format(L"��ѯʧ��:%s", e.ErrorMessage());
		AfxMessageBox(errormessage);//��ʾ������Ϣ
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
// timeInfoDlg ��Ϣ�������
BOOL timeInfoDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	timeC1.SetFormat(L"yyyy/MM/dd");
	//�������ݿ�
	try
	{
		CoInitialize(NULL);
		m_pConnection.CreateInstance(__uuidof(Connection));//��ʼ��ָ��
		m_pConnection->Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Informant.mdb", "", "", adModeUnknown);
	}
	catch (_com_error e)
	{
		CString errormessage;
		errormessage.Format(L"�������ݿ�ʧ��!\r\n������Ϣ:%s", e.ErrorMessage());
		AfxMessageBox(errormessage);//��ʾ������Ϣ
	}
	CRect rect;
	timeInfo.GetClientRect(&rect);
	timeInfo.SetExtendedStyle(timeInfo.GetExtendedStyle() | LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);
	timeInfo.InsertColumn(0, L"ʱ��", LVCFMT_CENTER, rect.Width() * 2 / 7);
	timeInfo.InsertColumn(1, L"����", LVCFMT_CENTER, rect.Width() * 2 / 7);
	timeInfo.InsertColumn(2, L"����", LVCFMT_CENTER, rect.Width() * 3 / 7);
	CHeaderCtrl* pHeaderCtrl = (CHeaderCtrl*)timeInfo.GetHeaderCtrl();
	pHeaderCtrl->EnableWindow(FALSE);
	//�����ݱ�
	m_pRecordset.CreateInstance(__uuidof(Recordset));
	Update(sql);

	return TRUE;
}
//���
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
			AfxMessageBox(_T("�����ݱ�ʧ��") + e.Description());//��ʾ������Ϣ
		}
		while (!m_pRecordset->adoEOF)
		{
			if (str1 == (CString)m_pRecordset->GetCollect(_T("date_"))
				&& str2 == (CString)m_pRecordset->GetCollect(_T("title_"))
				&& str3 == (CString)m_pRecordset->GetCollect(_T("detail_")))
			{
				i++;
				AfxMessageBox(_T("��Ϣ�Ѿ����ڣ�"));
				break;
			}
			else m_pRecordset->MoveNext();
		}
		if (i == 0)
		{
			CString	sq = _T("INSERT INTO tbl_time VALUES ('") + str1 + _T("','") + str2 + _T("','") + str3 + _T("')");
			m_pRecordset.CreateInstance(__uuidof(Recordset));
			try {       //����Open����ִ��SQL����      
				m_pRecordset->Open(_variant_t(sq),
					m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
				timeInfo.InsertItem(0, str1);
				timeInfo.SetItemText(0, 1, str2);
				timeInfo.SetItemText(0, 2, str3);
			}
			catch (_com_error e) {
				AfxMessageBox(_T("���ʧ��") + e.Description());
			}

		}
	}
	else if (str1 == "" || str2 == "" || str3 == "")
		AfxMessageBox(_T("��������Ч��Ϣ"));
}
//ɾ��
void timeInfoDlg::OnBnClickedButton2()
{
	POSITION pos = timeInfo.GetFirstSelectedItemPosition();
	if (Str1 != ""&&Str2 != ""&&Str3 != ""&&pos != NULL)
	{
		CString sq = _T("DELETE FROM tbl_time WHERE date_='") + Str1 + _T("' AND title_='") + Str2 + _T("' AND detail_='") + Str3 + _T("'");
		m_pRecordset.CreateInstance(__uuidof(Recordset));
		try {       //����Open����ִ��SQL����      
			m_pRecordset->Open(_variant_t(sq),
				m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
			//AfxMessageBox(_T("ɾ���ɹ���"));
		}
		catch (_com_error e)
		{
			AfxMessageBox(_T("ɾ��ʧ��") + e.Description());
		}
		timeInfo.DeleteItem((int)timeInfo.GetNextSelectedItem(pos));
	}
	else
		AfxMessageBox(_T("��ѡ�����ݣ�"));
}
//�޸�
void timeInfoDlg::OnBnClickedButton3()
{
	int i = 0;
	CString str1, str2, str3;
	GetDlgItemText(IDC_DATETIMEPICKER1, str1);
	GetDlgItemText(IDC_EDIT3, str2);
	GetDlgItemText(IDC_EDIT4, str3);
	POSITION pos = timeInfo.GetFirstSelectedItemPosition();
	if (str1 == "" || str2 == "" || str3 == "" || pos == NULL)
		AfxMessageBox(_T("��������Ч��Ϣ"));
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
						AfxMessageBox(_T("��Ϣ�Ѿ����ڣ�"));
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
				//AfxMessageBox(_T("���³ɹ���"));
			}
			catch (_com_error e)
			{
				AfxMessageBox(_T("����ʧ�ܣ�") + e.Description());
			}
			timeInfo.DeleteAllItems();
			Update(sql);
		}
	}
}
//���
void timeInfoDlg::OnBnClickedButton4()
{
	try
	{
		m_pRecordset->Open(_variant_t(sql), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	}
	catch (_com_error e)
	{
		CString errormessage;
		errormessage.Format(L"�����ݱ�ʧ��:%s", e.ErrorMessage());
		AfxMessageBox(errormessage);//��ʾ������Ϣ
	}
	if (m_pRecordset->adoEOF)
		AfxMessageBox(_T("������գ�"));
	else if (IDYES == AfxMessageBox(L"�Ƿ�������ݣ�", MB_YESNO))
	{
		m_pRecordset->Close();

		CString sq = _T("DELETE * FROM tbl_time");
		m_pRecordset.CreateInstance(__uuidof(Recordset));
		try {       //����Open����ִ��SQL����      
			m_pRecordset->Open(_variant_t(sq),
				m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
			//AfxMessageBox(_T("��ճɹ���"));
			timeInfo.DeleteAllItems();
		}
		catch (_com_error e)
		{
			AfxMessageBox(_T("���ʧ��"));
		}
	}
}
//�˳�
void timeInfoDlg::OnBnClickedButton5()
{
	//�ر����ݿ�
	::CoUninitialize(); //�ر�COM����
	CDialogEx::OnOK();//�˳�
}
//��ʾ����
void timeInfoDlg::OnLvnItemchangedList1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	NMLISTVIEW *pnml = (NMLISTVIEW*)pNMLV;
	if (-1 != pnml->iItem)
	{
		Str1 = timeInfo.GetItemText(pnml->iItem, 0);
		Str2 = timeInfo.GetItemText(pnml->iItem, 1);
		Str3 = timeInfo.GetItemText(pnml->iItem, 2);
		SetDlgItemText(IDC_STATIC1, _T("ʱ��:  ") + Str1
			+ "\n" + "���⣺ " + Str2
			+ "\n" + "���ݣ� " + Str3);
	}
	*pResult = 0;
}