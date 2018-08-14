// noteDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "Informant.h"
#include "noteDlg.h"
#include "afxdialogex.h"


// noteDlg �Ի���

IMPLEMENT_DYNAMIC(noteDlg, CDialogEx)

noteDlg::noteDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(DD_NOTE_DIALOG, pParent)
{

	EnableAutomation();

}

noteDlg::~noteDlg()
{
}

void noteDlg::OnFinalRelease()
{
	// �ͷ��˶��Զ�����������һ�����ú󣬽�����
	// OnFinalRelease��  ���ཫ�Զ�
	// ɾ���ö���  �ڵ��øû���֮ǰ�����������
	// ��������ĸ���������롣

	CDialogEx::OnFinalRelease();
}
//�������
void noteDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, noteLC);

	//noteDate.SetFormat(L"yyyy/MM/dd");

	CRect rect;
	noteLC.GetClientRect(&rect);
	noteLC.SetExtendedStyle(noteLC.GetExtendedStyle() | LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);
	noteLC.InsertColumn(0, _T("ʱ��"), LVCFMT_CENTER, rect.Width() * 2 / 7);
	noteLC.InsertColumn(1, _T("����"), LVCFMT_CENTER, rect.Width() * 2 / 7);
	noteLC.InsertColumn(2, _T("����"), LVCFMT_CENTER, rect.Width() * 3 / 7);
	CHeaderCtrl* pHeaderCtrl = (CHeaderCtrl*)noteLC.GetHeaderCtrl();
	pHeaderCtrl->EnableWindow(FALSE);


	try
	{
		CoInitialize(NULL);
		m_pConnection.CreateInstance(__uuidof(Connection));//��ʼ��ָ��
		m_pConnection->Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Informant.mdb", "", "", adModeUnknown);
	}
	catch (_com_error e)
	{
		CString errormessage;
		errormessage.Format(_T("�������ݿ�ʧ��!\r\n������Ϣ:%s"), e.ErrorMessage());
		AfxMessageBox(errormessage);//��ʾ������Ϣ
	}
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
		noteLC.InsertItem(0, _T(""));
		noteLC.SetItemText(0, 0, (CString)m_pRecordset->GetCollect(_T("date_")));
		noteLC.SetItemText(0, 1, (CString)m_pRecordset->GetCollect(_T("title_")));
		noteLC.SetItemText(0, 2, (CString)m_pRecordset->GetCollect(_T("detail_")));
		m_pRecordset->MoveNext();
	}
	m_pRecordset->Close();
	DDX_Control(pDX, IDC_DATETIMEPICKER1, noteDate);
}


BEGIN_MESSAGE_MAP(noteDlg, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON5, &noteDlg::OnBnClickedButton5)
	ON_BN_CLICKED(IDC_BUTTON1, &noteDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &noteDlg::OnBnClickedButton2)
	ON_NOTIFY(LVN_ITEMCHANGED, IDC_LIST1, &noteDlg::OnLvnItemchangedList1)
	ON_BN_CLICKED(IDC_BUTTON4, &noteDlg::OnBnClickedButton4)
	ON_BN_CLICKED(IDC_BUTTON3, &noteDlg::OnBnClickedButton3)
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(noteDlg, CDialogEx)
END_DISPATCH_MAP()

// ע��: ������� IID_InoteDlg ֧��
//  ��֧������ VBA �����Ͱ�ȫ�󶨡�  �� IID ����ͬ���ӵ� .IDL �ļ��е�
//  ���Ƚӿڵ� GUID ƥ�䡣

// {45081699-6EFE-4FE3-B830-CD43B7B96EFD}
static const IID IID_InoteDlg =
{ 0x45081699, 0x6EFE, 0x4FE3,{ 0xB8, 0x30, 0xCD, 0x43, 0xB7, 0xB9, 0x6E, 0xFD } };

BEGIN_INTERFACE_MAP(noteDlg, CDialogEx)
	INTERFACE_PART(noteDlg, IID_InoteDlg, Dispatch)
END_INTERFACE_MAP()


// noteDlg ��Ϣ�������

//�˳�
void noteDlg::OnBnClickedButton5()
{
	::CoUninitialize(); //�ر�COM����
	CDialogEx::OnOK();//�˳�
}

//����
void noteDlg::OnBnClickedButton1()
{
	int i = 0;
	CString str1, str2, str3;
	GetDlgItemText(IDC_DATETIMEPICKER1, str1);
	GetDlgItemText(IDC_EDIT2, str2);
	GetDlgItemText(IDC_EDIT3, str3);
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
		//�鿴�Ƿ�����ͬ�ļ�¼
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
		//û����ͬ��¼���
		if (i == 0)
		{
			CString	sq = _T("INSERT INTO tbl_note VALUES ('") + str1 + _T("','") + str2 + _T("','") + str3 + _T("')");
			m_pRecordset.CreateInstance(__uuidof(Recordset));
			try {       //����Open����ִ��SQL����      
				m_pRecordset->Open(_variant_t(sq),
					m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
				noteLC.InsertItem(0, str1);
				noteLC.SetItemText(0, 1, str2);
				noteLC.SetItemText(0, 2, str3);
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
void noteDlg::OnBnClickedButton2()
{
	POSITION pos = noteLC.GetFirstSelectedItemPosition();
	if (Str1 != ""&&Str2 != ""&&Str3 != ""&&pos != NULL)
	{

		CString sq = _T("DELETE FROM tbl_note WHERE date_='") + Str1 + _T("' AND title_='") + Str2 + _T("' AND detail_='") + Str3 + _T("'");

		try {       //����Open����ִ��SQL����   
			m_pRecordset.CreateInstance(__uuidof(Recordset));
			m_pRecordset->Open(_variant_t(sq),
				m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
			AfxMessageBox(_T("ɾ���ɹ���"));
		}
		catch (_com_error e)
		{
			AfxMessageBox(_T("ɾ��ʧ��") + e.Description());
		}

		noteLC.DeleteItem((int)noteLC.GetNextSelectedItem(pos));
	}
	else
		AfxMessageBox(_T("��ѡ�����ݣ�"));
}

//ѡ��
void noteDlg::OnLvnItemchangedList1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	NMLISTVIEW *pnml = (NMLISTVIEW*)pNMLV;
	if (-1 != pnml->iItem)
	{
		Str1 = noteLC.GetItemText(pnml->iItem, 0);
		Str2 = noteLC.GetItemText(pnml->iItem, 1);
		Str3 = noteLC.GetItemText(pnml->iItem, 2);
		SetDlgItemText(IDC_DATETIMEPICKER1, Str1);
		SetDlgItemText(IDC_EDIT2, Str2);
		SetDlgItemText(IDC_EDIT3, Str3);
		*pResult = 0;
	}
	*pResult = 0;
}

//���
void noteDlg::OnBnClickedButton4()
{
	try
	{
		m_pRecordset->Open(_variant_t(sql), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	}
	catch (_com_error e)
	{
		CString errormessage;
		errormessage.Format(_T("�����ݱ�ʧ��:%s"), e.ErrorMessage());
		AfxMessageBox(errormessage);//��ʾ������Ϣ
	}
	if (m_pRecordset->adoEOF)
		AfxMessageBox(_T("������գ�"));
	else if (IDYES == AfxMessageBox(_T("�Ƿ�������ݣ�"), MB_YESNO))
	{
		m_pRecordset->Close();

		CString sq = _T("DELETE * FROM tbl_note");
		m_pRecordset.CreateInstance(__uuidof(Recordset));
		try {       //����Open����ִ��SQL����      
			m_pRecordset->Open(_variant_t(sq),
				m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
			AfxMessageBox(_T("��ճɹ���"));
			noteLC.DeleteAllItems();
		}
		catch (_com_error e)
		{
			AfxMessageBox(_T("���ʧ��"));
		}
	}
}

//�޸�
void noteDlg::OnBnClickedButton3()
{
	int i = 0;
	CString str1, str2, str3;
	GetDlgItemText(IDC_DATETIMEPICKER1, str1);
	GetDlgItemText(IDC_EDIT2, str2);
	GetDlgItemText(IDC_EDIT3, str3);
	POSITION pos = noteLC.GetFirstSelectedItemPosition();
	if (str1 == "" || str2 == "" || str3 == "" || pos == NULL)
		AfxMessageBox(_T("��������Ч��Ϣ"));
	else if (str1 != ""&&str2 != ""&&str3 != "")
	{
		try
		{
			m_pRecordset.CreateInstance(__uuidof(Recordset));
			m_pRecordset->Open(_variant_t(sql), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
			m_pRecordset->MoveFirst();
		}
		catch (_com_error e)
		{
			AfxMessageBox(_T("�����ݱ�ʧ��") + e.Description());//��ʾ������Ϣ
		}
		//�鿴�Ƿ�����ͬ�ļ�¼
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
			m_pRecordset->MoveNext();
		}
		if (i == 0)
		{
			try
			{
				CString sq = _T("DELETE FROM tbl_note WHERE date_='") + Str1 + _T("' AND title_='") + Str2 + _T("' AND detail_='") + Str3 + _T("'");
				m_pRecordset.CreateInstance(__uuidof(Recordset));
				m_pRecordset->Open(_variant_t(sq),
					m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
				sq = _T("INSERT INTO tbl_note VALUES ('") + str1 + _T("','") + str2 + _T("','") + str3 + _T("')");
				noteLC.DeleteItem((int)noteLC.GetNextSelectedItem(pos));
				m_pRecordset.CreateInstance(__uuidof(Recordset));
				m_pRecordset->Open(_variant_t(sq),
					m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
				noteLC.InsertItem(0, str1);
				noteLC.SetItemText(0, 1, str2);
				noteLC.SetItemText(0, 2, str3);
				AfxMessageBox(_T("���³ɹ���"));
			}
			catch (_com_error e)
			{
				AfxMessageBox(_T("����ʧ�ܣ�") + e.Description());
			}
		}
	}
}


BOOL noteDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  �ڴ���Ӷ���ĳ�ʼ��
	noteDate.SetFormat(L"yyyy/MM/dd");

	return TRUE;  // return TRUE unless you set the focus to a control
				  // �쳣: OCX ����ҳӦ���� FALSE
}
