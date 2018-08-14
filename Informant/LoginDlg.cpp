// LoginDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "Informant.h"
#include "LoginDlg.h"
#include "afxdialogex.h"


// CLoginDlg �Ի���

IMPLEMENT_DYNAMIC(CLoginDlg, CDialogEx)

CLoginDlg::CLoginDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CLoginDlg::IDD, pParent)
{
	DA.OnInitADOConn();
	m_pRecordset = DA.m_pRecordset;
	m_pConnection = DA.m_pConnection;
}

CLoginDlg::~CLoginDlg()
{
	DA.ExitConnect();
}

void CLoginDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CLoginDlg, CDialogEx)
	ON_BN_CLICKED(IDC_LOGIN, &CLoginDlg::OnBnClickedLogin)
	ON_WM_CLOSE()
	ON_BN_CLICKED(IDC_BUTTON1, &CLoginDlg::OnBnClickedButton1)
END_MESSAGE_MAP()


// CLoginDlg ��Ϣ�������


void CLoginDlg::OnBnClickedLogin()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	CString str;
	GetDlgItem(IDC_EDIT1)->GetWindowText(str);
	CString str1;
	GetDlgItem(IDC_EDIT5)->GetWindowText(str1);
	if (str == _T("") || str1 == _T("")){
		AfxMessageBox(_T("�û��������벻��Ϊ�գ�"));
	}
	else{
		CString sql;
		sql.Format(_T("select * from tbl_user where username='%s'"), str);
		m_pRecordset = m_pConnection->Execute((_bstr_t)sql, NULL, adCmdText);
		CString password;
		if (!m_pRecordset->adoEOF){
			password = m_pRecordset->GetCollect("password");
			
		}
		if (password == str1)
			CDialogEx::OnOK();
		else
			AfxMessageBox(_T("�����û���������"));
	}
}




void CLoginDlg::OnBnClickedButton1()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	CString str;
	GetDlgItem(IDC_EDIT1)->GetWindowText(str);
	CString str1;
	GetDlgItem(IDC_EDIT5)->GetWindowText(str1);
	if (str == _T("") || str1 == _T(""))
	{
		AfxMessageBox(_T("�û���������������Ϊ�գ�"));
	}
	else
	{
		CString sqlll;
		sqlll.Format(_T("select * from tbl_user where username='%s'"), str);
		m_pRecordset = m_pConnection->Execute((_bstr_t)sqlll, NULL, adCmdText);
		if (!m_pRecordset->adoEOF)
			AfxMessageBox(_T("��Ҫ�ظ�ע��"));
		else
		{
			_bstr_t sql;
			sql = "select * from tbl_user";
			m_pRecordset.CreateInstance(__uuidof(Recordset));
			m_pRecordset->Open(sql, m_pConnection.GetInterfacePtr(),
				adOpenDynamic, adLockOptimistic, adCmdText);
			m_pRecordset->AddNew();
			m_pRecordset->PutCollect("username", (_bstr_t)str);
			m_pRecordset->PutCollect("password", (_bstr_t)str1);
			m_pRecordset->Update();
			AfxMessageBox(_T("ע��ɹ�!"));
		}
	}
}


void CLoginDlg::OnClose()
{
	// TODO:  �ڴ������Ϣ�����������/�����Ĭ��ֵ
	exit(0);
	CDialogEx::OnClose();
}