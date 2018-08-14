// LoginDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "Informant.h"
#include "LoginDlg.h"
#include "afxdialogex.h"


// CLoginDlg 对话框

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


// CLoginDlg 消息处理程序


void CLoginDlg::OnBnClickedLogin()
{
	// TODO:  在此添加控件通知处理程序代码
	CString str;
	GetDlgItem(IDC_EDIT1)->GetWindowText(str);
	CString str1;
	GetDlgItem(IDC_EDIT5)->GetWindowText(str1);
	if (str == _T("") || str1 == _T("")){
		AfxMessageBox(_T("用户名或密码不能为空！"));
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
			AfxMessageBox(_T("请检查用户名或密码"));
	}
}




void CLoginDlg::OnBnClickedButton1()
{
	// TODO:  在此添加控件通知处理程序代码
	// TODO:  在此添加控件通知处理程序代码
	CString str;
	GetDlgItem(IDC_EDIT1)->GetWindowText(str);
	CString str1;
	GetDlgItem(IDC_EDIT5)->GetWindowText(str1);
	if (str == _T("") || str1 == _T(""))
	{
		AfxMessageBox(_T("用户名或密码名不能为空！"));
	}
	else
	{
		CString sqlll;
		sqlll.Format(_T("select * from tbl_user where username='%s'"), str);
		m_pRecordset = m_pConnection->Execute((_bstr_t)sqlll, NULL, adCmdText);
		if (!m_pRecordset->adoEOF)
			AfxMessageBox(_T("不要重复注册"));
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
			AfxMessageBox(_T("注册成功!"));
		}
	}
}


void CLoginDlg::OnClose()
{
	// TODO:  在此添加消息处理程序代码和/或调用默认值
	exit(0);
	CDialogEx::OnClose();
}