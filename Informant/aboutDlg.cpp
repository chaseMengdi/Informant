// aboutDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "Informant.h"
#include "aboutDlg.h"
#include "afxdialogex.h"


// aboutDlg �Ի���

IMPLEMENT_DYNAMIC(aboutDlg, CDialogEx)

aboutDlg::aboutDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_ABOUTBOX, pParent)
{

}

aboutDlg::~aboutDlg()
{
}

void aboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(aboutDlg, CDialogEx)
END_MESSAGE_MAP()


// aboutDlg ��Ϣ�������
