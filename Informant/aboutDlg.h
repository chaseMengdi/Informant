#pragma once


// aboutDlg �Ի���

class aboutDlg : public CDialogEx
{
	DECLARE_DYNAMIC(aboutDlg)

public:
	aboutDlg(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~aboutDlg();

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
};
