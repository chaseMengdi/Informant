#pragma once
#include "CWMPPlayer4.h"
#include "afxcmn.h"


// remindDlg �Ի���

class remindDlg : public CDialogEx
{
	DECLARE_DYNAMIC(remindDlg)

public:
	remindDlg(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~remindDlg();

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_REMIND_DIALOG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
public:
	CWMPPlayer4 myPlayer;
	_ConnectionPtr pCon;
	_RecordsetPtr pRec;
	virtual BOOL OnInitDialog();
	CListCtrl infoTime;
	CListCtrl infoDate;
};
