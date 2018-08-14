#pragma once
#include "CWMPPlayer4.h"
#include "afxcmn.h"


// remindDlg 对话框

class remindDlg : public CDialogEx
{
	DECLARE_DYNAMIC(remindDlg)

public:
	remindDlg(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~remindDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_REMIND_DIALOG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	CWMPPlayer4 myPlayer;
	_ConnectionPtr pCon;
	_RecordsetPtr pRec;
	virtual BOOL OnInitDialog();
	CListCtrl infoTime;
	CListCtrl infoDate;
};
