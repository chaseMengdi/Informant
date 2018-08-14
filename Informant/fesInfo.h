#pragma once
#include "afxcmn.h"
#include "afxdtctl.h"


// fesInfo 对话框

class fesInfo : public CDialogEx
{
	DECLARE_DYNAMIC(fesInfo)

public:
	fesInfo(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~fesInfo();

	// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = DD_FESINFO_DIALOG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
	virtual BOOL OnInitDialog();
	CListCtrl fesList;
	_ConnectionPtr pCon;
	_CommandPtr pCom;
	_RecordsetPtr pRec;
	afx_msg void OnBnClickedButton2();
	afx_msg void OnBnClickedButton4();
	afx_msg void OnBnClickedButton5();
	afx_msg void OnBnClickedButton3();
	CDateTimeCtrl dateTimePicker;
};
