#pragma once
#include "afxcmn.h"
#include "Database.h"

// dateInfo 对话框

class dateInfo : public CDialogEx
{
	DECLARE_DYNAMIC(dateInfo)

public:
	dateInfo(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~dateInfo();

	virtual void OnFinalRelease();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = DD_DATEINFO_DIALOG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
public:
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton2();
	afx_msg void OnBnClickedButton3();
	afx_msg void OnBnClickedButton4();
	afx_msg void OnBnClickedButton5();
	virtual BOOL OnInitDialog();
	CListCtrl m_list;
	CDateTimeCtrl m_dateCtrl;

	Database DA;
	_RecordsetPtr m_pRecordset;
	_ConnectionPtr m_pConnection;

	void readAll();
	BOOL Find();
	afx_msg void OnNMClickList1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	UINT ThreadFun(LPVOID pParam);
};
