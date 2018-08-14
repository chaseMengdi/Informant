#pragma once
#include "afxcmn.h"
#include "afxdtctl.h"


// timeInfoDlg 对话框

class timeInfoDlg : public CDialogEx
{
	DECLARE_DYNAMIC(timeInfoDlg)

public:
	timeInfoDlg(CWnd* pParent = NULL);   // 标准构造函数
										 //	BOOL OnInitDialog();
	virtual ~timeInfoDlg();

	// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = DD_TIMEINFO_DIALOG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	_ConnectionPtr m_pConnection; //连接接口 
	_RecordsetPtr m_pRecordset;  //用来打开库内数据表的智能指针
	CListCtrl timeInfo;
	void Update(CString sql);
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton2();
	afx_msg void OnBnClickedButton3();
	afx_msg void OnBnClickedButton4();
	afx_msg void OnBnClickedButton5();
	afx_msg void OnLvnItemchangedList1(NMHDR *pNMHDR, LRESULT *pResult);
	
	CDateTimeCtrl timeC1;
};
