#pragma once
#include "afxcmn.h"
#include "afxdtctl.h"


// noteDlg 对话框

class noteDlg : public CDialogEx
{
	DECLARE_DYNAMIC(noteDlg)

public:
	noteDlg(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~noteDlg();
	CString sql = _T("SELECT *FROM tbl_note"), Str1, Str2, Str3;
	_ConnectionPtr m_pConnection; //连接接口 
	_RecordsetPtr m_pRecordset;  //用来打开库内数据表的智能指针
	virtual void OnFinalRelease();

	// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = DD_NOTE_DIALOG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
public:
	CListCtrl noteLC;
	afx_msg void OnBnClickedButton5();
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton2();
	afx_msg void OnLvnItemchangedList1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedButton4();
	afx_msg void OnBnClickedButton3();
	CDateTimeCtrl noteDate;
	virtual BOOL OnInitDialog();
};
