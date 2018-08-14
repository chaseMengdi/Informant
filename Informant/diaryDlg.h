#pragma once
#include "afxwin.h"
#include "afxdtctl.h"
#include "ATLComTime.h"
#include "afxcmn.h"
#include "Database.h"


// diaryDlg �Ի���

class diaryDlg : public CDialogEx
{
	DECLARE_DYNAMIC(diaryDlg)

public:
	diaryDlg(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~diaryDlg();

	virtual void OnFinalRelease();

	// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = DD_DIARY_DIALOG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
public:
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedAdd();
	CButton m_add;
	CButton m_change;
	//	CDateTimeCtrl m_datetime;
	COleDateTime m_datetime;
	CButton m_delete;
	CString m_detail;
	CButton m_empty;
	CButton m_exit;
	CListCtrl m_output;
	CString m_title;
	CString m_weather;
	virtual BOOL OnInitDialog();
	CDateTimeCtrl m_dateCtrl;
	Database Informant;
	_RecordsetPtr m_pRecordset;
	_ConnectionPtr m_pConnection;
	void readAll();
	void getData();
	BOOL Find();
	afx_msg void OnLvnItemchangedOutputlist(NMHDR *pNMHDR, LRESULT *pResult);
	//afx_msg void OnBnClickedChange();
	afx_msg void OnBnClickedDelete();
	afx_msg void OnBnClickedExit();
	afx_msg void OnBnClickedChange();
	afx_msg void OnBnClickedEmpty();
	CDateTimeCtrl dateTimePicker1;
};