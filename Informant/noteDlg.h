#pragma once
#include "afxcmn.h"
#include "afxdtctl.h"


// noteDlg �Ի���

class noteDlg : public CDialogEx
{
	DECLARE_DYNAMIC(noteDlg)

public:
	noteDlg(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~noteDlg();
	CString sql = _T("SELECT *FROM tbl_note"), Str1, Str2, Str3;
	_ConnectionPtr m_pConnection; //���ӽӿ� 
	_RecordsetPtr m_pRecordset;  //�����򿪿������ݱ������ָ��
	virtual void OnFinalRelease();

	// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = DD_NOTE_DIALOG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

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
