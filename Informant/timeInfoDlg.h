#pragma once
#include "afxcmn.h"
#include "afxdtctl.h"


// timeInfoDlg �Ի���

class timeInfoDlg : public CDialogEx
{
	DECLARE_DYNAMIC(timeInfoDlg)

public:
	timeInfoDlg(CWnd* pParent = NULL);   // ��׼���캯��
										 //	BOOL OnInitDialog();
	virtual ~timeInfoDlg();

	// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = DD_TIMEINFO_DIALOG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
public:
	_ConnectionPtr m_pConnection; //���ӽӿ� 
	_RecordsetPtr m_pRecordset;  //�����򿪿������ݱ������ָ��
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
