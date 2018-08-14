
// InformantDlg.h : 头文件
//

#pragma once
#include "afxcmn.h"


// CInformantDlg 对话框
class CInformantDlg : public CDialogEx
{
	// 构造
public:
	CInformantDlg(CWnd* pParent = NULL);	// 标准构造函数

											// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_INFORMANT_DIALOG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


														// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	CListCtrl infoMain;
	_ConnectionPtr pCon;
	_CommandPtr pCom;
	_RecordsetPtr pRec;
	afx_msg void On32773();
	afx_msg void On32774();
	afx_msg void On32775();
	afx_msg void On32777();
	afx_msg void On32778();
	afx_msg void On32771();
	afx_msg void On32776();
	afx_msg void On32779();
	afx_msg void On32772();
	CListCtrl infoDate;
	BOOL state;
protected:
	afx_msg LRESULT OnSystemtray(WPARAM wParam, LPARAM lParam);
	virtual void OnCancel();
public:
	virtual BOOL DestroyWindow();
	void readAll();
private:
	NOTIFYICONDATA NotifyIcon;  //系统托盘类
public:
	afx_msg void OnTimer(UINT_PTR nIDEvent);
};
