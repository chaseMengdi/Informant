#pragma once


// aboutDlg 对话框

class aboutDlg : public CDialogEx
{
	DECLARE_DYNAMIC(aboutDlg)

public:
	aboutDlg(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~aboutDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
};
