// remindDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "Informant.h"
#include "remindDlg.h"
#include "afxdialogex.h"


// remindDlg �Ի���

IMPLEMENT_DYNAMIC(remindDlg, CDialogEx)

remindDlg::remindDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_REMIND_DIALOG, pParent)
{

}

remindDlg::~remindDlg()
{
}

void remindDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_OCX1, myPlayer);
	DDX_Control(pDX, IDC_LIST1, infoTime);
	DDX_Control(pDX, IDC_LIST3, infoDate);
}


BEGIN_MESSAGE_MAP(remindDlg, CDialogEx)
END_MESSAGE_MAP()


// remindDlg ��Ϣ�������


BOOL remindDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  �ڴ���Ӷ���ĳ�ʼ��
	CRect rect;
	infoTime.GetClientRect(&rect);
	infoTime.SetExtendedStyle(infoTime.GetExtendedStyle() | LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);
	infoTime.InsertColumn(0, L"����", LVCFMT_CENTER, rect.Width()/2);
	infoTime.InsertColumn(1, L"����", LVCFMT_CENTER, rect.Width()/2);
	CHeaderCtrl* pHeaderCtrl = (CHeaderCtrl*)infoTime.GetHeaderCtrl();
	pHeaderCtrl->EnableWindow(FALSE);

	infoDate.GetClientRect(&rect);
	infoDate.SetExtendedStyle(infoDate.GetExtendedStyle() | LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);
	infoDate.InsertColumn(0, L"����", LVCFMT_CENTER, rect.Width()/3);
	infoDate.InsertColumn(1, L"����", LVCFMT_CENTER, rect.Width() / 3);
	infoDate.InsertColumn(2, L"����", LVCFMT_CENTER, rect.Width()/3);
	pHeaderCtrl = (CHeaderCtrl*)infoDate.GetHeaderCtrl();
	pHeaderCtrl->EnableWindow(FALSE);

	//���ݿ�����
	CoInitialize(NULL);
	pCon.CreateInstance(__uuidof(Connection));
	_bstr_t strCon =
		"Provider = Microsoft.Jet.OLEDB.4.0; Data Source = .\\Informant.mdb";
	//"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\\Informant.mdb";
	try {
		pCon->Open(strCon, "", "", adModeUnknown);
	}
	catch (_com_error e) {
		AfxMessageBox(L"��ʼ�����Ӵ���\n");
		AfxMessageBox(e.ErrorMessage() + e.Description());
	}
	//��ȡϵͳʱ��
	CTime now = CTime::GetCurrentTime();
	int nowY = now.GetYear(), nowM = now.GetMonth(), nowD = now.GetDay();
	COleDateTime now_1(nowY, nowM, nowD,0,0,0);
	CString nowS = now.Format("%Y/%m/%d");
	SetDlgItemText(IDC_STATIC1, L"������" + nowS + L"��");
	CString sql = L"SELECT * FROM tbl_time WHERE date_='"+nowS+L"'";
	pRec.CreateInstance(__uuidof(Recordset));
	try {       //����Open����ִ��SQL�����ò�ѯ�����¼�� 
		pRec->Open(_variant_t(sql),
			pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	}
	catch (_com_error e) {
		AfxMessageBox(L"��ѯ���ݴ���1\n");
		AfxMessageBox(e.ErrorMessage() + e.Description());
	}
	while (!pRec->adoEOF) {
		infoTime.InsertItem(0, (CString)pRec->GetCollect(L"title_"));
		infoTime.SetItemText(0, 1, (CString)pRec->GetCollect(L"detail_"));
		pRec->MoveNext();
	}
	//�����������ݻ�ȡ
	BOOL state = FALSE;
	CString date_, year_, month_, day_, cycle_;
	int year, month, day,sub;
	sql = L"SELECT * FROM tbl_dateInfo";
	pRec.CreateInstance(__uuidof(Recordset));
	try {       //����Open����ִ��SQL�����ò�ѯ�����¼�� 
		pRec->Open(_variant_t(sql),
			pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	}
	catch (_com_error e) {
		AfxMessageBox(L"��ѯ���ݴ���3\n");
		AfxMessageBox(e.ErrorMessage() + e.Description());
	}
	while (!pRec->adoEOF) {
		date_=(CString)pRec->GetCollect(L"begintime");
		cycle_=(CString)pRec->GetCollect(L"period");//��ȡ���ݿ��е����ں�����
		year_ = date_.Left(4);
		month_ = date_.Right(5);
		month_ = month_.Left(2);
		day_ = date_.Right(2);
		year=_ttoi(year_);
		month=_ttoi(month_);
		day=_ttoi(day_);
		COleDateTime time_1(year,month,day,0,0,0);
		sub = abs((int)(time_1-now_1).m_span);
		if (cycle_ == L"ÿ��")//�ж��Ƿ񵽴����ڣ��˴����˼򻯴���
			state = TRUE;        //δ����ƽ�����ꡢ��ͬ�·���������
		if (cycle_ == L"ÿ��" && (sub % 7 == 0))
			state = TRUE;
		if (cycle_ == L"ÿ��" && nowD==day && nowM!=month)
			state = TRUE;
		if (cycle_ == L"ÿ��"&& nowD==day&&nowM==month&&nowY!=year)
			state = TRUE;
		if (state) {//�����������������
			infoDate.InsertItem(0,(CString)pRec->GetCollect(L"title"));
			infoDate.SetItemText(0, 1, (CString)pRec->GetCollect(L"detail"));
			infoDate.SetItemText(0, 2, (CString)pRec->GetCollect(L"period"));
		}
		pRec->MoveNext();

		//�����������ݻ�ȡ
		sql = L"SELECT * FROM tbl_music";
		pRec.CreateInstance(__uuidof(Recordset));
		try {       //����Open����ִ��SQL�����ò�ѯ�����¼�� 
			pRec->Open(_variant_t(sql),
				pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
		}
		catch (_com_error e) {
			AfxMessageBox(L"��ѯ���ִ���\n");
			AfxMessageBox(e.ErrorMessage() + e.Description());
		}
		while (!pRec->adoEOF) {
			myPlayer.put_URL((CString)pRec->GetCollect(L"music"));
			pRec->MoveNext();
		}
	}

	//��������



	return TRUE;  // return TRUE unless you set the focus to a control
				  // �쳣: OCX ����ҳӦ���� FALSE
}
