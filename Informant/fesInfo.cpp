// fesInfo.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "Informant.h"
#include "fesInfo.h"
#include "afxdialogex.h"


// fesInfo �Ի���

IMPLEMENT_DYNAMIC(fesInfo, CDialogEx)

fesInfo::fesInfo(CWnd* pParent /*=NULL*/)
	: CDialogEx(DD_FESINFO_DIALOG, pParent)
{
}

fesInfo::~fesInfo()
{
}

void fesInfo::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, fesList);
	DDX_Control(pDX, IDC_DATETIMEPICKER1, dateTimePicker);
}


BEGIN_MESSAGE_MAP(fesInfo, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON1, &fesInfo::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &fesInfo::OnBnClickedButton2)
	ON_BN_CLICKED(IDC_BUTTON4, &fesInfo::OnBnClickedButton4)
	ON_BN_CLICKED(IDC_BUTTON5, &fesInfo::OnBnClickedButton5)
	ON_BN_CLICKED(IDC_BUTTON3, &fesInfo::OnBnClickedButton3)
END_MESSAGE_MAP()

BOOL fesInfo::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  �ڴ���Ӷ���ĳ�ʼ��
	//fesList��ʼ��
	dateTimePicker.SetFormat(L"yyyy/MM/dd");
	CRect rec;
	fesList.GetClientRect(&rec);
	fesList.SetExtendedStyle(fesList.GetExtendedStyle() | LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);
	fesList.InsertColumn(0, L"����ʱ��", LVCFMT_CENTER, rec.Width() * 5 / 32);
	fesList.InsertColumn(1, L"��������", LVCFMT_CENTER, rec.Width() / 4);
	fesList.InsertColumn(2, L"��ϸ��Ϣ", LVCFMT_CENTER, rec.Width() * 11 / 32);
	fesList.InsertColumn(3, L"������ģ�", LVCFMT_CENTER, rec.Width() / 4);
	CHeaderCtrl* pHeaderCtrl = (CHeaderCtrl*)fesList.GetHeaderCtrl();
	pHeaderCtrl->EnableWindow(FALSE);

	//���ݿ�����
	//AfxOleInit();
	CoInitialize(NULL);
	pCon.CreateInstance(__uuidof(Connection));
	_bstr_t strCon =
		"Provider = Microsoft.Jet.OLEDB.4.0; Data Source = .\\Informant.mdb";
	//"DRIVER={Microsoft Access Driver (*.mdb)};DBQ=Informant.mdb;";
	try {
		pCon->Open(strCon, "", "", adModeUnknown);
	}
	catch (_com_error e) {
		AfxMessageBox(e.ErrorMessage() + e.Description());
	}

	//��ѯ��ȡ�����ݵ�fesList
	CString sql = L"SELECT * FROM tbl_festival ORDER BY date";
	pRec.CreateInstance(__uuidof(Recordset));
	try {       //����Open����ִ��SQL�����ò�ѯ�����¼��         
				//pRec->Open(_variant_t(sql),
				//	pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
		pRec = pCon->Execute((_bstr_t)sql, NULL, adCmdText);
	}
	catch (_com_error e) {
		AfxMessageBox(e.ErrorMessage() + e.Description());
	}
	while (!pRec->adoEOF) {
		fesList.InsertItem(0, (CString)pRec->GetCollect(L"date"));
		fesList.SetItemText(0, 1, (CString)pRec->GetCollect(L"name_"));
		fesList.SetItemText(0, 2, (CString)pRec->GetCollect(L"remarks"));
		fesList.SetItemText(0, 3, (CString)pRec->GetCollect(L"change"));
		pRec->MoveNext();
	}

	return TRUE;  // return TRUE unless you set the focus to a control
				  // �쳣: OCX ����ҳӦ���� FALSE
}
// fesInfo ��Ϣ�������

//����µĽ�������
void fesInfo::OnBnClickedButton1()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CString date, name, detail, sql;

	GetDlgItem(IDC_DATETIMEPICKER1)->GetWindowText(date);
	GetDlgItem(IDC_EDIT2)->GetWindowTextW(name);
	GetDlgItem(IDC_EDIT3)->GetWindowTextW(detail);//��ȡ���������Ϣ
	if (name == L"" || detail == L"")
		AfxMessageBox(L"���ƻ���ϸ����Ϊ�գ�");
	else {//������Ϣ��Ϊ��ʱ
		try {       //ִ��SQL����      
			sql = L"SELECT * FROM tbl_festival WHERE date='" + date + L"' AND name_='" + name + L"' AND remarks='" + detail + L"' AND change='TRUE'";
			pRec.CreateInstance(__uuidof(Recordset));
			pRec->Open(_variant_t(sql),
				pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
		}
		catch (_com_error e) {
			AfxMessageBox(e.ErrorMessage() + e.Description());
			
		}
		if (pRec->adoEOF) {//���δ�ҵ�
			sql = L"INSERT INTO tbl_festival VALUES ('" + date + L"','" + name + L"','" + detail + L"', 'TRUE')";
			pRec.CreateInstance(__uuidof(Recordset));
			try {       //����Open����ִ��SQL����      
				pRec->Open(_variant_t(sql),
					pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
				//AfxMessageBox(L"��ӳɹ���");
				fesList.InsertItem(0, date);
				fesList.SetItemText(0, 1, name);
				fesList.SetItemText(0, 2, detail);
				fesList.SetItemText(0, 3, L"TRUE");
			}
			catch (_com_error e) {
				AfxMessageBox(e.ErrorMessage() + e.Description());
			}
		}
		else
			AfxMessageBox(L"�������ظ���ӣ�");
	}
}

//���½�������
void fesInfo::OnBnClickedButton3()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CString date, name, detail, change, sql, date_n, name_n, detail_n;
	POSITION pos = fesList.GetFirstSelectedItemPosition();
	int nId;
	if (pos != NULL) {//�����ѡ��ĳ��ʱ
		nId = (int)fesList.GetNextSelectedItem(pos);
		date = fesList.GetItemText(nId, 0);
		name = fesList.GetItemText(nId, 1);
		detail = fesList.GetItemText(nId, 2);
		change = fesList.GetItemText(nId, 3);//��ȡѡ�еĽ��������Ϣ
		if (change == L"TRUE") {
			GetDlgItem(IDC_DATETIMEPICKER1)->GetWindowText(date_n);
			GetDlgItem(IDC_EDIT2)->GetWindowTextW(name_n);
			GetDlgItem(IDC_EDIT3)->GetWindowTextW(detail_n);//��ȡ�����µ������Ϣ
															//sql = L"UPDATE tbl_festival  SET tbl_festival.date='" + date_n + L"' AND tbl_festival.name_='" + name_n + L"' AND tbl_festival.remarks='" + detail_n
															//	+ L"' WHERE tbl_festival.date='" + date + L"' AND tbl_festival.name_='" + name + L"' AND tbl_festival.remarks='" + detail + L"'";
															//GetDlgItem(IDC_EDIT2)->SetWindowTextW(sql);
			pRec.CreateInstance(__uuidof(Recordset));
			try {       //����Open����ִ��SQL����    
				sql = L"DELETE FROM tbl_festival WHERE date='" + date + L"' AND name_='" + name + L"' AND remarks='" + detail + L"' AND change='TRUE'";
				pRec->Open(_variant_t(sql),
					pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
				sql = L"INSERT INTO tbl_festival VALUES ('" + date_n + L"','" + name_n + L"','" + detail_n + L"','TRUE')";
				pRec->Open(_variant_t(sql),
					pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
				//AfxMessageBox(L"���³ɹ���");
				fesList.SetItemText(nId, 0, date_n);
				fesList.SetItemText(nId, 1, name_n);
				fesList.SetItemText(nId, 2, detail_n);
			}
			catch (_com_error e) {
				AfxMessageBox(e.ErrorMessage() + e.Description());
			}
		}
		else
			AfxMessageBox(L"ϵͳ���ղ�������ģ�");
	}
	else//���δѡ������
		AfxMessageBox(L"δѡ�н��գ�");
}

//��ս�������
void fesInfo::OnBnClickedButton4()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CString sql = L"DELETE  FROM tbl_festival WHERE change='TRUE'";
	pRec.CreateInstance(__uuidof(Recordset));
	try {       //����Open����ִ��SQL����      
		pRec->Open(_variant_t(sql),
			pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
		//AfxMessageBox(L"��ճɹ���");
		fesList.DeleteAllItems();//��յ�ǰfesList
								 //�ٴγ�ʼ��fesList
		sql = L"SELECT * FROM tbl_festival ORDER BY date";
		pRec.CreateInstance(__uuidof(Recordset));
		try {       //����Open����ִ��SQL�����ò�ѯ�����¼��         
					//pRec->Open(_variant_t(sql),
					//	pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
			pRec = pCon->Execute((_bstr_t)sql, NULL, adCmdText);
		}
		catch (_com_error e) {
			AfxMessageBox(e.ErrorMessage() + e.Description());
		}
		while (!pRec->adoEOF) {
			fesList.InsertItem(0, (CString)pRec->GetCollect(L"date"));
			fesList.SetItemText(0, 1, (CString)pRec->GetCollect(L"name_"));
			fesList.SetItemText(0, 2, (CString)pRec->GetCollect(L"remarks"));
			fesList.SetItemText(0, 3, (CString)pRec->GetCollect(L"change"));
			pRec->MoveNext();
		}

	}
	catch (_com_error e) {
		AfxMessageBox(e.ErrorMessage() + e.Description());
	}
}

//�˳��������Ѳ��ر����ݿ�����
void fesInfo::OnBnClickedButton5()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CDialogEx::OnOK();
	CoUninitialize();
	/*if (pCon->State){
	pCon->Close();
	pRec->Close();
	}*/
}

//ɾ��ѡ�еĽ�������
void fesInfo::OnBnClickedButton2()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CString date, name, detail, sql, change;
	POSITION pos = fesList.GetFirstSelectedItemPosition();
	int nId;
	if (pos != NULL) {//�����ѡ��ĳ��ʱ
		nId = (int)fesList.GetNextSelectedItem(pos);
		date = fesList.GetItemText(nId, 0);
		name = fesList.GetItemText(nId, 1);
		detail = fesList.GetItemText(nId, 2);
		change = fesList.GetItemText(nId, 3);//��ȡѡ�еĽ��������Ϣ
		if (change == L"TRUE") {//�û���ӽ��տ�ɾ��
			sql = L"DELETE FROM tbl_festival WHERE date='" + date + L"' AND name_='" + name + L"' AND remarks='" + detail + L"'";
			pRec.CreateInstance(__uuidof(Recordset));
			try {       //����Open����ִ��SQL����      
				pRec->Open(_variant_t(sql),
					pCon.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
				//AfxMessageBox(L"ɾ���ɹ���");
				fesList.DeleteItem(nId);
			}
			catch (_com_error e) {
				AfxMessageBox(e.ErrorMessage() + e.Description());
			}
		}
		else
			AfxMessageBox(L"ϵͳ���ղ�����ɾ����");
	}
	else//���δѡ������
		AfxMessageBox(L"δѡ�н��գ�");
}