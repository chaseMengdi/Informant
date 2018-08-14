#include "stdafx.h"
#include "Database.h"


Database::Database()
{
}


Database::~Database()
{
}


void Database::OnInitADOConn()
{
	try
	{
		CoInitialize(NULL);//≥ı ºªØcomø‚
		m_pConnection = _ConnectionPtr(__uuidof(Connection));
		m_pConnection->ConnectionString =
			"Provider = Microsoft.Jet.OLEDB.4.0; Data Source =.\\Informant.mdb";
		m_pConnection->Open("", "", "", adConnectUnspecified);
	}
	catch (_com_error e)
	{
		AfxMessageBox(e.Description());
	}
}


void Database::ExitConnect()
{
	if (m_pRecordset != NULL)
		m_pRecordset->Close();
	m_pConnection->Close();
	CoUninitialize();//–∂‘ÿcomø‚
}