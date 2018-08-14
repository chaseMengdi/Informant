#pragma once
class Database
{
public:
	Database();
	~Database();

	void OnInitADOConn();
	void ExitConnect();
	_ConnectionPtr m_pConnection;
	_RecordsetPtr m_pRecordset;
};

