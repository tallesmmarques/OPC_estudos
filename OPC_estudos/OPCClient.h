#pragma once

#include <objidl.h>
#include <vector>

#include "opcda.h"
#include "SOCDataCallback.h"

#define OPC_SERVER_NAME L"Matrikon.OPC.Simulation.1"

class OPCClient
{
private:
	IOPCServer* pIOPCServer = nullptr;
	IOPCItemMgt* pIOPCItemMgt = nullptr;
	OPCHANDLE hClientGroup = 0;
	OPCHANDLE hServerGroup = 0;
	IConnectionPoint* pIConnectionPoint = nullptr;
	SOCDataCallback* pSOCDataCallback = nullptr;
	DWORD dwCookie = 0;

	std::vector<OPCHANDLE> hServerItems;
	std::vector<double> ItemsValue;
	int _numItems = 0;

public:
	static HRESULT StartupCOM();
	static void ReleaseCOM();

	~OPCClient();

	void InstantiateServer();
	void AddGroup(LPCWSTR name);
	int AddItem(wchar_t item_id[]);
	float SyncReadItem(int item_num_id);
	void SyncWriteItem(int item_num_id, double value);

	// Async Read
	void StartupASyncRead();
	void SaveASyncReadItem(int item_num_id, double value);
	double GetASyncReadItem(int item_num_id);
	void CancelASyncRead();
};

