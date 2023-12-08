#pragma once

#include <objidl.h>
#include <vector>

#include "opcda.h"

#define OPC_SERVER_NAME L"Matrikon.OPC.Simulation.1"

class OPCClient
{
private:
	IOPCServer* pIOPCServer = nullptr;
	IOPCItemMgt* pIOPCItemMgt = nullptr;
	OPCHANDLE hClientGroup = 0;
	OPCHANDLE hServerGroup = 0;

	std::vector<OPCHANDLE> hServerItems;
	int _numItems = 0;

public:
	static HRESULT StartupCOM();
	static void ReleaseCOM();

	~OPCClient();

	void InstantiateServer();
	void AddGroup(LPCWSTR name);
	int AddItem(wchar_t item_id[]);
	float SyncReadItem(int item_num_id);
	void SyncWriteItem(int item_num_id, float value);
};

