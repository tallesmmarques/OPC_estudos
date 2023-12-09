#include <atlbase.h>
#include <iostream>
#include <objidl.h>

#include "opcda.h"
#include "opcerror.h"
#include "Main.h"
#include "SOCAdviseSink.h"
#include "SOCDataCallback.h"
#include "SOCWrapperFunctions.h"
#include "OPCClient.h"

using namespace std;

#define OPC_SERVER_NAME L"Matrikon.OPC.Simulation.1"
#define VT VT_R4

wchar_t SAW_ID[] = L"Saw-toothed Waves.Real4";
wchar_t TRIANGLE_ID[] = L"Triangle Waves.Rea14";
wchar_t RANDOM_ID[] = L"Random.Real4";
wchar_t BUCKET_ID[] = L"Bucket Brigade.Real4";

UINT OPC_DATA_TIME = RegisterClipboardFormat (_T("OPCSTMFORMATDATATIME"));

int main(void)
{
	HRESULT hr;

	printf("Inicializando ambiente COM...\n");
	hr = OPCClient::StartupCOM();
	_ASSERT(!FAILED(hr));

	OPCClient* pOPCClient = new OPCClient;

	printf("Instanciando servidor...\n");
	pOPCClient->InstantiateServer();

	printf("Adicionando grupo ao servidor...\n");
	pOPCClient->AddGroup(L"Group1");

	printf("Adicionando itens ao grupo...\n");
	int RandomItemID = pOPCClient->AddItem(RANDOM_ID);
	printf("ID do item %ls: %d\n", RANDOM_ID, RandomItemID);
	int TriandleItemID = pOPCClient->AddItem(TRIANGLE_ID);
	printf("ID do item %ls: %d\n", TRIANGLE_ID, TriandleItemID);
	int BucketItemID = pOPCClient->AddItem(BUCKET_ID);
	printf("ID do item %ls: %d\n", BUCKET_ID, BucketItemID);
	printf("\n");

	

	//printf("Valor lido de Random: %f\n", pOPCClient->SyncReadItem(RandomItemID));
	//Sleep(1000);
	//printf("Valor lido de Random: %f\n", pOPCClient->SyncReadItem(RandomItemID));
	//Sleep(1000);
	//printf("Valor lido de Random: %f\n", pOPCClient->SyncReadItem(RandomItemID));
	//Sleep(2000);
	//printf("\n");


	//printf("Valor lido de Bucket: %f\n", pOPCClient->SyncReadItem(BucketItemID));
	//Sleep(1000);
	//pOPCClient->SyncWriteItem(BucketItemID, 100);
	//Sleep(1000);
	//printf("Valor lido de Bucket: %f\n", pOPCClient->SyncReadItem(BucketItemID));
	//Sleep(1000);
	//pOPCClient->SyncWriteItem(BucketItemID, 50);
	//Sleep(1000);
	//printf("Valor lido de Bucket: %f\n", pOPCClient->SyncReadItem(BucketItemID));
	//printf("\n");

 
	pOPCClient->StartupASyncRead();

	int bRet;
	MSG msg;
	DWORD ticks1, ticks2;
	ticks1 = GetTickCount();
	printf("Aguardando notificacoes...\n");
	do {
		bRet = GetMessage( &msg, NULL, 0, 0 );
		if (!bRet){
			printf ("Failed to get windows message! Error code = %d\n", GetLastError());
			exit(0);
		}
		DispatchMessage(&msg);

		printf("R:%f T:%f B:%f\n",
			pOPCClient->GetASyncReadItem(RandomItemID),
			pOPCClient->GetASyncReadItem(TriandleItemID),
			pOPCClient->GetASyncReadItem(BucketItemID));

        ticks2 = GetTickCount();
	}
	while ((ticks2 - ticks1) < 10000);

	pOPCClient->CancelASyncRead();


	printf("Finalizando instancias...\n");
	delete pOPCClient;

	printf("Finalizando ambiente COM...\n");
	OPCClient::ReleaseCOM();

	return EXIT_SUCCESS;
}

int main_bkp(void)
{
	IOPCServer* pIOPCServer = nullptr;
	IOPCItemMgt* pIOPCItemMgt = nullptr;
	
	OPCHANDLE hServerGroup;
	OPCHANDLE hServerItemSaw;
	OPCHANDLE hServerItemTriangle;
	OPCHANDLE hServerItemRandom;

	int i;
	char buf[100];

	printf("Initializing the COM environment...\n");
	CoInitialize(NULL);

	/* ---------------------------------------------------------------------------------- */
	printf("Instantiating the MATRIKON OPC Server for Simulation...\n");
	CLSID CLSID_OPCServer;
	HRESULT hr;

	// get the CLSID from the OPC Server Name:
	hr = CLSIDFromString(OPC_SERVER_NAME, &CLSID_OPCServer);
	_ASSERT(!FAILED(hr));

	//queue of the class instances to create
	LONG cmq = 1; // nbr of class instance to create.
	MULTI_QI queue[1] = {{&IID_IOPCServer, NULL, 0}};

	// create an instance of the IOPCServer
	hr = CoCreateInstanceEx(CLSID_OPCServer, NULL, CLSCTX_SERVER, NULL, cmq, queue);
	_ASSERT(!hr);

	pIOPCServer = (IOPCServer*) queue[0].pItf;

	/* ---------------------------------------------------------------------------------- */
	printf("Adding a group in the INACTIVE state for the moment...\n");
	//AddTheGroup(pIOPCServer, pIOPCItemMgt, hServerGroup);
	DWORD dwUpdateRate = 0;
	OPCHANDLE hClientGroup = 0;

	// Add an OPC group and get a pointer to the IUnknown I/F:
    hr = pIOPCServer->AddGroup(
		/*szName*/ L"Group1",
		/*bActive*/ FALSE,
		/*dwRequestedUpdateRate*/ 1000,
		/*hClientGroup*/ hClientGroup,
		/*pTimeBias*/ 0,
		/*pPercentDeadband*/ 0,
		/*dwLCID*/0,
		/*phServerGroup*/&hServerGroup,
		&dwUpdateRate,
		/*riid*/ IID_IOPCItemMgt,
		/*ppUnk*/ (IUnknown**) &pIOPCItemMgt);
	_ASSERT(!FAILED(hr));

	/* ---------------------------------------------------------------------------------- */
	size_t m;
	// Saw-toothed Waves
	wcstombs_s(&m, buf, 100, SAW_ID, _TRUNCATE);
	printf("Adding the item %s to the group...\n", buf);

	// Triangle Waves
	wcstombs_s(&m, buf, 100, TRIANGLE_ID, _TRUNCATE);
	printf("Adding the item %s to the group...\n", buf);

	// Random
	wcstombs_s(&m, buf, 100, RANDOM_ID, _TRUNCATE);
	printf("Adding the item %s to the group...\n", buf);

	// Array of items to add:
	OPCITEMDEF ItemArray[3] =
	{{
		/*szAccessPath*/ L"",
		/*szItemID*/ SAW_ID,
		/*bActive*/ TRUE,
		/*hClient*/ 1,
		/*dwBlobSize*/ 0,
		/*pBlob*/ NULL,
		/*vtRequestedDataType*/ VT,
		/*wReserved*/0
	},{
		/*szAccessPath*/ L"",
		/*szItemID*/ TRIANGLE_ID,
		/*bActive*/ TRUE,
		/*hClient*/ 1,
		/*dwBlobSize*/ 0,
		/*pBlob*/ NULL,
		/*vtRequestedDataType*/ VT,
		/*wReserved*/0
	},{
		/*szAccessPath*/ L"",
		/*szItemID*/ RANDOM_ID,
		/*bActive*/ TRUE,
		/*hClient*/ 1,
		/*dwBlobSize*/ 0,
		/*pBlob*/ NULL,
		/*vtRequestedDataType*/ VT,
		/*wReserved*/0
	}};

	// Add Result:
	OPCITEMRESULT* pAddResult=NULL;
	HRESULT* pErrors = NULL;

	// Add an Item to the previous Group:
	hr = pIOPCItemMgt->AddItems(3, ItemArray, &pAddResult, &pErrors);
	if (hr != S_OK){
		printf("Failed call to AddItems function. Error code = %x\n", hr);
		exit(0);
	}

	// Server handle for the added item:
	hServerItemSaw = pAddResult[0].hServer;
	hServerItemTriangle = pAddResult[1].hServer;
	hServerItemRandom = pAddResult[2].hServer;

	// release memory allocated by the server:
	CoTaskMemFree(pAddResult->pBlob);

	CoTaskMemFree(pAddResult);
	pAddResult = NULL;

	CoTaskMemFree(pErrors);
	pErrors = NULL;

	/* ---------------------------------------------------------------------------------- */

	VARIANT varValue1, varValue2, varValue3; 
	VariantInit(&varValue1);
	VariantInit(&varValue2);
	VariantInit(&varValue3);

	printf ("Reading synchronously during 10 seconds...\n\n");
	OPCHANDLE hServerItems[] = { hServerItemSaw, hServerItemTriangle, hServerItemRandom };
	for (i=0; i<10; i++) {
		//ReadItem(pIOPCItemMgt, hServerItemSaw, varValue1);
		// value of the item:
		OPCITEMSTATE* pValue = NULL;
		IOPCSyncIO* pIOPCSyncIO;
		HRESULT* pErrors = NULL; //to store error code(s)

		//get a pointer to the IOPCSyncIOInterface:
		pIOPCItemMgt->QueryInterface(__uuidof(pIOPCSyncIO), (void**) &pIOPCSyncIO);

		// read the item value from the device:
		hr = pIOPCSyncIO->Read(OPC_DS_DEVICE, 3, hServerItems, &pValue, &pErrors);
		_ASSERT(!hr);
		_ASSERT(pValue!=NULL);

		varValue1 = pValue[0].vDataValue;
		varValue2 = pValue[1].vDataValue;
		varValue3 = pValue[2].vDataValue;

		//Release memeory allocated by the OPC server:
		CoTaskMemFree(pErrors);
		pErrors = NULL;

		CoTaskMemFree(pValue);
		pValue = NULL;

		// release the reference to the IOPCSyncIO interface:
		pIOPCSyncIO->Release();

		printf("Read values -> saw: %6.2f | triangle: %6.2f | random: %6.2f\n", 
			varValue1.fltVal, varValue2.fltVal, varValue3.fltVal);

		Sleep(1000);
	}
	printf("\n");

	/* ---------------------------------------------------------------------------------- */

	IConnectionPoint* pIConnectionPoint = NULL; //pointer to IConnectionPoint Interface
	DWORD dwCookie = 0;
	SOCDataCallback* pSOCDataCallback = new SOCDataCallback();
	pSOCDataCallback->AddRef();

	printf("Setting up the IConnectionPoint callback connection...\n");
	SetDataCallback(pIOPCItemMgt, pSOCDataCallback, pIConnectionPoint, &dwCookie);

	// Change the group to the ACTIVE state so that we can receive the
	// server´s callback notification
	printf("Changing the group state to ACTIVE...\n");
    SetGroupActive(pIOPCItemMgt); 

	// Enter again a message pump in order to process the server´s callback
	// notifications, for the same reason explained before.
		
	int bRet;
	MSG msg;
	DWORD ticks1, ticks2;

	ticks1 = GetTickCount();
	printf("Waiting for IOPCDataCallback notifications during 10 seconds...\n");
	do {
		bRet = GetMessage( &msg, NULL, 0, 0 );
		if (!bRet){
			printf ("Failed to get windows message! Error code = %d\n", GetLastError());
			exit(0);
		}
		TranslateMessage(&msg); // This call is not really needed ...
		DispatchMessage(&msg);  // ... but this one is!
        ticks2 = GetTickCount();
	}
	while ((ticks2 - ticks1) < 10000);

	// Cancel the callback and release its reference
	printf("Cancelling the IOPCDataCallback notifications...\n");
    CancelDataCallback(pIConnectionPoint, dwCookie);
	//pIConnectionPoint->Release();
	pSOCDataCallback->Release();


	/* ---------------------------------------------------------------------------------- */
	printf("Removing the OPC itens...\n");
	hr = pIOPCItemMgt->RemoveItems(3, hServerItems, &pErrors);
	_ASSERT(!hr);

	//release memory allocated by the server:
	CoTaskMemFree(pErrors);
	pErrors = NULL;

	/* ---------------------------------------------------------------------------------- */
	printf("Removing the OPC group object...\n");
    pIOPCItemMgt->Release();

	/* ---------------------------------------------------------------------------------- */
	printf("Removing the OPC server object...\n");
	hr = pIOPCServer->RemoveGroup(hServerGroup, FALSE);
	if (hr != S_OK){
		if (hr == OPC_S_INUSE)
			printf ("Failed to remove OPC group: object still has references to it.\n");
		else printf ("Failed to remove OPC group. Error code = %x\n", hr);
		exit(0);
	}
	pIOPCServer->Release();

	/* ---------------------------------------------------------------------------------- */
	printf ("Releasing the COM environment...\n");
	CoUninitialize();

	return 0;
}


