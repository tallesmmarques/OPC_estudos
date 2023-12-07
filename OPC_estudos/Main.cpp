#include <atlbase.h>
#include <iostream>
#include <objidl.h>

#include "opcda.h"
#include "opcerror.h"
#include "Main.h"

using namespace std;

#define OPC_SERVER_NAME L"Matrikon.OPC.Simulation.1"
#define VT VT_R4

wchar_t SAW_ID[] = L"Saw-toothed Waves.Real4";
wchar_t TRIANGLE_ID[] = L"Triangle Waves.Rea14";
wchar_t RANDOM_ID[] = L"Random.Real4";

int main(void)
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

	printf("Instantiating the MATRIKON OPC Server for Simulation...\n");
	pIOPCServer = InstantiateServer(OPC_SERVER_NAME);

	printf("Adding a group in the INACTIVE state for the moment...\n");
	AddTheGroup(pIOPCServer, pIOPCItemMgt, hServerGroup);

	size_t m;
	// Saw-toothed Waves
	wcstombs_s(&m, buf, 100, SAW_ID, _TRUNCATE);
	printf("Adding the item %s to the group...\n", buf);
    AddTheItem(pIOPCItemMgt, hServerItemSaw, SAW_ID);

	// Triangle Waves
	wcstombs_s(&m, buf, 100, TRIANGLE_ID, _TRUNCATE);
	printf("Adding the item %s to the group...\n", buf);
    AddTheItem(pIOPCItemMgt, hServerItemTriangle, TRIANGLE_ID);

	// Random
	wcstombs_s(&m, buf, 100, RANDOM_ID, _TRUNCATE);
	printf("Adding the item %s to the group...\n", buf);
    AddTheItem(pIOPCItemMgt, hServerItemRandom, RANDOM_ID);

	VARIANT varValue1, varValue2, varValue3; 
	VariantInit(&varValue1);
	VariantInit(&varValue2);
	VariantInit(&varValue3);

	printf ("Reading synchronously during 10 seconds...\n\n");
	for (i=0; i<10; i++) {
		ReadItem(pIOPCItemMgt, hServerItemSaw, varValue1);
		ReadItem(pIOPCItemMgt, hServerItemTriangle, varValue2);
		ReadItem(pIOPCItemMgt, hServerItemRandom, varValue3);
		printf("Read values -> saw: %6.2f | triangle: %6.2f | random: %6.2f\n", 
			varValue1.fltVal, varValue2.fltVal, varValue3.fltVal);

		Sleep(1000);
	}
	printf("\n");


	printf("Removing the OPC itens...\n");
	RemoveItem(pIOPCItemMgt, hServerItemSaw);
	RemoveItem(pIOPCItemMgt, hServerItemTriangle);
	RemoveItem(pIOPCItemMgt, hServerItemRandom);

	printf("Removing the OPC group object...\n");
    pIOPCItemMgt->Release();
	RemoveGroup(pIOPCServer, hServerGroup);

	printf("Removing the OPC server object...\n");
	pIOPCServer->Release();

	printf ("Releasing the COM environment...\n");
	CoUninitialize();

	return 0;
}

IOPCServer* InstantiateServer(wchar_t ServerName[])
{
	CLSID CLSID_OPCServer;
	HRESULT hr;

	// get the CLSID from the OPC Server Name:
	hr = CLSIDFromString(ServerName, &CLSID_OPCServer);
	_ASSERT(!FAILED(hr));


	//queue of the class instances to create
	LONG cmq = 1; // nbr of class instance to create.
	MULTI_QI queue[1] =
		{{&IID_IOPCServer,
		NULL,
		0}};

	//Server info:
	//COSERVERINFO CoServerInfo =
    //{
	//	/*dwReserved1*/ 0,
	//	/*pwszName*/ REMOTE_SERVER_NAME,
	//	/*COAUTHINFO*/  NULL,
	//	/*dwReserved2*/ 0
    //}; 

	// create an instance of the IOPCServer
	hr = CoCreateInstanceEx(CLSID_OPCServer, NULL, CLSCTX_SERVER,
		/*&CoServerInfo*/NULL, cmq, queue);
	_ASSERT(!hr);

	// return a pointer to the IOPCServer interface:
	return(IOPCServer*) queue[0].pItf;
}

void AddTheGroup(IOPCServer* pIOPCServer, IOPCItemMgt* &pIOPCItemMgt, 
				 OPCHANDLE& hServerGroup)
{
	DWORD dwUpdateRate = 0;
	OPCHANDLE hClientGroup = 0;

	// Add an OPC group and get a pointer to the IUnknown I/F:
    HRESULT hr = pIOPCServer->AddGroup(/*szName*/ L"Group1",
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
}
	
void AddTheItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE& hServerItem, wchar_t ITEM_ID[])
{
	HRESULT hr;

	// Array of items to add:
	OPCITEMDEF ItemArray[1] =
	{{
	/*szAccessPath*/ L"",
	/*szItemID*/ ITEM_ID,
	/*bActive*/ TRUE,
	/*hClient*/ 1,
	/*dwBlobSize*/ 0,
	/*pBlob*/ NULL,
	/*vtRequestedDataType*/ VT,
	/*wReserved*/0
	}};

	//Add Result:
	OPCITEMRESULT* pAddResult=NULL;
	HRESULT* pErrors = NULL;

	// Add an Item to the previous Group:
	hr = pIOPCItemMgt->AddItems(1, ItemArray, &pAddResult, &pErrors);
	if (hr != S_OK){
		printf("Failed call to AddItems function. Error code = %x\n", hr);
		exit(0);
	}

	// Server handle for the added item:
	hServerItem = pAddResult[0].hServer;

	// release memory allocated by the server:
	CoTaskMemFree(pAddResult->pBlob);

	CoTaskMemFree(pAddResult);
	pAddResult = NULL;

	CoTaskMemFree(pErrors);
	pErrors = NULL;
}

void ReadItem(IUnknown* pGroupIUnknown, OPCHANDLE hServerItem, VARIANT& varValue)
{
	// value of the item:
	OPCITEMSTATE* pValue = NULL;

	//get a pointer to the IOPCSyncIOInterface:
	IOPCSyncIO* pIOPCSyncIO;
	pGroupIUnknown->QueryInterface(__uuidof(pIOPCSyncIO), (void**) &pIOPCSyncIO);

	// read the item value from the device:
	HRESULT* pErrors = NULL; //to store error code(s)
	HRESULT hr = pIOPCSyncIO->Read(OPC_DS_DEVICE, 1, &hServerItem, &pValue, &pErrors);
	_ASSERT(!hr);
	_ASSERT(pValue!=NULL);

	varValue = pValue[0].vDataValue;

	//Release memeory allocated by the OPC server:
	CoTaskMemFree(pErrors);
	pErrors = NULL;

	CoTaskMemFree(pValue);
	pValue = NULL;

	// release the reference to the IOPCSyncIO interface:
	pIOPCSyncIO->Release();
}

void RemoveItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE hServerItem)
{
	// server handle of items to remove:
	OPCHANDLE hServerArray[1];
	hServerArray[0] = hServerItem;
	
	//Remove the item:
	HRESULT* pErrors; // to store error code(s)
	HRESULT hr = pIOPCItemMgt->RemoveItems(1, hServerArray, &pErrors);
	_ASSERT(!hr);

	//release memory allocated by the server:
	CoTaskMemFree(pErrors);
	pErrors = NULL;
}

void RemoveGroup (IOPCServer* pIOPCServer, OPCHANDLE hServerGroup)
{
	// Remove the group:
	HRESULT hr = pIOPCServer->RemoveGroup(hServerGroup, FALSE);
	if (hr != S_OK){
		if (hr == OPC_S_INUSE)
			printf ("Failed to remove OPC group: object still has references to it.\n");
		else printf ("Failed to remove OPC group. Error code = %x\n", hr);
		exit(0);
	}
}

