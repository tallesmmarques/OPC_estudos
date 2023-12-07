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


