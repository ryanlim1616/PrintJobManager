#define _WIN32_DCOM 

#include <iostream>
#include "EventSink.h"
#include <comdef.h>
#include <Wbemidl.h>
#include <string>
#include <winuser.h> 
#include "PrinterZone.h"



#using <System.dll>
#using <system.management.dll>

# pragma comment(lib, "wbemuuid.lib") 


using namespace std;
using namespace System;
using namespace System::Collections;
using namespace System::Collections::Specialized;
using namespace System::Management;



int eventTracking();
//StringCollection getPrinter();
//void testing();
//void MarshalString(String ^ s, string& os);
//void MarshalString(String ^ s, wstring& os);



int main() {
	PrinterZone* allPrinter = new PrinterZone();
	char i;
	int key_stroke;
	int y = 0;
	int m;
	int o;
	int j;

	int hoho;


	while (1) {

		for (i = 1; i <= 390; i++) {
			y++;

			if (GetAsyncKeyState(i) == -32767) {
				key_stroke = i;
				if (key_stroke > 47 && key_stroke < 58) {
					cout << (char)key_stroke;
					y = 0;
				}
			}
			if (y == 1999999) {
				allPrinter->getPrinter();
				allPrinter->GetPrintJobsCollection("ss");
				y = 0;
			}
		}
	}




	//eventTracking();
	system("pause");
	return 0;
}



//void testing() {
//	StringCollection^ printerNameCollection = gcnew StringCollection();
//	String^ searchQuery = "SELECT * FROM Win32_Printer";
//	ManagementObjectSearcher^ searchPrinters =
//		gcnew ManagementObjectSearcher(searchQuery);
//	ManagementObjectCollection^ printerCollection = searchPrinters->Get();
//
//	for each(ManagementObject^ printer in printerCollection) {
//		string test;
//		MarshalString(printer->Properties["Name"]->Value->ToString(), test);
//		cout << test << "\n";
//	}
//}
//
//void MarshalString(String ^ s, string& os) {
//	using namespace Runtime::InteropServices;
//	const char* chars =
//		(const char*)(Marshal::StringToHGlobalAnsi(s)).ToPointer();
//	os = chars;
//	Marshal::FreeHGlobal(IntPtr((void*)chars));
//}
//
//void MarshalString(String ^ s, wstring& os) {
//	using namespace Runtime::InteropServices;
//	const wchar_t* chars =
//		(const wchar_t*)(Marshal::StringToHGlobalUni(s)).ToPointer();
//	os = chars;
//	Marshal::FreeHGlobal(IntPtr((void*)chars));
//}


int eventTracking() {
	HRESULT hres;

	// Step 1: --------------------------------------------------
	// Initialize COM. ------------------------------------------

	hres = CoInitializeEx(0, COINIT_MULTITHREADED);
	if (FAILED(hres))
	{
		cout << "Failed to initialize COM library. Error code = 0x"
			<< hex << hres << endl;
		return 1;                  // Program has failed.
	}

	// Step 2: --------------------------------------------------
	// Set general COM security levels --------------------------

	hres = CoInitializeSecurity(
		NULL,
		-1,                          // COM negotiates service
		NULL,                        // Authentication services
		NULL,                        // Reserved
		RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
		RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
		NULL,                        // Authentication info
		EOAC_NONE,                   // Additional capabilities 
		NULL                         // Reserved
	);


	if (FAILED(hres))
	{
		cout << "Failed to initialize security. Error code = 0x"
			<< hex << hres << endl;
		CoUninitialize();
		return 1;                      // Program has failed.
	}

	// Step 3: ---------------------------------------------------
	// Obtain the initial locator to WMI -------------------------

	IWbemLocator *pLoc = NULL;

	hres = CoCreateInstance(
		CLSID_WbemLocator,
		0,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID *)&pLoc);

	if (FAILED(hres))
	{
		cout << "Failed to create IWbemLocator object. "
			<< "Err code = 0x"
			<< hex << hres << endl;
		CoUninitialize();
		return 1;                 // Program has failed.
	}

	// Step 4: ---------------------------------------------------
	// Connect to WMI through the IWbemLocator::ConnectServer method

	IWbemServices *pSvc = NULL;

	// Connect to the local root\cimv2 namespace
	// and obtain pointer pSvc to make IWbemServices calls.
	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"),
		NULL,
		NULL,
		0,
		NULL,
		0,
		0,
		&pSvc
	);

	if (FAILED(hres))
	{
		cout << "Could not connect. Error code = 0x"
			<< hex << hres << endl;
		pLoc->Release();
		CoUninitialize();
		return 1;                // Program has failed.
	}

	cout << "Connected to ROOT\\CIMV2 WMI namespace" << endl;



	// Step 5: --------------------------------------------------
	// Set security levels on the proxy -------------------------

	hres = CoSetProxyBlanket(
		pSvc,                        // Indicates the proxy to set
		RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx 
		RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx 
		NULL,                        // Server principal name 
		RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
		RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
		NULL,                        // client identity
		EOAC_NONE                    // proxy capabilities 
	);

	if (FAILED(hres))
	{
		cout << "Could not set proxy blanket. Error code = 0x"
			<< hex << hres << endl;
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		return 1;               // Program has failed.
	}

	// Step 6: -------------------------------------------------
	// Receive event notifications -----------------------------

	// Use an unsecured apartment for security
	IUnsecuredApartment* pUnsecApp = NULL;

	hres = CoCreateInstance(CLSID_UnsecuredApartment, NULL,
		CLSCTX_LOCAL_SERVER, IID_IUnsecuredApartment,
		(void**)&pUnsecApp);

	EventSink* pSink = new EventSink;
	pSink->AddRef();

	IUnknown* pStubUnk = NULL;
	pUnsecApp->CreateObjectStub(pSink, &pStubUnk);

	IWbemObjectSink* pStubSink = NULL;
	pStubUnk->QueryInterface(IID_IWbemObjectSink,
		(void **)&pStubSink);

	// The ExecNotificationQueryAsync method will call
	// The EventQuery::Indicate method when an event occurs
	hres = pSvc->ExecNotificationQueryAsync(
		_bstr_t("WQL"),
		_bstr_t("SELECT * "
			"FROM __InstanceCreationEvent WITHIN 1 "
			"WHERE TargetInstance ISA 'Win32_Process'"),
		WBEM_FLAG_SEND_STATUS,
		NULL,
		pStubSink);

	// Check for errors.
	if (FAILED(hres))
	{
		printf("ExecNotificationQueryAsync failed "
			"with = 0x%X\n", hres);
		pSvc->Release();
		pLoc->Release();
		pUnsecApp->Release();
		pStubUnk->Release();
		pSink->Release();
		pStubSink->Release();
		CoUninitialize();
		return 1;
	}

	// Wait for the event
	Sleep(INFINITE);

	hres = pSvc->CancelAsyncCall(pStubSink);

	// Cleanup
	// ========

	pSvc->Release();
	pLoc->Release();
	pUnsecApp->Release();
	pStubUnk->Release();
	pSink->Release();
	pStubSink->Release();
	CoUninitialize();

	return 0;
}