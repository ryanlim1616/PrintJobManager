#ifndef PrinterZone_H
#define PrinterZone_H

#include <iostream>
#include <comdef.h>
#include <Wbemidl.h>
#include <string>


#using <System.dll>
#using <system.management.dll>

# pragma comment(lib, "wbemuuid.lib") 


using namespace std;
using namespace System;
using namespace System::Collections;
using namespace System::Collections::Specialized;
using namespace System::Management;

using namespace std;

class PrinterZone {
private:
	int a = 0;
public:
	PrinterZone() { };
	void getPrinter();
	void MarshalString(String ^ s, string& os);
	void MarshalString(String ^ s, wstring& os);
	void GetPrintJobsCollection(string printerName);
};

#endif