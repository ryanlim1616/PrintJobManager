#include "PrinterZone.h"

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

void PrinterZone::getPrinter() {
	StringCollection^ printerNameCollection = gcnew StringCollection();
	String^ searchQuery = "SELECT * FROM Win32_Printer";
	ManagementObjectSearcher^ searchPrinters =
		gcnew ManagementObjectSearcher(searchQuery);
	ManagementObjectCollection^ printerCollection = searchPrinters->Get();

	for each(ManagementObject^ printer in printerCollection) {
		string test;
		MarshalString(printer->Properties["Name"]->Value->ToString(), test);
		cout << test << "\n";
	}
}

void PrinterZone::MarshalString(String ^ s, string& os) {
	using namespace Runtime::InteropServices;
	const char* chars =
		(const char*)(Marshal::StringToHGlobalAnsi(s)).ToPointer();
	os = chars;
	Marshal::FreeHGlobal(IntPtr((void*)chars));
}

void PrinterZone::MarshalString(String ^ s, wstring& os) {
	using namespace Runtime::InteropServices;
	const wchar_t* chars =
		(const wchar_t*)(Marshal::StringToHGlobalUni(s)).ToPointer();
	os = chars;
	Marshal::FreeHGlobal(IntPtr((void*)chars));
}



void PrinterZone::GetPrintJobsCollection(string printerName) {

	StringCollection^ printJobCollection = gcnew StringCollection();
	String^ searchQuery = "SELECT * FROM Win32_PrintJob";

	/*searchQuery can also be mentioned with where Attribute,
	but this is not working in Windows 2000 / ME / 98 machines
	and throws Invalid query error*/
	ManagementObjectSearcher^ searchPrintJobs =
		gcnew ManagementObjectSearcher(searchQuery);
	ManagementObjectCollection^ prntJobCollection = searchPrintJobs->Get();


	for each(ManagementObject^ prntJob in prntJobCollection)
	{
		string test;
		MarshalString(prntJob->Properties["Document"]->Value->ToString(), test);
		cout << test << "\n";
		//System.String jobName = prntJob.Properties["Name"].Value.ToString();

		////Job name would be of the format [Printer name], [Job ID]
		//char[] splitArr = new char[1];
		//splitArr[0] = Convert.ToChar(",");
		//string prnterName = jobName.Split(splitArr)[0];
		//string documentName = prntJob.Properties["Document"].Value.ToString();
		//if (String.Compare(prnterName, printerName, true) == 0)
		//{
		//	printJobCollection.Add(documentName);
		//}
	}
	//return printJobCollection;
}