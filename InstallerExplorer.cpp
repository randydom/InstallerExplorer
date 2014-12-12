// InstallerExplorer.cpp : main project file.

#include "stdafx.h"
#include "ProductsForm.h"

using namespace InstallerExplorer;

[STAThreadAttribute]
int main(array<System::String ^> ^args)
{
	// Enabling Windows XP visual effects before any controls are created
	Application::EnableVisualStyles();
	Application::SetCompatibleTextRenderingDefault(false); 

	// Create the main window and run it
	Application::Run(gcnew ProductsForm());
	return 0;
}
