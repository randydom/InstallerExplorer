#pragma once

using namespace System;
using namespace System::Windows::Forms;

namespace InstallerExplorer {

	public ref class ListViewItemComparer : public System::Collections::IComparer {
	private:
		int  col;
		bool reverse;

	public:
		ListViewItemComparer() {
			col     = 0;
			reverse = false;
		}

		ListViewItemComparer(int column) {
			col     = column;
			reverse = false;
		}

		ListViewItemComparer(int column, bool reverseSort) {
			col     = column;
			reverse = reverseSort;
		}

		virtual int Compare(System::Object^ x, System::Object^ y) {
			if(reverse)
				return String::Compare(((ListViewItem^)y)->SubItems[col]->Text, ((ListViewItem^)x)->SubItems[col]->Text);
			else
				return String::Compare(((ListViewItem^)x)->SubItems[col]->Text, ((ListViewItem^)y)->SubItems[col]->Text);
		}
	};
}