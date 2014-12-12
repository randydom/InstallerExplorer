#pragma once

#include "Comparer.h"

using namespace System;
using namespace System::Collections;
using namespace System::Text;
using namespace System::Windows::Forms;

namespace InstallerExplorer {
	public class ListUtils {
	protected:
		int  lastColumnClicked;
		bool reverseSort;

	public:
		ListUtils() {
			lastColumnClicked = 0;
			reverseSort       = false;
		}

		System::Void SelectAll(ListView^ lv) {
			for(int i = 0; i < lv->Items->Count; i++)
				lv->Items[i]->Selected = true;
		}

		System::Void Copy(ListView^ lv, int cColumns, LPWSTR wcsHeader) {
			StringBuilder^ sb = gcnew StringBuilder(gcnew String(wcsHeader));
			String^ tab = gcnew String("\t");

			sb->AppendLine();

			if(0 == lv->SelectedItems->Count)
				SelectAll(lv);

			for(int i = 0; i < lv->Items->Count; i++) {
				if(lv->Items[i]->Selected) {
					bool first = true;
					for(int j = 0; j < cColumns; j++) {
						if(!first)
							sb->Append(tab);
						first = false;
						sb->Append(lv->Items[i]->SubItems[j]->Text);
					}
					sb->AppendLine();
				}
			}
			Clipboard::Clear();
			Clipboard::SetDataObject(sb->ToString());
		}

		void DoKeyDown(KeyEventArgs^ e, ListView^ lv, int columns, LPWSTR headers) {
			if(e->Control) {
				if(e->KeyCode == Keys::C) {
					Copy(lv, columns, headers);
					e->Handled = true;
				}
				else if(e->KeyCode == Keys::A) {
					SelectAll(lv);
					e->Handled = true;
				}
			}
		}

		void DoColumnClick(int column, ListView^ lv) {
			if(column == lastColumnClicked)
				reverseSort = !reverseSort;
			else
				reverseSort = false;
			lastColumnClicked = column;
			lv->ListViewItemSorter = gcnew InstallerExplorer::ListViewItemComparer(column, reverseSort);
		}
	};
}