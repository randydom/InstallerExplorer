#pragma once

using namespace System;
using namespace System::Windows::Forms;

namespace InstallerExplorer {

	/// <summary>
	/// Dispaly the components of the given product
	/// </summary>
	public ref class ComponentsForm : public Form {
	protected:
		InstallerExplorer::ListUtils* listUtils;

	public:
		ComponentsForm(String^ productID) {
			listUtils = new InstallerExplorer::ListUtils();
			InitializeComponent();

			WCHAR wszProduct[39] = {0};
			pin_ptr<const wchar_t> wchProductID = PtrToStringChars(productID);
			wcscpy_s(wszProduct, wchProductID);

			this->SuspendLayout();
			EnumComponents(wszProduct);
			this->ResumeLayout();
		}

	protected:
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~ComponentsForm() {
			if (components)
				delete components;
			if(listUtils)
				delete listUtils;
		}


#pragma region Event Handlers

		System::Void ComponentsForm_KeyDown(System::Object^ sender, KeyEventArgs^ e) {
			listUtils->DoKeyDown(e, this->listView1, 2, L"Component ID\tComponent Path");
		}

		System::Void listView1_ColumnClick(System::Object^ sender, ColumnClickEventArgs^ e) {
			listUtils->DoColumnClick(e->Column, this->listView1);
		}

#pragma endregion

#pragma region Enumeration Code

		void EnumComponents(WCHAR* pszProductCode) {
			WCHAR wszComponent[39] = {0};
			WCHAR wszOwningProduct[39] = {0};
			int i     = 0;
			int count = 0;
			HRESULT hr = S_OK;

			do {
				hr = MsiEnumComponents (i++, wszComponent);
				if(S_OK == hr) {
					int ii = 0;
					do {
						hr = MsiEnumClients(wszComponent, ii++, wszOwningProduct);
						if(S_OK == hr) {
							if(0 == wcscmp(pszProductCode, wszOwningProduct)) {
								count++;
								AddComponent(pszProductCode, wszComponent);
							}
						}
					} while(S_OK == hr);
					hr = S_OK;
				}
			} while(S_OK == hr);

			WCHAR wcsCount[32] = {0};
			_itow_s(count, wcsCount, 10);
			this->lblCount->Text = gcnew String(wcsCount);
		}

		void AddComponent(WCHAR* pwsProductCode, WCHAR* pwsComponentCode) {
			DWORD dwCount = 2048;
			WCHAR* lpPath = new WCHAR[2049];

			HRESULT hr = MsiGetComponentPath (pwsProductCode, pwsComponentCode, lpPath, &dwCount);
			if(SUCCEEDED(hr)) {
				ListViewItem^  listViewItem1 = (gcnew ListViewItem(gcnew cli::array< System::String^  >(2) {gcnew String(pwsComponentCode), gcnew String(lpPath)}, -1));
				this->listView1->Items->AddRange(gcnew cli::array<ListViewItem^  >(1) {listViewItem1});
			}
		}

#pragma endregion

#pragma region Form Initialization

	private: System::Windows::Forms::Label^  lblCount;
	protected: 
	private: System::Windows::Forms::Label^  label1;
	private: System::Windows::Forms::Button^  btnOK;
	private: System::Windows::Forms::ListView^  listView1;
	private: System::Windows::Forms::ColumnHeader^  columnHeader1;
	private: System::Windows::Forms::ColumnHeader^  columnHeader2;



	private:
		/// <summary>
		/// Required designer variable.
		/// </summary>
		System::ComponentModel::Container ^components;
#pragma endregion

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		void InitializeComponent(void)
		{
			this->lblCount = (gcnew System::Windows::Forms::Label());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->btnOK = (gcnew System::Windows::Forms::Button());
			this->listView1 = (gcnew System::Windows::Forms::ListView());
			this->columnHeader1 = (gcnew System::Windows::Forms::ColumnHeader());
			this->columnHeader2 = (gcnew System::Windows::Forms::ColumnHeader());
			this->SuspendLayout();
			// 
			// lblCount
			// 
			this->lblCount->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((System::Windows::Forms::AnchorStyles::Bottom | System::Windows::Forms::AnchorStyles::Left));
			this->lblCount->AutoSize = true;
			this->lblCount->Location = System::Drawing::Point(57, 226);
			this->lblCount->Name = L"lblCount";
			this->lblCount->Size = System::Drawing::Size(13, 13);
			this->lblCount->TabIndex = 7;
			this->lblCount->Text = L"0";
			// 
			// label1
			// 
			this->label1->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((System::Windows::Forms::AnchorStyles::Bottom | System::Windows::Forms::AnchorStyles::Left));
			this->label1->AutoSize = true;
			this->label1->Location = System::Drawing::Point(12, 226);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(38, 13);
			this->label1->TabIndex = 6;
			this->label1->Text = L"Count:";
			// 
			// btnOK
			// 
			this->btnOK->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((System::Windows::Forms::AnchorStyles::Bottom | System::Windows::Forms::AnchorStyles::Right));
			this->btnOK->DialogResult = System::Windows::Forms::DialogResult::OK;
			this->btnOK->Location = System::Drawing::Point(506, 217);
			this->btnOK->Name = L"btnOK";
			this->btnOK->Size = System::Drawing::Size(75, 23);
			this->btnOK->TabIndex = 5;
			this->btnOK->Text = L"OK";
			this->btnOK->UseVisualStyleBackColor = true;
			// 
			// listView1
			// 
			this->listView1->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((((System::Windows::Forms::AnchorStyles::Top | System::Windows::Forms::AnchorStyles::Bottom) 
				| System::Windows::Forms::AnchorStyles::Left) 
				| System::Windows::Forms::AnchorStyles::Right));
			this->listView1->Columns->AddRange(gcnew cli::array< System::Windows::Forms::ColumnHeader^  >(2) {this->columnHeader1, this->columnHeader2});
			this->listView1->FullRowSelect = true;
			this->listView1->GridLines = true;
			this->listView1->HideSelection = false;
			this->listView1->Location = System::Drawing::Point(12, 12);
			this->listView1->Name = L"listView1";
			this->listView1->Size = System::Drawing::Size(569, 199);
			this->listView1->Sorting = System::Windows::Forms::SortOrder::Ascending;
			this->listView1->TabIndex = 4;
			this->listView1->UseCompatibleStateImageBehavior = false;
			this->listView1->View = System::Windows::Forms::View::Details;
			this->listView1->ColumnClick += gcnew System::Windows::Forms::ColumnClickEventHandler(this, &ComponentsForm::listView1_ColumnClick);
			// 
			// columnHeader1
			// 
			this->columnHeader1->Text = L"Component ID";
			this->columnHeader1->Width = 240;
			// 
			// columnHeader2
			// 
			this->columnHeader2->Text = L"Component Path";
			this->columnHeader2->Width = 306;
			// 
			// ComponentsForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(593, 248);
			this->Controls->Add(this->lblCount);
			this->Controls->Add(this->label1);
			this->Controls->Add(this->btnOK);
			this->Controls->Add(this->listView1);
			this->KeyPreview = true;
			this->Name = L"ComponentsForm";
			this->StartPosition = System::Windows::Forms::FormStartPosition::CenterParent;
			this->Text = L"Installed Components";
			this->KeyDown += gcnew System::Windows::Forms::KeyEventHandler(this, &ComponentsForm::ComponentsForm_KeyDown);
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
};
}
