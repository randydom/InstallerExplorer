#pragma once

using namespace System;
using namespace System::Windows::Forms;

#include "ListUtils.h"

namespace InstallerExplorer {
	/// <summary>
	/// Display all features of the given product
	/// </summary>
	public ref class FeaturesForm : public Form {
	protected:
		InstallerExplorer::ListUtils* listUtils;

	public:
		FeaturesForm(String^ productID) {
			listUtils = new InstallerExplorer::ListUtils();
			InitializeComponent();

			this->SuspendLayout();
			EnumFeatures(productID);
			this->ResumeLayout();
		}

	protected:
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~FeaturesForm() {
			if (components)
				delete components;
			if(listUtils)
				delete listUtils;
		}

#pragma region Event Handlers

		System::Void FeaturesForm_KeyDown(System::Object^ sender, KeyEventArgs^ e) {
			listUtils->DoKeyDown(e, this->listView1, 2, L"Feature Name\tState");
		}

		System::Void listView1_ColumnClick(System::Object^ sender, ColumnClickEventArgs^ e) {
			listUtils->DoColumnClick(e->Column, this->listView1);
		}

#pragma endregion

#pragma region Enumeration Code

		void EnumFeatures(String^ productID) {
			const int cchGUID = 38;
			WCHAR wszFeature[cchGUID+1] = {0};
			WCHAR wszParent[cchGUID+1]  = {0};
			WCHAR wszProduct[cchGUID+1] = {0};

			pin_ptr<const wchar_t> wchProductID = PtrToStringChars(productID);
			wcscpy_s(wszProduct, wchProductID);

			int i = 0;
			HRESULT hr = S_OK;
			do {
				hr = MsiEnumFeatures(wszProduct, i++, wszFeature, wszParent);
				if(S_OK == hr)
					AddFeature(wszProduct, wszFeature);
			} while(S_OK == hr);

			WCHAR wcsCount[32] = {0};
			_itow_s(i-1, wcsCount, 10);
			this->lblCount->Text = gcnew String(wcsCount);
		}

		void AddFeature(WCHAR* pwsProductCode, WCHAR* pwsFeatureCode) {
			INSTALLSTATE state = MsiQueryFeatureState(pwsProductCode, pwsFeatureCode);
			String^ stateStr = gcnew String("?");
			switch(state) {
				case INSTALLSTATE_ABSENT:
					stateStr = L"Absent";
					break;
				case INSTALLSTATE_ADVERTISED:
					stateStr = "Advertised";
					break;
				case INSTALLSTATE_LOCAL:
					stateStr = "Local";
					break;
				case INSTALLSTATE_SOURCE:
					stateStr = "Source";
					break;
			}
			ListViewItem^  listViewItem1 = (gcnew ListViewItem(gcnew cli::array< System::String^  >(2) {gcnew String(pwsFeatureCode), stateStr}, -1));
			this->listView1->Items->AddRange(gcnew cli::array<ListViewItem^  >(1) {listViewItem1});
		}

#pragma endregion

#pragma region Form Initialization

	private: System::Windows::Forms::ListView^  listView1;
	private: System::Windows::Forms::Button^  btnOK;
	private: System::Windows::Forms::ColumnHeader^  columnHeader1;
	private: System::Windows::Forms::Label^  label1;
	private: System::Windows::Forms::Label^  lblCount;
	private: System::Windows::Forms::ColumnHeader^  columnHeader2;

	protected: 

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
			this->listView1 = (gcnew System::Windows::Forms::ListView());
			this->columnHeader1 = (gcnew System::Windows::Forms::ColumnHeader());
			this->columnHeader2 = (gcnew System::Windows::Forms::ColumnHeader());
			this->btnOK = (gcnew System::Windows::Forms::Button());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->lblCount = (gcnew System::Windows::Forms::Label());
			this->SuspendLayout();
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
			this->listView1->Size = System::Drawing::Size(374, 233);
			this->listView1->Sorting = System::Windows::Forms::SortOrder::Ascending;
			this->listView1->TabIndex = 0;
			this->listView1->UseCompatibleStateImageBehavior = false;
			this->listView1->View = System::Windows::Forms::View::Details;
			this->listView1->ColumnClick += gcnew System::Windows::Forms::ColumnClickEventHandler(this, &FeaturesForm::listView1_ColumnClick);
			// 
			// columnHeader1
			// 
			this->columnHeader1->Text = L"Feature Name";
			this->columnHeader1->Width = 272;
			// 
			// columnHeader2
			// 
			this->columnHeader2->Text = L"State";
			this->columnHeader2->Width = 81;
			// 
			// btnOK
			// 
			this->btnOK->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((System::Windows::Forms::AnchorStyles::Bottom | System::Windows::Forms::AnchorStyles::Right));
			this->btnOK->DialogResult = System::Windows::Forms::DialogResult::OK;
			this->btnOK->Location = System::Drawing::Point(311, 251);
			this->btnOK->Name = L"btnOK";
			this->btnOK->Size = System::Drawing::Size(75, 23);
			this->btnOK->TabIndex = 1;
			this->btnOK->Text = L"OK";
			this->btnOK->UseVisualStyleBackColor = true;
			// 
			// label1
			// 
			this->label1->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((System::Windows::Forms::AnchorStyles::Bottom | System::Windows::Forms::AnchorStyles::Left));
			this->label1->AutoSize = true;
			this->label1->Location = System::Drawing::Point(13, 260);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(38, 13);
			this->label1->TabIndex = 2;
			this->label1->Text = L"Count:";
			// 
			// lblCount
			// 
			this->lblCount->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((System::Windows::Forms::AnchorStyles::Bottom | System::Windows::Forms::AnchorStyles::Left));
			this->lblCount->AutoSize = true;
			this->lblCount->Location = System::Drawing::Point(58, 260);
			this->lblCount->Name = L"lblCount";
			this->lblCount->Size = System::Drawing::Size(13, 13);
			this->lblCount->TabIndex = 3;
			this->lblCount->Text = L"0";
			// 
			// FeaturesForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(398, 286);
			this->Controls->Add(this->lblCount);
			this->Controls->Add(this->label1);
			this->Controls->Add(this->btnOK);
			this->Controls->Add(this->listView1);
			this->KeyPreview = true;
			this->Name = L"FeaturesForm";
			this->StartPosition = System::Windows::Forms::FormStartPosition::CenterParent;
			this->Text = L"Installed Features";
			this->KeyDown += gcnew System::Windows::Forms::KeyEventHandler(this, &FeaturesForm::FeaturesForm_KeyDown);
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
};
}
