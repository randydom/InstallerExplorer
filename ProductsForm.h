#pragma once

using namespace System;
using namespace System::Windows::Forms;

#include "FeaturesForm.h"
#include "PatchesForm.h"
#include "ComponentsForm.h"
#include "ListUtils.h"

namespace InstallerExplorer {
	/// <summary>
	/// List all installed products
	/// Handle commands to display details
	/// </summary>
	public ref class ProductsForm : public Form {
	protected:
		InstallerExplorer::ListUtils* listUtils;

	public:
		ProductsForm(void) {
			listUtils = new InstallerExplorer::ListUtils();
			InitializeComponent();
			
			this->SuspendLayout();
			EnumProducts();
			this->ResumeLayout();
		}

	protected:
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~ProductsForm() {
			if (components)
				delete components;
			if(listUtils)
				delete listUtils;
		}

#pragma region Event Handlers

		System::Void btnFeatures_Click(System::Object^ sender, System::EventArgs^ e) {
			if(1 == this->listView1->SelectedIndices->Count) {
			   String^ str = this->listView1->SelectedItems[0]->SubItems[1]->Text;
			   InstallerExplorer::FeaturesForm ^ ff = gcnew InstallerExplorer::FeaturesForm(str);
			   ff->ShowDialog();
			}
		}

		System::Void btnComponents_Click(System::Object^ sender, System::EventArgs^ e) {
			if(1 == this->listView1->SelectedIndices->Count) {
			   String^ str = this->listView1->SelectedItems[0]->SubItems[1]->Text;
			   InstallerExplorer::ComponentsForm ^ cf = gcnew InstallerExplorer::ComponentsForm(str);
			   cf->ShowDialog();
			}
		}

		System::Void btnPatches_Click(System::Object^ sender, System::EventArgs^ e) {
			if(1 == this->listView1->SelectedIndices->Count) {
			   String^ str = this->listView1->SelectedItems[0]->SubItems[1]->Text;
			   InstallerExplorer::PatchesForm ^ pf = gcnew InstallerExplorer::PatchesForm(str);
			   pf->ShowDialog();
			}
		}

		System::Void listView1_SelectedIndexChanged(System::Object^ sender, System::EventArgs^ e) {
			bool enable = (this->listView1->SelectedItems->Count == 1);
			this->btnComponents->Enabled = enable;
			this->btnFeatures->Enabled = enable;
			this->btnPatches->Enabled = enable;
		}

		System::Void Form1_KeyDown(System::Object^ sender, KeyEventArgs^ e) {
			listUtils->DoKeyDown(e, this->listView1, 3, L"Product Name\tProduct Code\tState");
		}

		System::Void listView1_ColumnClick(System::Object^ sender, ColumnClickEventArgs^ e) {
			listUtils->DoColumnClick(e->Column, this->listView1);
		}

#pragma endregion

#pragma region Enumeration Code

	private:
		void EnumProducts() {
			const int cchGUID = 38;
			WCHAR wszProductCode[cchGUID+1] = {0};

			int i = 0;
			HRESULT hr = S_OK;
			do{
				hr = MsiEnumProducts(i++, wszProductCode);
				if(S_OK == hr)
				{
					AddProduct(wszProductCode);
				}
			} while(S_OK == hr);

			WCHAR wcsCount[32] = {0};
			_itow_s(i-1, wcsCount, 10);
			this->lblCount->Text = gcnew String(wcsCount);
		}

		void AddProduct(WCHAR* pwsProductCode) {
			DWORD cchProductName = MAX_PATH;
			WCHAR* lpProductName = new WCHAR[cchProductName];
			if (!lpProductName)
				return;

			// obtain the user friendly name of the product
			HRESULT hr = MsiGetProductInfo(pwsProductCode,INSTALLPROPERTY_PRODUCTNAME,lpProductName,&cchProductName);
			if (ERROR_MORE_DATA == hr) {
				// try again, but with a larger product name buffer
				delete [] lpProductName;

				// returned character count does not include
				// terminating NULL
				++cchProductName;

				lpProductName = new WCHAR[cchProductName];
				if (!lpProductName)
					hr = ERROR_OUTOFMEMORY;

				hr = MsiGetProductInfo(pwsProductCode,INSTALLPROPERTY_PRODUCTNAME,lpProductName,&cchProductName);
			}
			if(S_OK == hr) {
				INSTALLSTATE state = MsiQueryProductState(pwsProductCode);
				String^ stateStr = gcnew String("?");
				switch(state) {
					case INSTALLSTATE_ABSENT:
						stateStr = L"Absent";
						break;
					case INSTALLSTATE_ADVERTISED:
						stateStr = "Advertised";
						break;
					case INSTALLSTATE_DEFAULT:
						stateStr = "Installed";
						break;
				}
				ListViewItem^  listViewItem1 = (gcnew ListViewItem(gcnew cli::array< System::String^  >(3) {gcnew String(lpProductName), gcnew String(pwsProductCode), stateStr}, -1));
				this->listView1->Items->AddRange(gcnew cli::array<ListViewItem^  >(1) {listViewItem1});
			}
		}
		
#pragma endregion

#pragma region Form Initialization

	private: System::Windows::Forms::ListView^  listView1;
	private: System::Windows::Forms::ColumnHeader^  columnHeader1;
	private: System::Windows::Forms::ColumnHeader^  columnHeader2;
	private: System::Windows::Forms::Button^  btnFeatures;
	private: System::Windows::Forms::Button^  btnComponents;
	private: System::Windows::Forms::Button^  btnPatches;
	private: System::Windows::Forms::Label^  lblCount;
	private: System::Windows::Forms::Label^  label1;
	private: System::Windows::Forms::ColumnHeader^  columnHeader3;


	protected: 

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
			this->columnHeader3 = (gcnew System::Windows::Forms::ColumnHeader());
			this->btnFeatures = (gcnew System::Windows::Forms::Button());
			this->btnComponents = (gcnew System::Windows::Forms::Button());
			this->btnPatches = (gcnew System::Windows::Forms::Button());
			this->lblCount = (gcnew System::Windows::Forms::Label());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->SuspendLayout();
			// 
			// listView1
			// 
			this->listView1->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((((System::Windows::Forms::AnchorStyles::Top | System::Windows::Forms::AnchorStyles::Bottom) 
				| System::Windows::Forms::AnchorStyles::Left) 
				| System::Windows::Forms::AnchorStyles::Right));
			this->listView1->Columns->AddRange(gcnew cli::array< System::Windows::Forms::ColumnHeader^  >(3) {this->columnHeader1, this->columnHeader2, 
				this->columnHeader3});
			this->listView1->FullRowSelect = true;
			this->listView1->GridLines = true;
			this->listView1->HideSelection = false;
			this->listView1->Location = System::Drawing::Point(0, 0);
			this->listView1->Name = L"listView1";
			this->listView1->Size = System::Drawing::Size(675, 318);
			this->listView1->Sorting = System::Windows::Forms::SortOrder::Ascending;
			this->listView1->TabIndex = 0;
			this->listView1->UseCompatibleStateImageBehavior = false;
			this->listView1->View = System::Windows::Forms::View::Details;
			this->listView1->SelectedIndexChanged += gcnew System::EventHandler(this, &ProductsForm::listView1_SelectedIndexChanged);
			this->listView1->ColumnClick += gcnew System::Windows::Forms::ColumnClickEventHandler(this, &ProductsForm::listView1_ColumnClick);
			// 
			// columnHeader1
			// 
			this->columnHeader1->Text = L"Product Name";
			this->columnHeader1->Width = 332;
			// 
			// columnHeader2
			// 
			this->columnHeader2->Text = L"Product Code";
			this->columnHeader2->Width = 262;
			// 
			// columnHeader3
			// 
			this->columnHeader3->Text = L"State";
			// 
			// btnFeatures
			// 
			this->btnFeatures->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((System::Windows::Forms::AnchorStyles::Bottom | System::Windows::Forms::AnchorStyles::Right));
			this->btnFeatures->Enabled = false;
			this->btnFeatures->Location = System::Drawing::Point(588, 324);
			this->btnFeatures->Name = L"btnFeatures";
			this->btnFeatures->Size = System::Drawing::Size(75, 23);
			this->btnFeatures->TabIndex = 1;
			this->btnFeatures->Text = L"Features";
			this->btnFeatures->UseVisualStyleBackColor = true;
			this->btnFeatures->Click += gcnew System::EventHandler(this, &ProductsForm::btnFeatures_Click);
			// 
			// btnComponents
			// 
			this->btnComponents->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((System::Windows::Forms::AnchorStyles::Bottom | System::Windows::Forms::AnchorStyles::Right));
			this->btnComponents->Enabled = false;
			this->btnComponents->Location = System::Drawing::Point(507, 324);
			this->btnComponents->Name = L"btnComponents";
			this->btnComponents->Size = System::Drawing::Size(75, 23);
			this->btnComponents->TabIndex = 2;
			this->btnComponents->Text = L"Components";
			this->btnComponents->UseVisualStyleBackColor = true;
			this->btnComponents->Click += gcnew System::EventHandler(this, &ProductsForm::btnComponents_Click);
			// 
			// btnPatches
			// 
			this->btnPatches->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((System::Windows::Forms::AnchorStyles::Bottom | System::Windows::Forms::AnchorStyles::Right));
			this->btnPatches->Enabled = false;
			this->btnPatches->Location = System::Drawing::Point(426, 324);
			this->btnPatches->Name = L"btnPatches";
			this->btnPatches->Size = System::Drawing::Size(75, 23);
			this->btnPatches->TabIndex = 3;
			this->btnPatches->Text = L"Patches";
			this->btnPatches->UseVisualStyleBackColor = true;
			this->btnPatches->Click += gcnew System::EventHandler(this, &ProductsForm::btnPatches_Click);
			// 
			// lblCount
			// 
			this->lblCount->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((System::Windows::Forms::AnchorStyles::Bottom | System::Windows::Forms::AnchorStyles::Left));
			this->lblCount->AutoSize = true;
			this->lblCount->Location = System::Drawing::Point(57, 329);
			this->lblCount->Name = L"lblCount";
			this->lblCount->Size = System::Drawing::Size(13, 13);
			this->lblCount->TabIndex = 9;
			this->lblCount->Text = L"0";
			// 
			// label1
			// 
			this->label1->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((System::Windows::Forms::AnchorStyles::Bottom | System::Windows::Forms::AnchorStyles::Left));
			this->label1->AutoSize = true;
			this->label1->Location = System::Drawing::Point(12, 329);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(38, 13);
			this->label1->TabIndex = 8;
			this->label1->Text = L"Count:";
			// 
			// Form1
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(675, 359);
			this->Controls->Add(this->lblCount);
			this->Controls->Add(this->label1);
			this->Controls->Add(this->btnPatches);
			this->Controls->Add(this->btnComponents);
			this->Controls->Add(this->btnFeatures);
			this->Controls->Add(this->listView1);
			this->KeyPreview = true;
			this->Name = L"Form1";
			this->Text = L"Installed Products";
			this->KeyDown += gcnew System::Windows::Forms::KeyEventHandler(this, &ProductsForm::Form1_KeyDown);
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
};
}

