#pragma once

#include <iostream>
#include <conio.h>
#include <cmath>
#include <string>
#include <ctime>
#include <msclr/marshal.h>
#include <msclr/marshal_cppstd.h>
#include <libxl.h>

namespace MainCorrectionCalculate {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace System::IO;
	using namespace System::IO::Ports;
	using namespace System::Configuration;
	using namespace msclr;
	using namespace msclr::interop;
	using namespace libxl;

	/// <summary>
	/// Сводка для MyForm
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
			//
			//TODO: добавьте код конструктора
			//
		}

	protected:
		/// <summary>
		/// Освободить все используемые ресурсы.
		/// </summary>
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Button^  button1;
	protected:
	public:
		Stream^ myStream;
		OpenFileDialog^ openFileDialog1 = gcnew OpenFileDialog;

	public:

	public:
	private:
		/// <summary>
		/// Обязательная переменная конструктора.
		/// </summary>
		System::ComponentModel::Container ^components;

		String^ convWchartToSystemStr(const wchar_t* value)
		{
			String^ rValue;
			/*String^ UnicodeValue;*/
			/*const wchar_t* unicode_value = (L"≤");*/
			marshal_context^ ctx = gcnew marshal_context();
			/*UnicodeValue = ctx->marshal_as<String^>(unicode_value);*/
			rValue = ctx->marshal_as<String^>(value)->Replace("от","/*от")
				->Replace("вкл.","вкл.*/")
				->Replace("/*от 20 до 50 вкл.*/","/*от 20 до 50 вкл.*/\t")
				->Replace("/*от 50 до 500 вкл.*/", "/*от 50 до 500 вкл.*/\t")
				->Replace("/*от 500 до 1000 вкл.*/", "/*от 500 до 1000 вкл.*/\t");
			delete ctx;
			return rValue;
		}
		inline double setDoublePrecis(double x)
		{
			int precision = 10000;//точность до 5 знака после запятой
			return round(x * precision) / precision;
		}

		std::string currentDateTime() {
			time_t rawtime;
			struct tm  tstruct;
			char       buf[80];
			time(&rawtime);
			localtime_s(&tstruct, &rawtime);
			strftime(buf, sizeof(buf), "%Y-%m-%d %X", &tstruct);
			return buf;
		}

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Требуемый метод для поддержки конструктора — не изменяйте 
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		void InitializeComponent(void)
		{
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->SuspendLayout();
			// 
			// button1
			// 
			this->button1->Location = System::Drawing::Point(42, 3);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(124, 23);
			this->button1->TabIndex = 0;
			this->button1->Text = L"SelectXlsFile";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Click += gcnew System::EventHandler(this, &MyForm::button1_Click);
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(213, 34);
			this->Controls->Add(this->button1);
			this->Name = L"MyForm";
			this->StartPosition = System::Windows::Forms::FormStartPosition::CenterScreen;
			this->Text = L"Main Correction";
			this->ResumeLayout(false);

		}
#pragma endregion
	
	private: System::Void button1_Click(System::Object^  sender, System::EventArgs^  e) 
	{
		String^ filename;
		String^ ctbfilename;
		String^ ctb;
		String^ Result;
		String^ DataTime;
		int cable = 1,power=4;
		const wchar_t *filename1;
		openFileDialog1->InitialDirectory = "C:\\Desktop";
		openFileDialog1->Filter = "All files (*.*)|*.*|xls files(*.xls) | *.xls";
		openFileDialog1->FilterIndex = 2;
		openFileDialog1->RestoreDirectory = true;

		if (openFileDialog1->ShowDialog() == System::Windows::Forms::DialogResult::OK)
		{
			if ((myStream = openFileDialog1->OpenFile()) != nullptr)
			{
				filename = openFileDialog1->FileName;
				ctb = openFileDialog1->SafeFileName;
				myStream->Close();
			}
		}
		ctbfilename =ctb->Replace(".xls", ".ctb");
		Book* book = xlCreateBook();
		StreamWriter ^crctbfile = gcnew StreamWriter(ctbfilename);
		marshal_context^ ctx = gcnew marshal_context();
		filename1 = ctx->marshal_as<const wchar_t*>(filename);
		DataTime +="//"+ctx->marshal_as<String^>(currentDateTime());
		Result +="//"+ filename+"\n"+DataTime+"\n"+"\n";
		Result += "//файл содержит различные поправки, касающиеся всего прибора ИПЗО" + "\n";
		Result += "//поправочные коэффициенты на кабели (по двум каналам, на трех частотах)" + "\n"+"\n";
		Result += "float CABLE_KOEF_01dB[][2][3] = {"+"\n"+"\t"+"/*на разъеме"+"*/"+"\t"+"{";
		if (book)
		{
			if (book->load(filename1))
			{
				Sheet* sheet =book->getSheet(2); //поправки по мощности
				Sheet* sheet1 = book->getSheet(3); //коррекция мощности
				Sheet* sheet2 = book->getSheet(1); //поправки по чувствительности
				if (sheet)
				{
					for (int i = 4; i <= 24;)
					{
						
						Result += "{"+Convert::ToString(round(sheet->readNum(i, 7)))+","+"\t";
						Result += Convert::ToString(round(sheet->readNum(i, 9))) + ","+"\t";
						Result += Convert::ToString(round(sheet->readNum(i, 11))) +"\t"+"},"+"\t";
						i++;
						Result += "{" + Convert::ToString(round(sheet->readNum(i, 7))) + ","+"\t";
						Result += Convert::ToString(round(sheet->readNum(i, 9))) + "," +"\t";
						if (i <= 24)
						{
							Result += Convert::ToString(round(sheet->readNum(i, 11))) + "}},";
							if (cable == 10)
							{
								Result += "\n" + "\t" + "/*кабель №" + Convert::ToString(cable) + "*/" + "\t" + "{";
							}
							else
							{
								Result += "\n" + "\t" + "/*кабель №" + Convert::ToString(cable) + "\t" "*/" + "\t" + "{";
							}
							cable++;
							i++;
						}
						else
						{
							Result += Convert::ToString(round(sheet->readNum(i, 11))) + "}}";
						}
					}
					cable = 1;
					Result += "\n" + "};"+"\n";
					Result += "\n" + "float CABLE_KOEF[][2][3] = {"+"\n"+"\t" +"/*на разъеме"+"*/"+"\t"+"{";
					for (int i = 4; i <= 24;)
					{
						Result += "{" + Convert::ToString(setDoublePrecis(sheet2->readNum(i, 5)))->Replace(",",".") + "," + "\t";
						Result += Convert::ToString(setDoublePrecis(sheet2->readNum(i, 7)))->Replace(",", ".") + "," + "\t";
						Result += Convert::ToString(setDoublePrecis(sheet2->readNum(i, 9)))->Replace(",", ".")+"\t"+ "}," + "\t";
						i++;
						Result += "{" + Convert::ToString(setDoublePrecis(sheet2->readNum(i, 5)))->Replace(",", ".") + "," + "\t";
						Result += Convert::ToString(setDoublePrecis(sheet2->readNum(i, 7)))->Replace(",", ".") + "," + "\t";
						if (i <= 24)
						{
							Result += Convert::ToString(setDoublePrecis(sheet2->readNum(i, 9)))->Replace(",", ".") + "}},";
							if (cable == 10)
							{
								Result += "\n" + "\t" + "/*кабель №" + Convert::ToString(cable) + "*/" + "\t" + "{";
							}
							else
							{
								Result += "\n" + "\t" + "/*кабель №" + Convert::ToString(cable) + "\t" "*/" + "\t" + "{";
							}
							cable++;
							i++;
						}
						else
						{
							Result += Convert::ToString(setDoublePrecis(sheet2->readNum(i, 9)))->Replace(",", ".") + "}}";
						}
					}
					cable = 1;
					Result += "\n" + "};" + "\n";
					Result += "\n" + "float POWER_KOEF[2][3]"+ "\t"+ "="+"\t" + "\t" +"{";
					Result += "{" + Convert::ToString(setDoublePrecis(sheet2->readNum(29, 4)))->Replace(",", ".") + "," + "\t";
					Result += Convert::ToString(setDoublePrecis(sheet2->readNum(29, 5)))->Replace(",", ".") + "," + "\t";
					Result += Convert::ToString(setDoublePrecis(sheet2->readNum(29, 6)))->Replace(",", ".") + "}," + "\t";
					Result += "{" + Convert::ToString(setDoublePrecis(sheet2->readNum(30, 4)))->Replace(",", ".") + "," + "\t";
					Result += Convert::ToString(setDoublePrecis(sheet2->readNum(30, 5)))->Replace(",", ".") + "," + "\t";
					Result += Convert::ToString(setDoublePrecis(sheet2->readNum(30, 6)))->Replace(",", ".") + "}};";

					Result += "\n" + "float SENS_KOEF_01dB[2][3] =" + "\t" + "{";
					Result += "{" + Convert::ToString(setDoublePrecis(sheet->readNum(3, 16)))->Replace(",", ".") + "," + "\t";
					Result += Convert::ToString(setDoublePrecis(sheet->readNum(3, 17)))->Replace(",", ".") + "," + "\t";
					Result += Convert::ToString(setDoublePrecis(sheet->readNum(3, 18)))->Replace(",", ".") + "}," + "\t";
					Result += "{" + Convert::ToString(setDoublePrecis(sheet->readNum(4, 16)))->Replace(",", ".") + "," + "\t";
					Result += Convert::ToString(setDoublePrecis(sheet->readNum(4, 17)))->Replace(",", ".") + "," + "\t";
					Result += Convert::ToString(setDoublePrecis(sheet->readNum(4, 18)))->Replace(",", ".") + "}};";

					Result +="\n"+"\n" + "float POWER_CORRECT[][2][3] = {" + "\n";
					for (int i = 4; i <= 33;)
					{
						Result += "\t"+convWchartToSystemStr(sheet1->readStr(power, 2))+"\t";
						Result += "{" + Convert::ToString(setDoublePrecis(sheet1->readNum(i, 8)))->Replace(",", ".") + "," + "\t";
						Result += Convert::ToString(setDoublePrecis(sheet1->readNum(i, 9)))->Replace(",", ".") + "," + "\t";
						Result += Convert::ToString(setDoublePrecis(sheet1->readNum(i, 10)))->Replace(",", ".")+ "}," + "\t";
						i++;
						Result += "{" + Convert::ToString(setDoublePrecis(sheet1->readNum(i, 8)))->Replace(",", ".") + "," + "\t";
						Result += Convert::ToString(setDoublePrecis(sheet1->readNum(i, 9)))->Replace(",", ".") + "," + "\t";
						if (i < 33)
						{
							Result += Convert::ToString(setDoublePrecis(sheet1->readNum(i, 10)))->Replace(",", ".") + "}},";
						}
						else
						{
							Result += Convert::ToString(setDoublePrecis(sheet1->readNum(i, 10)))->Replace(",", ".") + "}}";
						}
						i++;
						Result += "\n";
						power += 2;
					}
					Result += "};" + "\n";

					Result += "\n" + "float HF_MOD_ATTEN[2][3] =" + "\t" + "{";
					Result += "{" + Convert::ToString(setDoublePrecis(sheet->readNum(29, 2)))->Replace(",", ".") + "," + "\t";
					Result += Convert::ToString(setDoublePrecis(sheet->readNum(29, 3)))->Replace(",", ".") + "," + "\t";
					Result += Convert::ToString(setDoublePrecis(sheet->readNum(29, 4)))->Replace(",", ".") + "}," + "\t";
					Result += "{" + Convert::ToString(setDoublePrecis(sheet->readNum(30, 2)))->Replace(",", ".") + "," + "\t";
					Result += Convert::ToString(setDoublePrecis(sheet->readNum(30, 3)))->Replace(",", ".") + "," + "\t";
					Result += Convert::ToString(setDoublePrecis(sheet->readNum(30, 4)))->Replace(",", ".") + "}};";
				}
			}
			book->release();
			crctbfile->Write(Result);
			crctbfile->WriteLine();
			crctbfile->Close();
		}
		delete ctx;
	}
	};
}
