///////////////////////////////////////
// MyDialog.cpp

#include "stdafx.h"
#include "MyDialog.h"
#include "resource.h"

//SoftwarePicker.cpp���\�[�X�t�@�C���Ƃ���include����Ƒ��d��`�G���[�ɂȂ�
//�w�b�_�t�@�C���Ɉڂ��Ċg���q��.h�Ƃ���
#include "SoftwarePicker.h"

//�L�������̃t�H���g�̐ݒ�
//�t�H���g�̓_�C�A���O��j������܂ŕK�v�Ȃ̂ł����Ő錾
//http://akky.xrea.jp/mfc/listfont.html
CFont m_font;

//�E�B���h�E�n���h��
HWND hwnd;
//�C���X�^���X�n���h��
HINSTANCE hinst;

//LPSTR lpCmdLine2;

//�����̐ݒ�
//resource.h�Őݒ肷��
//int periodYear  = 2014;
//int periodMonth = 9;
//int periodDay   = 30;
int periodYear  = VER_INT_PERIOD_YEAR;
int periodMonth = VER_INT_PERIOD_MONTH;
int periodDay   = VER_INT_PERIOD_DAY;
int periodYearDisplay  = VER_INT_PERIOD_YEAR;
int periodMonthDisplay = VER_INT_PERIOD_MONTH;
int periodDayDisplay   = VER_INT_PERIOD_DAY;
wstring productVersion = VER_STR_PRODUCTVERSION;
wstring customer = VER_STR_CUSTOMER;

InventoryManager currentInventoryManager;


// Definitions for the CMyDialog class
CMyDialog::CMyDialog(UINT nResID, CWnd* pParent)
	: CDialog(nResID, pParent)
{
	m_hInstRichEdit = ::LoadLibrary(_T("RICHED32.DLL"));
    if (!m_hInstRichEdit)
 		::MessageBox(NULL, _T("CMyDialog::CRichView  Failed to load RICHED32.DLL"), _T(""), MB_ICONWARNING);
}

CMyDialog::~CMyDialog()
{
	::FreeLibrary(m_hInstRichEdit);
}

INT_PTR CMyDialog::DialogProc(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
//	switch (uMsg)
//	{
		//Additional messages to be handled go here
//	}

	// Pass unhandled messages on to parent DialogProc
	return DialogProcDefault(uMsg, wParam, lParam);
}

BOOL CMyDialog::OnCommand(WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);

	//switch (LOWORD(wParam))
    //{
	//case IDC_BUTTON1:
	//	OnButton();
	//	return TRUE;
	//case IDC_RADIO1:
	//	OnRadio1();
	//	return TRUE;
	//case IDC_RADIO2:
	//	OnRadio2();
	//	return TRUE;
	//case IDC_RADIO3:
	//	OnRadio3();
	//	return TRUE;
	//case IDC_CHECK1:
	//	OnCheck1();
	//	return TRUE;
	//case IDC_CHECK2:
	//	OnCheck2();
	//	return TRUE;
	//case IDC_CHECK3:
	//	OnCheck3();
	//	return TRUE;
    //} //switch (LOWORD(wParam))

	return FALSE;
}

BOOL CMyDialog::OnInitDialog()
{

	//�R�}���h���C�������̎擾
	//http://www.wabiapp.com/WabiSampleSource/windows/get_command_line.html

	// �������X�g�̐�
	int nArgc = 0;
 
	// �������X�g�̎擾
	TCHAR* lpCmdLine = GetCommandLine();
	TCHAR** lppArgv = CommandLineToArgvW(lpCmdLine, &nArgc);
 
	// �������X�g�̕\��
	for(int i = 0; i < nArgc; i++){
		//MessageBox(lppArgv[i], VER_STR_PRODUCTNAME, MB_OK);
		if(_tcscmp(StrToLower(lppArgv[i]), L"/silent") == 0){
			//MessageBox(L"silent mode", VER_STR_PRODUCTNAME, MB_OK);
			silentMode = TRUE;
		}else if(_tcscmp(StrToLower(lppArgv[i]), L"/ini") == 0){
			if(i+1 < nArgc){
				iniPath = lppArgv[i+1];
			}
		}
	}
 
	// �������X�g�̉��
	LocalFree(lppArgv);

	// Set the Icon
	SetIconLarge(IDW_MAIN);
	SetIconSmall(IDW_MAIN);

	//�L�������̃t�H���g�ύX
	//http://www.crimson-systems.com/tips/t026c.htm
	//m_font.CreatePointFont(70,L"MS Shell Dlg");
	//m_font.CreatePointFont(120,L"Arial");
	//GetDlgItem(IDC_STATIC5)->SetFont(&m_font);

	//���ݓ����̎擾
	struct tm now;
	__time64_t long_time;
	_time64( &long_time ); 
	_localtime64_s( &now, &long_time );

	//�������`�F�b�N
	BOOL overdue = FALSE;
	if(now.tm_year + 1900 > periodYear){
		overdue = TRUE;
	}else if(now.tm_year + 1900 == periodYear){
		if(now.tm_mon + 1 > periodMonth){
			overdue = TRUE;
		}else if(now.tm_mon + 1 == periodMonth){
			if(now.tm_mday > periodDay){
				overdue = TRUE;
			}
		}
	}
	//�������߂��Ă���΃��b�Z�[�W�{�b�N�X���o���ďI��
	if(overdue){
		if(silentMode){
			FILE* fp = NULL;
			errno_t error;
			//UTF-8�ŕۑ�
			error = _wfopen_s( &fp, L"message.txt", L"w,ccs=UTF-8" );
			if(error == 0){
				_ftprintf(fp, L"���C�Z���X�̗L�����Ԃ��I�����܂����B���g�p���肪�Ƃ��������܂����B\n");
			}
			fclose(fp);
		}else{
			MessageBox(L"���C�Z���X�̗L�����Ԃ��I�����܂����B���g�p���肪�Ƃ��������܂����B", VER_STR_PRODUCTNAME, MB_OK);
			//CDialog::OnOK()��destroy�����遁�I���Ƃ����킯�ł͂Ȃ�
			CDialog::OnOK();
		}

		exit(0);
	}

	//INI�t�@�C�����烉�x���e�L�X�g��ǂݍ���
	TCHAR iniChar[4096];
	DWORD ret = 0;
	ret = GetPrivateProfileString(L"form", L"label_hardware_no", L"", iniChar, sizeof(iniChar), iniPath.c_str());
	if(ret > 0){
		labelHardwareNo = iniChar;
		if(labelHardwareNo == L"[none]"){
			labelHardwareNo = L"";
		}
		SetDlgItemText(IDC_STATIC_HARDWARE_NO, labelHardwareNo.c_str());
	}
	ret = GetPrivateProfileString(L"form", L"label1", L"", iniChar, sizeof(iniChar), iniPath.c_str());
	if(ret > 0){
		label1 = iniChar;
		if(label1 == L"[none]"){
			label1 = L"";
		}
		SetDlgItemText(IDC_STATIC1, label1.c_str());
	}
	ret = GetPrivateProfileString(L"form", L"label2", L"", iniChar, sizeof(iniChar), iniPath.c_str());
	if(ret > 0){
		label2 = iniChar;
		if(label2 == L"[none]"){
			label2 = L"";
		}
		SetDlgItemText(IDC_STATIC2, label2.c_str());
	}
	ret = GetPrivateProfileString(L"form", L"label3", L"", iniChar, sizeof(iniChar), iniPath.c_str());
	if(ret > 0){
		label3 = iniChar;
		if(label3 == L"[none]"){
			label3 = L"";
		}
		SetDlgItemText(IDC_STATIC3, label3.c_str());
	}
	ret = GetPrivateProfileString(L"form", L"label4", L"", iniChar, sizeof(iniChar), iniPath.c_str());
	if(ret > 0){
		label4 = iniChar;
		if(label4 == L"[none]"){
			label4 = L"";
		}
		SetDlgItemText(IDC_STATIC4, label4.c_str());
	}
	ret = GetPrivateProfileString(L"form", L"label5", L"", iniChar, sizeof(iniChar), iniPath.c_str());
	if(ret > 0){
		label5 = iniChar;
		if(label5 == L"[none]"){
			label5 = L"";
		}
		SetDlgItemText(IDC_STATIC5, label5.c_str());
	}
	ret = GetPrivateProfileString(L"form", L"label6", L"", iniChar, sizeof(iniChar), iniPath.c_str());
	if(ret > 0){
		label6 = iniChar;
		if(label6 == L"[none]"){
			label6 = L"";
		}
		SetDlgItemText(IDC_STATIC6, label6.c_str());
	}

	//Edit�R���g���[����ǂݎ���p�ɂ���
	//http://www.geocities.jp/chiakifujimon/control/control1.html
	wstring iniString;
	ret = GetPrivateProfileString(L"form", L"input_hardware_no", L"", iniChar, sizeof(iniChar), iniPath.c_str());
	if(ret > 0){
		iniString = iniChar;
		if(iniString == L"[invalid]"){
			CEdit* pedt = (CEdit*)GetDlgItem(IDC_EDIT_HARDWARE_NO);
			pedt->SetReadOnly(TRUE);
			isInputHardwareNoValid = false;
		}
	}

	ret = GetPrivateProfileString(L"form", L"input1", L"", iniChar, sizeof(iniChar), iniPath.c_str());
	if(ret > 0){
		iniString = iniChar;
		if(iniString == L"[invalid]"){
			CEdit* pedt = (CEdit*)GetDlgItem(IDC_EDIT1);
			pedt->SetReadOnly(TRUE);
			isInput1Valid = false;
		}
	}

	ret = GetPrivateProfileString(L"form", L"input2", L"", iniChar, sizeof(iniChar), iniPath.c_str());
	if(ret > 0){
		iniString = iniChar;
		if(iniString == L"[invalid]"){
			CEdit* pedt = (CEdit*)GetDlgItem(IDC_EDIT2);
			pedt->SetReadOnly(TRUE);
			isInput2Valid = false;
		}
	}

	ret = GetPrivateProfileString(L"form", L"input3", L"", iniChar, sizeof(iniChar), iniPath.c_str());
	if(ret > 0){
		iniString = iniChar;
		if(iniString == L"[invalid]"){
			CEdit* pedt = (CEdit*)GetDlgItem(IDC_EDIT3);
			pedt->SetReadOnly(TRUE);
			isInput3Valid = false;
		}
	}

	ret = GetPrivateProfileString(L"form", L"input4", L"", iniChar, sizeof(iniChar), iniPath.c_str());
	if(ret > 0){
		iniString = iniChar;
		if(iniString == L"[invalid]"){
			CEdit* pedt = (CEdit*)GetDlgItem(IDC_EDIT4);
			pedt->SetReadOnly(TRUE);
			isInput4Valid = false;
		}
	}

	ret = GetPrivateProfileString(L"form", L"input5", L"", iniChar, sizeof(iniChar), iniPath.c_str());
	if(ret > 0){
		iniString = iniChar;
		if(iniString == L"[invalid]"){
			CEdit* pedt = (CEdit*)GetDlgItem(IDC_EDIT5);
			pedt->SetReadOnly(TRUE);
			isInput5Valid = false;
		}
	}

	ret = GetPrivateProfileString(L"form", L"input6", L"", iniChar, sizeof(iniChar), iniPath.c_str());
	if(ret > 0){
		iniString = iniChar;
		if(iniString == L"[invalid]"){
			CEdit* pedt = (CEdit*)GetDlgItem(IDC_EDIT6);
			pedt->SetReadOnly(TRUE);
			isInput6Valid = false;
		}
	}


	//INI�t�@�C������o�̓t�H���_��ǂݍ���
	ret = GetPrivateProfileString(L"output", L"path", L"", iniChar, sizeof(iniChar), iniPath.c_str());
	outputPath = iniChar;
	if(outputPath.size() > 0 && outputPath.substr(outputPath.size()-1,1) != L"\\"){
		outputPath += L"\\";
	}
	//outputPath.substr(outputPath.size()-1, 1);
	//outputPath = L"12345";
	//MessageBox(outputPath.c_str(), VER_STR_PRODUCTNAME, MB_OK);
	//MessageBox(outputPath.substr(outputPath.size()-1,1).c_str(), VER_STR_PRODUCTNAME, MB_OK);

	//INI�t�@�C������n�[�h�E�F�A�Ǘ��ԍ��̃v���t�B�b�N�X��ǂݍ���
	ret = GetPrivateProfileString(L"input", L"hardware_no_prefix", L"", iniChar, sizeof(iniChar), iniPath.c_str());
	if(ret > 0){
		hardwareNoPrefix = iniChar;
	}

	//INI�t�@�C������n�[�h�E�F�A�Ǘ��ԍ��̌�����ǂݍ���
	ret = GetPrivateProfileString(L"input", L"hardware_no_digits", L"", iniChar, sizeof(iniChar), iniPath.c_str());
	if(ret > 0){
		hardwareNoDigits = _wtoi(iniChar);
	}

	//INI�t�@�C��������s��R�}���h��ǂݍ���
	ret = GetPrivateProfileString(L"output", L"program", L"", iniChar, sizeof(iniChar), iniPath.c_str());
	if(ret > 0){
		iniProgram = iniChar;
	}
	ret = GetPrivateProfileString(L"output", L"arguments", L"", iniChar, sizeof(iniChar), iniPath.c_str());
	if(ret > 0){
		iniProgramArguments = iniChar;
	}

	//INI�t�@�C������Oracle�Ή��̗L����ǂݍ���
	ret = GetPrivateProfileString(L"output", L"opatch_check", L"", iniChar, sizeof(iniChar), iniPath.c_str());
	if(ret > 0){
		iniOracleCompatible = _wtoi(iniChar);
	}



	//MessageBox(getWMI(L"SELECT * FROM Win32_OperatingSystem", L"Caption").c_str(), VER_STR_PRODUCTNAME, MB_OK);
	//MessageBox(getWMI(L"SELECT * FROM Win32_Printer", L"Name").c_str(), VER_STR_PRODUCTNAME, MB_OK);
	//MessageBox(getWMI(L"SELECT * FROM Win32_Printer", L"PrinterState").c_str(), VER_STR_PRODUCTNAME, MB_OK);
	//MessageBox(getWMI(L"SELECT * FROM Win32_NetworkAdapterConfiguration", L"IPAddress").c_str(), VER_STR_PRODUCTNAME, MB_OK);
	//MessageBox(getWMI(L"SELECT * FROM Win32_Processor", L"CurrentClockSpeed").c_str(), VER_STR_PRODUCTNAME, MB_OK);
	//MessageBox(getWMI(L"SELECT * FROM Win32_Processor", L"NumberOfCores").c_str(), VER_STR_PRODUCTNAME, MB_OK);
	//MessageBox(getWMI(L"SELECT * FROM Win32_Processor", L"NumberOfLogicalProcessors").c_str(), VER_STR_PRODUCTNAME, MB_OK);

	//default.csv�Ƃ̓ˍ��̂��߂ɂ����Ńn�[�h�E�F�A�����擾���Ă���
	currentInventoryManager.getHardwareInfo();


	//�t�H�[�}�b�g��AdvancedManager�`���̏ꍇ
	if(outputFileFormat == L"AdvancedManager"){

		//�f�t�H���g�l�����[�h
		currentInventoryManager.loadDefaultSetting();

		//�O��̓��͏����擾
		currentInventoryManager.load();

		//�T�C�����g���[�h�̏ꍇ
		if(silentMode){
			//�O��̓��͏�񂪂���Ȃ�O��̓��͏��𗬗p���ď����W���A��ʂ�\�������ɏI��
			if(currentInventoryManager.getPreviousHardwareNoValue() != L""){
				currentInventoryManager.setInputInfo(
					currentInventoryManager.getPreviousHardwareNoValue(),
					currentInventoryManager.getPreviousValue1(),
					currentInventoryManager.getPreviousValue2(),
					currentInventoryManager.getPreviousValue3(),
					currentInventoryManager.getPreviousValue4(),
					currentInventoryManager.getPreviousValue5(),
					currentInventoryManager.getPreviousValue6()
				);
				long res  = currentInventoryManager.getHardwareInfo();
				long res2 = currentInventoryManager.listUp();
				if(res == 0 && res2 == 0){
					currentInventoryManager.load();
					currentInventoryManager.save();
					wstring outputFileName = currentInventoryManager.output();
					//wchar_t messageChars [1024];
					//swprintf(messageChars, 1024, L"�T�C�����g���[�h�ŃC���x���g������ۑ����܂����B\n�ۑ��t�@�C��: \n%s", outputFileName.c_str());
					//MessageBox(messageChars, VER_STR_PRODUCTNAME, MB_OK);
					
					//default.csv�̃��l�[��
					currentInventoryManager.cleanupDefaultSetting();
				}
				//�O���v���O�����̎��s
				if(iniProgram != L""){
					RunProcess(iniProgram, iniProgramArguments, false);
				}

				//�I��
				CDialog::OnOK();
			}

		}

		SetDlgItemText(IDC_EDIT_HARDWARE_NO, currentInventoryManager.getPreviousHardwareNoValue().c_str());
		if(isInput1Valid){
			SetDlgItemText(IDC_EDIT1, currentInventoryManager.getPreviousValue1().c_str());
		}
		if(isInput2Valid){
			SetDlgItemText(IDC_EDIT2, currentInventoryManager.getPreviousValue2().c_str());
		}
		if(isInput3Valid){
			SetDlgItemText(IDC_EDIT3, currentInventoryManager.getPreviousValue3().c_str());
		}
		if(isInput4Valid){
			SetDlgItemText(IDC_EDIT4, currentInventoryManager.getPreviousValue4().c_str());
		}
		if(isInput5Valid){
			SetDlgItemText(IDC_EDIT5, currentInventoryManager.getPreviousValue5().c_str());
		}
		if(isInput6Valid){
			SetDlgItemText(IDC_EDIT6, currentInventoryManager.getPreviousValue6().c_str());
		}

	}

		//SetDlgItemText(IDC_EDIT_HARDWARE_NO, currentInventoryManager.getPreviousHardwareNoValue().c_str());
		//SetDlgItemText(IDC_EDIT_HARDWARE_NO, L"123");
		//SetDlgItemText(IDC_EDIT6, L"123");


	//�G�f�B�b�g�{�b�N�X�Ƀf�t�H���g�l��ݒ肷���
	// Put some text in the edit boxes
	//SetDlgItemText(IDC_EDIT1, _T("Edit Control"));
	//SetDlgItemText(IDC_EDIT2, _T("Edit Control"));
	//SetDlgItemText(IDC_EDIT3, _T("Edit Control"));
	//SetDlgItemText(IDC_RICHEDIT1, _T("Rich Edit Window"));

	//�G�f�B�b�g�{�b�N�X�̒l���擾�����
	//wstring edit3 = GetDlgItemText(IDC_EDIT3);

	//���X�g�{�b�N�X�ɒl��ݒ肷��ꍇ�̗�
	//for (int i = 0 ; i < 8 ; i++)
	//	SendDlgItemMessage(IDC_LIST1, LB_ADDSTRING, 0, (LPARAM) _T("List Box"));

	return true;
}

void CMyDialog::OnOK(){
	wstring editHardwareNo = GetDlgItemText(IDC_EDIT_HARDWARE_NO);
	wstring edit1 = GetDlgItemText(IDC_EDIT1);
	wstring edit2 = GetDlgItemText(IDC_EDIT2);
	wstring edit3 = GetDlgItemText(IDC_EDIT3);
	wstring edit4 = GetDlgItemText(IDC_EDIT4);
	wstring edit5 = GetDlgItemText(IDC_EDIT5);
	wstring edit6 = GetDlgItemText(IDC_EDIT6);
	
	if(editHardwareNo == L"deepdebuglog=true"){
		MessageBox(L"�\�t�g�E�F�A���擾���A�������O���o�͂��܂��B���͗��ɓ��͌�AOK���N���b�N���Ă��������B", VER_STR_PRODUCTNAME, MB_OK);
		SetDlgItemText(IDC_EDIT_HARDWARE_NO, _T(""));
		debugLevel = 3;		
		return;
	}
	if(editHardwareNo == L"debuglog=true"){
		MessageBox(L"�\�t�g�E�F�A���擾���A�������O���o�͂��܂��B���͗��ɓ��͌�AOK���N���b�N���Ă��������B", VER_STR_PRODUCTNAME, MB_OK);
		SetDlgItemText(IDC_EDIT_HARDWARE_NO, _T(""));
		//PrintDebugLog(L"123");
		debugLevel = 2;		
		return;
	}
	if(editHardwareNo == L"log=true"){
		MessageBox(L"�\�t�g�E�F�A���擾���A�������O���o�͂��܂��B���͗��ɓ��͌�AOK���N���b�N���Ă��������B", VER_STR_PRODUCTNAME, MB_OK);
		SetDlgItemText(IDC_EDIT_HARDWARE_NO, _T(""));
		debugLevel = 1;
		return;
	}
	//if(edit1.size() == 0 || edit2.size() == 0 || edit3.size() == 0){
	if(editHardwareNo.size() == 0){
		wstring msg = L"����͂��Ă��������B";
		msg = VER_STR_LABEL_HARDWARE_NO + msg;
		MessageBox(msg.c_str(), VER_STR_PRODUCTNAME, MB_OK);
		return;
	}

	wstring hardwareNoInfix = L"";
	//�n�[�h�E�F�A�Ǘ��ԍ��̐ړ�����INI�Őݒ肳��Ă���ꍇ
	if(hardwareNoPrefix.size() > 0){
		//���͂��ꂽ�n�[�h�E�F�A�Ǘ��ԍ��ɐړ������t���Ă���ꍇ
		if(editHardwareNo.substr(0, hardwareNoPrefix.size()) == hardwareNoPrefix){
			hardwareNoInfix = editHardwareNo.substr(hardwareNoPrefix.size(), editHardwareNo.size() - hardwareNoPrefix.size());
		}else{
			hardwareNoInfix = editHardwareNo;
		}
	}else{
		hardwareNoInfix = editHardwareNo;
	}

	//MessageBox(hardwareNoInfix.c_str(), VER_STR_PRODUCTNAME, MB_OK);


	//�n�[�h�E�F�A�Ǘ��ԍ��̃[���l�ߌ�����INI�Őݒ肳��Ă���ꍇ
	if(hardwareNoDigits > 0){
		//�ړ����𔲂����n�[�h�E�F�A�Ǘ��ԍ��������łȂ�or�������������遁�����ɍ����Ă��Ȃ�
		if(!IsNumericString(hardwareNoInfix)){
			wstring msg = L"�̏������w�肳�ꂽ���̂ƈ�v���܂���B";
			msg = VER_STR_LABEL_HARDWARE_NO + msg;
			MessageBox(msg.c_str(), VER_STR_PRODUCTNAME, MB_OK);
			return;
		}else if(hardwareNoInfix.size() > hardwareNoDigits){
			wstring msg = L"�̏������w�肳�ꂽ���̂ƈ�v���܂���B";
			msg = VER_STR_LABEL_HARDWARE_NO + msg;
			MessageBox(msg.c_str(), VER_STR_PRODUCTNAME, MB_OK);
			return;
		}
		//0�l��
		wstring padding(hardwareNoDigits - hardwareNoInfix.size(), '0');
		hardwareNoInfix = padding + hardwareNoInfix;
	}

	
	//�e�L�X�g�{�b�N�X�ɏ����߂�
	if(hardwareNoPrefix.size() > 0){
		editHardwareNo = hardwareNoPrefix + hardwareNoInfix;
	}else{
		editHardwareNo = hardwareNoInfix;
	}

	SetDlgItemText(IDC_EDIT_HARDWARE_NO, editHardwareNo.c_str());
	//MessageBox(editHardwareNo.c_str(), VER_STR_PRODUCTNAME, MB_OK);

	//InventoryManager testInventoryManager;
	//inventory testInventory;
	//testInventory = testInventoryManager.getHardwareInfo();
	//SetDlgItemText(IDC_EDIT1, testInventory.computerName.c_str());
	//SetDlgItemText(IDC_EDIT2, testInventory.userName.c_str());
	//SetDlgItemText(IDC_EDIT2, testInventory.MACAddress.c_str());
	//SetDlgItemText(IDC_EDIT3, testInventory.IPAddress.c_str());

	if(debugLevel > 0){
		//�f�o�b�O���O�̕ۑ��t�@�C����
		//�R���s���[�^���̎擾
		WORD wVersionRequested;
		WSADATA wsaData;
		char computerName[255];
		_TCHAR wcomputerName[255];
		//PHOSTENT hostinfo;
		wVersionRequested = MAKEWORD( 1, 1 );
		//char *ip;

		if ( WSAStartup( wVersionRequested, &wsaData ) == 0 ){
			if( gethostname ( computerName, sizeof(computerName)) == 0){
				//printf("Host name: %s\n", name);
				mbstowcs(wcomputerName, computerName, sizeof(computerName));
			}
		}
		//����
		struct tm timestamp;
		__time64_t long_time;
		_time64( &long_time ); 
		_localtime64_s( &timestamp, &long_time );
		//�t�@�C�����̌���
		wchar_t outputFileNameChars [100];
		swprintf(outputFileNameChars, 100,  _T( "%s%04d%02d%02d%02d%02d%02ddebug.log" ), wcomputerName, timestamp.tm_year + 1900, timestamp.tm_mon + 1, timestamp.tm_mday, timestamp.tm_hour, timestamp.tm_min, timestamp.tm_sec);
		debugLogFileName = outputFileNameChars;
		
		//�ۑ��t�@�C�����J��
		LONG status;
		status = _wfopen_s( &debugFP, debugLogFileName.c_str(), L"a+,ccs=UTF-8" );

		//���[�h�A�g�p�����̏�������
		wstring productName = VER_STR_PRODUCTNAME;
		PrintDebugLog(L"productName: ");
		PrintDebugLog(productName);
		PrintDebugLog(L"\n");
		wstring productVersion = VER_STR_PRODUCTVERSION;
		PrintDebugLog(L"productVersion: ");
		PrintDebugLog(productVersion);
		PrintDebugLog(L"\n");
		PrintDebugLog(L"outputFileEncode: ");
		PrintDebugLog(outputFileEncode);
		PrintDebugLog(L"\n");
		PrintDebugLog(L"outputFileFormat: ");
		PrintDebugLog(outputFileFormat);
		PrintDebugLog(L"\n");
		PrintDebugLog(L"outputFileOption: ");
		PrintDebugLog(outputFileOption);
		PrintDebugLog(L"\n");
		PrintDebugLog(L"Expiry Date: ");
		wchar_t stringBuffer[32];
		_stprintf(stringBuffer, TEXT("%ld/%ld/%ld"), periodYear, periodMonth, periodDay);
		PrintDebugLog(stringBuffer);
		PrintDebugLog(L"\n");
		
	}

	currentInventoryManager.setInputInfo(editHardwareNo, edit1, edit2, edit3, edit4, edit5, edit6);
	long res  = currentInventoryManager.getHardwareInfo();
	long res2 = currentInventoryManager.listUp();

	if(res == 0 && res2 == 0){
		//�t�H�[�}�b�g��AdvancedManager�`���̏ꍇ
		if(outputFileFormat == L"AdvancedManager"){
			currentInventoryManager.load();
			currentInventoryManager.save();
		}

		wstring returnString = currentInventoryManager.output();
		wchar_t messageChars [1024];
		//swprintf(messageChars, 1024, L"�C���x���g������ۑ����܂����B\n�ۑ��t�@�C��: \n%s", outputFileName.c_str());
		MessageBox(returnString.c_str(), VER_STR_PRODUCTNAME, MB_OK);
		//MessageBox(_T("�\�t�g�E�F�A���̏o�͂��������܂����B"), edit3.c_str(), MB_OK);

		if(debugLevel > 0){
			swprintf(messageChars, 1024, L"�������O��ۑ����܂����B\n�ۑ��t�@�C��: \n%s", debugLogFileName.c_str());
			MessageBox(messageChars, VER_STR_PRODUCTNAME, MB_OK);
		}

		//default.csv�̃��l�[��
		currentInventoryManager.cleanupDefaultSetting();

		//�O���v���O�����̎��s
		if(iniProgram != L""){
			//wstring processString = L"";
			RunProcess(iniProgram, iniProgramArguments, true);
			//MessageBox(processString.c_str(), VER_STR_PRODUCTNAME, MB_OK);
		}

		//�I��
		CDialog::OnOK();
	}

	if(debugLevel > 0){
		fclose(debugFP);
	}
}

//�F��ς���
//Resource.rc�Őݒ肵���X�^�e�B�b�N�e�L�X�g��ς�����@�i���G�Ȃ̂ō̗p�����j
//http://social.msdn.microsoft.com/Forums/vstudio/en-US/d4bc8cfa-b593-48aa-ad32-a3012363fba9/mfc-clistbox-font-and-color?forum=vcgeneral
//�����F��ݒ肵���e�L�X�g��ǉ�������@
//http://gurigumi.s349.xrea.com/programming/visualcpp/sdk_textcolor.html
//http://www.eonet.ne.jp/~maeda/winc/textcor.htm

void CMyDialog::OnDraw(CDC* pDC)
{
	//�`��̈�̃T�C�Y���w��
	//GetClientRect�Ŏ����I�ɕ`��̈�̃T�C�Y���擾���悤�Ƃ��Ă����܂������Ȃ�
	//���MessageBox���ĂԂƂȂ������܂�����
	/*
	int a = GetClientRect().Height();
	wchar_t messageChars [1024];
	swprintf(messageChars, 1024, L"height: %d", a);
	MessageBox(messageChars, _T("LIP"), MB_OK);
	a = GetClientRect().Height();
	messageChars [1024];
	swprintf(messageChars, 1024, L"height: %d", a);
	MessageBox(messageChars, _T("LIP"), MB_OK);
	*/
	CRect rc = GetClientRect();
	//CRect rc(0,0,300,325);
	CFont font;
	//�{�茧�Ɛ_�ˎs�͈˗��ɂ��t�H���g�T�C�Y��傫��
	//LOGFONT lf = {12, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY,VARIABLE_PITCH | FF_DONTCARE, NULL}; 
	//���ꌧ�d�Z�Z���^�[�͓���؂�Ȃ��̂ňȍ~�͂܂�������
	LOGFONT lf = {10, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY,VARIABLE_PITCH | FF_DONTCARE, NULL}; 
	font.CreateFontIndirect(&lf); 
	pDC->SelectObject(&font);

	//pDC->SetTextColor(::GetSysColor(COLOR_WINDOWTEXT));
	pDC->SetBkColor(GetSysColor(COLOR_BTNFACE));
	pDC->SetTextColor(RGB(150, 150, 150));

	//CString cs = LoadString(IDW_MAIN);
	wchar_t stringBuffer[1024];
	//_stprintf(stringBuffer, TEXT("v%s %s Expiry Date: %ld/%ld/%ld"), productVersion.c_str(), customer.c_str(), periodYearDisplay, periodMonthDisplay, periodDayDisplay);
	_stprintf(stringBuffer, TEXT("v%s %s  Expiry Date: %ld/%ld/%ld"), productVersion.c_str(), customer.c_str(), periodYearDisplay, periodMonthDisplay, periodDayDisplay);
	CString cs = stringBuffer;
	pDC->DrawText(cs, cs.GetLength(), rc, DT_RIGHT|DT_BOTTOM|DT_SINGLELINE);
	//CString cs2 = productVersion.c_str();
	//pDC->DrawText(cs2, cs2.GetLength(), rc, DT_LEFT|DT_BOTTOM|DT_SINGLELINE);
}

/*
LRESULT CALLBACK WinProc(HWND hwnd,UINT msg,WPARAM wp,LPARAM lp)
{
	HDC hdc;
	PAINTSTRUCT ps;
	RECT rect;
	//���オ(10,10�j�A�E�����i200,200)�̗̈���w��
	rect.left=10;
	rect.top=10;
	rect.right=200;
	rect.bottom=200;
	switch(msg){
		case WM_DESTROY:
			PostQuitMessage(0);
			return 0;
		case WM_PAINT:
			hdc=BeginPaint(hwnd,&ps);
			SetTextColor(hdc,RGB(0,0,255));
			//�`��
			DrawText(hdc,L"�e�X�g�łǂ��܂ŕ`�悳��āA���s�����̂��`�F�b�N���Ă��܂��B",-1,&rect,DT_LEFT | DT_WORDBREAK);
			EndPaint(hwnd,&ps);
			return 0;
	}
	return DefWindowProc(hwnd,msg,wp,lp);
}
*/
/*
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    //int wmId, wmEvent;
    PAINTSTRUCT ps;
    HDC hdc;
    LPCTSTR str = TEXT("�e�L�X�g");

    switch (message)
    {
        case WM_PAINT:
            hdc = BeginPaint(hWnd, &ps);
            SetBkColor(hdc, RGB(0, 225, 225));
            SetTextColor(hdc, RGB(255, 0, 0));
            TextOut(hdc, 10, 10, str, _tcslen(str));
            EndPaint(hWnd, &ps);
            break;
        case WM_DESTROY:
            PostQuitMessage(0);
            break;
        default:
            return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}
*/
/*
void CMyDialog::OnButton()
{
	SetDlgItemText(IDC_STATIC3, _T("Button Pressed"));
	TRACE(_T("Button Pressed\n"));
}
*/
/*
void CMyDialog::OnCheck1()
{
	SetDlgItemText(IDC_STATIC3, _T("Check Box 1"));
	TRACE(_T("Check Box 1\n"));
}

void CMyDialog::OnCheck2()
{
	SetDlgItemText(IDC_STATIC3, _T("Check Box 2"));
	TRACE(_T("Check Box 2\n"));
}

void CMyDialog::OnCheck3()
{
	SetDlgItemText(IDC_STATIC3, _T("Check Box 3"));
	TRACE(_T("Check Box 3\n"));
}

void CMyDialog::OnRadio1()
{
	SetDlgItemText(IDC_STATIC3, _T("Radio 1"));
	TRACE(_T("Radio 1\n"));
}

void CMyDialog::OnRadio2()
{
	SetDlgItemText(IDC_STATIC3, _T("Radio 2"));
	TRACE(_T("Radio 2\n"));
}

void CMyDialog::OnRadio3()
{
	SetDlgItemText(IDC_STATIC3, _T("Radio 3"));
	TRACE(_T("Radio 3\n"));
}
*/
