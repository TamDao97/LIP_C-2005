//��������Ă����Ȃ��ƁAmap���g���Ƃ��Ɉȉ��̃G���[���o��
//Run-Time Check Failure #2 - Stack around the variable '_Lk' was corrupted.
//PlattformSDK��crt�ɂ���xtree���ǂ܂�Ă��邱�Ƃ������炵��
//�Q�l:http://unkar.org/r/tech/1204045410
#pragma warning(disable:4786)

//CUI�A�v���ƈႢ�A�����ǉ����Ȃ��ƈȉ��̃G���[���o��
//fatal error C1010: �v���R���p�C�� �w�b�_�[���������ɕs���� EOF ��������܂����B'#include "stdafx.h"' ���\�[�X�ɒǉ����܂�����?
#include "stdafx.h"

#include <windows.h>
#include <stdio.h>
#include <locale.h>
#include <tchar.h>
#include <string.h>
#include <ctype.h>
#include <iostream>
#include <map>
#include <string>
#include <vector>
//tramsform�ɕK�v
#include <algorithm>
#include <cctype>
#include <cstdio>
//�t�@�C�������o���ɕK�v
#include <iostream>
#include <fstream>
//�G���^�[�L�[���������ɓ��̓L�[���󂯕t����Ƃ��ɕK�v
#include <conio.h>
//�t�@�C�����ɓ�����t����Ƃ��ɕK�v
#include <time.h>
//�z�X�g���擾
#include <stdio.h>
#include <WinSock.h>
#pragma comment(lib, "wsock32.lib")
//�l�b�g���[�N�A�_�v�^�擾
#include <Windows.h>
#include <Iphlpapi.h>
#include <Assert.h>
#pragma comment(lib, "iphlpapi.lib")

//WMI�擾
#include <comdef.h>
//#include <comutil.h>
#include <Wbemidl.h>

//_mkdir�ɕK�v
#include <direct.h>

//WMI����CPU�����擾����Ƃ�SocketDesignation�̏d�������J�E���g�ɕK�v
#include <set>

//ORACLE�擾���̃e�L�X�g�t�@�C����͂ɕK�v
//VC2005�ɂ͊܂܂�Ă��炸VC2010�ȍ~�łȂ��ƃ_���@boost�𓱓�����Ǝg���邪�����܂ł��邩�ǂ���
//#include <regex>


//UUID����
//#include <rpc.h>
//#pragma comment(lib ,"rpcrt4.lib")
//https://stackoverflow.com/questions/24365331/how-can-i-generate-uuid-in-c-without-using-boost-library
//http://koolgeeks.seesaa.net/article/117666391.html

//���ꂪ�Ȃ��Ɩ��O�ɂ��������ړ���"std::"������ �C�����Ȃ���΂Ȃ�Ȃ�
//�Ǝ��ɖ��O��Ԃ��`���Ȃ��Ȃ珑���Ă���
//�Q�l:http://www.amris.co.jp/cpp/c16.html
using namespace std;

//OS���ʂɒǉ�
//http://pmakino.jp/tdiary/20090627.html
//GetVersionEx() �Ŏg����萔 (VC++2005 ����`)
#define VER_SUITE_WH_SERVER 0x00008000
 
//GetSystemMetrics() �Ŏg����萔 (VC++2005 ����`)
#define SM_TABLETPC     86
#define SM_MEDIACENTER  87
#define SM_STARTER      88
#define SM_SERVERR2     89
 
//GetProductInfo() �Ŏg����萔 (VC++2005 ����`)
#define PRODUCT_UNDEFINED                           0x00000000 // An unknown product
#define PRODUCT_ULTIMATE                            0x00000001 // Ultimate Edition
#define PRODUCT_HOME_BASIC                          0x00000002 // Home Basic Edition
#define PRODUCT_HOME_PREMIUM                        0x00000003 // Home Premium Edition
#define PRODUCT_ENTERPRISE                          0x00000004 // Enterprise Edition / Windows 10 Enterprise
#define PRODUCT_HOME_BASIC_N                        0x00000005 // Home Basic Edition
#define PRODUCT_BUSINESS                            0x00000006 // Business Edition
#define PRODUCT_STANDARD_SERVER                     0x00000007 // Server Standard Edition
#define PRODUCT_DATACENTER_SERVER                   0x00000008 // Server Datacenter Edition
#define PRODUCT_SMALLBUSINESS_SERVER                0x00000009 // Small Business Server
#define PRODUCT_ENTERPRISE_SERVER                   0x0000000A // Server Enterprise Edition
#define PRODUCT_STARTER                             0x0000000B // Starter Edition
#define PRODUCT_DATACENTER_SERVER_CORE              0x0000000C // Server Datacenter Edition (core installation)
#define PRODUCT_STANDARD_SERVER_CORE                0x0000000D // Server Standard Edition (core installation)
#define PRODUCT_ENTERPRISE_SERVER_CORE              0x0000000E // Server Enterprise Edition (core installation)
#define PRODUCT_ENTERPRISE_SERVER_IA64              0x0000000F // Server Enterprise Edition for Itanium-based Systems
#define PRODUCT_BUSINESS_N                          0x00000010 // Business Edition
#define PRODUCT_WEB_SERVER                          0x00000011 // Web Server Edition
#define PRODUCT_CLUSTER_SERVER                      0x00000012 // Cluster Server Edition
#define PRODUCT_HOME_SERVER                         0x00000013 // Home Server Edition
#define PRODUCT_STORAGE_EXPRESS_SERVER              0x00000014 // Storage Server Express Edition
#define PRODUCT_STORAGE_STANDARD_SERVER             0x00000015 // Storage Server Standard Edition
#define PRODUCT_STORAGE_WORKGROUP_SERVER            0x00000016 // Storage Server Workgroup Edition
#define PRODUCT_STORAGE_ENTERPRISE_SERVER           0x00000017 // Storage Server Enterprise Edition
#define PRODUCT_SERVER_FOR_SMALLBUSINESS            0x00000018 // Server for Small Business Edition
#define PRODUCT_SMALLBUSINESS_SERVER_PREMIUM        0x00000019 // Small Business Server Premium Edition
#define PRODUCT_HOME_PREMIUM_N                      0x0000001A // Home Premium Edition
#define PRODUCT_ENTERPRISE_N                        0x0000001B // Enterprise Edition
#define PRODUCT_ULTIMATE_N                          0x0000001C // Ultimate Edition
#define PRODUCT_WEB_SERVER_CORE                     0x0000001D // Web Server Edition (core installation)
#define PRODUCT_MEDIUMBUSINESS_SERVER_MANAGEMENT    0x0000001E // Windows Essential Business Server Management Server
#define PRODUCT_MEDIUMBUSINESS_SERVER_SECURITY      0x0000001F // Windows Essential Business Server Security Server
#define PRODUCT_MEDIUMBUSINESS_SERVER_MESSAGING     0x00000020 // Windows Essential Business Server Messaging Server
#define PRODUCT_SMALLBUSINESS_SERVER_PRIME          0x00000021 // Server Foundation
#define PRODUCT_HOME_PREMIUM_SERVER                 0x00000022 // Home Premium Server Edition?
#define PRODUCT_SERVER_FOR_SMALLBUSINESS_V          0x00000023 // Server 2008 without Hyper-V for Windows Essential Server Solutions
#define PRODUCT_STANDARD_SERVER_V                   0x00000024 // Server Standard Edition without Hyper-V
#define PRODUCT_DATACENTER_SERVER_V                 0x00000025 // Server Datacenter Edition without Hyper-V
#define PRODUCT_ENTERPRISE_SERVER_V                 0x00000026 // Server Enterprise Edition without Hyper-V
#define PRODUCT_DATACENTER_SERVER_CORE_V            0x00000027 // Server Datacenter Edition without Hyper-V (core installation)
#define PRODUCT_STANDARD_SERVER_CORE_V              0x00000028 // Server Standard Edition without Hyper-V (core installation)
#define PRODUCT_ENTERPRISE_SERVER_CORE_V            0x00000029 // Server Enterprise Edition without Hyper-V (core installation)
#define PRODUCT_HYPERV                              0x0000002A // Microsoft Hyper-V Server
#define PRODUCT_STORAGE_EXPRESS_SERVER_CORE         0x0000002B // Storage Server Express Edition (core installation)
#define PRODUCT_STORAGE_STANDARD_SERVER_CORE        0x0000002C // Storage Server Standard Edition (core installation)
#define PRODUCT_STORAGE_WORKGROUP_SERVER_CORE       0x0000002D // Storage Server Workgroup Edition (core installation)
#define PRODUCT_STORAGE_ENTERPRISE_SERVER_CORE      0x0000002E // Storage Server Enterprise Edition (core installation)
#define PRODUCT_STARTER_N                           0x0000002F // Starter Edition
#define PRODUCT_PROFESSIONAL                        0x00000030 // Professional Edition
#define PRODUCT_PROFESSIONAL_N                      0x00000031 // Professional Edition
#define PRODUCT_SB_SOLUTION_SERVER                  0x00000032
#define PRODUCT_SERVER_FOR_SB_SOLUTIONS             0x00000033
#define PRODUCT_STANDARD_SERVER_SOLUTIONS           0x00000034
#define PRODUCT_STANDARD_SERVER_SOLUTIONS_CORE      0x00000035 // (core installation)
#define PRODUCT_SB_SOLUTION_SERVER_EM               0x00000036
#define PRODUCT_SERVER_FOR_SB_SOLUTIONS_EM          0x00000037
#define PRODUCT_SOLUTION_EMBEDDEDSERVER             0x00000038
#define PRODUCT_SOLUTION_EMBEDDEDSERVER_CORE        0x00000039 // (core installation)
#define PRODUCT_SMALLBUSINESS_SERVER_PREMIUM_CORE   0x0000003F // (core installation)
#define PRODUCT_ESSENTIALBUSINESS_SERVER_MGMT       0x0000003B
#define PRODUCT_ESSENTIALBUSINESS_SERVER_ADDL       0x0000003C
#define PRODUCT_ESSENTIALBUSINESS_SERVER_MGMTSVC    0x0000003D
#define PRODUCT_ESSENTIALBUSINESS_SERVER_ADDLSVC    0x0000003E
#define PRODUCT_CLUSTER_SERVER_V                    0x00000040 // Cluster Server Edition without Hyper-V
#define PRODUCT_EMBEDDED                            0x00000041
//��`�R���ǉ�
#define PRODUCT_MULTIPOINT_STANDARD_SERVER          0x0000004C
#define PRODUCT_MULTIPOINT_PREMIUM_SERVER           0x0000004D
#define PRODUCT_CORE                                0x00000065
#define PRODUCT_CORE_N                              0x00000062
#define PRODUCT_PROFESSIONAL_WMC                    0x00000067
//��`�R���ǉ� 2015-08-24
#define PRODUCT_CORE_COUNTRYSPECIFIC                0x00000063 // Windows 10 Home China
#define PRODUCT_CORE_SINGLELANGUAGE                 0x00000064 // Windows 10 Home Single Language
#define PRODUCT_MOBILE_CORE                         0x00000068 // Windows 10 Mobile
#define PRODUCT_MOBILE_ENTERPRISE                   0x00000085 // Windows 10 Mobile Enterprise
#define PRODUCT_EDUCATION                           0x00000079 // Windows 10 Education
#define PRODUCT_EDUCATION_N                         0x0000007A // Windows 10 Education N
#define PRODUCT_DATACENTER_EVALUATION_SERVER        0x00000050 // Server Datacenter (evaluation installation)
#define PRODUCT_ENTERPRISE_E                        0x00000046 // Windows 10 Enterprise E
#define PRODUCT_ENTERPRISE_N_EVALUATION             0x00000054 // Windows 10 Enterprise N (evaluation installation)
#define PRODUCT_ENTERPRISE_EVALUATION               0x00000048 // Windows 10 Enterprise (evaluation installation)

typedef void (WINAPI *PGNSI)(LPSYSTEM_INFO);

//http://msdn.microsoft.com/en-us/library/windows/desktop/ms724358(v=vs.85).aspx
typedef BOOL (WINAPI *PGPI)(DWORD, DWORD, DWORD, DWORD, PDWORD);

//http://msdn.microsoft.com/en-us/library/windows/desktop/bb773263(v=vs.85).aspx
/*
typedef struct _DllVersionInfo {
  DWORD cbSize;
  DWORD dwMajorVersion;
  DWORD dwMinorVersion;
  DWORD dwBuildNumber;
  DWORD dwPlatformID;
} DLLVERSIONINFO;
*/
/*
typedef struct DLLVERSIONINFO {
  DWORD cbSize;
  DWORD dwMajorVersion;
  DWORD dwMinorVersion;
  DWORD dwBuildNumber;
  DWORD dwPlatformID;
};
*/
//http://msdn.microsoft.com/en-us/library/windows/desktop/bb776404(v=vs.85).aspx
//typedef HRESULT (CALLBACK* DLLGETVERSIONINFOPROC)(DLLVERSIONINFO *);

//_tolower�Ŏg��
//#include <afx.h>

//WMI�擾�̂��ߒǉ�
#ifndef _WIN32_DCOM
#   define _WIN32_DCOM
#endif
//#pragma comment(lib, "wbemuuid.lib")
//���ŃG���[���o��
//fatal error LNK1103: debugging information corrupt; recompile module Error executing link.exe.
//�Â�SDK�łȂ���VC++2005�ɂ͑Ή����Ă��Ȃ��͗l
//��������ł��Ȃ��̂ŁA���ڏ���������
//http://stackoverflow.com/questions/6568467/vc-compatibility-problem
GUID MY_CLSID_WbemLocator = {0x4590f811, 0x1d3a, 0x11d0, {0x89, 0x1f, 0x00, 0xaa, 0x00, 0x4b, 0x2e, 0x24}};
GUID MY_IID_IWbemLocator = {0xdc12a687, 0x737f, 0x11cf, {0x88, 0x4d, 0x00, 0xaa, 0x00, 0x4b, 0x2e, 0x24}};
 
_COM_SMARTPTR_TYPEDEF(IWbemLocator, __uuidof(IWbemLocator));
_COM_SMARTPTR_TYPEDEF(IWbemServices, __uuidof(IWbemServices));
_COM_SMARTPTR_TYPEDEF(IEnumWbemClassObject, __uuidof(IEnumWbemClassObject));
_COM_SMARTPTR_TYPEDEF(IWbemClassObject, __uuidof(IWbemClassObject));



//���O���[�o���ϐ�

//OS��64bit��32bit�����x�������������������ʂ�ۊ�
//-1:������
// 0:32bit
// 1:64bit
long isWow64Flag = -1;

//�f�o�b�O�i�R���\�[���ɏڂ��������o�͂���j���[�h���ۂ�
//0:�f�o�b�O���[�h�ł͂Ȃ��ʏ탂�[�h
//1:�f�o�b�O���[�h
long debugLevel = 0;
//�ۑ��t�@�C�����i���Ƃŏ��������j�����
//wstring outputFileName = L"InventoryManager.csv";
//�f�o�b�O���O�̕ۑ��t�@�C����
wstring debugLogFileName = L"debug.log";
//�f�o�b�O���O�̕ۑ��t�@�C���̃|�C���^
FILE* debugFP = NULL;

//�C���x���g����PC���ɕۑ�����ۂ̃L�[���[�h
unsigned int inventorySaveStringKey = 116; 

//SoftwarePicker.h�p�̏o�͐ݒ�
//�ۑ��t�@�C���̃G���R�[�h�i�f�t�H���g��UTF-8�A���̑�UTF-8N�ASJIS���j
wstring outputFileEncode = VER_STR_OUTPUTFILEENCODE;
//wstring outputFileEncode = L"SJIS";
//�ۑ��t�@�C���̌`���i"SIMPLE","SARMS"�j��SARMS�̂Ƃ��A�G���R�[�h��SJIS����
wstring outputFileFormat = VER_STR_OUTPUTFILEFORMAT;
//wstring outputFileFormat = L"SIMPLE";
//MORE:�ۑ��t�@�C����SIMPLE�`���̂Ƃ��APC�x���_�[�APC�@��ACPU�^�C�v�̗��ǉ�����
//wstring outputFileOption = L"MORE";
wstring outputFileOption = VER_STR_OUTPUTFILEOPTION;

//���샂�[�h�i�f�t�H���g��FALSE�j
BOOL silentMode = FALSE;

//�E�B���h�E�̃��x��
wstring labelHardwareNo = VER_STR_LABEL_HARDWARE_NO;
wstring label1 = VER_STR_LABEL1;
wstring label2 = VER_STR_LABEL2;
wstring label3 = VER_STR_LABEL3;
wstring label4 = VER_STR_LABEL4;
wstring label5 = VER_STR_LABEL5;
wstring label6 = VER_STR_LABEL6;

//���͗��̗L���E����
bool isInputHardwareNoValid = true;
bool isInput1Valid = true;
bool isInput2Valid = true;
bool isInput3Valid = true;
bool isInput4Valid = true;
bool isInput5Valid = true;
bool isInput6Valid = true;

//ini�t�@�C���̃p�X
wstring iniPath = L".\\LIP.ini";
//ini�t�@�C���̃f�[�^�̃p�X
wstring outputPath = L".\\";
wstring hardwareNoPrefix = L"";
unsigned int hardwareNoDigits = 0;
wstring iniProgram = L"";
wstring iniProgramArguments = L"";
unsigned int iniOracleCompatible = 0;

//default.csv�t�@�C���̃p�X
wstring defaultCSVPath = L".\\default.csv";



/*
�Q�l

RegEnumValue�̃T���v��
http://d.hatena.ne.jp/s-kita/20120331/1333174926

RegQueryValueEx�̃T���v��
http://d.hatena.ne.jp/s-kita/20120408/1333895565

RegEnumKeyEx�̃T���v��
http://d.hatena.ne.jp/s-kita/20120401/1333290793

64bitOS���ʂ̃T���v��
http://msdn.microsoft.com/en-us/library/ms684139(VS.85).aspx

�R�}���h���C���I�v�V������͂̃T���v��
http://masudahp.web.fc2.com/cl/tool/tool01.html
http://codezine.jp/article/detail/4086
http://www.opensource.apple.com/source/awk/awk-1.2/main.c
http://c-faq.com/misc/argv.html
http://sa.eei.eng.osaka-u.ac.jp/eeisa003/tani_prog/org_option.html

�}���`�o�C�g�֘A
http://cx5software.com/article_vcpp_unicode/
http://victreal.com/Junk/_T/

���������[�N�΍�
http://www.aerith.net/cpp/safe-coding-j.html

���W�X�g���̍ő啶�������̏��
http://support.microsoft.com/kb/256986/ja

LocalizedString
http://msdn.microsoft.com/en-us/library/windows/desktop/dd374120(v=vs.85).aspx

LocalizedString�̓ǂݍ���
http://stackoverflow.com/questions/13862278/getting-string-from-a-dll-string-table
http://stackoverflow.com/questions/9324563/how-does-one-read-a-localized-name-from-the-registry

�v���O�����Ƌ@�\�̏o�́iVBS�j
http://social.technet.microsoft.com/Forums/windows/en-US/e1599843-f5c8-45ca-9cdf-56d5e5ccf43b/addremove-programs-export-to-text-file?forum=itproxpsp


*/
/*
//�ŐV�ł�Windows Platform SDK�iV7.1�j��Microsoft SDKs\Windows\v7.1\Include\WinReg.h����ڐA
WINADVAPI
LSTATUS
APIENTRY
RegLoadMUIStringW (
                    __in                    HKEY        hKey,
                    __in_opt                LPCWSTR    pszValue,
                    __out_bcount_opt(cbOutBuf)  LPWSTR     pszOutBuf,
                    __in                    DWORD       cbOutBuf,
                    __out_opt               LPDWORD     pcbData,
                    __in                    DWORD       Flags,   
                    __in_opt                LPCWSTR    pszDirectory
                    );

*/

//�f�o�b�O���O���o��
//http://program.station.ez-net.jp/special/vc/basic/function/stdarg.asp
//http://www.kde.cs.tut.ac.jp/~atsushi/?p=117
void PrintDebugLog(wstring text){
	/*
    FILE* fp = NULL;
	//BOM�t����UTF-8�ŏo��
	//a+���[�h�Ńt�@�C�������ɒǉ� http://msdn.microsoft.com/ja-jp/library/z5hh6ee9(v=vs.90).aspx
	LONG status;
	status = _wfopen_s( &fp, debugLogFileName.c_str(), L"a+,ccs=UTF-8" );
	//�t�@�C���������ƊJ�������Ԃ�l���`�F�b�N���Ă����Ȃ��Ɠr���ŃG���[��������
	//���������������ނ̂���肩�H�o�b�t�@�ɗ��߂Ă��珑�����񂾂ق����������ǂ���
	if(status == 0){
		_ftprintf(fp, L"%s", text.c_str());
		//���Ă����Ȃ��Ɓu���̃v���Z�X�Ŏg�p���v�ƂȂ��ăG���[����
		fclose(fp);
	}
	*/
	_ftprintf(debugFP, L"%s", text.c_str());
}

//�����񒆂̉p�啶�������������������ɕϊ�
_TCHAR *StrToLower(_TCHAR *acString){
    for (_TCHAR* pc = acString; *pc; pc++) {
        *pc = tolower(*pc);
    }
    return (acString);
}

//�����񒆂̉p��������啶���ɕϊ�
_TCHAR *StrToUpper(_TCHAR *acString){
    //_TCHAR acString[] = TEXT("aBcDe");
    for (_TCHAR* pc = acString; *pc; pc++) {
        *pc = _toupper(*pc);
    }
    //_tprintf(TEXT("%s\n"), acString);
	return (acString);
}

//������̑啶�����������ɕϊ�
void toLowerCase(wstring& string){
	transform(string.begin (), string.end (), string.begin (), towlower);
}

//������̑啶���������𖳎�������r
bool compareToIgnoreCase(wstring string1, wstring string2){
	toLowerCase(string1);
	toLowerCase(string2);
	if(string1 == string2){
		return true;
	}else{
		return false;
	}
}


//������̒u��
//http://www.geocities.jp/eneces_jupiter_jp/cpp1/010-055.html
//�g����
//wstring String = L"abc012abc012abc012";
//String = Replace(String, L"012", L"defg");
wstring Replace(wstring String1, wstring String2, wstring String3){
    wstring::size_type Pos(String1.find(String2));

    while(Pos != wstring::npos)
    {
        String1.replace(Pos, String2.length(), String3);
        Pos = String1.find(String2, Pos + String3.length());
    }

    return String1;
}


//���p�J�i��S�p�J�i�ɒu��
wstring HankakuKana2ZenkakuKana(wstring targetString){
	//���p�A�S�p�̎Q�l https://gist.github.com/punytan/379007
	wstring hankakuKana1 = L"���������������������������������������������������������������";
	wstring zenkakuKana1 = L"�B�u�v�A�E���@�B�D�F�H�������b�[�A�C�E�G�I�J�L�N�P�R�T�V�X�Z�\�^�`�c�e�g�i�j�k�l�m�n�q�t�w�z�}�~���������������������������J�K";
	wstring hankakuKana2 = L"�޶޷޸޹޺޻޼޽޾޿�������������������������������";
	wstring zenkakuKana2 = L"���K�M�O�Q�S�U�W�Y�[�]�_�a�d�f�h�o�r�u�x�{�p�s�v�y�|";
	wstring tempHankakuKana;
	wstring tempZenkakuKana;

	for(int i = 0; i < (int)zenkakuKana2.size(); ++i){
		tempHankakuKana = hankakuKana2.substr(i*2, 2);
		tempZenkakuKana = zenkakuKana2[i];
		targetString = Replace(targetString, tempHankakuKana, tempZenkakuKana);
	}
	
	for(int i = 0; i < (int)zenkakuKana1.size(); ++i){
		tempHankakuKana = hankakuKana1[i];
		tempZenkakuKana = zenkakuKana1[i];
		targetString = Replace(targetString, tempHankakuKana, tempZenkakuKana);
	}
	
    return targetString;
}

bool IsNumericString(wstring string){
	bool isNumericString = true;
	for(int i=0; i<(int)string.size(); ++i){
		if(48 <= string[i] && string[i] <= 57){
			//�������Ȃ�
		}else{
			//���l����Ȃ��Ȃ�����I��
			isNumericString = false;
			break;
		}
	}
	return isNumericString;
}
/*
//�����񒆂̍ŏ��ɏo�Ă������l����w�茅���Ń[���l�߂���
wstring FirstNumericInStringZeroPadding(wstring targetString, wstring digits){
	//digits�������݂̂ō\�����ꂽ�����񂩃`�F�b�N
	//http://qiita.com/edo1z/items/da66e28e206d2b01157e
	//https://msdn.microsoft.com/en-us/library/fcc4ksh8(v=vs.80).aspx
	if(all_of(digits.cbegin(), digits.cend(), _istdigit)){
		//�������Ȃ�
	}else{
		//�����݂̂ō\�����ꂽ������ł͂Ȃ��ꍇ�A"1"�Ƃ���
		digits=L"1";
	}
	//sprintfSetting = L"%010d"
	wstring sprintfSetting = L"%0" + digits + L"d";

	int numericFirstIndex = -1;
	int numericLastIndex = -1;
	for(int i=0; i<(int)targetString.size(); ++i){
		//if(48 <= targetString[i].c_str()[0] && targetString[i].c_str()[0] <= 57){
		if(48 <= targetString[i] && targetString[i] <= 57){
			//���߂Đ��l���o�Ă�����FirstIndex�ɃZ�b�g
			if(numericFirstIndex == -1){
				numericFirstIndex = i;
			}
			//���l�Ȃ�LastIndex�ɂ��Z�b�g
			numericLastIndex = i;
		}else{
			//���l����Ȃ��Ȃ�����I��
			if(numericFirstIndex > -1){
				break;
			}
		}
	}
	int substringInteger = _wtoi(targetString.substr(numericFirstIndex, numericLastIndex - numericFirstIndex +1).c_str());
    //_tprintf(TEXT("test: %d\n"), substringInteger);

	wchar_t zeroPaddingSubstringIntegerChars [12];
	swprintf(zeroPaddingSubstringIntegerChars, 12, sprintfSetting, substringInteger);
    //_tprintf(TEXT("test: %s\n"), zeroPaddingSubstringIntegerChars);
	targetString.replace(numericFirstIndex, numericLastIndex - numericFirstIndex + 1, zeroPaddingSubstringIntegerChars);
	
    return targetString;
}
*/
//�����񒆂̐��l���S��10���Ń[���l�߂���@�G���[��������̂Œ������ƁI�I
wstring NumericInStringZeroPadding(wstring targetString){
	int numericFirstIndex = -1;
	int numericLastIndex = -1;
	for(int i=(int)targetString.size()-1; i>=0; --i){
	//for(int i=0; i<(int)targetString.size(); ++i){
		//if(48 <= targetString[i].c_str()[0] && targetString[i].c_str()[0] <= 57){
		if(48 <= targetString[i] && targetString[i] <= 57){
			//���߂Đ��l���o�Ă�����LastIndex�ɃZ�b�g
			if(numericLastIndex == -1){
				numericLastIndex = i;
			}
			//���l�Ȃ�FirstIndex�ɂ��Z�b�g
			numericFirstIndex = i;
		}else{
			//���l����Ȃ��Ȃ�����I��
			if(numericLastIndex > -1){
				int substringInteger = _wtoi(targetString.substr(numericFirstIndex, numericLastIndex - numericFirstIndex +1).c_str());
				//_tprintf(TEXT("test: %d\n"), substringInteger);

				wchar_t zeroPaddingSubstringIntegerChars [12];
				swprintf(zeroPaddingSubstringIntegerChars, 12, L"%010d", substringInteger);
				//_tprintf(TEXT("test: %s\n"), zeroPaddingSubstringIntegerChars);
				targetString.replace(numericFirstIndex, numericLastIndex - numericFirstIndex + 1, zeroPaddingSubstringIntegerChars);
				numericFirstIndex = -1;
				numericLastIndex = -1;
			}
		}
	}
	int substringInteger = _wtoi(targetString.substr(numericFirstIndex, numericLastIndex - numericFirstIndex +1).c_str());
    //_tprintf(TEXT("test: %d\n"), substringInteger);

	wchar_t zeroPaddingSubstringIntegerChars [12];
	swprintf(zeroPaddingSubstringIntegerChars, 12, L"%010d", substringInteger);
    _tprintf(TEXT("test: %s\n"), zeroPaddingSubstringIntegerChars);
	targetString.replace(numericFirstIndex, numericLastIndex - numericFirstIndex +1, zeroPaddingSubstringIntegerChars);
	
    return targetString;
}

//�����񂪐��������ō\������Ă��邩�`�F�b�N
BOOL isNumericString(wstring targetString)
{
	for(int i=0; i<(int)targetString.size(); ++i){
		//if(48 <= targetString[i].c_str()[0] && targetString[i].c_str()[0] <= 57){
		if(48 <= targetString[i] && targetString[i] <= 57){
			//�����Ȃ̂�OK
		}else{
			return FALSE;
		}
	}
	return TRUE;
}

//�����񂪉p���������ō\������Ă��邩�`�F�b�N
BOOL isAlphanumericString(wstring targetString)
{
	for(int i=0; i<(int)targetString.size(); ++i){
		//if(48 <= targetString[i].c_str()[0] && targetString[i].c_str()[0] <= 57){
		if(48 <= targetString[i] && targetString[i] <= 57){
			//�����Ȃ̂�OK
		}else if(65 <= targetString[i] && targetString[i] <= 90){
			//�p���啶���Ȃ̂�OK
		}else if(97 <= targetString[i] && targetString[i] <= 122){
			//�p���������Ȃ̂�OK
		}else{
			return FALSE;
		}
	}
	return TRUE;
}

//�O���v���O�����̎��s
//http://tooljp.com/language/C-Languate/sample-code/CreateProcess-sample-code.html
//http://www.wabiapp.com/WabiSampleSource/windows/create_process.html
//�R�}���h�A�����̓n����
//https://support.microsoft.com/en-us/help/175986/info-understanding-createprocess-and-command-line-arguments
void RunProcess(wstring programString, wstring argumentsString, bool showWindow){

	if(debugLevel > 1){
		PrintDebugLog(L"RunProcess programString=");
		PrintDebugLog(programString);
		PrintDebugLog(L"\n");
	}
	PROCESS_INFORMATION p;
	STARTUPINFO s;
	ZeroMemory(&s,sizeof(s));
	s.cb = sizeof(s);
	//LPTSTR lpArg;
	//lpArg = commandString;
	argumentsString = programString + argumentsString;

	int ret;
	if(showWindow){
		ret = CreateProcess((LPTSTR)programString.c_str(),(LPTSTR)argumentsString.c_str(),NULL,NULL,FALSE,NORMAL_PRIORITY_CLASS,NULL,NULL,&s,&p);
	}else{
		ret = CreateProcess((LPTSTR)programString.c_str(),(LPTSTR)argumentsString.c_str(),NULL,NULL,FALSE,NORMAL_PRIORITY_CLASS | CREATE_NO_WINDOW,NULL,NULL,&s,&p);
	}
	if(ret == 0){
		if(debugLevel > 1){
			PrintDebugLog(L"Error:CreateProcess error.\n");
		}
	}else{
		CloseHandle(p.hThread);
		DWORD ret2 = WaitForSingleObject(p.hProcess,1000);
		if(ret2 == WAIT_TIMEOUT){
			if(debugLevel > 1){
				PrintDebugLog(L"Warning:CreateProcess WAIT_TIMEOUT.\n");
			}
		}else if(ret2 == WAIT_FAILED){
			if(debugLevel > 1){
				PrintDebugLog(L"Error:CreateProcess WAIT_FAILED.\n");
			}
		}else if(ret2 == WAIT_OBJECT_0){
			if(debugLevel > 1){
				PrintDebugLog(L"CreateProcess WAIT_OBJECT_0.\n");
			}
		}
		CloseHandle(p.hProcess);
	}

}
/*
//�O���v���O���������s���Č��ʂ𓾂�
//https://ameblo.jp/cpp-prg/entry-10232970547.html
wstring RunProcess2(wstring programString, wstring argumentsString){

	if(debugLevel > 1){
		PrintDebugLog(L"RunProcess programString=");
		PrintDebugLog(programString);
		PrintDebugLog(L"\n");
	}
	//�p�C�v�쐬
	SECURITY_ATTRIBUTES sa={0}; 
	sa.nLength = sizeof( SECURITY_ATTRIBUTES );
	sa.lpSecurityDescriptor = NULL;
	sa.bInheritHandle = TRUE;
	HANDLE hReadPipe, hWritePipe ;
	CreatePipe( &hReadPipe, &hWritePipe, &sa, 50000 );
	//�N�����ݒ�
	STARTUPINFO si={0}; 
	si.cb = sizeof( STARTUPINFO );
	si.dwFlags = STARTF_USESTDHANDLES; 
	si.wShowWindow = SW_HIDE; 
	si.hStdOutput = hWritePipe; 
	PROCESS_INFORMATION pi;
	//PROCESS_INFORMATION p;
	//STARTUPINFO s;
	//ZeroMemory(&s,sizeof(s));
	//s.cb = sizeof(s);
	argumentsString = programString + argumentsString;

	//if(!CreateProcess(NULL, &cmd[0] , NULL, NULL, TRUE, CREATE_NO_WINDOW, NULL, NULL, &si, &pi ))
	//int ret = CreateProcess((LPTSTR)programString.c_str(),(LPTSTR)argumentsString.c_str(),NULL,NULL,FALSE,NORMAL_PRIORITY_CLASS,NULL,NULL,&si,&pi);
	int ret = CreateProcess((LPTSTR)programString.c_str(),(LPTSTR)argumentsString.c_str(),NULL,NULL,TRUE,CREATE_NO_WINDOW,NULL,NULL,&si,&pi);
	if(ret == 0){
		if(debugLevel > 1){
			PrintDebugLog(L"Error:CreateProcess error.\n");
		}
	}else{

		//�o�͂�ۑ�
		string output;
		unsigned long n, m=0, r, errcnt=0, flg=0;
		while (!flg){
		if(WaitForSingleObject( pi.hProcess, 100)==WAIT_OBJECT_0)flg=1;
		PeekNamedPipe( hReadPipe, NULL, 0, NULL, &n, NULL );
		if( n == 0 ) { errcnt++; if(errcnt>30) break; else continue; }
		errcnt=0;
		if( n+m > output.capacity() ) output.reserve(2*(n>m?n:m)); 
		output.resize(m+n);
		ReadFile( hReadPipe, &output[m], n, &r, NULL );
		m+=r; }
		//�I������
		CloseHandle( pi.hThread );
		CloseHandle( hWritePipe );
		CloseHandle( hReadPipe );

		//CloseHandle(p.hThread);
		DWORD ret2 = WaitForSingleObject(pi.hProcess,1000);
		if(ret2 == WAIT_TIMEOUT){
			if(debugLevel > 1){
				PrintDebugLog(L"Warning:CreateProcess WAIT_TIMEOUT.\n");
			}
		}else if(ret2 == WAIT_FAILED){
			if(debugLevel > 1){
				PrintDebugLog(L"Error:CreateProcess WAIT_FAILED.\n");
			}
		}else if(ret2 == WAIT_OBJECT_0){
			if(debugLevel > 1){
				PrintDebugLog(L"CreateProcess WAIT_OBJECT_0.\n");
			}
		}
		CloseHandle( pi.hProcess );
		//CloseHandle(p.hProcess);
		wstring output2 = L"";
		output2 = output.c_str();
		return output2;
	}

}
*/
typedef BOOL (WINAPI *LPFN_ISWOW64PROCESS) (HANDLE, PBOOL);

LPFN_ISWOW64PROCESS fnIsWow64Process;

//64bitOS���ۂ��̔���
BOOL IsWow64(){
	//�O���[�o���ϐ��Ɍ��ʂ��c���Ă���΂�����g�p
	if(isWow64Flag == 1){
		return TRUE;
	}else if(isWow64Flag == 0){
		return FALSE;
	//�O���[�o���ϐ��Ɍ��ʂ��c���Ă��Ȃ��ꍇ
	}else{
	    BOOL bIsWow64 = FALSE;

		//IsWow64Process is not available on all supported versions of Windows.
		//Use GetModuleHandle to get a handle to the DLL that contains the function
		//and GetProcAddress to get a pointer to the function if available.

		fnIsWow64Process = (LPFN_ISWOW64PROCESS) GetProcAddress(GetModuleHandle(TEXT("kernel32")),"IsWow64Process");

		if(NULL != fnIsWow64Process){
			if (!fnIsWow64Process(GetCurrentProcess(),&bIsWow64)){
				//handle error
				if(debugLevel > 1){
					PrintDebugLog(L"Error:fnIsWow64Process error.\n");
				}
			}
		}
		/*
		if(debugLevel > 1){
			if(bIsWow64){
				PrintDebugLog(L"IsWow64():TRUE.\n");
			}else{
				PrintDebugLog(L"IsWow64():FALSE.\n");
			}
		}
		*/

		if(bIsWow64){
			isWow64Flag = 1;
		}else{
			isWow64Flag = 0;
		}

		return bIsWow64;
	}
}


//WMI�̎擾
wstring getWMI(wstring WMIQueryString, wstring WMIObjectString){

	//�V���A���ԍ��̎擾
	//http://www.programmershare.com/1579747/ OS���擾�܂ŃR�[�h�����Ă��邪�뎚�������C������Ԃ�����
	//http://togarasi.wordpress.com/2009/12/25/c-%E3%81%A7-wmi-%E3%82%AF%E3%83%A9%E3%82%B9%E3%82%92%E4%BD%BF%E3%81%86/�@�����x�������G���[
	//http://msdn.microsoft.com/en-us/library/aa390423(v=vs.85).aspx

	//�u�T�C�h�o�C�T�C�h�\�����������Ȃ����߁A�A�v���P�[�V�������J�n�ł��܂���ł����v�G���[���\�������
	//�X�^�e�B�b�N�����N���邱�Ƃŉ���
	//http://d.hatena.ne.jp/soappp/20110717/1310914195
	//���W�F�N�g�̃v���p�e�B�ŁA[C/C++]-[�R�[�h����]-[�����^�C�����C�u����]�ŁA[�}���`�X���b�h DLL (/MD)] �̑���� [�}���`�X���b�h (/MT)] ��I��

	//_CrtDbgReportW�]�X�̃G���[���o��
	//http://dixq.net/forum/viewtopic.php?f=3&t=1228
	//�f�o�b�O���Ȃ�/MT�ł͂Ȃ�/MTd�Ƃ��邱�ƂŎ���
    HRESULT hres;

    // Step 1: --------------------------------------------------
    // Initialize COM. ------------------------------------------
    hres =  CoInitializeEx(0, COINIT_MULTITHREADED); 
    if (FAILED(hres)){
        //cout << "Failed to initialize COM library. Error code = 0x" 
        //    << hex << hres << endl;
		if(debugLevel > 1){
			PrintDebugLog(L"Error:Failed to initialize COM library. Error code :");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%d"), hres);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
        return L"";               // Program has failed.
    }

    // Step 2: --------------------------------------------------
    // Set general COM security levels --------------------------
    hres =  CoInitializeSecurity(
        NULL, 
        -1,                          // COM authentication
        NULL,                        // Authentication services
        NULL,                        // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
        RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
        NULL,                        // Authentication info
        EOAC_NONE,                   // Additional capabilities 
        NULL                         // Reserved
        );

	//RPC_E_TOO_LATE -2147417831
	//http://forums.codeguru.com/showthread.php?404520-Can-RPC_E_TOO_LATE-be-ignored
    if (hres != -2147417831 && FAILED(hres)){
        //cout << "Failed to initialize security. Error code = 0x" 
        //    << hex << hres << endl;
		if(debugLevel > 1){
			PrintDebugLog(L"Error:Failed to initialize security. Error code :");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%d"), hres);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
        CoUninitialize();
        return L"";               // Program has failed.
    }
    
    // Step 3: ---------------------------------------------------
    // Obtain the initial locator to WMI -------------------------
    IWbemLocator *pLoc = NULL;

    hres = CoCreateInstance(
        MY_CLSID_WbemLocator,             
        0, 
        CLSCTX_INPROC_SERVER, 
        MY_IID_IWbemLocator, (LPVOID *) &pLoc);
 
    if (FAILED(hres)){
        //cout << "Failed to create IWbemLocator object."
        //    << " Err code = 0x"
        //    << hex << hres << endl;
		if(debugLevel > 1){
			PrintDebugLog(L"Error:Failed to create IWbemLocator object. Error code :");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%d"), hres);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
        CoUninitialize();
        return L"";               // Program has failed.
    }

    // Step 4: -----------------------------------------------------
    // Connect to WMI through the IWbemLocator::ConnectServer method
    IWbemServices *pSvc = NULL;
 
    // Connect to the root\cimv2 namespace with
    // the current user and obtain pointer pSvc
    // to make IWbemServices calls.
    hres = pLoc->ConnectServer(
         _bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
         NULL,                    // User name. NULL = current user
         NULL,                    // User password. NULL = current
         0,                       // Locale. NULL indicates current
         NULL,                    // Security flags.
         0,                       // Authority (for example, Kerberos)
         0,                       // Context object 
         &pSvc                    // pointer to IWbemServices proxy
         );
    
    if (FAILED(hres)){
        //cout << "Could not connect. Error code = 0x" 
        //     << hex << hres << endl;
		if(debugLevel > 1){
			PrintDebugLog(L"Error:Could not connect. Error code :");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%d"), hres);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
        pLoc->Release();     
        CoUninitialize();
        return L"";               // Program has failed.
    }

    //cout << "Connected to ROOT\\CIMV2 WMI namespace" << endl;

    // Step 5: --------------------------------------------------
    // Set security levels on the proxy -------------------------
    hres = CoSetProxyBlanket(
       pSvc,                        // Indicates the proxy to set
       RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
       RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
       NULL,                        // Server principal name 
       RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
       RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
       NULL,                        // client identity
       EOAC_NONE                    // proxy capabilities 
    );

    if (FAILED(hres)){
        //cout << "Could not set proxy blanket. Error code = 0x" 
        //    << hex << hres << endl;
		if(debugLevel > 1){
			PrintDebugLog(L"Error:Could not set proxy blanket. Error code :");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%d"), hres);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
        pSvc->Release();
        pLoc->Release();     
        CoUninitialize();
        return L"";               // Program has failed.
    }

    // Step 6: --------------------------------------------------
    // Use the IWbemServices pointer to make requests of WMI ----
    IEnumWbemClassObject* pEnumerator = NULL;
	//wstring WMIQuery = L"SELECT * FROM Win32_ComputerSystemProduct";
/*
	hres = pSvc->ExecQuery(
        bstr_t("WQL"), 
        bstr_t("SELECT * FROM Win32_ComputerSystemProduct"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, 
        NULL,
        &pEnumerator);
*/
	hres = pSvc->ExecQuery(
        bstr_t("WQL"), 
        bstr_t(WMIQueryString.c_str()),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, 
        NULL,
        &pEnumerator);
    
    if (FAILED(hres)){
        //cout << "Query for operating system name failed."
        //    << " Error code = 0x" 
        //    << hex << hres << endl;
		if(debugLevel > 1){
			PrintDebugLog(L"Error:Query for operating system name failed. Error code :");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%d"), hres);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return L"";               // Program has failed.
    }

    // Step 7: -------------------------------------------------
    // Get the data from the query in step 6 -------------------
    IWbemClassObject *pclsObj;
    ULONG uReturn = 0;
	wstringstream wsstream;

    while (pEnumerator){
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, 
            &pclsObj, &uReturn);

        if(0 == uReturn){
            break;
        }

        VARIANT vtProp;

        // Get the value of the Name property
		//wstring WMIObjectName = L"IdentifyingNumber";
		hr = pclsObj->Get(BSTR(WMIObjectString.c_str()), 0, &vtProp, 0, 0);
		//�G���[�R�[�h�̔�������Ă���
		//https://www.codeproject.com/Articles/19433/WMI-Sample-Get-IP-address
		if(!FAILED(hr)){ // && vtProp.bstrVal){
			//bstr��NULL�̂Ƃ��Awstring�ɕϊ�����ƃG���[��������̂Ŕ��������
			//http://www.sutosoft.com/room/archives/000355.html
			//���ł͂Ȃ������̗p
			//https://stackoverflow.com/questions/22525890/c-getting-wmi-array-data-from-the-local-computer
			if(vtProp.vt==VT_NULL || vtProp.vt==VT_EMPTY){
				//�������Ȃ�
			}else if(vtProp.vt & VT_ARRAY){
				//�z��̂Ƃ��̏����͌����_�ŕK�v�Ȃ��̂ŉ������Ȃ��@�����ꂪ�Ȃ��ƃG���[��������
			}else{
				//outputString += vtProp.bstrVal;
				//outputString += L";";
				_bstr_t bstr = vtProp;
				wsstream << bstr << L";";
			}
		}
        //wcout << " IdentifyingNumber : " << vtProp.bstrVal << endl;
        VariantClear(&vtProp);

        pclsObj->Release();
    }

	wstring outputString = wsstream.str();
	if(outputString.size() > 0 && outputString.substr(outputString.size()-1, 1) == L";"){
		outputString = outputString.substr(0, outputString.size()-1);
	}

    // Cleanup
    pSvc->Release();
    pLoc->Release();
    pEnumerator->Release();
    if(!pclsObj) pclsObj->Release();
    CoUninitialize();

    return outputString;   // Program successfully completed.
}

//WMI����CPU�Ɋւ�����̎擾
//CPU���h����\�P�b�g��NumberOfSockets�͎擾���@������
long getCPUInfo(wstring &ProcessorName, wstring &ProcessorMaxClockSpeed, wstring &NumberOfProcessors, wstring &NumberOfCores, wstring &NumberOfLogicalProcessors){
    HRESULT hres;

	// Initialize COM. ------------------------------------------
    hres =  CoInitializeEx(0, COINIT_MULTITHREADED); 
    if (FAILED(hres)){
		if(debugLevel > 1){
			PrintDebugLog(L"Error:Failed to initialize COM library. Error code :");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%d"), hres);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
        return -1;               // Program has failed.
    }

    // Set general COM security levels --------------------------
    hres =  CoInitializeSecurity(
        NULL, 
        -1,                          // COM authentication
        NULL,                        // Authentication services
        NULL,                        // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
        RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
        NULL,                        // Authentication info
        EOAC_NONE,                   // Additional capabilities 
        NULL                         // Reserved
        );

	//RPC_E_TOO_LATE -2147417831
	//http://forums.codeguru.com/showthread.php?404520-Can-RPC_E_TOO_LATE-be-ignored
    if (hres != -2147417831 && FAILED(hres)){
		if(debugLevel > 1){
			PrintDebugLog(L"Error:Failed to initialize security. Error code :");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%d"), hres);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
        CoUninitialize();
        return -1;               // Program has failed.
    }
    
    // Obtain the initial locator to WMI -------------------------
    IWbemLocator *pLoc = NULL;

    hres = CoCreateInstance(
        MY_CLSID_WbemLocator,             
        0, 
        CLSCTX_INPROC_SERVER, 
        MY_IID_IWbemLocator, (LPVOID *) &pLoc);
 
    if (FAILED(hres)){
		if(debugLevel > 1){
			PrintDebugLog(L"Error:Failed to create IWbemLocator object. Error code :");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%d"), hres);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
        CoUninitialize();
        return -1;               // Program has failed.
    }

    // Connect to WMI through the IWbemLocator::ConnectServer method
    IWbemServices *pSvc = NULL;
    hres = pLoc->ConnectServer(
         _bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
         NULL,                    // User name. NULL = current user
         NULL,                    // User password. NULL = current
         0,                       // Locale. NULL indicates current
         NULL,                    // Security flags.
         0,                       // Authority (for example, Kerberos)
         0,                       // Context object 
         &pSvc                    // pointer to IWbemServices proxy
         );
    
    if (FAILED(hres)){
		if(debugLevel > 1){
			PrintDebugLog(L"Error:Could not connect. Error code :");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%d"), hres);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
        pLoc->Release();     
        CoUninitialize();
        return -1;               // Program has failed.
    }

    // Set security levels on the proxy -------------------------
    hres = CoSetProxyBlanket(
       pSvc,                        // Indicates the proxy to set
       RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
       RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
       NULL,                        // Server principal name 
       RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
       RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
       NULL,                        // client identity
       EOAC_NONE                    // proxy capabilities 
    );

    if (FAILED(hres)){
		if(debugLevel > 1){
			PrintDebugLog(L"Error:Could not set proxy blanket. Error code :");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%d"), hres);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
        pSvc->Release();
        pLoc->Release();     
        CoUninitialize();
        return -1;               // Program has failed.
    }

    // Use the IWbemServices pointer to make requests of WMI ----
	//Win32_Processor��������擾����
	//Win32_ComputerSystem�����NumberOfCores�ANumberOfLogicalProcessors�����邪�A�Â�Windows���ƒl������
	//�܂�Win32_Processor�Ɠ������ʂ����Ȃ����ߎg���Ȃ�
	//https://www.symantec.com/connect/downloads/identifying-physical-hyperthreaded-and-multicore-processors-windows
    IEnumWbemClassObject* pEnumerator = NULL;
	hres = pSvc->ExecQuery(
        bstr_t("WQL"), 
        bstr_t("SELECT * FROM Win32_Processor"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, 
        NULL,
        &pEnumerator);
    
    if (FAILED(hres)){
		if(debugLevel > 1){
			PrintDebugLog(L"Error:Query for operating system name failed. Error code :");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%d"), hres);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return -1;               // Program has failed.
    }

    // Get the data from the query in step 6 -------------------
    IWbemClassObject *pclsObj;
    ULONG uReturn = 0;
	wstringstream sProcessorName;
	wstringstream sProcessorMaxClockSpeed;
	wstring currentProcessorName;
	wstring currentProcessorMaxClockSpeed;
	int iHotfixAfterNumberOfProcessors = 0;
	int iHotfixAfterNumberOfCores = 0;
	int iHotfixAfterNumberOfLogicalProcessors = 0;
	set<BSTR> socketDesignationSet;
	int iHotfixBeforeNumberOfLogicalProcessors = 0;

    while (pEnumerator){
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if(0 == uReturn){
            break;
        }
        VARIANT vtProp;

        //�l�̎擾
		hr = pclsObj->Get(BSTR(L"Name"), 0, &vtProp, 0, 0);
		if(!FAILED(hr)){
			if(vtProp.vt==VT_NULL || vtProp.vt==VT_EMPTY){
				//�������Ȃ�
			}else{
				_bstr_t bstr = vtProp;
				currentProcessorName = bstr;
				if(sProcessorName.str() != currentProcessorName){
					if(sProcessorName.str() != L""){
						sProcessorName << L";";
					}
					sProcessorName << currentProcessorName;
				}else{
					//�����Ȃ�ǉ����Ȃ�
				}
			}
		}

		hr = pclsObj->Get(BSTR(L"MaxClockSpeed"), 0, &vtProp, 0, 0);
		if(!FAILED(hr)){
			if(vtProp.vt==VT_NULL || vtProp.vt==VT_EMPTY){
				//�������Ȃ�
			}else{
				_bstr_t bstr = vtProp;
				currentProcessorMaxClockSpeed = bstr;
				if(sProcessorMaxClockSpeed.str() != currentProcessorMaxClockSpeed){
					if(sProcessorMaxClockSpeed.str() != L""){
						sProcessorMaxClockSpeed << L";";
					}
					sProcessorMaxClockSpeed << currentProcessorMaxClockSpeed;
				}else{
					//�����Ȃ�ǉ����Ȃ�
				}
			}
		}

		//CPU�i�v���Z�b�T�A�΁j�̃J�E���g
		//https://community.spiceworks.com/topic/126129-how-to-determine-physical-processor-count-with-windows-os
		hr = pclsObj->Get(BSTR(L"NumberOfCores"), 0, &vtProp, 0, 0);
		//NumberOfCores�̎擾���������遁hotfix�ς݂�WMI
		if(!FAILED(hr)){
			//�G���g���̐���CPU�̐�
			iHotfixAfterNumberOfProcessors += 1;
			if(vtProp.vt==VT_NULL || vtProp.vt==VT_EMPTY){
				//
			}else{
				//NumberOfCores�̍��v���R�A�̍��v
				iHotfixAfterNumberOfCores += vtProp.intVal;
			}
			hr = pclsObj->Get(BSTR(L"NumberOfLogicalProcessors"), 0, &vtProp, 0, 0);
			if(!FAILED(hr)){
				if(vtProp.vt==VT_NULL || vtProp.vt==VT_EMPTY){
					//
				}else{
					//NumberOfLogicalProcessors�̍��v���_���v���Z�b�T�̍��v
					iHotfixAfterNumberOfLogicalProcessors += vtProp.intVal;
				}
			}			
		//NumberOfCores�̎擾�����s���遁hotfix�O��WMI
		}else{
			hr = pclsObj->Get(BSTR(L"SocketDesignation"), 0, &vtProp, 0, 0);
			if(!FAILED(hr)){
				if(vtProp.vt==VT_NULL || vtProp.vt==VT_EMPTY){
					//
				}else{
					//SocketDesignation�̏d���������������v��CPU��
					socketDesignationSet.insert(vtProp.bstrVal);
				}
			}			
			//�G���g���̐����_���v���Z�b�T�̐�
			iHotfixBeforeNumberOfLogicalProcessors += 1;
		}

		VariantClear(&vtProp);
        pclsObj->Release();
    }

	ProcessorName = sProcessorName.str();
	ProcessorMaxClockSpeed = sProcessorMaxClockSpeed.str();
	if(iHotfixAfterNumberOfProcessors > 0){
		wstringstream sHotfixAfterNumberOfProcessors;
		sHotfixAfterNumberOfProcessors << iHotfixAfterNumberOfProcessors;
		NumberOfProcessors = sHotfixAfterNumberOfProcessors.str();

		wstringstream sHotfixAfterNumberOfCores;
		sHotfixAfterNumberOfCores << iHotfixAfterNumberOfCores;
		NumberOfCores = sHotfixAfterNumberOfCores.str();

		wstringstream sHotfixAfterNumberOfLogicalProcessors;
		sHotfixAfterNumberOfLogicalProcessors << iHotfixAfterNumberOfLogicalProcessors;
		NumberOfLogicalProcessors = sHotfixAfterNumberOfLogicalProcessors.str();
	}else{
		wstringstream sHotfixBeforeNumberOfProcessors;
		sHotfixBeforeNumberOfProcessors << socketDesignationSet.size();
		NumberOfProcessors = sHotfixBeforeNumberOfProcessors.str();

		NumberOfCores = L"";

		wstringstream sHotfixBeforeNumberOfLogicalProcessors;
		sHotfixBeforeNumberOfLogicalProcessors << iHotfixBeforeNumberOfLogicalProcessors;
		NumberOfLogicalProcessors = sHotfixBeforeNumberOfLogicalProcessors.str();
	}
	//MessageBox(NULL,ProcessorName.c_str(),L"test",MB_OK);
	//MessageBox(NULL,ProcessorMaxClockSpeed.c_str(),L"test",MB_OK);
	//MessageBox(NULL,NumberOfProcessors.c_str(),L"test",MB_OK);
	//MessageBox(NULL,NumberOfCores.c_str(),L"test",MB_OK);
	//MessageBox(NULL,NumberOfLogicalProcessors.c_str(),L"test",MB_OK);

	// Cleanup
    pSvc->Release();
    pLoc->Release();
    pEnumerator->Release();
    if(!pclsObj) pclsObj->Release();
    CoUninitialize();

    return 0;
}


//OS���̎擾
wstring getOSName(){
	//���C������
	//OS�����G�f�B�V�����A�o�[�W�����܂߂Ď擾����
	//���i�̎Q�l http://www.gesource.jp/programming/bcb/77.html
	//����̎Q�l http://stackoverflow.com/questions/1268178/how-to-check-in-delphi-the-os-version-windows-7-or-server-2008-r2
	//95�̎Q�l http://chokuto.ifdef.jp/advanced/version.html
	//NT�̎Q�l http://www.delphimaster.net/view/7-33832/all
	//XP�̎Q�l https://handle2.uc3m.es/nestk/trunk/deps/openni/Source/OpenNI/Win32/XnOSWin32.cpp
	//���ƂŋC�������ڂ�����C++�̃T���v�� http://pmakino.jp/tdiary/20090627.html
	//�������ڂ��� http://ht-deko.minim.ne.jp/tech002.html
	wstring OSName = L"";

	// GetVersionEx �ɂ���b����
	//http://pmakino.jp/tdiary/20090627.html
	OSVERSIONINFOEX info;
	ZeroMemory(&info, sizeof(OSVERSIONINFOEX)); // �K�v����̂�?���Ə����Ă�����
	info.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);
	BOOL bOsVersionInfoEx;
	if ((bOsVersionInfoEx = GetVersionEx((OSVERSIONINFO*)&info)) == FALSE) {
		// Windows NT 4.0 SP5 �ȑO�� Windows 9x
		info.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
		if (!GetVersionEx((OSVERSIONINFO *)&info)) return FALSE;
	}

	// Windows XP �ȍ~�ł� GetNativeSystemInfo ���A����ȊO�ł� GetSystemInfo ���g�������̃Z�N�V�����̕K�v���s��
	//http://pmakino.jp/tdiary/20090627.html
	SYSTEM_INFO sysinfo;
	ZeroMemory(&sysinfo, sizeof(SYSTEM_INFO)); // �K�v����̂�?���Ə����Ă�����
	PGNSI pGNSI = (PGNSI)GetProcAddress(GetModuleHandle(TEXT("kernel32.dll")), "GetNativeSystemInfo");
	if (pGNSI) pGNSI(&sysinfo);
	else GetSystemInfo(&sysinfo);
	//------------------
	wchar_t path[200] = L"C:\\Windows\\System32\\kernel32.dll";
	DWORD dwDummy;
	DWORD dwFVISize = GetFileVersionInfoSize(path, &dwDummy);
	LPBYTE lpVersionInfo = new BYTE[dwFVISize];
	if (GetFileVersionInfo(path, 0, dwFVISize, lpVersionInfo) == 0)
	{
		return FALSE;
	}

	UINT uLen;
	VS_FIXEDFILEINFO* lpFfi;
	BOOL bVer = VerQueryValue(lpVersionInfo, L"\\", (LPVOID*)&lpFfi, &uLen);

	if (!bVer || uLen == 0)
	{
		return L"Unknown windows";
	}

	DWORD dwProductVersionMS = lpFfi->dwProductVersionMS;
	info.dwMajorVersion = HIWORD(dwProductVersionMS);
	info.dwMinorVersion = LOWORD(dwProductVersionMS);


	HKeyHolder currentVersion;
	if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, LR"(SOFTWARE\Microsoft\Windows NT\CurrentVersion)", 0, KEY_QUERY_VALUE, &currentVersion) != ERROR_SUCCESS)
		return L"Unknown windows";
	DWORD valueType;
	BYTE buffer[256];
	DWORD bufferSize = 256;

	if (RegQueryValueExW(currentVersion, L"CurrentBuild", nullptr, &valueType, buffer, &bufferSize) != ERROR_SUCCESS)
		return false;

	if (valueType != REG_SZ)
		return false;

	int version;
	std::wistringstream versionStream(reinterpret_cast<wchar_t*>(buffer));
	versionStream >> version;
	info.dwBuildNumber = version;
	//------------------
/*
	OSVERSIONINFO info;
	ZeroMemory(&info, sizeof(OSVERSIONINFO));
	info.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);

	if (!GetVersionEx(&info)) {
		puts("error");
		return 0;
	}
*/

	// Vista �ȍ~�ł� GetProductInfo ���g����
	DWORD type;
	//DWORD dwordCSDVersion;
	if(info.dwMajorVersion >= 6){
		PGPI pGPI = (PGPI)GetProcAddress(GetModuleHandle(TEXT("kernel32.dll")), "GetProductInfo");
		pGPI(info.dwMajorVersion, info.dwMinorVersion, info.wServicePackMajor, info.wServicePackMinor, &type);
		
		//Technical Preview�̔���̂��߂�info.szCSDVersion�𐔒l�ɕϊ�
		//���ԈႢ�@�K�v�Ȃ̂�info.dwBuildNumber�Ȃ̂Ō����琔�l�ɂȂ��Ă���
		//dwordCSDVersion = wcstod(info.szCSDVersion, _T('\0'));
	}

	//Windows8.1��GetVersionEx()��Windows8�ɋU�������l��Ԃ���

	//�΍�1 shell32.dll�͐������l��Ԃ��̂�shell32.dll�ɖ₢���킹��
	//http://www.all.undo.jp/asr/Ver35-4/10.html
	//DLLVERSIONINFO dvi;
	//if(info.dwMajorVersion >= 6){
	//	DLLGETVERSIONINFOPROC pDGV = (DLLGETVERSIONINFOPROC)GetProcAddress(GetModuleHandle(TEXT("shell32.dll")), "DllGetVersion");
	//	dvi.cbSize = sizeof(dvi);
	//	pDGV(&dvi);
	//}
	//���x���ǂݍ��݂̌`�ŏ������Ƃ������G���[��������̂Œ��߂đ΍�2���̗p

	//�΍�2 Windows8.1�ɑΉ����Ă��邱�Ƃ������}�j�t�F�X�g�𖄂ߍ��ށiMS������������@�j
	//http://msdn.microsoft.com/en-us/library/windows/desktop/dn302074(v=vs.85).aspx
	//http://www.inasoft.org/talk/h201310a.html
	//�}�j�t�F�X�g�̖��ߍ��ݕ��@
	//http://www.g-ishihara.com/vc_wi_01.htm
	//�}�j�t�F�X�g��exe���́A�}�j�t�F�X�g�Ǝ��ۂ�exe��������Ă��Ă����Ȃ����ۂ�

	//���ؗp�ɒl��\������
	/*
	_tprintf(TEXT("OSVERSIONINFO.dwMajorVersion         : %d\n"), info.dwMajorVersion);
	_tprintf(TEXT("OSVERSIONINFO.dwMinorVersion         : %d\n"), info.dwMinorVersion);
	_tprintf(TEXT("OSVERSIONINFO.dwBuildNumber          : %d\n"), info.dwBuildNumber);
	_tprintf(TEXT("OSVERSIONINFO.dwPlatformId           : %d\n"), info.dwPlatformId);
	_tprintf(TEXT("OSVERSIONINFO.szCSDVersion           : %s\n"), info.szCSDVersion);
	//_tprintf(TEXT("OSVERSIONINFO.dwPlatformId: %#x\n"), info.dwPlatformId);

	if(info.dwMajorVersion >= 5){
		_tprintf(TEXT("OSVERSIONINFOEX.wServicePackMajor    : %d\n"), info.wServicePackMajor);
		_tprintf(TEXT("OSVERSIONINFOEX.wServicePackMinor    : %d\n"), info.wServicePackMinor);
		_tprintf(TEXT("OSVERSIONINFOEX.wSuiteMask           : %#09x\n"), info.wSuiteMask);
		_tprintf(TEXT("OSVERSIONINFOEX.wProductType         : %#09x\n"), info.wProductType);
	}		

	if(info.dwMajorVersion >= 6){
		_tprintf(TEXT("GetProductInfo.pdwReturnedProductType: %#010x\n"), type);
	}		
	*/

	if(debugLevel > 1){
		wchar_t stringBuffer[32];
		_stprintf(stringBuffer, TEXT("OSVERSIONINFO.dwMajorVersion         : %d\n"), info.dwMajorVersion);
		PrintDebugLog(stringBuffer);
		_stprintf(stringBuffer, TEXT("OSVERSIONINFO.dwMinorVersion         : %d\n"), info.dwMinorVersion);
		PrintDebugLog(stringBuffer);
		_stprintf(stringBuffer, TEXT("OSVERSIONINFO.dwBuildNumber          : %d\n"), info.dwBuildNumber);
		PrintDebugLog(stringBuffer);
		_stprintf(stringBuffer, TEXT("OSVERSIONINFO.dwPlatformId           : %d\n"), info.dwPlatformId);
		PrintDebugLog(stringBuffer);
		_stprintf(stringBuffer, TEXT("OSVERSIONINFO.szCSDVersion           : %s\n"), info.szCSDVersion);
		PrintDebugLog(stringBuffer);
		if(info.dwMajorVersion >= 5){
			_stprintf(stringBuffer, TEXT("OSVERSIONINFOEX.wServicePackMajor    : %d\n"), info.wServicePackMajor);
			PrintDebugLog(stringBuffer);
			_stprintf(stringBuffer, TEXT("OSVERSIONINFOEX.wServicePackMinor    : %d\n"), info.wServicePackMinor);
			PrintDebugLog(stringBuffer);
			_stprintf(stringBuffer, TEXT("OSVERSIONINFOEX.wSuiteMask           : %#09x\n"), info.wSuiteMask);
			PrintDebugLog(stringBuffer);
			_stprintf(stringBuffer, TEXT("OSVERSIONINFOEX.wProductType         : %#09x\n"), info.wProductType);
			PrintDebugLog(stringBuffer);
		}
		if(info.dwMajorVersion >= 6){
			_stprintf(stringBuffer, TEXT("GetProductInfo.pdwReturnedProductType: %#010x\n"), type);
			PrintDebugLog(stringBuffer);
		}
	}

	//��������OS�̃o�[�W�����A�G�f�B�V�������f
	if(info.dwPlatformId == VER_PLATFORM_WIN32s){
		OSName += L"Microsoft Win32s";
	// Windows 9x�n
	}else if(info.dwPlatformId == VER_PLATFORM_WIN32_WINDOWS){
		if(info.dwMajorVersion == 4 && info.dwMinorVersion == 0){
			OSName += L"Windows 95";
			if(info.szCSDVersion[1] = 'B' || info.szCSDVersion[1] == 'C'){
				OSName += L" OEM Service Release 2";
			}
		}else if(info.dwMajorVersion == 4 && info.dwMinorVersion == 10){
			OSName += L"Windows 98";
			if(info.szCSDVersion[1] == 'A'){
				OSName += L" Second Edition";
			}
		}else if(info.dwMajorVersion == 4 && info.dwMinorVersion == 90){
			OSName += L"Windows Me";
		}else{
			OSName += L"Windows 95 family";
		}
	}else if(info.dwPlatformId == VER_PLATFORM_WIN32_NT){
		//���̃Z�N�V�����ŏ��Ɉړ�
/*
		//2000�ȍ~�ŏڂ������ʂ��s�����߂�OSVERSIONINFOEX�����
		//OSVERSIONINFOEX��2000�ȍ~�łȂ��ƌĂяo���G���[�ɂȂ�
		//http://chokuto.ifdef.jp/advanced/version.html
		OSVERSIONINFOEX infoex;
		ZeroMemory(&infoex, sizeof(OSVERSIONINFOEX));
		infoex.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);
		
		//64bitOS�����ʂ���
		SYSTEM_INFO sysinfo = {0};
		GetSystemInfo(&sysinfo);
*/
		if(info.dwMajorVersion == 3){
			OSName += L"Windows NT";

			//NT�̃G�f�B�V�����̓��W�X�g�������Ȃ��Ƃ����Ȃ�
			//http://support.microsoft.com/kb/92395/ja
			//http://support.microsoft.com/kb/152078/ja
/*
			RegList productOptionsRegList;
			productOptionsRegList.setHKeyAndSubKey(L"HKEY_LOCAL_MACHINE", L"SYSTEM\\CurrentControlSet\\Control\\ProductOptions", FALSE);
			//productOptionsRegList.displayAll();
			regValue productTypeValue = productOptionsRegList.getValue(L"ProductType");
			wstring compareProductTypeValue = productTypeValue.data.c_str();
			transform(compareProductTypeValue.begin (), compareProductTypeValue.end (), compareProductTypeValue.begin (), towlower);
			//_tprintf(TEXT("value: %s\n"), compareProductTypeValue.c_str());
*/
			HKEY hKey;
			TCHAR szProductType[9] = TEXT("");
			DWORD dwBufLen = sizeof(TCHAR) * sizeof(szProductType);
			LONG lRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, TEXT("SYSTEM\\CurrentControlSet\\Control\\ProductOptions"), 0, KEY_QUERY_VALUE, &hKey);
			if (lRet == ERROR_SUCCESS) {
				lRet = RegQueryValueEx(hKey, TEXT("ProductType"), NULL, NULL, (LPBYTE)szProductType, &dwBufLen);
				RegCloseKey(hKey);
			}
			if(info.dwMinorVersion == 1){
				if(lstrcmpi(szProductType, L"LANMANNT") == 0){
					OSName += L" Advanced Server 3.1";
				}else if(lstrcmpi(szProductType, L"WINNT") == 0){
					OSName += L" Workstation 3.1";
				}else{
					OSName += L" 3.1 (Unknown Edition)";
				}
			}else if(info.dwMinorVersion == 5){
				if(lstrcmpi(szProductType, L"SERVERNT") == 0){
					OSName += L" Server 3.5";
				}else if(lstrcmpi(szProductType, L"WINNT") == 0){
					OSName += L" Workstation 3.5";
				}else{
					OSName += L" 3.5 (Unknown Edition)";
				}
			}else if(info.dwMinorVersion == 51){
				if(lstrcmpi(szProductType, L"SERVERNT") == 0){
					OSName += L" Server 3.51";
				}else if(lstrcmpi(szProductType, L"WINNT") == 0){
					OSName += L" Workstation 3.51";
				}else{
					OSName += L" 3.51 (Unknown Edition)";
				}
			}else{
				if(lstrcmpi(szProductType, L"SERVERNT") == 0){
					OSName += L" Server 3.x";
				}else if(lstrcmpi(szProductType, L"WINNT") == 0){
					OSName += L" Workstation 3.x";
				}else{
					OSName += L" 3.x (Unknown Edition)";
				}
			}
		//NT4.0
		//http://tebl.homelinux.com/files/project_54/sysctl/osversion.cpp
		//wikipedia�Ƃ�microsoft�T�C�g�Ƃ����������o���o��������
		//�ȉ��̃T�|�[�g���C�t�T�C�N���̃y�[�W�̏������Ƃ���
		//http://support.microsoft.com/lifecycle/?LN=en-us&p1=3194&x=2&y=14
		//http://support.microsoft.com/lifecycle/?LN=en-us&p1=3188&x=12&y=11
		//�ȉ��͔��ʕ��@��������Ȃ������̂Ŗ��Ή�
		//Windows NT Workstation 4.0 Developer Edition
		}else if(info.dwMajorVersion == 4){
			OSName += L"Windows NT";
			if(info.wProductType == VER_NT_WORKSTATION){
				OSName += L" Workstation 4.0";
			}else if(info.wProductType == VER_NT_SERVER || info.wProductType == VER_NT_DOMAIN_CONTROLLER){
				if(info.wSuiteMask & VER_SUITE_ENTERPRISE){
					OSName += L" Server 4.0 Enterprise Edition";
				}else if(info.wSuiteMask & VER_SUITE_TERMINAL){
					OSName += L" Server 4.0, Terminal Server Edition";
				}else{
					OSName += L" Server 4.0 Standard Edition";
				}
			}else{
				OSName += L" 4.0 (Unknown Edition)";
			}
		//2000�̃G�f�B�V����
		//http://support.microsoft.com/lifecycle/search/?sort=PN&alpha=Windows+2000&Filter=FilterNO
		//�����AWindows 2000 Professional Edition��Windows 2000 Professional�Ƃ���
		}else if(info.dwMajorVersion == 5 && info.dwMinorVersion == 0){
			OSName += L"Windows 2000";
			if(info.wProductType == VER_NT_WORKSTATION){
				OSName += L" Professional";
			}else if(info.wProductType == VER_NT_SERVER || info.wProductType == VER_NT_DOMAIN_CONTROLLER){
				if(info.wSuiteMask & VER_SUITE_ENTERPRISE){
					OSName += L" Advanced Server";
				}else if(info.wSuiteMask & VER_SUITE_DATACENTER){
					OSName += L" DataCenter Server";
				}else{
					OSName += L" Server";
				}
			}else{
				OSName += L" (Unknown Edition)";
			}
		//XP�̃G�f�B�V����
		//http://support.microsoft.com/lifecycle/?c2=1173
		//��������������K�G�f�B�V�����̎��� http://support.microsoft.com/kb/922474
		}else if(info.dwMajorVersion == 5 && info.dwMinorVersion == 1){
			//http://msdn.microsoft.com/en-us/library/windows/desktop/ms724385(v=vs.85).aspx
			if(info.wProductType == VER_NT_WORKSTATION){			
				OSName += L"Windows XP";
				if(GetSystemMetrics(SM_TABLETPC) != 0){
					OSName += L" Tablet PC Edition";
				}else if(GetSystemMetrics(SM_STARTER) != 0){
					OSName += L" Starter Edition";
				}else if(GetSystemMetrics(SM_MEDIACENTER) != 0){
					OSName += L" Media Center Edition";
				}else if(info.wSuiteMask & VER_SUITE_PERSONAL){
					OSName += L" Home Edition";
				//http://support.microsoft.com/lifecycle/search/default.aspx?sort=PN&alpha=Embedded&Filter=FilterNO
				//http://msdn.microsoft.com/en-us/library/windows/desktop/ms724833(v=vs.85).aspx
				}else if(info.wSuiteMask & VER_SUITE_EMBEDDEDNT){
					OSName += L" Embedded";
				}else{
					OSName += L" Professional";
				}
			}else{
				OSName += L"Windows 5.1 (Not Workstation)";
			}
		}else if(info.dwMajorVersion == 5 && info.dwMinorVersion == 2){
			if(info.wProductType == VER_NT_WORKSTATION){
				if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
					OSName += L"Windows XP Professional x64 Edition";
				}else{
					OSName += L"Windows 5.2 (Workstation, Not AMD64)";
				}
			//2003�̃G�f�B�V����
			//http://support.microsoft.com/lifecycle/?p1=3198
			//http://support.microsoft.com/lifecycle/?c2=1163
			}else{
				OSName += L"Windows Server 2003";
				if(GetSystemMetrics(SM_SERVERR2) == 0){
					if(info.wSuiteMask & VER_SUITE_DATACENTER){
						if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
							OSName += L", Datacenter x64 Edition";
						}else if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_IA64){
							OSName += L", Datacenter Edition for Itanium-Based Systems";
						}else{
							OSName += L", Datacenter Edition";
						}
					}else if(info.wSuiteMask & VER_SUITE_ENTERPRISE){
						if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
							OSName += L", Enterprise x64 Edition";
						}else if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_IA64){
							OSName += L", Enterprise Edition for Itanium-Based Systems";
						}else{
							OSName += L", Enterprise Edition";
						}
					}else if(info.wSuiteMask & VER_SUITE_BLADE){
						OSName += L", Web Edition";
					//http://msdn.microsoft.com/en-us/library/windows/desktop/ms724833(v=vs.85).aspx
					}else if(info.wSuiteMask & VER_SUITE_STORAGE_SERVER){
						OSName = L"Windows Storage Server 2003";
					//http://support.microsoft.com/lifecycle/search/default.aspx?sort=PN&alpha=Cluster&Filter=FilterNO
					//http://msdn.microsoft.com/en-us/library/windows/desktop/ms724833(v=vs.85).aspx
					}else if(info.wSuiteMask & VER_SUITE_COMPUTE_SERVER){
						OSName += L" Compute Cluster Edition";
					}else{
						if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
							OSName += L", Standard x64 Edition";
						}else{
							OSName += L", Standard Edition";
						}
					}
				}else{
					if(info.wSuiteMask & VER_SUITE_DATACENTER){
						if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
							OSName += L" R2 Datacenter x64 Edition";
						//}else if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_IA64){
						//	OSName += L" R2 Datacenter Edition for Itanium-Based Systems";
						}else{
							OSName += L" R2 Datacenter Edition";
						}
					}else if(info.wSuiteMask & VER_SUITE_ENTERPRISE){
						if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
							OSName += L" R2 Enterprise x64 Edition";
						//}else if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_IA64){
						//	OSName += L" R2 Enterprise Edition for Itanium-Based Systems";
						}else{
							OSName += L" R2 Enterprise Edition";
						}
					//}else if(info.wSuiteMask & VER_SUITE_BLADE){
					//	OSName += L", Web Edition";
					//http://msdn.microsoft.com/en-us/library/windows/desktop/ms724833(v=vs.85).aspx
					}else if(info.wSuiteMask & VER_SUITE_STORAGE_SERVER){
						OSName = L"Windows Storage Server 2003 R2";
					//http://msdn.microsoft.com/en-us/library/windows/desktop/ms724833(v=vs.85).aspx
					}else if(info.wSuiteMask & VER_SUITE_WH_SERVER){
						OSName = L"Windows Home Server";
					}else{
						if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
							OSName += L", Standard x64 Edition";
						}else{
							OSName += L", Standard Edition";
						}
					}
				}
			}
		//Vista�܂���2008
		}else if(info.dwMajorVersion == 6 && info.dwMinorVersion == 0){
			//Vista�̃G�f�B�V����
			//http://support.microsoft.com/lifecycle/?c2=11732
			//�؍��łƂ���K�G�f�B�V�����AKN�G�f�B�V����������炵�������C�t�T�C�N���\�ɂ͂Ȃ��B�܂����ʎ�����������Ȃ��B
			if(info.wProductType == VER_NT_WORKSTATION){
				OSName += L"Windows Vista";
				if(type == PRODUCT_BUSINESS){
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" Business 64-bit Edition";
					}else{
						OSName += L" Business";
					}
				}else if(type == PRODUCT_BUSINESS_N){
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" Business N 64-bit Edition";
					}else{
						OSName += L" Business N";
					}
				}else if(type == PRODUCT_ENTERPRISE){
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" Enterprise 64-bit Edition";
					}else{
						OSName += L" Enterprise";
					}
				}else if(type == PRODUCT_HOME_BASIC){
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" Home Basic 64-bit Edition";
					}else{
						OSName += L" Home Basic";
					}
				}else if(type == PRODUCT_HOME_BASIC_N){
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" Home Basic N 64-bit Edition";
					}else{
						OSName += L" Home Basic N";
					}
				}else if(type == PRODUCT_HOME_PREMIUM){
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" Home Premium 64-bit Edition";
					}else{
						OSName += L" Home Premium";
					}
				//Vista�ɂ�Home Preminum N�͂Ȃ�  Home Basic��Business����
				//http://ja.wikipedia.org/wiki/Microsoft_Windows_Vista
				//}else if(type == PRODUCT_HOME_PREMIUM_N){
				//	if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
				//		OSName += L" Home Premium N 64-bit Edition";
				//	}else{
				//		OSName += L" Home Premium N";
				//	}
				}else if(type == PRODUCT_STARTER){
					//if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
					//	OSName += L" Home Premium 64-bit Edition";
					//}else{
						OSName += L" Starter";
					//}
				}else if(type == PRODUCT_ULTIMATE){
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" Ultimate 64-bit Edition";
					}else{
						OSName += L" Ultimate";
					}
				}else{
					OSName += L" (Unknown Edition)";
				}
			//2008�̃G�f�B�V����
			//http://support.microsoft.com/lifecycle/search/?sort=PN&alpha=Windows+Server+2008&Filter=FilterNO
			//�����Ƒ���Ȃ�
			//http://support.microsoft.com/lifecycle/search/default.aspx?sort=PN&alpha=Server+2008&Filter=FilterNO
			//���C�t�T�C�N���\�ł�64��32�̋�ʂ��Ȃ�����ʂ�t����
			//�������������� Windows Storage Server 2008 Basic�̔��ʕ��@��������Ȃ��ctype�ł�basic���ۂ����̂��Ȃ�
			//�ȉ��̃V���[�Y������
			//Windows Storage Server 2008 Basic
			//Windows Storage Server 2008 Basic 32-bit
			//Windows Storage Server 2008 Basic Embedded
			//Windows Storage Server 2008 Basic Embedded 32-bit
			//�ȉ���������
			//Windows Storage Server 2008 Enterprise Embedded
			//Windows Storage Server 2008 Standard Embedded
			//Windows Storage Server 2008 Workgroup Embedded
			}else{
				OSName += L"Windows Server 2008";
				if(type == PRODUCT_DATACENTER_SERVER || type == PRODUCT_DATACENTER_SERVER_CORE){
					OSName += L" Datacenter";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_DATACENTER_SERVER_V || type == PRODUCT_DATACENTER_SERVER_CORE_V){
					OSName += L" Datacenter without Hyper-V";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_ENTERPRISE_SERVER || type == PRODUCT_ENTERPRISE_SERVER_CORE){
					OSName += L" Enterprise";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					//}else if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_IA64){
					//	OSName += L" for Itanium-Based Systems";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_ENTERPRISE_SERVER_V || type == PRODUCT_ENTERPRISE_SERVER_CORE_V){
					OSName += L" Enterprise without Hyper-V";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_ENTERPRISE_SERVER_IA64){
					//���i����Entreprise���t���Ȃ���type�ɂ�Entreprise���t��
					//�Q�l http://www.microsoft.com/en-us/download/details.aspx?id=12794
					OSName += L" for Itanium-based Systems";
				}else if(type == PRODUCT_SERVER_FOR_SMALLBUSINESS){
					//����64bit�ł����Ȃ�
					//�Q�l http://en.wikipedia.org/wiki/Windows_Server_Essentials
					//Windows Small Business Server 2008 is only available for the x86-64 (64-bit) architecture
					OSName += L" for Windows Essential Server Solutions";
				}else if(type == PRODUCT_SERVER_FOR_SMALLBUSINESS_V){
					//���C�t�T�C�N���\��
					//Windows Server 2008 for Windows Essential Server Solutions without Hyper-V
					//http://msdn.microsoft.com/en-us/library/windows/desktop/ms724358(v=vs.85).aspx �͌ꏇ���t��
					//Windows Server 2008 without Hyper-V for Windows Essential Server Solutions
					//���C�t�T�C�N���\���̗p
					OSName += L" for Windows Essential Server Solutions without Hyper-V";
				}else if(type == PRODUCT_SERVER_FOUNDATION){
					//64bit�ł����Ȃ�
					OSName += L" Foundation";
				}else if(type == PRODUCT_STANDARD_SERVER || type == PRODUCT_STANDARD_SERVER_CORE){
					OSName += L" Standard";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_STANDARD_SERVER_V || type == PRODUCT_STANDARD_SERVER_CORE_V){
					OSName += L" Standard without Hyper-V";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_SMALLBUSINESS_SERVER_PREMIUM || type == PRODUCT_SMALLBUSINESS_SERVER_PREMIUM_CORE){
					//64bit�ł����Ȃ�
					//�Q�l http://en.wikipedia.org/wiki/Windows_Server_Essentials
					//Windows Small Business Server 2008 is only available for the x86-64 (64-bit) architecture
					OSName = L"Windows Small Business Server 2008 Premium";
				}else if(type == PRODUCT_SMALLBUSINESS_SERVER){
					//64bit�ł����Ȃ�
					//�Q�l http://en.wikipedia.org/wiki/Windows_Server_Essentials
					//Windows Small Business Server 2008 is only available for the x86-64 (64-bit) architecture
					OSName = L"Windows Small Business Server 2008 Standard";
				//http://en.wikipedia.org/wiki/Windows_Essential_Business_Server_2008
				//http://msdn.microsoft.com/en-us/library/windows/desktop/ms724358(v=vs.85).aspx
				//http://support.microsoft.com/lifecycle/search/default.aspx?sort=PN&alpha=Essential+Business+Server&Filter=FilterNO
				}else if(type == PRODUCT_MEDIUMBUSINESS_SERVER_MANAGEMENT || type == PRODUCT_MEDIUMBUSINESS_SERVER_MESSAGING || type == PRODUCT_MEDIUMBUSINESS_SERVER_SECURITY){
					//���C�t�T�C�N���\�ɂ͖����Standard�����邪�A�ǂ���ʕt���邩�킯�킩��
					OSName = L"Windows Essential Business Server 2008";
				}else if(type == PRODUCT_STORAGE_ENTERPRISE_SERVER || type == PRODUCT_STORAGE_ENTERPRISE_SERVER_CORE){
					OSName = L"Windows Storage Server 2008 Enterprise";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_STORAGE_STANDARD_SERVER || type == PRODUCT_STORAGE_STANDARD_SERVER_CORE){
					OSName = L"Windows Storage Server 2008 Standard";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_STORAGE_WORKGROUP_SERVER || type == PRODUCT_STORAGE_WORKGROUP_SERVER_CORE){
					OSName = L"Windows Storage Server 2008 Workgroup";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_WEB_SERVER || type == PRODUCT_WEB_SERVER_CORE){
					OSName = L"Windows Web Server 2008";
				}else if(type == PRODUCT_CLUSTER_SERVER){
					OSName = L"Windows HPC Server 2008";

				}else if(type == PRODUCT_CLUSTER_SERVER_V){
					//type�̐������uServer Hyper Core V�v�Ƃ����������X������uwithout Hyper-V�v�Ɣ��f
					OSName = L"Windows HPC Server 2008 without Hyper-V";
				}else{
					OSName += L" (Unknown Edition)";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}

			}
		//7�܂���2008R2(MultiPoint Server 2010/2011�AHome Server 2011�܂�)
		}else if(info.dwMajorVersion == 6 && info.dwMinorVersion == 1){
			//7�̃G�f�B�V����
			//http://support.microsoft.com/lifecycle/search/?sort=PN&alpha=Windows+7&Filter=FilterNO
			//���{���wikipedia�ł�N�G�f�B�V������Home Premium�AProfessional�AEnterprise�AUltimate�ƂȂ��Ă���
			//http://ja.wikipedia.org/wiki/Microsoft_Windows_7
			//MS�̃y�[�W�ł�Starter, Home Premium, Professional, Enterprise, and Ultimate�ƂȂ��Ă���iStarter���v���X�j
			//http://windows.microsoft.com/en-us/windows7/products/what-is-windows-7-n-edition
			//���C�t�T�C�N���\�ł�Home Premium��N�G�f�B�V�������Ȃ��c
			//http://support.microsoft.com/lifecycle/search/?sort=PN&alpha=Windows+7&Filter=FilterNO
			//type�̃y�[�W�ł�PRODUCT_HOME_BASIC_N������c
			//http://msdn.microsoft.com/en-us/library/windows/desktop/ms724358(v=vs.85).aspx
			//�؍��łƂ���KN�G�f�B�V����������炵�������C�t�T�C�N���\�ɂ͂Ȃ��B�������B
			//http://en.wikipedia.org/wiki/Windows_7_editions#Special-purpose_editions
			//���C�t�T�C�N���\�ł�64��32�̋�ʂ��Ȃ�����ʂ�t����
			if(info.wProductType == VER_NT_WORKSTATION){
				OSName += L"Windows 7";
				if(type == PRODUCT_ENTERPRISE){
					OSName += L" Enterprise";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_ENTERPRISE_N){
					OSName += L" Enterprise N";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_HOME_BASIC){
					OSName += L" Home Basic";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_HOME_PREMIUM){
					OSName += L" Home Premium";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_HOME_PREMIUM_N){
					//���C�t�T�C�N���\�ɂ͂Ȃ����c
					OSName += L" Home Premium N";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_PROFESSIONAL){
					OSName += L" Professional";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_PROFESSIONAL_N){
					OSName += L" Professional N";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_STARTER){
					OSName += L" Starter";
				}else if(type == PRODUCT_STARTER_N){
					OSName += L" Starter N";
				}else if(type == PRODUCT_ULTIMATE){
					OSName += L" Ultimate";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_ULTIMATE_N){
					OSName += L" Ultimate N";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else{
					OSName += L" (Unknown Edition)";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}

			//Windows Server 2008 R2�̃G�f�B�V����
			//���C�t�T�C�N���\ http://support.microsoft.com/lifecycle/search/default.aspx?sort=PN&alpha=Server+2008&Filter=FilterNO
			//wikipedia http://ja.wikipedia.org/wiki/Microsoft_Windows_Server_2008_R2
			//MS�y�[�W https://www.microsoft.com/ja-jp/server-cloud/local/windows-server/2008/r2/editions/features.aspx
			//MS�y�[�W��wikipedia�ɂ�Foundation������ ���C�t�T�C�N���\�ɂ͂Ȃ�
			//Windows Server ����h������MultiPoint Server�Ȃ���̂�����(#�M�D�L)�Ʉ���;
			//http://social.msdn.microsoft.com/Forums/vstudio/en-US/25e6c227-3586-47f1-aadb-298e462ddfa0/how-to-differentiate-between-windows-server-2008-r2-and-windows-multipoint-server-2011?forum=vcgeneral
			//2008R2��64bit�ł����Ȃ�

			//MultiPoint Server 2010/2011�̃G�f�B�V����
			//Windows MultiPoint Server 2010
			//Windows MultiPoint Server 2010 Academic���A�J�f�~�b�N�𔻒f������@��������Ȃ� ������
			//Windows MultiPoint Server 2011 Premium
			//Windows MultiPoint Server 2011 Standard
			//���C�t�T�C�N���\ http://support.microsoft.com/lifecycle/search/default.aspx?sort=PN&alpha=multipoint&Filter=FilterNO
			//wikipedia http://en.wikipedia.org/wiki/Windows_MultiPoint_Server
			//MS�y�[�W http://www.microsoft.com/ja-jp/education/multipoint.aspx

			//Home Server 2011
			}else{
				OSName += L"Windows Server 2008 R2";
				if(type == PRODUCT_DATACENTER_SERVER || type == PRODUCT_DATACENTER_SERVER_CORE){
					OSName += L" Datacenter";
				}else if(type == PRODUCT_ENTERPRISE_SERVER || type == PRODUCT_ENTERPRISE_SERVER_CORE){
					OSName += L" Enterprise";
				}else if(type == PRODUCT_ENTERPRISE_SERVER_IA64){
					//���i����Entreprise���t���Ȃ���type�ɂ�Entreprise���t��
					//�Q�l http://www.microsoft.com/en-us/download/details.aspx?id=12794
					OSName += L" for Itanium-based Systems";
				}else if(type == PRODUCT_SERVER_FOUNDATION){
					//MS�y�[�W��wikipedia�ɂ�Foundation������ ���C�t�T�C�N���\�ɂ͂Ȃ�
					OSName += L" Foundation";
				}else if(type == PRODUCT_STANDARD_SERVER || type == PRODUCT_STANDARD_SERVER_CORE){
					OSName += L" Standard";
				}else if(type == PRODUCT_STORAGE_STANDARD_SERVER || type == PRODUCT_STORAGE_STANDARD_SERVER_CORE){
					//2008R2��Storage Server�ɂ�Standard���Ȃ��A���󂪂���
					//type������ō����Ă��邩�͍ڂ��Ă���T�C�g�Ȃ������̂Ŏ��@�Ō��؂��邵���Ȃ�����
					OSName = L"Windows Storage Server 2008 R2";
				}else if(type == PRODUCT_HOME_SERVER){
					OSName = L"Windows Storage Server 2008 R2 Essentials";
				}else if(type == PRODUCT_WEB_SERVER || type == PRODUCT_WEB_SERVER_CORE){
					OSName = L"Windows Web Server 2008 R2";
				}else if(type == PRODUCT_CLUSTER_SERVER){
					OSName = L"Windows HPC Server 2008 R2";
				}else if(type == PRODUCT_SOLUTION_EMBEDDEDSERVER){
					OSName = L"Windows MultiPoint Server 2010";
				}else if(type == PRODUCT_MULTIPOINT_PREMIUM_SERVER){
					OSName = L"Windows MultiPoint Server 2011 Premium";
				}else if(type == PRODUCT_MULTIPOINT_STANDARD_SERVER){
					OSName = L"Windows MultiPoint Server 2011 Standard";
				//http://ja.wikipedia.org/wiki/Microsoft_Windows_Home_Server_2011
				}else if(type == PRODUCT_HOME_PREMIUM_SERVER){
					OSName = L"Windows Home Server 2011";
				}else{
					OSName += L" (Unknown Edition)";
				}

			}
		//8�܂���2012
		}else if(info.dwMajorVersion == 6 && info.dwMinorVersion == 2){
			//8�̃G�f�B�V����
			//�G�f�B�V�����́A����APro�AEnterprise�̎O��� ���ꂼ���N�G�f�B�V���������� Windows RT�͂Ƃ肠�����l�����Ȃ�
			//���C�t�T�C�N���\ http://support.microsoft.com/lifecycle/search/?sort=PN&alpha=Windows+8&Filter=FilterNO
			//wikipwdia http://ja.wikipedia.org/wiki/Microsoft_Windows_8
			//�؍��łƂ���KN�G�f�B�V����������炵�������C�t�T�C�N���\�ɂ͂Ȃ��B�������B
			//http://support.microsoft.com/kb/2703761/ja
			//type�̕\�ɂ́uWindows 8 China�v�����邪�A�������Ă��悭�킩��Ȃ��̂ŃG�f�B�V�����݂��Ȃ�
			//http://msdn.microsoft.com/en-us/library/windows/desktop/ms724358(v=vs.85).aspx
			//���C�t�T�C�N���\�ł�64��32�̋�ʂ��Ȃ�����ʂ�t����
			if(info.wProductType == VER_NT_WORKSTATION){
				OSName += L"Windows 8";
				if(type == PRODUCT_ENTERPRISE){
					OSName += L" Enterprise";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_ENTERPRISE_N){
					OSName += L" Enterprise N";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_CORE){
					//����
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_CORE_N){
					//�����N�G�f�B�V����
					OSName += L" N";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_PROFESSIONAL){
					OSName += L" Pro";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_PROFESSIONAL_N){
					OSName += L" Pro N";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				//Pro��Media Center Pack��K�p�����type��PRODUCT_PROFESSIONAL_WMC�ɕς��炵��(#�M�D�L)�Ʉ���;
				//http://blog.uskanda.com/2012/11/05/windows-8-pro-media-center-product-inf/
				//Media Center Pack�͗L���i�ꎞ���̃L�����y�[�����͖����j�Ȃ̂ŁA�ʃG�f�B�V�����Ƃ��Ĉ���
				//MS�y�[�W�����������G�f�B�V�����ۂ������Ă��� http://windows.microsoft.com/ja-jp/windows-8/feature-packs
				//Pro N��Media Center Pack�����ꂽ��ǂ�����Ĕ��f���邩�͖�����
				}else if(type == PRODUCT_PROFESSIONAL_WMC){
					OSName += L" Pro with Media Center";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else{
					OSName += L" (Unknown Edition)";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}
			//Windows Server 2012�̃G�f�B�V����
			//�G�f�B�V������Datacenter�AStandard�AEssentials�AFoundation
			//���C�t�T�C�N���\ http://support.microsoft.com/lifecycle/search/default.aspx?sort=PN&alpha=Server+2012&Filter=FilterNO
			//wikipedia http://ja.wikipedia.org/wiki/Microsoft_Windows_Server_2012
			//MS�y�[�W http://download.microsoft.com/download/B/F/4/BF474812-BE9E-41CE-9F5F-6C6E2F0B5B22/WS2012_Licensing-Pricing_Datasheet_ja.pdf
			//2012��64bit�ł����Ȃ�
			//MultiPoint Server 2012�̃G�f�B�V����
			//Windows MultiPoint Server 2012 Premium
			//Windows MultiPoint Server 2012 Standard
			//���C�t�T�C�N���\ http://support.microsoft.com/lifecycle/search/default.aspx?sort=PN&alpha=multipoint&Filter=FilterNO
			//wikipedia http://en.wikipedia.org/wiki/Windows_MultiPoint_Server
			//MS�y�[�W http://www.microsoft.com/ja-jp/education/multipoint.aspx
			}else{
				OSName += L"Windows Server 2012";
				if(type == PRODUCT_DATACENTER_SERVER || type == PRODUCT_DATACENTER_SERVER_CORE){
					OSName += L" Datacenter";
				}else if(type == PRODUCT_SERVER_FOUNDATION){
					OSName += L" Foundation";
				}else if(type == PRODUCT_STANDARD_SERVER || type == PRODUCT_STANDARD_SERVER_CORE){
					OSName += L" Standard";
				//Windows Server 2012 Essentials��type�����Əo��̂����x���v������̂��Ȃ��s��
				//���u�����Ă���
				}else if(type == PRODUCT_SERVER_FOR_SMALLBUSINESS){
					OSName += L" Essentials";
				}else if(type == PRODUCT_MULTIPOINT_PREMIUM_SERVER){
					OSName = L"Windows MultiPoint Server 2012 Premium";
				}else if(type == PRODUCT_MULTIPOINT_STANDARD_SERVER){
					OSName = L"Windows MultiPoint Server 2012 Standard";
				}else{
					OSName += L" (Unknown Edition)";
				}

			}
		//8.1�܂���2012R2
		//�}�j�t�F�X�g�����Ȃ���8�ɋU�������
		//http://msdn.microsoft.com/en-us/library/windows/desktop/dn302074(v=vs.85).aspx
		//http://www.inasoft.org/talk/h201310a.html
		//�}�j�t�F�X�g�̖��ߍ��ݕ��@
		//http://www.g-ishihara.com/vc_wi_01.htm
		//�}�j�t�F�X�g��exe���́A�}�j�t�F�X�g�Ǝ��ۂ�exe��������Ă��Ă����Ȃ����ۂ�
		}else if(info.dwMajorVersion == 6 && info.dwMinorVersion == 3){
			//8.1�̃G�f�B�V����
			//�G�f�B�V������8�Ɠ������A����APro�AEnterprise�̎O��� ���ꂼ���N�G�f�B�V����������
			//���C�t�T�C�N���\ http://support.microsoft.com/lifecycle/search/?sort=PN&alpha=Windows+8&Filter=FilterNO
			//wikipwdia http://ja.wikipedia.org/wiki/Microsoft_Windows_8#Windows_8.1
			//�؍��łƂ���KN�G�f�B�V����������炵�������C�t�T�C�N���\�ɂ͂Ȃ��B�������B
			//http://support.microsoft.com/kb/2835517
			//���C�t�T�C�N���\�ł�64��32�̋�ʂ��Ȃ�����ʂ�t����
			if(info.wProductType == VER_NT_WORKSTATION){
				OSName += L"Windows 8.1";
				if(type == PRODUCT_ENTERPRISE){
					OSName += L" Enterprise";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_ENTERPRISE_N){
					OSName += L" Enterprise N";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_CORE){
					//����
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_CORE_N){
					//�����N�G�f�B�V����
					OSName += L" N";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_PROFESSIONAL){
					OSName += L" Pro";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_PROFESSIONAL_N){
					OSName += L" Pro N";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				//Pro��Media Center Pack��K�p�����type��PRODUCT_PROFESSIONAL_WMC�ɕς��炵��(#�M�D�L)�Ʉ���;
				//http://blog.uskanda.com/2012/11/05/windows-8-pro-media-center-product-inf/
				//Media Center Pack�͗L���i�ꎞ���̃L�����y�[�����͖����j�Ȃ̂ŁA�ʃG�f�B�V�����Ƃ��Ĉ���
				//MS�y�[�W�����������G�f�B�V�����ۂ������Ă��� http://windows.microsoft.com/ja-jp/windows-8/feature-packs
				//Pro N��Media Center Pack�����ꂽ��ǂ�����Ĕ��f���邩�͖�����
				}else if(type == PRODUCT_PROFESSIONAL_WMC){
					OSName += L" Pro with Media Center";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else{
					OSName += L" (Unknown Edition)";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}
			//2012R2�̃G�f�B�V����
			//���C�t�T�C�N���\ http://support.microsoft.com/lifecycle/search/default.aspx?sort=PN&alpha=Server+2012&Filter=FilterNO
			//wikipedia http://ja.wikipedia.org/wiki/Microsoft_Windows_Server_2012
			//MS�y�[�W http://download.microsoft.com/download/B/F/4/BF474812-BE9E-41CE-9F5F-6C6E2F0B5B22/WS2012_Licensing-Pricing_Datasheet_ja.pdf
			//�G�f�B�V������2012�Ɠ�����Datacenter�AStandard�AEssentials�AFoundation
			//2012��64bit�ł����Ȃ�
			}else{
				OSName += L"Windows Server 2012 R2";
				if(type == PRODUCT_DATACENTER_SERVER || type == PRODUCT_DATACENTER_SERVER_CORE){
					OSName += L" Datacenter";
				}else if(type == PRODUCT_SERVER_FOUNDATION){
					OSName += L" Foundation";
				}else if(type == PRODUCT_STANDARD_SERVER || type == PRODUCT_STANDARD_SERVER_CORE){
					OSName += L" Standard";
				//Windows Server 2012 R2 Essentials��type�����Əo��̂����x���v������̂��Ȃ��s��
				//���u�����Ă���
				}else if(type == PRODUCT_SERVER_FOR_SMALLBUSINESS){
					OSName += L" Essentials";
				}else{
					OSName += L" (Unknown Edition)";
				}
			}
		//Windows10 Technical Preview�̓r���܂�
		//GUID���Y��ver��Windows�ɓ�������Ă���wscript.exe�����\�[�X�G�f�B�^�ŊJ���Ċm�F
		//{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}
		//http://social.msdn.microsoft.com/Forums/azure/en-US/07cbfc3a-bced-45b7-80d2-a9d32a7c95d4/supportedos-manifest-for-windows-10?forum=windowsgeneraldevelopmentissues
		//�}�j�t�F�X�g�����Ȃ���8.1�ɋU�������
		//http://blogs.msdn.com/b/chuckw/archive/2013/09/10/manifest-madness.aspx
		//�}�j�t�F�X�g�̖��ߍ��ݕ��@
		//http://www.g-ishihara.com/vc_wi_01.htm
		//�}�j�t�F�X�g��exe���́A�}�j�t�F�X�g�Ǝ��ۂ�exe��������Ă��Ă����Ȃ����ۂ�
		}else if(info.dwMajorVersion == 6 && info.dwMinorVersion == 4){
			//Windows10 Technical Preview build 9841����9879�܂�
			//http://it.srad.jp/story/14/11/23/0357207/
			//�G�f�B�V������Technical Preview�̒i�K�ł͖���iPro�����j��Enterprise��2���
			//���C�t�T�C�N���\�ł�64��32�̋�ʂ��Ȃ�����ʂ�t����
			if(info.wProductType == VER_NT_WORKSTATION){
				//dwordCSDVersion
				OSName += L"Windows 10 Technical Preview";
				//Technical Preview
				//Technical Preview��Build 10240�����i�łɂȂ���
				//http://www.tenforums.com/windows-insider/1946-download-windows-10-insider-preview-iso-file.html
				//http://tattu1902.blog136.fc2.com/blog-entry-99.html
				if(type == PRODUCT_ENTERPRISE){
					//Technical Preview��Enterprise�ł�for���t��
					//http://www.microsoft.com/en-us/evalcenter/evaluate-windows-technical-preview-for-enterprise
					OSName += L" for Enterprise";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_PROFESSIONAL){
					//Technical Preview�̖���ł͓����I�ɂ�Pro����
					//http://windows.microsoft.com/ja-jp/windows/preview
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else{
					OSName += L" (Unknown Edition)";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}
			}
		//Windows10�iTechnical Preview�̓r�����琳���ł܂Łj�܂��͎����T�[�oOS
		//GUID���Y��ver��Windows�ɓ�������Ă���wscript.exe�����\�[�X�G�f�B�^�ŊJ���Ċm�F
		//{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}
		//https://msdn.microsoft.com/en-us/library/windows/desktop/dn481241(v=vs.85).aspx
		//�}�j�t�F�X�g�����Ȃ���Ver6.2��Windows8�ɋU�������
		//https://msdn.microsoft.com/en-us/library/windows/desktop/ms724833(v=vs.85).aspx
		//�}�j�t�F�X�g�̖��ߍ��ݕ��@
		//http://www.g-ishihara.com/vc_wi_01.htm
		//�}�j�t�F�X�g��exe���́A�}�j�t�F�X�g�Ǝ��ۂ�exe��������Ă��Ă����Ȃ����ۂ�
		}else if(info.dwMajorVersion == 10 && info.dwMinorVersion == 0){
			//-----------------
			if (info.dwBuildNumber >= 22000) {
				OSName += L"Windows 11";
				///		Windows11�̃G�f�B�V����
					//	�G�f�B�V������7��ށ����[���b�p�̔��g���X�g�@�ɑΉ����邽�߁AMediaPlayer12���O����N�������̂�����
					/*Windows 11 Home
						Windows 11 Education
						Windows 11 Pro
						Windows 11 Pro Education
						Windows 11 Pro for Workstations
						Windows 11 IoT Enterprise
						Windows 11 Enterprise*/
						///	�����łŊm�F���鎖���F
						//	�EN�G�f�B�V������������̂�����B
						//	���C�t�T�C�N���\ https://docs.microsoft.com/ja-jp/lifecycle/faq/windows
				if (info.wProductType == VER_NT_WORKSTATION) {
					//	dwordCSDVersion
					/*OSName += L"Windows 11";*/
					//		Windows11��BuildNumber��22000����n�܂��Ă���B
					//		http://www.tenforums.com/windows-insider/1946-download-windows-10-insider-preview-iso-file.html
					//		http://tattu1902.blog136.fc2.com/blog-entry-99.html
					{
						if (type == PRODUCT_CORE) {
							//	Home�����͖���
							OSName += L" Home";
							if (sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64) {
								OSName += L" (64bit)";
							}
							else {
								OSName += L" (32bit)";
							}
						}
						else if (type == PRODUCT_CORE_N) {
							///		Home�����͖���
							OSName += L" Home N";
							OSName += L" N";
							if (sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64) {
								OSName += L" (64bit)";
							}
							else {
								OSName += L" (32bit)";
							}
						}
						else if (type == PRODUCT_CORE_COUNTRYSPECIFIC) {
							OSName += L" Home China";
							OSName += L" China";
							if (sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64) {
								OSName += L" (64bit)";
							}
							else {
								OSName += L" (32bit)";
							}
						}
						else if (type == PRODUCT_CORE_SINGLELANGUAGE) {
							OSName += L" Home Single Language";
							//OSName += L" Single Language";
							if (sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64) {
								OSName += L" (64bit)";
							}
							else {
								OSName += L" (32bit)";
							}
						}
						else if (type == PRODUCT_MOBILE_CORE) {
							OSName += L" Mobile";
						}
						else if (type == PRODUCT_PROFESSIONAL) {
							OSName += L" Pro";
							if (sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64) {
								OSName += L" (64bit)";
							}
							else {
								OSName += L" (32bit)";
							}
						}
						else if (type == PRODUCT_PROFESSIONAL_N) {
							OSName += L" Pro N";
							if (sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64) {
								OSName += L" (64bit)";
							}
							else {
								OSName += L" (32bit)";
							}
						}
						else if (type == PRODUCT_ENTERPRISE) {
							OSName += L" Enterprise";
							if (sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64) {
								OSName += L" (64bit)";
							}
							else {
								OSName += L" (32bit)";
							}
						}
						else if (type == PRODUCT_ENTERPRISE_EVALUATION) {
							OSName += L" Enterprise (evaluation installation)";
							if (sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64) {
								OSName += L" (64bit)";
							}
							else {
								OSName += L" (32bit)";
							}
						}
						else if (type == PRODUCT_ENTERPRISE_E) {
							///	Windows 10 Enterprise E�Ƃ͉����悭������Ȃ����ꉞ�ǉ�
							OSName += L" Enterprise E";
							if (sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64) {
								OSName += L" (64bit)";
							}
							else {
								OSName += L" (32bit)";
							}
						}
						else if (type == PRODUCT_ENTERPRISE_N) {
							OSName += L" Enterprise N";
							if (sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64) {
								OSName += L" (64bit)";
							}
							else {
								OSName += L" (32bit)";
							}
						}
						else if (type == PRODUCT_ENTERPRISE_N_EVALUATION) {
							OSName += L" Enterprise N (evaluation installation)";
							if (sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64) {
								OSName += L" (64bit)";
							}
							else {
								OSName += L" (32bit)";
							}
						}
						else if (type == PRODUCT_EDUCATION) {
							OSName += L" Education";
							if (sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64) {
								OSName += L" (64bit)";
							}
							else {
								OSName += L" (32bit)";
							}
						}
						else if (type == PRODUCT_EDUCATION_N) {
							OSName += L" Education N";
							if (sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64) {
								OSName += L" (64bit)";
							}
							else {
								OSName += L" (32bit)";
							}
						}
						else if (type == PRODUCT_MOBILE_ENTERPRISE) {
							OSName += L" Mobile Enterprise";
						}
						else {
							OSName += L" (Unknown Edition)";
							if (sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64) {
								OSName += L" (64bit)";
							}
							else {
								OSName += L" (32bit)";
							}
						}
					}
				}
			}
			//-----------------
			
			//Windows10�̃G�f�B�V����
			//�G�f�B�V������6���
			//�����łŊm�F���鎖���F
			//�E���ꂼ���N�G�f�B�V���������邩������͗l
			//�EPro�����ł͂Ȃ����󂪂��邩������ł͂Ȃ�Home
			//�Ewith Bing�����邩�����m�F
			//�E�؍��łƂ���KN�G�f�B�V���������邩 http://support.microsoft.com/kb/2835517
			//���C�t�T�C�N���\ http://support.microsoft.com/lifecycle/search/?sort=PN&alpha=Windows+10&Filter=FilterNO
			//wikipwdia https://en.wikipedia.org/wiki/Windows_10_editions
			//���C�t�T�C�N���\�ł�64��32�̋�ʂ��Ȃ�����ʂ�t����
			if(info.wProductType == VER_NT_WORKSTATION){
				//dwordCSDVersion
				OSName += L"Windows 10";
				//Technical Preview��Build 10240�����i�łɂȂ������A���p�łׂ̍����G�f�B�V�������m�F���Ă��Ȃ��̂�Build Number�Ŕ��f����̂͂�߂�
				//http://www.tenforums.com/windows-insider/1946-download-windows-10-insider-preview-iso-file.html
				//http://tattu1902.blog136.fc2.com/blog-entry-99.html
				//if(info.dwBuildNumber < 10240){
				if(type == PRODUCT_CORE){
					//Home�����͖���
					//OSName += L" Home";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_CORE_N){
					//Home�����͖���
					//OSName += L" Home N";
					OSName += L" N";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_CORE_COUNTRYSPECIFIC){
					//OSName += L" Home China";
					OSName += L" China";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_CORE_SINGLELANGUAGE){
					//OSName += L" Home Single Language";
					OSName += L" Single Language";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_MOBILE_CORE){
					OSName += L" Mobile";
				}else if(type == PRODUCT_PROFESSIONAL){
					OSName += L" Pro";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_PROFESSIONAL_N){
					OSName += L" Pro N";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_ENTERPRISE){
					OSName += L" Enterprise";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_ENTERPRISE_EVALUATION){
					OSName += L" Enterprise (evaluation installation)";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_ENTERPRISE_E){
					//Windows 10 Enterprise E�Ƃ͉����悭������Ȃ����ꉞ�ǉ�
					OSName += L" Enterprise E";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_ENTERPRISE_N){
					OSName += L" Enterprise N";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_ENTERPRISE_N_EVALUATION){
					OSName += L" Enterprise N (evaluation installation)";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_EDUCATION){
					OSName += L" Education";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_EDUCATION_N){
					OSName += L" Education N";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_MOBILE_ENTERPRISE){
					OSName += L" Mobile Enterprise";
				}else{
					OSName += L" (Unknown Edition)";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}
			//Server 2016�ȍ~�̃G�f�B�V����
			//���C�t�T�C�N���\ http://support.microsoft.com/lifecycle/search/default.aspx?sort=PN&alpha=Server+2012&Filter=FilterNO
			//wikipedia http://ja.wikipedia.org/wiki/Microsoft_Windows_Server_2012
			//MS�y�[�W http://download.microsoft.com/download/B/F/4/BF474812-BE9E-41CE-9F5F-6C6E2F0B5B22/WS2012_Licensing-Pricing_Datasheet_ja.pdf
			//�G�f�B�V������2012�Ɠ�����Datacenter�AStandard�AEssentials�AFoundation
			}else{
				//Server 2016��2019�̔���̓��W�X�g�����K�v
				HKEY hKey;
				TCHAR szReleaseId[9] = TEXT("");
				DWORD dwBufLen = sizeof(TCHAR) * sizeof(szReleaseId);
				LONG lRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion"), 0, KEY_QUERY_VALUE, &hKey);
				if (lRet == ERROR_SUCCESS) {
					lRet = RegQueryValueEx(hKey, TEXT("ReleaseId"), NULL, NULL, (LPBYTE)szReleaseId, &dwBufLen);
					RegCloseKey(hKey);
				}
				int ReleaseId = _wtoi(szReleaseId);
				
				//ReleaseId�Ŕ���
				//https://techcommunity.microsoft.com/t5/Windows-Server-Insiders/Windows-Server-2019-version-info/td-p/234472
				if(ReleaseId < 1803){
					//Server 2016 Standard -> 1607
					OSName += L"Windows Server 2016";
				}else{
					//Server 2019 Standard -> 1803
					//Server 2019 Datacenter -> 1809
					OSName += L"Windows Server 2019";
				}
				if(type == PRODUCT_DATACENTER_SERVER || type == PRODUCT_DATACENTER_SERVER_CORE){
					OSName += L" Datacenter";
				}else if(type == PRODUCT_SERVER_FOUNDATION){
					OSName += L" Foundation";
				}else if(type == PRODUCT_STANDARD_SERVER || type == PRODUCT_STANDARD_SERVER_CORE){
					OSName += L" Standard";
				//Windows Server 2012 R2 Essentials��type�����Əo��̂����x���v������̂��Ȃ��s��
				//���u�����Ă���
				}else if(type == PRODUCT_SERVER_FOR_SMALLBUSINESS){
					OSName += L" Essentials";
				}else{
					OSName += L" (Unknown Edition)";
				}
			}
		}else{
			wchar_t OSNameBuffer[64];
			_stprintf(OSNameBuffer, TEXT("Unknown Windows (Version %d.%d)"), info.dwMajorVersion, info.dwMinorVersion);
			OSName = OSNameBuffer;
		}
	}

	//_tprintf(TEXT("OS: %s\n"), OSName.c_str());
	if(debugLevel > 1){
		PrintDebugLog(L"getOSName(). OSName:");
		PrintDebugLog(OSName);
		PrintDebugLog(L"\n");
	}

	return OSName;
}

//LocalizedString�̎擾
wstring getLocalizedString(wstring inputString){
	//HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\DW WLAN Card Utility
	//@C:\Windows\system32\bcmwlrc.dll,-4001
	//HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{3E29EE6C-963A-4aae-86C1-DC237C4A49FC}
	//@C:\Program Files (x86)\Intel\Intel(R) Rapid Storage Technology\Uninstall\Setup.exe,-2018

	//LocalizedString�͈ȉ��̌`���œn�����
	//"@<PE-path>,-<stringID>[;<comment>]"
	//http://msdn.microsoft.com/en-us/library/windows/desktop/dd374120(v=vs.85).aspx
	//������","�͂Ȃ����Ƃ�����
	//�܂��͋�؂蕶���̈ʒu�𒲂ׂ�
	int index1 = (int)inputString.find(L"@", 0);
	if(debugLevel > 1 && index1 != 0){
		//fwprintf_s(stderr, TEXT("getLocalizedString error. inputString:%s"),inputString.c_str());
		PrintDebugLog(L"Error:getLocalizedString \"@\" not found. inputString:");
		PrintDebugLog(inputString);
		PrintDebugLog(L", pos:");
		wchar_t stringBuffer[32];
		_stprintf(stringBuffer, TEXT("%ld"), index1);
		PrintDebugLog(stringBuffer);
		PrintDebugLog(L"\n");
		return L"";
	}
	//int index2 = (int)inputString.find(L",-", 0); //���X
	int index2 = (int)inputString.find(L",-", 0); //debuglog����R�s�y
	//if(debugLevel > 1 && index2 != string::npos){
	if(debugLevel > 1 && index2 == string::npos){
		//fwprintf_s(stderr, TEXT("getLocalizedString error. inputString:%s"),inputString.c_str());
		PrintDebugLog(L"Error:getLocalizedString \",-\" not found. inputString:");
		PrintDebugLog(inputString);
		PrintDebugLog(L", pos:");
		wchar_t stringBuffer[32];
		_stprintf(stringBuffer, TEXT("%ld"), index2);
		PrintDebugLog(stringBuffer);
		PrintDebugLog(L"\n");
		//MessageBox(NULL,index1,_T("MAC"),MB_OK);
		return L"";
	}
	int index3 = (int)inputString.find(L";", 0);
	if(debugLevel > 1 && index3 != string::npos && index2 > index3){
	//if(debugLevel > 1 && (index3 == string::npos || index2 > index3)){
	//if(debugLevel > 1 && index2 > index3){
		//fwprintf_s(stderr, TEXT("getLocalizedString error. inputString:%s"),inputString.c_str());
		PrintDebugLog(L"Error:getLocalizedString \";\" pos bad. inputString:");
		PrintDebugLog(inputString);
		PrintDebugLog(L", pos:");
		wchar_t stringBuffer[32];
		_stprintf(stringBuffer, TEXT("%ld"), index3);
		PrintDebugLog(stringBuffer);
		PrintDebugLog(L"\n");
		return L"";
	}

	wstring path;
	wstring stringID;

	path = inputString.substr(1, index2 - index1 - 1);
	//��؂蕶���u;�v���Ȃ��ꍇ�i��stringID�̂��Ƃ�comment���Ȃ��ꍇ�j
	if(index3 == string::npos){
		stringID = inputString.substr(index2 + 2, inputString.size() - index2);
	//��؂蕶���u;�v������ꍇ�i��stringID�̂��Ƃ�comment������ꍇ�j
	}else{
		stringID = inputString.substr(index2 + 2, index3 - index2 - 1);
	}

	//_tprintf(TEXT("path: %s\n"), path.c_str());
	//_tprintf(TEXT("stringID: %s\n"), stringID.c_str());

	//�������當����̃��[�h
	//LoadLibraryEx�����邪�A�Ƃ肠����LoadLibrary��
	//LoadLibraryEx�Ȃ�A�p�X�w��Ńt�@�C����������ς�����
	//http://msdn.microsoft.com/ja-jp/library/cc429243.aspx
	//�ǂ��炪���K�؂��͖�����

	//64bitOS��32bit�A�v���P�[�V�����ŃA�N�Z�X�����Ƃ�system32���J���Ȃ���肠��
	//���̏ꍇ�ASysnative�ɃA�N�Z�X���邱�Ƃŉ���ł���
	//http://yamori-jp.blogspot.jp/2011/04/64bitwindows32bitsystem32.html
	//http://msdn.microsoft.com/ja-jp/library/aa384187(v=vs.85).aspx
	//�ڍׂȋ����͖�����

	//64bitOS��
	if(IsWow64()){
		//��r�p�ɏ������ɂ��镶�����p��
		wstring pathLower(path.c_str());
		transform(pathLower.begin (), pathLower.end (), pathLower.begin (), tolower);
		//path��system32���܂܂�Ă��邩�`�F�b�N
		int indexSystem32 = (int)pathLower.find(L"\\system32\\", 0);
		//�����Sysnative�ɒu��
		if(indexSystem32 != string::npos){
			path.replace(indexSystem32, 10, L"\\Sysnative\\");
			//_tprintf(TEXT("path: %s\n"), path.c_str());
		}
    }
	HMODULE hModule = LoadLibrary(path.c_str());
	if(hModule == NULL){
		if(debugLevel > 1){
			//fwprintf_s(stderr, TEXT("LoadLibrary error. path:%s"),path.c_str());
			PrintDebugLog(L"Error:LoadLibrary error. path:");
			PrintDebugLog(path);
			PrintDebugLog(L"\n");
		}
		return L"";
	}

	//LoadString
	//http://msdn.microsoft.com/ja-jp/library/cc410872.aspx
	//�o�b�t�@�𒴂��钷���̕�����͐؂�̂Ă���̂Ńo�b�t�@�͑傫�߂Ɏ���Ă���
	//_TCHAR pszValue[1024] = { 0 };
	_TCHAR pszValue[4096];
	int nResult = LoadString(hModule, _wtoi(stringID.c_str()), pszValue, _countof(pszValue));
	wstring pszVstring = pszValue;
	//_tprintf(TEXT("getLocalizedString: %s\n"), pszVstring.c_str());
	return pszVstring;

}

struct subKey{
	//_TCHAR* name;
	//_TCHAR* fullPathName;
	//_TCHAR* lastWriteTime;
	wstring name;
	wstring fullPathName;
	wstring lastWriteTime;
};
struct regValue{
	//_TCHAR* name;
	//_TCHAR* data;
	//_TCHAR* type;
	wstring name;
	wstring lowerName;
	wstring data;
	wstring type;
	//�\�[�g�ł���悤�ɔ�r���Z�q���`
	//http://d.hatena.ne.jp/minus9d/20130501/1367415668
	//�������Ă�����
	//sort(data_array.begin(), data_array.end());�̂悤�ɂł���
	//�Ō��const��Y����"instantiated from here"�Ƃ����G���[���o�ăR���p�C���ł��Ȃ��̂Œ���
	//bool operator<( const data_t& right ) const {
		//return num == right.num ? str < right.str : num < right.num;
	//	return name < right.name;
	//}
};
struct hardwareInfo{
	wstring HardwareNoTitle;
	wstring HardwareNoValue;
	wstring HardwareNoValueUpdateFlag;
	wstring Title1;
	wstring Value1;
	wstring Value1UpdateFlag;
	wstring Title2;
	wstring Value2;
	wstring Value2UpdateFlag;
	wstring Title3;
	wstring Value3;
	wstring Value3UpdateFlag;
	wstring Title4;
	wstring Value4;
	wstring Value4UpdateFlag;
	wstring Title5;
	wstring Value5;
	wstring Value5UpdateFlag;
	wstring Title6;
	wstring Value6;
	wstring Value6UpdateFlag;
	wstring ComputerName;
	wstring ComputerNameUpdateFlag;
	wstring UserName;
	wstring UserNameUpdateFlag;
	wstring IPAddress;
	wstring IPAddressUpdateFlag;
	wstring MACAddress;
	wstring MACAddressUpdateFlag;
	wstring OSName;
	wstring OSNameUpdateFlag;
	wstring SystemProductName;
	wstring SystemProductNameUpdateFlag;
	wstring SystemManufacturer;
	wstring SystemManufacturerUpdateFlag;
	wstring IdentifyingNumber;
	wstring IdentifyingNumberUpdateFlag;
	wstring ProcessorName;
	wstring ProcessorNameUpdateFlag;
	wstring ProcessorMaxClockSpeed;
	wstring ProcessorMaxClockSpeedUpdateFlag;
	wstring NumberOfSockets;
	wstring NumberOfSocketsUpdateFlag;
	wstring NumberOfProcessors;
	wstring NumberOfProcessorsUpdateFlag;
	wstring NumberOfCores;
	wstring NumberOfCoresUpdateFlag;
	wstring NumberOfLogicalProcessors;
	wstring NumberOfLogicalProcessorsUpdateFlag;
	wstring Virtualization;
	wstring VirtualizationUpdateFlag;
	wstring TotalPhysicalMemory;
	wstring TotalPhysicalMemoryUpdateFlag;
	wstring Timestamp;
	wstring TimestampUpdateFlag;
	wstring Timestamp2;
	wstring ToolVersion;
	wstring ToolVersionUpdateFlag;
};

struct softwareInfo{
	wstring Name;
	wstring Vendor;
	wstring Version;
	wstring InstallDate;
	wstring ScanRegistryRange;
	wstring CompareName;
	wstring NameCodePoint;
	wstring RegistoryKey;
	wstring UpdateFlag;
	//�\�[�g�ł���悤�ɔ�r���Z�q���`
	//http://d.hatena.ne.jp/minus9d/20130501/1367415668
	//�������Ă�����
	//sort(data_array.begin(), data_array.end());�̂悤�ɂł���
	//�Ō��const��Y����"instantiated from here"�Ƃ����G���[���o�ăR���p�C���ł��Ȃ��̂Œ���
	//bool operator<( const data_t& right ) const {
		//return num == right.num ? str < right.str : num < right.num;
	//	return name < right.name;
	//}
	//http://homepage2.nifty.com/well/sort.html �����������̗p
	//softwareInfo::nameCompare
	//���ёւ��̃��[���i�ڍׂ͖����؁j
	//�p���͑啶������������
	//�Ђ炪�ȃJ�^�J�i�̈Ⴂ�͖���
	//������UTF-8�ł͂Ȃ�JIS��
	static bool nameCompare(const softwareInfo& left, const softwareInfo& right){
		//��r�p�������p��
		wstring compareLeft = left.CompareName.c_str();
		wstring compareRight = right.CompareName.c_str();

		return compareLeft < compareRight;
	}
};

struct patchInfo{
	wstring Name;
	wstring ParentName;
	wstring KB;
	wstring Vendor;
	wstring Version;
	wstring InstallDate;
	wstring ScanRegistryRange;
	wstring CompareName;
	wstring NameCodePoint;
	wstring RegistoryKey;
	wstring UpdateFlag;
	//�\�[�g�ł���悤�ɔ�r���Z�q���`
	//http://d.hatena.ne.jp/minus9d/20130501/1367415668
	static bool nameCompare(const patchInfo& left, const patchInfo& right){
		//��r�p�������p��
		wstring compareLeft = left.CompareName.c_str();
		wstring compareRight = right.CompareName.c_str();
		return compareLeft < compareRight;
	}
};


//�w�肳�ꂽ�L�[�����̒l�ƃT�u�L�[���擾����N���X
class RegList{
	private:
		//�L�[�@�@�F�l�̖��O
		//�l�@�@�@�F�l�̒��g
		//map<_TCHAR*, _TCHAR*> regValues;
		//�L�[�@�@�F�l�̖��O
		//�l�@�@�@�F�l�̃f�[�^�^
		//map<_TCHAR*, _TCHAR*> regTypes;
		//map<_TCHAR*, _TCHAR* > data;
		//�T�u�L�[�̔z��
		//_TCHAR** subKeys;���錾���ɂ̓G���[�ɂȂ�Ȃ����g���Ƃ�exe���G���[��������
		//vector<_TCHAR*> subKeys;
		unsigned int subKeysIndex;
		vector<regValue> regValues;
		vector<subKey> subKeys;

    //����
    //data[TEXT("Environment")][TEXT("ProductName")] = TEXT("Name");
	public:
		//�����o�֐��̐錾
		//�K�v���i��L�[�ƃT�u�L�[�j���󂯎���ă��W�X�g�����X�L��������֐�
		long setHKeyAndSubKey(wstring hKey, wstring subKey, BOOL scan64bit);
		//�擾�����f�[�^��W���o�͂ɏo���֐�
		void displayAll();
		//�w�肳�ꂽ���O�̒l�̒��g�ƌ^��Ԃ��֐�
		//���O�͑啶������������
		regValue getValue(wstring name);
		//���������T�u�L�[��񋓂���֐��i�Ă΂��x��1���Q�Ɠn�����čŌ�܂ŏo����߂�l��-1�j
		subKey getNextSubKey();
};

//�K�v���i��L�[�ƃT�u�L�[�j���󂯎���ă��W�X�g�����X�L��������֐�
long RegList::setHKeyAndSubKey(wstring hKeyChar, wstring subKeyChar, BOOL scan64bit){
	subKeysIndex = 0;
	//clear���Ȃ��ƍēxsetHKeyAndSubKey���Ăяo�����Ƃ��ɈȑO�̃f�[�^�������Ă��������Ȃ�
	regValues.clear();
	subKeys.clear();

	HKEY hKey;
	
	//�啶���ɂ���̂̓G���[�����������Ƃ�����̂łƂ肠������~�@�����ǋ��͂܂�
	//StrToUpper(hKeyChar);
	if(hKeyChar == TEXT("HKEY_CLASSES_ROOT")){
		hKey = HKEY_CLASSES_ROOT;
	}else if(hKeyChar == TEXT("HKEY_CURRENT_CONFIG")){
		hKey = HKEY_CURRENT_CONFIG;
	}else if(hKeyChar == TEXT("HKEY_CURRENT_USER")){
		hKey = HKEY_CURRENT_USER;
	}else if(hKeyChar == TEXT("HKEY_LOCAL_MACHINE")){
		hKey = HKEY_LOCAL_MACHINE;
	}else if(hKeyChar == TEXT("HKEY_USERS")){
		hKey = HKEY_USERS;
	}else{
		if(debugLevel > 1){
			//fwprintf_s(stderr, TEXT("HKEY unmatch. hKeyChar:%s\n"),hKeyChar.c_str());
			PrintDebugLog(L"Error:HKEY unmatch. hKeyChar:");
			PrintDebugLog(hKeyChar);
			PrintDebugLog(L"\n");
		}
		return -1;
	}	
	
	//strcpy(name,ss);
	//regValues.insert( map<_TCHAR*, map<_TCHAR*, _TCHAR*> >::value_type( 10, "aaa" ) );= TEXT("Name");
	//_tprintf(TEXT("setHKeyAndSubKey\n"));
	//_tprintf(TEXT(map[hKey][subKey]));
	//map<_TCHAR*, map<_TCHAR*, _TCHAR*> > regValues;
	//_tprintf(regValues[hKey][subKey]);
	//regValues[hKeyChar][subKey] = TEXT("ccc");

	HKEY phkResult;
    TCHAR szValueName[MAX_PATH];
    DWORD dwValueNameSize;
    BYTE lpData[1024*4];
    DWORD dwDataSize;
    DWORD dwType;
    LONG Status;
	//���̂悤�ɂ���ƁA�u�Ԑڑ���̃��x�����قȂ�܂��v�Ƃ����G���[���o��
	//ULONG_PTR hKey = HKEY_LOCAL_MACHINE;
	//������LONG�ł��Ȃ�
	//�悭�������Ă��Ȃ��̂ő΍��s��
	//LONG hKey = HKEY_LOCAL_MACHINE;

    //Status = RegOpenKeyEx(hKey, subKeyChar, 0, KEY_READ  | (IsWow64() ? KEY_WOW64_64KEY : 0), &phkResult);
	if(IsWow64() == TRUE){
		//64bitOS��64bit�L�[���X�L��������ꍇ
		if(scan64bit == TRUE){
			Status = RegOpenKeyEx(hKey, subKeyChar.c_str(), 0, KEY_READ | KEY_WOW64_64KEY, &phkResult);
			if(debugLevel > 1){
				PrintDebugLog(L"RegOpenKeyEx, KEY_READ, KEY_WOW64_64KEY. hKey:");
				PrintDebugLog(hKeyChar);
				PrintDebugLog(L", subKey:");
				PrintDebugLog(subKeyChar);
				PrintDebugLog(L", Status:");
				wchar_t stringBuffer[32];
				_stprintf(stringBuffer, TEXT("%ld"), Status);
				PrintDebugLog(stringBuffer);
				PrintDebugLog(L"\n");
			}
		//64bitOS��32bit�L�[���X�L��������ꍇ
		}else{
			Status = RegOpenKeyEx(hKey, subKeyChar.c_str(), 0, KEY_READ | KEY_WOW64_32KEY, &phkResult);
			if(debugLevel > 1){
				PrintDebugLog(L"RegOpenKeyEx, KEY_READ, KEY_WOW64_32KEY. hKey:");
				PrintDebugLog(hKeyChar);
				PrintDebugLog(L", subKey:");
				PrintDebugLog(subKeyChar);
				PrintDebugLog(L", Status:");
				wchar_t stringBuffer[32];
				_stprintf(stringBuffer, TEXT("%ld"), Status);
				PrintDebugLog(stringBuffer);
				PrintDebugLog(L"\n");
			}
		}
	}else{
		//32bitOS��64bit�L�[���X�L��������ꍇ
		if(scan64bit == TRUE){
			if(debugLevel > 1){
				//fwprintf_s(stderr, TEXT("Do Not Scan 64bitKey in 32bitOS.\n"));
				PrintDebugLog(L"Error:Do Not Scan 64bitKey in 32bitOS.");
				PrintDebugLog(L"\n");
			}
		}else{
			Status = RegOpenKeyEx(hKey, subKeyChar.c_str(), 0, KEY_READ, &phkResult);
			if(debugLevel > 1){
				PrintDebugLog(L"RegOpenKeyEx, KEY_READ. hKey:");
				PrintDebugLog(hKeyChar);
				PrintDebugLog(L", subKey:");
				PrintDebugLog(subKeyChar);
				PrintDebugLog(L", Status:");
				wchar_t stringBuffer[32];
				_stprintf(stringBuffer, TEXT("%ld"), Status);
				PrintDebugLog(stringBuffer);
				PrintDebugLog(L"\n");
			}
		}
	}

    if (Status != ERROR_SUCCESS) {
		if(debugLevel > 1){
			//fwprintf_s(stderr, TEXT("RegOpenKeyEx error. hKeyChar;%s, subKeyChar:%s\n"),hKeyChar.c_str(), subKeyChar.c_str());
			PrintDebugLog(L"Error:RegOpenKeyEx error. hKeyChar:");
			PrintDebugLog(hKeyChar);
			PrintDebugLog(L", subKeyChar:");
			PrintDebugLog(subKeyChar);
			PrintDebugLog(L"\n");
		}
		RegCloseKey(phkResult);
		return -2;
    }

    for (DWORD dwIndex = 0; ; dwIndex++) {
        
        dwValueNameSize = sizeof(szValueName)/sizeof(szValueName[0]);
        dwDataSize = sizeof(lpData)/sizeof(lpData[0]);
		
		//wchar_t buffer[200];
		//wchar_t szValueNameBuffer[MAX_PATH];
		//wchar_t *buffer;
		//wchar_t *szValueNameBuffer;
		//new�œ��I�Ɋm�ۂ��Ȃ��Ə㏑�����Ă��܂�
		//http://www.geocities.jp/ky_webid/cpp/language/012.html
		//buffer = new wchar_t[2048];
		//szValueNameBuffer = new wchar_t[MAX_PATH];
		wchar_t regDataCharBuffer[2048];
		//DWORD regDataDwordBuffer;

        Status = RegEnumValue(
			phkResult, 
            dwIndex, 
            szValueName, 
            &dwValueNameSize, 
            NULL, 
            &dwType, 
            lpData, 
            &dwDataSize);

        if (Status == ERROR_NO_MORE_ITEMS) {
            //puts("\nERROR_NO_MORE_ITEMS");
            break;
        } else if (Status != ERROR_SUCCESS) {
			if(debugLevel > 1){
				//fwprintf_s(stderr, TEXT("RegEnumValue error. hKeyChar;%s, subKeyChar:%s\n"),hKeyChar.c_str(), subKeyChar.c_str());
				PrintDebugLog(L"Error:RegEnumValue error. hKeyChar:");
				PrintDebugLog(hKeyChar);
				PrintDebugLog(L", subKeyChar:");
				PrintDebugLog(subKeyChar);
				PrintDebugLog(L"\n");
			}
			RegCloseKey(phkResult);
            return -3;
        } else {//Status == ERROR_SUCCESS

            //_tprintf(TEXT("-----------------------------------------------------------\n"));
            //_tprintf(TEXT("ValueName: %s\n"), szValueName);
			
			//_tcscpy_s(szValueNameBuffer, MAX_PATH, szValueName);

			//������擾����l�ɂ���regValue�\���̂�p�ӂ���
			regValue currentRegValue;
			//currentRegValue.name = szValueNameBuffer;
			//tchar��wstring�ɑ��
			currentRegValue.name = szValueName;
			//���W�X�g���̃L�[���͑啶�������������Ȃ̂ŏ������ɂ������O���p��
			//tolower�Ŕ�r����ƁA���{��Ȃǂŕ�����������@towlower���g��
			currentRegValue.lowerName = szValueName;
			transform(currentRegValue.lowerName.begin (), currentRegValue.lowerName.end (), currentRegValue.lowerName.begin (), towlower);

            switch(dwType){

            case REG_BINARY:
                //�o�C�i���l�́AHex�_���v����
                //puts("REG_BINARY");
				//regTypes[szValueNameBuffer] = TEXT("REG_BINARY");
				//regValues[szValueNameBuffer] = TEXT("");
				currentRegValue.type = L"REG_BINARY";
				//currentRegValue.data�ɋ󕶎�����Z�b�g����ƕ������ǉ������^�C�~���O�ŃA�N�Z�X�ᔽ�̗�O�ƂȂ�
				//currentRegValue.data = TEXT("");
				//currentRegValue.data = new wchar_t[2048];

				//_TCHAR* tempString = TEXT("");
                for (DWORD i = 0; i < dwDataSize; i++) {
                    //printf("%02X ", lpData[i]);
					//_stprintf(buffer, TEXT("%02X "), lpData[i]);
					_stprintf(regDataCharBuffer, TEXT("%02X "), lpData[i]);
					//_tcscat(regValues[szValueNameBuffer], buffer);
					//_tcscat(currentRegValue.data, buffer);
					currentRegValue.data += regDataCharBuffer;
					//wcscat_s(currentRegValue.data, 1024, buffer);
					//strcat(tempString, _stprintf(TEXT("%02X "), lpData[i]));
                }
				//regValues[TEXT("")][szValueName] = tempString;
                //putchar('\n');
                break;

            case REG_DWORD:
                //puts("REG_DWORD");
				//regTypes[szValueNameBuffer] = TEXT("REG_DWORD");
				currentRegValue.type = L"REG_DWORD";
                DWORD dwValue;
                memcpy(&dwValue, lpData, sizeof(dwValue));
				_stprintf(regDataCharBuffer, TEXT("%d"), dwValue);
				//regValues[szValueNameBuffer] = buffer;
				currentRegValue.data = regDataCharBuffer;
                //printf("%d\n", dwValue);
                break;

            case REG_EXPAND_SZ:
                //ExpandEnvironmentStrings�Ŋ��ϐ��������W�J���Ă���\������
                //puts("REG_EXPAND_SZ");
				//regTypes[szValueNameBuffer] = TEXT("REG_EXPAND_SZ");
				currentRegValue.type = TEXT("REG_EXPAND_SZ");
                TCHAR szDst[MAX_PATH];
                ExpandEnvironmentStrings( (LPTSTR)lpData, szDst, sizeof(szDst)/sizeof(szDst[0]) );
                //_tprintf(TEXT("%s\n"), szDst);
				_stprintf(regDataCharBuffer, TEXT("%s"), szDst);
				//regValues[szValueNameBuffer] = buffer;
				currentRegValue.data = regDataCharBuffer;
                break;

            case REG_QWORD:
                //puts("REG_QWORD");
				//regTypes[szValueNameBuffer] = TEXT("REG_QWORD");
 				currentRegValue.type = TEXT("REG_QWORD");
               __int64 n;
                memcpy(&n, lpData, sizeof(n));
				//_tprintf(TEXT("%lld\n"), n);
				_stprintf(regDataCharBuffer, TEXT("%lld"), n);
				//regValues[szValueNameBuffer] = buffer;
				currentRegValue.data = regDataCharBuffer;
                break;

            case REG_SZ:
				//if(currentRegValue.name)
                //puts("REG_SZ");
				//regTypes[szValueNameBuffer] = TEXT("REG_SZ");
 				currentRegValue.type = TEXT("REG_SZ");
				//_tprintf(TEXT("%s\n"), lpData);
				_stprintf(regDataCharBuffer, TEXT("%s"), lpData);
				//regValues[szValueNameBuffer] =buffer; 
				currentRegValue.data = regDataCharBuffer;
                break;

            case REG_MULTI_SZ:
 				currentRegValue.type = TEXT("REG_MULTI_SZ");
				//REG_MULTI_SZ�̓k�������i\0�j�ŋ�؂��A�I�[�̓k���������A���i\0\0�j���Ă���B�p��ł�double-null-terminated list
				//REG_SZ�̕��@�����̂܂ܓK�p����ƁA�ŏ��̃k�������܂ł����擾�ł��Ȃ�
				//_stprintf(regDataCharBuffer, TEXT("%s"), lpData);
				//currentRegValue.data = regDataCharBuffer;

				//�T�C�Y���w�肵�Ă����܂��������i�ڍז����؁j
				//swprintf(regDataCharBuffer, dwDataSize, TEXT("%s"), lpData);
				
				//wstring�̏������ɂ��̂܂ܓn���Ă����܂��������i�ڍז����؁j
				//wstring regDataStrBuffer(char(lpData), 0, dwDataSize);
				//currentRegValue.data = regDataStrBuffer;

				//�k�����������s�ɒu�����Ă��珈��
				//����ɂ��
				//36 00 38 00 41 00 42 00 36 00 37 00 43 00 41 00 37 00 44 00 41 00 37 00 46 00 46 00 46 00 46 00 35 00 32 00 30 00 35 00 41 00 37 00 43 00 38 00 30 00 34 00 31 00 30 00 31 00 30 00 36 00 31 00 00 00 31 00 32 00 33 00 33 00 00 00 00 00 
				//��
				//36 00 38 00 41 00 42 00 36 00 37 00 43 00 41 00 37 00 44 00 41 00 37 00 46 00 46 00 46 00 46 00 35 00 32 00 30 00 35 00 41 00 37 00 43 00 38 00 30 00 34 00 31 00 30 00 31 00 30 00 36 00 31 00 0A 00 31 00 32 00 33 00 33 00 0A 00 00 00 
				for (DWORD i = 0; i < dwDataSize; i+=2) {
					if(lpData[i] == '\0' && lpData[i+1] == '\0' && i < dwDataSize-2){
						lpData[i] = '\n';
					}
				}
				_stprintf(regDataCharBuffer, TEXT("%s"), lpData);
				currentRegValue.data = regDataCharBuffer;
				break;
 
            default:
				//regTypes[szValueNameBuffer] = TEXT("Unknown");
				currentRegValue.type = TEXT("Unknown");
				//_tprintf(TEXT("Unknown\n"));
				//regValues[szValueNameBuffer] = TEXT("");
 				currentRegValue.data = TEXT("");
               break;
            }

			regValues.push_back(currentRegValue);

			//���W�X�g���̒l�f�[�^���f�o�b�O���O�Ƃ��ĕۑ����悤�Ƃ���Ǝ��Ԃ������肷����̂Ń��x���i�グ���߂�
			if(debugLevel > 1){
				PrintDebugLog(L"RegEnumValue. name:");
				PrintDebugLog(currentRegValue.name);
				PrintDebugLog(L", type:");
				PrintDebugLog(currentRegValue.type);
				PrintDebugLog(L", data:");
				PrintDebugLog(currentRegValue.data);
				PrintDebugLog(L", Status:");
				wchar_t stringBuffer[32];
				_stprintf(stringBuffer, TEXT("%ld"), Status);
				PrintDebugLog(stringBuffer);
				PrintDebugLog(L"\n");
			}
        }
    }


	//�w��L�[�z���̃T�u�L�[�����擾

	TCHAR szSubkeyName[MAX_PATH];
    DWORD dwSubkeyNameSize;
    FILETIME ftLastWriteTime;

    for (DWORD dwIndex = 0; TRUE; dwIndex++) {
        
        DWORD dwResult;

        dwSubkeyNameSize = sizeof(szSubkeyName)/sizeof(szSubkeyName[0]);


		dwResult = RegEnumKeyEx(
			phkResult, 
            dwIndex, 
            szSubkeyName, 
            &dwSubkeyNameSize, 
            NULL, 
            NULL, 
            NULL, 
            &ftLastWriteTime);
    
        if (dwResult == ERROR_SUCCESS) {

			//subKey�\���̂�p�ӂ���
			subKey currentSubKey;

 			_TCHAR *subKeyNameBuffer;
			//new�œ��I�Ɋm�ۂ��Ȃ��Ə㏑�����Ă��܂�
			//http://www.geocities.jp/ky_webid/cpp/language/012.html
			subKeyNameBuffer = new _TCHAR[MAX_PATH];

			_TCHAR *subKeyFullPathNameBuffer;
			subKeyFullPathNameBuffer = new _TCHAR[MAX_PATH];

			FILETIME ftLocalTime;
            SYSTEMTIME LastWriteTime;
			//new�œ��I�Ɋm�ۂ��Ȃ��Ə㏑�����Ă��܂�
			//http://www.geocities.jp/ky_webid/cpp/language/012.html
            //TCHAR szLastWriteTime[32];
			_TCHAR *szLastWriteTime;
			szLastWriteTime = new _TCHAR[32];
            
            ::FileTimeToLocalFileTime(&ftLastWriteTime, &ftLocalTime);
            ::FileTimeToSystemTime(&ftLocalTime, &LastWriteTime);

            //_sntprintf(
			//	szLastWriteTime, 
            //    sizeof(szLastWriteTime)/sizeof(szLastWriteTime[0]),
            //    TEXT("%04d/%02d/%02d %02d:%02d:%02d"), 
            //    LastWriteTime.wYear, LastWriteTime.wMonth, LastWriteTime.wDay, 
            //    LastWriteTime.wHour, LastWriteTime.wMinute, LastWriteTime.wSecond);
            _stprintf(
				szLastWriteTime, 
                TEXT("%04d/%02d/%02d %02d:%02d:%02d"), 
                LastWriteTime.wYear, LastWriteTime.wMonth, LastWriteTime.wDay, 
                LastWriteTime.wHour, LastWriteTime.wMinute, LastWriteTime.wSecond);

            //_tprintf(TEXT("SubkeyName: %s, LastWriteTime: %s\n"), szSubkeyName, szLastWriteTime);
			_tcscpy_s(subKeyNameBuffer, MAX_PATH, szSubkeyName);
			//subKeys.push_back(subKeyNameBuffer);
			//currentSubKey.name = subKeyNameBuffer;
			currentSubKey.name = szSubkeyName;
			//_stprintf(subKeyFullPathNameBuffer, TEXT("%s\\%s"), subKeyChar, subKeyNameBuffer);
			//currentSubKey.fullPathName = subKeyFullPathNameBuffer;
			currentSubKey.fullPathName = subKeyChar + L"\\" + szSubkeyName;
			currentSubKey.lastWriteTime = szLastWriteTime;
			subKeys.push_back(currentSubKey);

			//���W�X�g���̗񋓃f�[�^���f�o�b�O���O�Ƃ��ĕۑ����悤�Ƃ���Ǝ��Ԃ������肷����̂Ń��x���i�グ���߂�
			if(debugLevel > 1){
				PrintDebugLog(L"RegEnumKeyEx. fullPathName:");
				PrintDebugLog(currentSubKey.fullPathName);
				PrintDebugLog(L", dwResult:");
				wchar_t stringBuffer[32];
				_stprintf(stringBuffer, TEXT("%ld"), dwResult);
				PrintDebugLog(stringBuffer);
				PrintDebugLog(L"\n");
			}

		} else if (dwResult == ERROR_NO_MORE_ITEMS) {
			//�񋓂��I����RegEnumKeyEx��ERROR_NO_MORE_ITEMS��Ԃ�
            //puts("ERROR_NO_MORE_ITEMS");
            break;
        } else {
			if(debugLevel > 0){
				//fwprintf_s(stderr, TEXT("RegEnumKeyEx error. hKeyChar;%s, subKeyChar:%s\n"),hKeyChar, subKeyChar);
				PrintDebugLog(L"Error:RegEnumKeyEx error. hKeyChar:");
				PrintDebugLog(hKeyChar);
				PrintDebugLog(L", subKeyChar:");
				PrintDebugLog(subKeyChar);
				PrintDebugLog(L"\n");
			}
			RegCloseKey(phkResult);
            return -4;
        }
 
    }
    RegCloseKey(phkResult);
	return 0;

}

//�擾�����f�[�^��W���o�͂ɏo���֐�
void RegList::displayAll(){
	//�l���X�g�̏o��
	vector<regValue>::iterator regValuesIterator = regValues.begin();  // �C�e���[�^�̃C���X�^���X��
	while(regValuesIterator != regValues.end()){
		_tprintf(TEXT("Value - Name:%s\tType:%s\tData:%s\n"), (*regValuesIterator).name.c_str(), (*regValuesIterator).type.c_str(), (*regValuesIterator).data.c_str());  // *���Z�q�ŊԐڎQ��
		++regValuesIterator;                 // �C�e���[�^���P�i�߂�
	}

	//�T�u�L�[���̏o��
	vector<subKey>::iterator subKeysIterator = subKeys.begin();  // �C�e���[�^�̃C���X�^���X��
	while(subKeysIterator != subKeys.end()){
		_tprintf(TEXT("Subkey - Name:%s\tFullPathName:%s\tLastWriteTime:%s\n"), (*subKeysIterator).name.c_str(), (*subKeysIterator).fullPathName.c_str(), (*subKeysIterator).lastWriteTime.c_str());  // *���Z�q�ŊԐڎQ��
		++subKeysIterator;                 // �C�e���[�^���P�i�߂�
	}
}

//�w�肳�ꂽ���O�̒l�̒��g�ƌ^��Ԃ��֐�
regValue RegList::getValue(wstring name){
//�߂�l�𕡐��Ԃ������̂ŁA�Q�Ɠn�����g���Ă���
//http://dixq.net/forum/viewtopic.php?f=3&t=3058
//http://www.geocities.jp/ky_webid/cpp/language/015.html

	//�������ɂ���
	//tolower�Ŕ�r����ƁA���{��Ȃǂŕ�����������@towlower���g��
	transform(name.begin (), name.end (), name.begin (), towlower);

	vector<regValue>::iterator regValuesIterator = regValues.begin();  // �C�e���[�^�̃C���X�^���X��
	while(regValuesIterator != regValues.end()){
		//_tprintf(TEXT("Name:%s\tType:%s\tData:%s\n"), (*regValuesIterator).name, (*regValuesIterator).type, (*regValuesIterator).data);  // *���Z�q�ŊԐڎQ��
		//if(_tcscmp(name, (*regValuesIterator).name.c_str()) == 0){
		if(name == (*regValuesIterator).lowerName){
			//*value = (*regValuesIterator).data.c_str();
			//*type = (*regValuesIterator).type;
			//_tprintf(TEXT("value: %s\n"), value);
			//���v������̂������0��Ԃ��ďI��
			return (*regValuesIterator);
		}
		// �C�e���[�^���P�i�߂�
		++regValuesIterator;
	}

/*

	//for ( vector<regValue>::iterator position = a.begin();
	//	(position = find_if(position, a.end(), &age5)) != a.end(); // �R�R�Ō���
	//	++position ) {
	//	cout << *position << endl;
	//}

	//_tprintf(TEXT("name: [%s]\n"), name);
	map<_TCHAR*, _TCHAR*>::iterator regValuesIterator = regValues.begin();
	while(regValuesIterator != regValues.end()){
		//�����Ƃ��Ďw�肳�ꂽ�l�̖��O�ƁAregValues�ɕۊǂ���Ă���l�̖��O�Ƃ��r
		//�����z��Ȃ̂�.find�͎g���Ȃ�
		if(_tcscmp(name, regValuesIterator->first) == 0){
			*value = regValuesIterator->second;
			//_tprintf(TEXT("value: %s\n"), value);
			//�^���X�g����Y��������̂𒲍�
			map<_TCHAR*, _TCHAR*>::iterator regTypesIterator = regTypes.find(regValuesIterator->first);
			if (regTypesIterator != regTypes.end()) {
				*type = regTypesIterator->second;
				//_tprintf(TEXT("type: %s\n"), type);
			}else{
				//�l������̂Ɍ^���ۑ�����Ă��Ȃ��ꍇ��-2��Ԃ�
				return -2;
			}
			//���v������̂������0��Ԃ��ďI��
			return 0;
		}
		++regValuesIterator;
	}
*/
	//�l��������Ȃ��ꍇ��-1��Ԃ�
	//return -1;
	regValue regValue;
	return regValue;
}

//���������T�u�L�[��񋓂���֐��i�Ă΂��x��1���Q�Ɠn�����čŌ�܂ŏo����߂�l��-1�j
subKey RegList::getNextSubKey()
{
	if(subKeysIndex >= subKeys.size()){
		subKey blankSubkey;
		return blankSubkey;
	}
	//*name = (subKeys.at(subKeysIndex)).name;
	//*fullPathName = (subKeys.at(subKeysIndex)).fullPathName;
	//++subKeysIndex;
	return subKeys.at(subKeysIndex++);
}

//�\�t�g�E�F�A���Ȃǂ̃��X�g���擾����N���X
class InventoryManager{
	private:
		//�L�[�@�@�F�l�̖��O
		//�l�@�@�@�F�l�̒��g
		//map<_TCHAR*, _TCHAR*> regValues;
		//�L�[�@�@�F�l�̖��O
		//�l�@�@�@�F�l�̃f�[�^�^
		//map<_TCHAR*, _TCHAR*> regTypes;
		//map<_TCHAR*, _TCHAR* > data;
		//�T�u�L�[�̔z��
		//_TCHAR** subKeys;���錾���ɂ̓G���[�ɂȂ�Ȃ����g���Ƃ�exe���G���[��������
		//vector<_TCHAR*> softwares;
		//int subKeysIndex;

		struct tm timestamp;

		hardwareInfo defaultHardwareInfo;
		hardwareInfo previousHardwareInfo;
		hardwareInfo currentHardwareInfo;

		vector<softwareInfo> previousSoftwareInfos;
		vector<softwareInfo> currentSoftwareInfos;

		vector<patchInfo> previousPatchInfos;
		vector<patchInfo> currentPatchInfos;


    //����
    //data[TEXT("Environment")][TEXT("ProductName")] = TEXT("Name");
	public:
		//�R���X�g���N�^
		InventoryManager(){

			if(debugLevel > 1){
				PrintDebugLog(L"InventoryManager::constructor\n");
			}

			//���ݓ����̎擾�i�o�̓t�@�C���ɏ������ށj
			//http://soundengine.jp/wordpress/tips/tutorial/176/
			//http://msdn.microsoft.com/en-us/library/a442x3ye(v=vs.80).aspx
			__time64_t long_time;
			_time64( &long_time ); 
			_localtime64_s( &timestamp, &long_time );

			wchar_t timestampChars [100];
			swprintf(timestampChars, 100, L"%04d%02d%02d%02d%02d%02d", timestamp.tm_year + 1900, timestamp.tm_mon + 1, timestamp.tm_mday, timestamp.tm_hour, timestamp.tm_min, timestamp.tm_sec);
			currentHardwareInfo.Timestamp = timestampChars;

			swprintf(timestampChars, 100, L"%04d/%02d/%02d %02d:%02d:%02d", timestamp.tm_year + 1900, timestamp.tm_mon + 1, timestamp.tm_mday, timestamp.tm_hour, timestamp.tm_min, timestamp.tm_sec);
			currentHardwareInfo.Timestamp2 = timestampChars;

			currentHardwareInfo.ToolVersion = VER_STR_PRODUCTVERSION;

		}

		//�����o�֐��̐錾
		//���͂��ꂽ����ۊǂ���֐�
		long setInputInfo(wstring editHardwareNo, wstring edit1, wstring edit2, wstring edit3, wstring edit4, wstring edit5, wstring edit6);
		//�n�[�h�E�F�A�̕K�v���i�R���s���[�^���A���[�U���AIP�AMAC���j���擾����֐�
		long getHardwareInfo();
		//�\�t�g�E�F�A���𒲂ׂ�֐��i���b�p�j
		long listUp();
		//�w�肳�ꂽ���W�X�g������\�t�g�E�F�A���𒲂ׂ�֐��i�{�́j
		long scanRegistry(BOOL scan64bit, BOOL scanHKCU);
		//WMI����p�b�`�����擾����֐��i�{�́j
		long getPatchFromWMI();
		//���e��S�o�͂���֐�
		wstring output();
		wstring outputSARMS();
		wstring outputSIMPLE();
		wstring outputAdvancedManager();

		//�\�t�g�E�F�A���̍\���̂�񋓂���֐�
		//softwareInfo getNextSoftware();
		//�C���x���g����ۑ�����֐�
		long save();
		//�O��̃C���x���g�������[�h����֐�
		long load();
		//wstring���t�@�C���ɏ������ފ֐�
		long fwriteWString(FILE* fp, wstring string);
		//wstring���t�@�C������ǂݍ��ފ֐�
		long freadWString(FILE* fp, wstring &string);
		//�O��擾���������擾����֐�
		//default.csv���珉���l�����[�h����֐�
		long loadDefaultSetting();
		//default.csv�����l�[������֐�
		long cleanupDefaultSetting();
		wstring getPreviousHardwareNoValue();
		wstring getPreviousValue1();
		wstring getPreviousValue2();
		wstring getPreviousValue3();
		wstring getPreviousValue4();
		wstring getPreviousValue5();
		wstring getPreviousValue6();

};

//���͂��ꂽ����ۊǂ���֐�
long InventoryManager::setInputInfo(wstring editHardwareNo, wstring edit1, wstring edit2, wstring edit3, wstring edit4, wstring edit5, wstring edit6){

	currentHardwareInfo.HardwareNoTitle = labelHardwareNo;
	currentHardwareInfo.Title1 = label1;
	currentHardwareInfo.Title2 = label2;
	currentHardwareInfo.Title3 = label3;
	currentHardwareInfo.Title4 = label4;
	currentHardwareInfo.Title5 = label5;
	currentHardwareInfo.Title6 = label6;
	currentHardwareInfo.HardwareNoValue = editHardwareNo;
	currentHardwareInfo.Value1 = edit1;
	currentHardwareInfo.Value2 = edit2;
	currentHardwareInfo.Value3 = edit3;
	currentHardwareInfo.Value4 = edit4;
	currentHardwareInfo.Value5 = edit5;
	currentHardwareInfo.Value6 = edit6;

	if(debugLevel > 1){
		PrintDebugLog(L"InventoryManager::setInputInfo ->\n");
		PrintDebugLog(currentHardwareInfo.HardwareNoTitle);
		PrintDebugLog(L":");
		PrintDebugLog(currentHardwareInfo.HardwareNoValue);
		PrintDebugLog(L"\n");
		PrintDebugLog(currentHardwareInfo.Title1);
		PrintDebugLog(L":");
		PrintDebugLog(currentHardwareInfo.Value1);
		PrintDebugLog(L"\n");
		PrintDebugLog(currentHardwareInfo.Title1);
		PrintDebugLog(L":");
		PrintDebugLog(currentHardwareInfo.Value1);
		PrintDebugLog(L"\n");

		PrintDebugLog(currentHardwareInfo.Title2);
		PrintDebugLog(L":");
		PrintDebugLog(currentHardwareInfo.Value2);
		PrintDebugLog(L"\n");
		PrintDebugLog(currentHardwareInfo.Title3);
		PrintDebugLog(L":");
		PrintDebugLog(currentHardwareInfo.Value3);
		PrintDebugLog(L"\n");
		PrintDebugLog(currentHardwareInfo.Title4);
		PrintDebugLog(L":");
		PrintDebugLog(currentHardwareInfo.Value4);
		PrintDebugLog(L"\n");
		PrintDebugLog(currentHardwareInfo.Title5);
		PrintDebugLog(L":");
		PrintDebugLog(currentHardwareInfo.Value5);
		PrintDebugLog(L"\n");
		PrintDebugLog(currentHardwareInfo.Title6);
		PrintDebugLog(L":");
		PrintDebugLog(currentHardwareInfo.Value6);
		PrintDebugLog(L"\n");
	}

	return 0;
}

//�n�[�h�E�F�A�̕K�v���i�R���s���[�^���A���[�U���AIP�AMAC���j���擾����֐�
long InventoryManager::getHardwareInfo(){

	if(debugLevel > 1){
		PrintDebugLog(L"InventoryManager::getHardwareInfo() ->\n");
	}

	//�R���s���[�^���̎擾
	//IP�A�h���X�͂��̕��@�Ŏ��Ǝ��Ȃ����Ƃ�����̂ŕʂ̕��@��
	//http://www.codeguru.com/csharp/csharp/cs_network/article.php/c6045/Obtain-all-IP-addresses-of-local-machine.htm

	WORD wVersionRequested;
	WSADATA wsaData;
	char computerName[255];
	//wchar_t *wname;
	_TCHAR wcomputerName[255];
	//PHOSTENT hostinfo;
	wVersionRequested = MAKEWORD( 1, 1 );
	//char *ip;

	if ( WSAStartup( wVersionRequested, &wsaData ) == 0 ){
		if( gethostname ( computerName, sizeof(computerName)) == 0){
			//printf("Host name: %s\n", computerName);
			mbstowcs(wcomputerName, computerName, sizeof(computerName));
			currentHardwareInfo.ComputerName = wcomputerName;
			if(debugLevel > 1){
				PrintDebugLog(L"gethostname. computerName:");
				PrintDebugLog(currentHardwareInfo.ComputerName);
				PrintDebugLog(L"\n");
			}
			/*IP�A�h���X�͂��̕��@�Ŏ��Ǝ��Ȃ����Ƃ�����̂ŕʂ̕��@�Ł@���O��������i���΍�j
			if((hostinfo = gethostbyname(computerName)) != NULL)
			{
				int nCount = 0;
				wchar_t stringBuffer[32];
				while(hostinfo->h_addr_list[nCount])
				{
					ip = inet_ntoa (*(struct in_addr *)hostinfo->h_addr_list[nCount]);
					_stprintf(stringBuffer, TEXT("#%d:%s\n"), ++nCount, ip);
					if(debugLevel > 1){
						PrintDebugLog(L"gethostbyname. IP:");
						PrintDebugLog(stringBuffer);
						PrintDebugLog(L"\n");
					}
				}
			}
			*/
			
		}
	}

	//MAC�A�h���X�AIP�A�h���X�̎擾
	IP_ADAPTER_INFO AdapterInfo[16];			// Allocate information for up to 16 NICs
	DWORD dwBufLen = sizeof(AdapterInfo);		// Save the memory size of buffer

	DWORD dwStatus = GetAdaptersInfo(			// Call GetAdapterInfo
		AdapterInfo,							// [out] buffer to receive data
		&dwBufLen);								// [in] size of receive data buffer
	assert(dwStatus == ERROR_SUCCESS);			// Verify return value is valid, no buffer overflow

	wchar_t MACAddressPart [10];
	DWORD tempAdapterIndex = ULONG_MAX;
	_TCHAR wIPAddress[16];
	_TCHAR wAdapterName[300];
	_TCHAR wDescription[300];
	_TCHAR wGatewayList[16];

	wstring tempMACAddress = L"";
	wstring tempIPAddress = L"";

	PIP_ADAPTER_INFO pAdapter = AdapterInfo;// Contains pointer to current adapter info
	do {
		//PrintMACaddress(pAdapter->Address);	// Print MAC address
        //printf("\tAdapter Addr: \t%ld\n", pAdapter->Address);
		//swprintf(MACAddress, 100,  _T( "%ld" ), pAdapter->Address);
		
		//MAC�A�h���X�擾�@���O�o�͎��̂��߂ɂ����Ŏ擾�@
		tempMACAddress = L"";
		for (UINT i = 0; i < pAdapter->AddressLength; i++) {
			//MAC�A�h���X���n�C�t����؂肵�Ȃ�
			//if (i == (pAdapter->AddressLength - 1)){
			//	swprintf(MACAddressPart, 10, L"%.2X",(int)pAdapter->Address[i]);
			//}
			//else{
			//	swprintf(MACAddressPart, 10, L"%.2X-",(int)pAdapter->Address[i]);
			//}
			swprintf(MACAddressPart, 10, L"%.2X",(int)pAdapter->Address[i]);
			tempMACAddress += MACAddressPart;

		}
		//IP�A�h���X�擾
		//tempIPAddress = pAdapter->IpAddressList.IpAddress.String;
		//char���܂�tchar�ɕϊ����Ă���wstring�ɑ��
		mbstowcs(wIPAddress, pAdapter->IpAddressList.IpAddress.String, sizeof(pAdapter->IpAddressList.IpAddress.String));
		tempIPAddress = wIPAddress;

		//���̑��擾
		mbstowcs(wAdapterName, pAdapter->AdapterName, sizeof(pAdapter->AdapterName));
		mbstowcs(wDescription, pAdapter->Description, sizeof(pAdapter->Description));
		mbstowcs(wGatewayList, pAdapter->GatewayList.IpAddress.String, sizeof(pAdapter->GatewayList.IpAddress.String));

		if(debugLevel > 0){
			PrintDebugLog(L"AdapterInfo. Index:");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%lu"), pAdapter->Index);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n MACAddress:");
			PrintDebugLog(tempMACAddress);
			PrintDebugLog(L"\n IPAddress:");
			PrintDebugLog(wIPAddress);
			PrintDebugLog(L"\n AdapterName:");
			PrintDebugLog(wAdapterName);
			PrintDebugLog(L"\n Description:");
			PrintDebugLog(wDescription);
			PrintDebugLog(L"\n Gateway:");
			PrintDebugLog(wGatewayList);
			PrintDebugLog(L"\n Type:");
			_stprintf(stringBuffer, TEXT("%u"), pAdapter->Type);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}

		//pAdapter->Index����ԏ��������̂�Primary Adapter�I
		if(pAdapter->Index < tempAdapterIndex){
			if(debugLevel > 1){
				PrintDebugLog(L" Index < tempIndex\n");
			}

			//����LAN��0.0.0.0�ɂȂ�b��΍�
			if(tempIPAddress != L"0.0.0.0"){
				tempAdapterIndex = pAdapter->Index;
				currentHardwareInfo.MACAddress = tempMACAddress;
				currentHardwareInfo.IPAddress = tempIPAddress;
				if(debugLevel > 1){
					PrintDebugLog(L" IPAddress != 0.0.0.0\n");
				}
			}
		}
        //MessageBox(NULL,MACAddress,_T("MAC"),MB_OK);
		pAdapter = pAdapter->Next;		// Progress through linked list
	}
	while(pAdapter);						// Terminate if last adapter

	//���݂̃��[�U���擾
	//http://okwave.jp/qa/q2597801.html
	//http://jehupc.exblog.jp/13175145/

	//_TCHAR wuserName[1024];�Ȃǂƒ����ɂ���ĕ�������������󕶎���ɂȂ����肷��
	//524�A64�ł��_��
	//MyDialog.cpp����Ăяo���ƂȂ�����������
	//DWORD�ɏ����l��ݒ肷��Ƒ��v�Ƃ����L����������
	//http://www.geocities.jp/midorinopage/Tips/getusername.html
	//�m���ɂ���ő��v�ɂȂ����c
	_TCHAR wuserName[1024];
	DWORD dwUserSize = 1024; // �擾�������[�U���̕�����̒���
	if ( ! GetUserName(wuserName,&dwUserSize) ){
		//return -1;
	}
	//MessageBox(NULL,wuserName,_T("���[�U��"),MB_OK);
	currentHardwareInfo.UserName = wuserName;

	//OS�̎擾
	currentHardwareInfo.OSName = getOSName();

	//�@��A�x���_�[�̎擾�@�~���W�X�g������WMI�ɐ؂�ւ�
	
	HKEY hKey;
	TCHAR szSystemProductName[1024] = TEXT("");
	TCHAR szSystemManufacturer[1024] = TEXT("");
	TCHAR szProcessorName[1024] = TEXT("");
	DWORD dwBufLenSP = sizeof(TCHAR) * sizeof(szSystemProductName);
	DWORD dwBufLenSM = sizeof(TCHAR) * sizeof(szSystemManufacturer);
	DWORD dwBufLenPN = sizeof(TCHAR) * sizeof(szProcessorName);
	LONG lRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, TEXT("HARDWARE\\DESCRIPTION\\System\\BIOS"), 0, KEY_QUERY_VALUE, &hKey);
	if (lRet == ERROR_SUCCESS) {
		lRet = RegQueryValueEx(hKey, TEXT("SystemManufacturer"), NULL, NULL, (LPBYTE)szSystemManufacturer, &dwBufLenSM);
		lRet = RegQueryValueEx(hKey, TEXT("SystemProductName"), NULL, NULL, (LPBYTE)szSystemProductName, &dwBufLenSP);
		RegCloseKey(hKey);
	}
	currentHardwareInfo.SystemManufacturer = szSystemManufacturer;
	currentHardwareInfo.SystemProductName = szSystemProductName;
	
	//currentHardwareInfo.SystemManufacturer = getWMI(L"SELECT * FROM Win32_ComputerSystem", L"Manufacturer");
	//currentHardwareInfo.SystemProductName = getWMI(L"SELECT * FROM Win32_ComputerSystem", L"Model");

	//���z�����ۂ�
	//Hyper-V
	if(currentHardwareInfo.SystemProductName == L"Virtual Machine"){
		currentHardwareInfo.Virtualization = L"1";
	//VMware
	}else if(currentHardwareInfo.SystemProductName == L"VMWare Virtual Platform"){
		currentHardwareInfo.Virtualization = L"1";
	//Oracle
	}else if(currentHardwareInfo.SystemProductName == L"Virtual Box"){
		currentHardwareInfo.Virtualization = L"1";
	//Xen
	}else if(currentHardwareInfo.SystemProductName == L"HVM domU"){
		currentHardwareInfo.Virtualization = L"1";
	//Qemu with KVM
	}else if(currentHardwareInfo.SystemProductName == L"KVM"){
		currentHardwareInfo.Virtualization = L"1";
	//Qemu (emulated)
	}else if(currentHardwareInfo.SystemProductName == L"Bochs"){
		currentHardwareInfo.Virtualization = L"1";
	}else{
		currentHardwareInfo.Virtualization = L"0";
	}

	//CPU���̎擾�@���W�X�g������WMI�ɐ؂�ւ�
	/*
	lRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, TEXT("HARDWARE\\DESCRIPTION\\System\\CentralProcessor\\0"), 0, KEY_QUERY_VALUE, &hKey);
	if (lRet == ERROR_SUCCESS) {
		lRet = RegQueryValueEx(hKey, TEXT("ProcessorNameString"), NULL, NULL, (LPBYTE)szProcessorName, &dwBufLenPN);
		RegCloseKey(hKey);
	}
	//MessageBox(NULL,szProcessorName,_T("pn"),MB_OK);
	currentHardwareInfo.ProcessorName = szProcessorName;
	*/
	//currentHardwareInfo.ProcessorName = getWMI(L"SELECT * FROM Win32_Processor", L"Name");

	//CPU�֘A���̎擾
	getCPUInfo(
		currentHardwareInfo.ProcessorName,
		currentHardwareInfo.ProcessorMaxClockSpeed,
		currentHardwareInfo.NumberOfProcessors,
		currentHardwareInfo.NumberOfCores,
		currentHardwareInfo.NumberOfLogicalProcessors
	);

	//�V���A���ԍ��̎擾
	currentHardwareInfo.IdentifyingNumber = getWMI(L"SELECT * FROM Win32_ComputerSystemProduct", L"IdentifyingNumber");

	//�������e�ʂ̎擾
	currentHardwareInfo.TotalPhysicalMemory = getWMI(L"SELECT * FROM Win32_ComputerSystem", L"TotalPhysicalMemory");

	//wstring test = getWMI(L"SELECT * FROM Win32_ComputerSystem", L"Model");
	//MessageBox(NULL,test.c_str(),L"test",MB_OK);

	return 0;

}

//�\�t�g�E�F�A���𒲂ׂ�֐��i���b�p�j
long InventoryManager::listUp(){

	if(debugLevel > 1){
		PrintDebugLog(L"InventoryManager::listUp() ->\n");
	}
	//clear���Ȃ��ƍēxlistUp���Ăяo�����Ƃ��ɈȑO�̃f�[�^�������Ă��������Ȃ�
	currentSoftwareInfos.clear();
	currentPatchInfos.clear();

	long status;
	if(debugLevel > 1){
		PrintDebugLog(L"scanRegistry(FALSE, FALSE) ->\n");
	}
	status = scanRegistry(FALSE, FALSE);

	if(status != 0){
		if(debugLevel > 1){
			//fwprintf_s(stderr, TEXT("scanRegistry(scan64bit=FALSE, scanHKCU=FALSE) error. status:%d\n"),status);
			PrintDebugLog(L"Error:scanRegistry(scan64bit=FALSE, scanHKCU=FALSE) error. status:");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%ld"), status);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
	}
	if(debugLevel > 1){
		PrintDebugLog(L"scanRegistry(FALSE, TRUE) ->\n");
	}
	status = scanRegistry(FALSE, TRUE);

	if(status != 0){
		if(debugLevel > 1){
			//fwprintf_s(stderr, TEXT("scanRegistry(scan64bit=FALSE, scanHKCU=TRUE) error. status:%d\n"),status);
			PrintDebugLog(L"Error:scanRegistry(scan64bit=FALSE, scanHKCU=TRUE) error. status:");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%ld"), status);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
	}
	//64bitOS�Ȃ�64bit�L�[�����ׂ�
	if(IsWow64()){
		if(debugLevel > 1){
			PrintDebugLog(L"scanRegistry(TRUE, FALSE) ->\n");
		}
		status = scanRegistry(TRUE, FALSE);
		if(status != 0){
			if(debugLevel > 1){
				//fwprintf_s(stderr, TEXT("scanRegistry(scan64bit=TRUE, scanHKCU=FALSE) error. status:%d\n"),status);
				PrintDebugLog(L"Error:scanRegistry(scan64bit=TRUE, scanHKCU=FALSE) error. status:");
				wchar_t stringBuffer[32];
				_stprintf(stringBuffer, TEXT("%ld"), status);
				PrintDebugLog(stringBuffer);
				PrintDebugLog(L"\n");
			}
		}
		//HKCU��32bit�ł�64bit�ł��ς��Ȃ��̂Œ��ׂȂ�
		//status = scanRegistry(TRUE, TRUE);
		//if(status != 0){
		//	if(debugLevel > 0){
		//		fwprintf_s(stderr, TEXT("scanRegistry(scan64bit=TRUE, scanHKCU=TRUE) error. status:%d\n"),status);
		//	}
		//}
	}


	//WMI����Windows�p�b�`���̎擾
	//WMI����擾����ƒx��
	//Windows Update Agent API(WUI)�o�R�Ŏ����@�����邪�A��v������̂����Ȃ��H
	//����WMI���������߂炵��
	//https://social.msdn.microsoft.com/Forums/vstudio/en-US/2002ed7c-fc40-43ac-9300-f1be27f16bbd/c-understand-that-windows-update-installed?forum=vcgeneral
	if(debugLevel > 1){
		PrintDebugLog(L"HotFix from WMI ->\n");
	}
	status = getPatchFromWMI();
	if(status != 0){
		if(debugLevel > 1){
			PrintDebugLog(L"Error:getPatchFromWMI() error. status:");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%ld"), status);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
	}

	//softwareInfos�̃\�[�g
	//�\�[�g�p�t�B�[���h(compareName)�̏���
	vector<softwareInfo>::iterator softwareInfoIterator = currentSoftwareInfos.begin();  // �C�e���[�^�̃C���X�^���X��
	while(softwareInfoIterator != currentSoftwareInfos.end()){
		//��r�p�������p��
		wstring compareName = (*softwareInfoIterator).Name.c_str();

		//�ꕶ�����؂�o���ăR�[�h�|�C���g��t���Ă���
		//�i����R�[�h�����㍬�����Ă����Ƃ��Ɋm�F���邽�߁j
		//�ꕶ�����؂�o�����@ http://d.hatena.ne.jp/minus9d/20120517/1337265190
		//const wchar_t * wstrtest = compareName.c_str();
		//wchar_t buffer [100];
		//for (int i = 0; i < wcslen(wstrtest); ++i) {
		//	swprintf(buffer, 100, L"%c:U+%lX ", wstrtest[i], wstrtest[i]);
		//	(*softwareInfoIterator).nameCodePoint += buffer;
		//}
		wchar_t buffer [100];
		for(int i = 0; i < (int)compareName.size(); ++i){
			swprintf(buffer, 100, L"%c:U+%lX ", compareName[i], compareName[i]);
			(*softwareInfoIterator).NameCodePoint += buffer;
		}
		//wprintf(L"\n");

		//wstring compareName = L"�ɂق񂲓��{��123�P�Q�R�J�^�J�i";
		//��r�p��������������ɂ���
		//tolower�Ŕ�r����ƁA���{��Ȃǂŕ�����������@towlower���g��
		transform(compareName.begin (), compareName.end (), compareName.begin (), towlower);

		//�v���O�����Ƌ@�\�̃\�[�g���[���͈ȉ��̂Ƃ���
		//(1)�Ђ炪�ȁA�J�^�J�i�A���p�J�^�J�i�͋�ʂ��Ȃ�
		//(2)�p���̑啶���A�������͋�ʂ��Ȃ�
		//(3)�������茾���Ɛ������p�ꁨ���ȁ������̏��ɕ���
		//(4)�����������Ă���ƁA�����͐��l�Ƃ��ă\�[�g����
		//   ��2000�܂ł͕��ʂɕ������������̂�XP�ȍ~���̃��[���ɂȂ����炵��
		//     �G�N�X�v���[���̃t�@�C�����t�@�C�����Ń\�[�g�����Ƃ��Ɠ��l
		//     1�̃t�@�C�����̒��ɕ����̐��l����������ꍇ�́A�ŏ��̂��̂��ΏۂƂȂ�
		//     �Q�l http://support.microsoft.com/default.aspx?kbid=319827
		//          http://www.atmarkit.co.jp/fwin2k/win2ktips/342xpsort/xpsort.html
		//     ������������邽�߂̊֐��Ƃ���StrCmpLogicalW���񋟂���Ă��邪�AXP�ȍ~�ł����g���Ȃ��̂œƎ���������
		//     http://msdn.microsoft.com/en-us/library/bb759947%28VS.85%29.aspx
		//����ŁA�\�[�g���͈ȉ��̂Ƃ���
		//���p���������p�p�����S�p�Ђ炪�ȁ��S�p���������p�J�i���S�p�������S�p�p��
		//���p�J�i�΍�Ƃ���LCMAP_FULLWIDTH����ƁA�p���̑O�ɂЂ炪�ȂƊ��������Ă��܂��̂�NG
		//�t��LCMAP_HALFWIDTH����ƁA�Ђ炪�Ȃ̑O�Ɋ��������Ă��܂��Ă����NG
		//LCMAP_FULLWIDTH�ALCMAP_HALFWIDTH�͈��e�����傫���̂ŁA���p�J�i�͓Ǝ������őS�p�J�i�ɕϊ�����
		//��LCMAP_FULLWIDTH�ŁuEPSON�������ײ�ޥհè�è�v���u�����������v�����^�h���C�o�E���[�e�B���e�B�������v�ƁA�����Ɂu�������v���Ȃ����t���Ă��܂���������B
		//  �������ꂽcsv���o�C�i���G�f�B�^�Ō��Ă��A����R�[�h��DEL�͕t���Ă��Ȃ��悤�Ȃ̂Ō����s��

		//��������������
		//���\�[�g�͉��ł���Ă��邩�A�C���X�g�[�����H����Ƃ��o�[�W�����H���m�F

		compareName = HankakuKana2ZenkakuKana(compareName);

		_TCHAR katakanaTcharBuffer[1024];
		wstring katakanaWstringBuffer;
		LCMapString(LOCALE_USER_DEFAULT, LCMAP_HIRAGANA, compareName.c_str(), (int)compareName.size(), katakanaTcharBuffer, 1024);
		//LCMapString(LOCALE_USER_DEFAULT, LCMAP_HIRAGANA | LCMAP_FULLWIDTH, compareName.c_str(), compareName.size(), katakanaTcharBuffer, 1024);
		//LCMapString(LOCALE_USER_DEFAULT, LCMAP_KATAKANA | LCMAP_HALFWIDTH, compareName.c_str(), compareName.size(), katakanaTcharBuffer, 1024);
		//LCMapString(LOCALE_USER_DEFAULT, LCMAP_HALFWIDTH, compareName.c_str(), compareName.size(), katakanaTcharBuffer, 1024);
		//LCMapString(LOCALE_USER_DEFAULT, NORM_IGNORECASE | NORM_IGNOREKANATYPE | NORM_IGNOREWIDTH, compareName.c_str(), compareName.size(), katakanaTcharBuffer, 1024);
		katakanaWstringBuffer = katakanaTcharBuffer;
		compareName = katakanaWstringBuffer.substr(0,compareName.size());

		(*softwareInfoIterator).CompareName = compareName;
		++softwareInfoIterator;                 // �C�e���[�^���P�i�߂�
	}
	sort(currentSoftwareInfos.begin(), currentSoftwareInfos.end(), softwareInfo::nameCompare);

	//Oracle�Ή�
	//Oracle���v���O�����Ƌ@�\�Ɍ���Ȃ��Ă�ini�Őݒ肳��Ă�����擾����
	if(iniOracleCompatible > 0){
		//Oracle�̃R�}���h����͂��Aoracle.inv�Ƃ��Ĉ�x�ۑ�

		//_wgetenv�ł͑��݂��Ȃ����̂̊��ϐ������悤�Ƃ���ƃG���[��������̂ň��S��_wgetenv_s���g���悤�ɕύX
		//http://answers.flyppdevportal.com/MVC/Post/Thread/aa519567-8913-431a-bbf8-4e88254f0cb7?category=vcgeneral
		//https://github.com/MicrosoftDocs/cpp-docs/blob/master/docs/c-runtime-library/reference/getenv-s-wgetenv-s.md
		size_t requiredSize;
		//wstring env = L"OneDrive";
		wstring env = L"ORACLE_HOME";
		_wgetenv_s(&requiredSize, NULL, 0, env.c_str()); //OneDrive
		if (requiredSize == 0){
			//printf("ORACLE_HOME doesn't exist!\n");
			//exit(1);
			//wstring msg = L"ORACLE_HOME not found";
			//MessageBox(NULL,msg.c_str(),L"test",MB_OK);
		}else{
			//wstring msg = L"ORACLE_HOME found";
			//MessageBox(NULL,msg.c_str(),L"test",MB_OK);
			TCHAR* libvar;
			libvar = (TCHAR*) malloc(requiredSize * sizeof(TCHAR));
			if (!libvar){
			  //printf("Failed to allocate memory!\n");
			  //exit(1);
			}else{
				_wgetenv_s( &requiredSize, libvar, requiredSize, env.c_str());
				wstring program = L"C:\\Windows\\System32\\cmd.exe";
				wstring programArguments = libvar;
				wstring oracle_output_file = L"OracleInventory.dat";
				programArguments = L"/K " + programArguments + L"\\OPatch\\opatch lsinventory > " + oracle_output_file;
				//MessageBox(NULL,program.c_str(),L"test",MB_OK);
				//MessageBox(NULL,programArguments.c_str(),L"test",MB_OK);
				RunProcess(program, programArguments, false);

				//�ȉ��������Ă�SJIS�̓��{�ꂪ������������
				//_tsetlocale(LC_ALL, L".ACP");
				//_tsetlocale(LC_ALL, L"japanese");
				//wcout.imbue(std::locale("Japanese"));
				//wcout.imbue(std::locale(""));

				//�ȉ���������SJIS�̓��{��𕶎����������ɓǂݍ��܂�����
				//https://nullnull.hatenablog.com/entry/20120629/1340935277
				//�K�v�ŏ����ȊO�̓R�����g�A�E�g
				//ios_base::sync_with_stdio(false);
				locale default_loc("");
				locale::global(default_loc);
				//locale ctype_default(locale::classic(), default_loc, locale::ctype); //��
				//wcout.imbue(ctype_default);
				//wcin.imbue(ctype_default);

				wifstream ifs(oracle_output_file.c_str());
				//�t�@�C�������݂���Ȃ璆�g���擾
				if(ifs.is_open()){
					//wstring str((istreambuf_iterator<TCHAR>(ifs)), std::istreambuf_iterator<TCHAR>());
					wstring line;
					bool oracle_products_found = false;
					int sprit_pos;
					wstring oracle_products_name;
					wstring oracle_products_ver;
					while (getline(ifs, line)){
						//���i���񋓌�Ƀt���O��false��
						if(line.find(L"�̐��i���C���X�g�[������Ă��܂��B") != std::string::npos) {
							oracle_products_found = false;
						}else if(line.find(L"products installed in this Oracle Home.") != std::string::npos) {
							oracle_products_found = false;
						}
						//���i���擾
						if(oracle_products_found && line != L""){
							sprit_pos = line.find(L"    ");
							if(sprit_pos != std::string::npos){
								oracle_products_name = line.substr(0,sprit_pos);
								//oracle_products_name = L"\"" + oracle_products_name + L"\"";
								//MessageBox(NULL,oracle_products_name.c_str(),L"test",MB_OK);
								oracle_products_ver = line.substr(sprit_pos, line.size()-sprit_pos);
								oracle_products_ver = Replace(oracle_products_ver, L" ", L"");
								//oracle_products_ver = L"\"" + oracle_products_ver + L"\"";
								//MessageBox(NULL,oracle_products_ver.c_str(),L"test",MB_OK);
							}else{
								oracle_products_name = line;
								oracle_products_ver = L"";
							}
							//softwareInfo�\���̂�p��
							softwareInfo currentSoftwareInfo;
							currentSoftwareInfo.Name = oracle_products_name;
							currentSoftwareInfo.Vendor = L"Oracle Corporation";
							currentSoftwareInfo.Version = oracle_products_ver;
							currentSoftwareInfos.push_back(currentSoftwareInfo);
						}
						//���i���񋓒��O�Ƀt���O��true��
						if(line.find(L"�C���X�g�[�����ꂽ�ŏ�ʐ��i") != std::string::npos) {
							oracle_products_found = true;
						}else if(line.find(L"Installed Top-level Products") != std::string::npos) {
							oracle_products_found = true;
						}
												
					}
				}
				//�ꎞ�t�@�C���폜�@�������Aini��2���w�肳��Ă���Έꎞ�t�@�C�����c��
				if(iniOracleCompatible != 2){
					DeleteFile(oracle_output_file.c_str());
				}
			}
		}

	}


	//patchInfos�̃\�[�g
	//�\�[�g�p�t�B�[���h(compareName)�̏���
	vector<patchInfo>::iterator patchInfoIterator = currentPatchInfos.begin();  // �C�e���[�^�̃C���X�^���X��
	while(patchInfoIterator != currentPatchInfos.end()){
		//��r�p�������p��
		wstring compareName = (*patchInfoIterator).ParentName + (*patchInfoIterator).Name;

		//�ꕶ�����؂�o���ăR�[�h�|�C���g��t���Ă���
		//�i����R�[�h�����㍬�����Ă����Ƃ��Ɋm�F���邽�߁j
		//�ꕶ�����؂�o�����@ http://d.hatena.ne.jp/minus9d/20120517/1337265190
		wchar_t buffer [100];
		for(int i = 0; i < (int)compareName.size(); ++i){
			swprintf(buffer, 100, L"%c:U+%lX ", compareName[i], compareName[i]);
			(*patchInfoIterator).NameCodePoint += buffer;
		}

		//wstring compareName = L"�ɂق񂲓��{��123�P�Q�R�J�^�J�i";
		//��r�p��������������ɂ���
		//tolower�Ŕ�r����ƁA���{��Ȃǂŕ�����������@towlower���g��
		transform(compareName.begin (), compareName.end (), compareName.begin (), towlower);

		//�v���O�����Ƌ@�\�̃\�[�g���[���͈ȉ��̂Ƃ���
		//(1)�Ђ炪�ȁA�J�^�J�i�A���p�J�^�J�i�͋�ʂ��Ȃ�
		//(2)�p���̑啶���A�������͋�ʂ��Ȃ�
		//(3)�������茾���Ɛ������p�ꁨ���ȁ������̏��ɕ���
		//(4)�����������Ă���ƁA�����͐��l�Ƃ��ă\�[�g����
		//   ��2000�܂ł͕��ʂɕ������������̂�XP�ȍ~���̃��[���ɂȂ����炵��
		//     �G�N�X�v���[���̃t�@�C�����t�@�C�����Ń\�[�g�����Ƃ��Ɠ��l
		//     1�̃t�@�C�����̒��ɕ����̐��l����������ꍇ�́A�ŏ��̂��̂��ΏۂƂȂ�
		//     �Q�l http://support.microsoft.com/default.aspx?kbid=319827
		//          http://www.atmarkit.co.jp/fwin2k/win2ktips/342xpsort/xpsort.html
		//     ������������邽�߂̊֐��Ƃ���StrCmpLogicalW���񋟂���Ă��邪�AXP�ȍ~�ł����g���Ȃ��̂œƎ���������
		//     http://msdn.microsoft.com/en-us/library/bb759947%28VS.85%29.aspx
		//����ŁA�\�[�g���͈ȉ��̂Ƃ���
		//���p���������p�p�����S�p�Ђ炪�ȁ��S�p���������p�J�i���S�p�������S�p�p��
		//���p�J�i�΍�Ƃ���LCMAP_FULLWIDTH����ƁA�p���̑O�ɂЂ炪�ȂƊ��������Ă��܂��̂�NG
		//�t��LCMAP_HALFWIDTH����ƁA�Ђ炪�Ȃ̑O�Ɋ��������Ă��܂��Ă����NG
		//LCMAP_FULLWIDTH�ALCMAP_HALFWIDTH�͈��e�����傫���̂ŁA���p�J�i�͓Ǝ������őS�p�J�i�ɕϊ�����
		//��LCMAP_FULLWIDTH�ŁuEPSON�������ײ�ޥհè�è�v���u�����������v�����^�h���C�o�E���[�e�B���e�B�������v�ƁA�����Ɂu�������v���Ȃ����t���Ă��܂���������B
		//  �������ꂽcsv���o�C�i���G�f�B�^�Ō��Ă��A����R�[�h��DEL�͕t���Ă��Ȃ��悤�Ȃ̂Ō����s��

		//��������������
		//���\�[�g�͉��ł���Ă��邩�A�C���X�g�[�����H����Ƃ��o�[�W�����H���m�F

		compareName = HankakuKana2ZenkakuKana(compareName);

		_TCHAR katakanaTcharBuffer[1024];
		wstring katakanaWstringBuffer;
		LCMapString(LOCALE_USER_DEFAULT, LCMAP_HIRAGANA, compareName.c_str(), (int)compareName.size(), katakanaTcharBuffer, 1024);
		//LCMapString(LOCALE_USER_DEFAULT, LCMAP_HIRAGANA | LCMAP_FULLWIDTH, compareName.c_str(), compareName.size(), katakanaTcharBuffer, 1024);
		//LCMapString(LOCALE_USER_DEFAULT, LCMAP_KATAKANA | LCMAP_HALFWIDTH, compareName.c_str(), compareName.size(), katakanaTcharBuffer, 1024);
		//LCMapString(LOCALE_USER_DEFAULT, LCMAP_HALFWIDTH, compareName.c_str(), compareName.size(), katakanaTcharBuffer, 1024);
		//LCMapString(LOCALE_USER_DEFAULT, NORM_IGNORECASE | NORM_IGNOREKANATYPE | NORM_IGNOREWIDTH, compareName.c_str(), compareName.size(), katakanaTcharBuffer, 1024);
		katakanaWstringBuffer = katakanaTcharBuffer;
		compareName = katakanaWstringBuffer.substr(0,compareName.size());

		(*patchInfoIterator).CompareName = compareName;
		++patchInfoIterator;                 // �C�e���[�^���P�i�߂�
	}
	sort(currentPatchInfos.begin(), currentPatchInfos.end(), patchInfo::nameCompare);


	return 0;
}

//�w�肳�ꂽ���W�X�g������\�t�g�E�F�A���𒲂ׂ�֐��i�{�́j
long InventoryManager::scanRegistry(BOOL scan64bit, BOOL scanHKCU){

	//_tprintf(TEXT("scanRegistry\n"));
	//Uninstall�L�[�̎擾
	wstring uninstallHKeyName;
	wstring uninstallSubKeyName;
	RegList uninstallRegList;
	long status;
	if(scanHKCU == FALSE){
		uninstallHKeyName = L"HKEY_LOCAL_MACHINE";
	}else{
		uninstallHKeyName = L"HKEY_CURRENT_USER";
	}
	//uninstallSubKeyChar = TEXT("Microsoft");
	uninstallSubKeyName = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall";
	status = uninstallRegList.setHKeyAndSubKey(uninstallHKeyName, uninstallSubKeyName, scan64bit);
	if(status != 0){
		if(debugLevel > 0){
			//fwprintf_s(stderr, TEXT("uninstallRegList.setHKeyAndSubKey error. uninstallHKeyName;%s, uninstallSubKeyName:%s, scan64bit:%s\n"),uninstallHKeyName, uninstallSubKeyName, scan64bit);
			PrintDebugLog(L"Error:uninstallRegList.setHKeyAndSubKey error. uninstallHKeyName:");
			PrintDebugLog(uninstallHKeyName);
			PrintDebugLog(L", uninstallSubKeyName:");
			PrintDebugLog(uninstallSubKeyName);
			PrintDebugLog(L", scan64bit:");
			if(scan64bit){
				PrintDebugLog(L"TRUE");
			}else{
				PrintDebugLog(L"FALSE");
			}
			PrintDebugLog(L"\n");
		}
		return -1;
	}
	//uninstallRegList.setaHKeyAndSubKey(TEXT("HKEY_LOCAL_MACHINE"), TEXT("SOFTWARE\\Protector Suite QL"), TRUE);

	// HKEY_LOCAL_MACHINE /subkey "SOFTWARE\\Protector Suite QL"

	//return -1;

	//Uninstall�L�[�z���̃T�u�L�[���Ƃɏ���
	//_TCHAR* uninstallSubKeySubkeyName=TEXT("");
	//_TCHAR* uninstallSubKeySubkeyFullPathName=TEXT("");
	subKey uninstallSubKeySubkey;
	RegList uninstallSubkeyRegList;
	//long SystemComponentExist;
	//_TCHAR* uninstallSubkeyRegValue=TEXT("");
	//_TCHAR* uninstallSubkeyRegType=TEXT("");
	//regValue SystemComponentValue;

	//while(uninstallRegList.getNextSubKey(&uninstallSubKeySubkeyName, &uninstallSubKeySubkeyFullPathName) == 0){
	while(TRUE){
		//������擾����l�ɂ���softwareInfo�\���̂�p�ӂ���
		softwareInfo currentSoftwareInfo;
		patchInfo currentPatchInfo;
		if(scan64bit){
			if(scanHKCU){
				currentSoftwareInfo.ScanRegistryRange = L"64bit/HKCU";
				currentPatchInfo.ScanRegistryRange = L"64bit/HKCU";
			}else{
				currentSoftwareInfo.ScanRegistryRange = L"64bit/HKLM";
				currentPatchInfo.ScanRegistryRange = L"64bit/HKLM";
			}
		}else{
			if(scanHKCU){
				currentSoftwareInfo.ScanRegistryRange = L"32bit/HKCU";
				currentPatchInfo.ScanRegistryRange = L"32bit/HKCU";
			}else{
				currentSoftwareInfo.ScanRegistryRange = L"32bit/HKLM";
				currentPatchInfo.ScanRegistryRange = L"32bit/HKLM";
			}
		}

		if(debugLevel > 2){
			PrintDebugLog(L"InventoryManager::scanRegistry() main sequence->\n");
		}

		uninstallSubKeySubkey = uninstallRegList.getNextSubKey();
		//�T�u�L�[������Ȃ�I��
		if(uninstallSubKeySubkey.name == L""){
			if(debugLevel > 2){
				PrintDebugLog(L"break. uninstallSubKeySubkey.name == \"\"\n");
			}
			break;
		}
		//_tprintf(TEXT("subKey: %s\n"), uninstallSubKeySubkeyFullPathName);
		//RegList uninstallSubkeyRegList = new RegList;
		uninstallSubkeyRegList.setHKeyAndSubKey(uninstallHKeyName, uninstallSubKeySubkey.fullPathName, scan64bit);
		//#SystemComponent�Ƃ������W�X�g���l������A���^��REG_DWORD�ł��̒l��1�̏ꍇ�A�������I��
		regValue SystemComponentValue = uninstallSubkeyRegList.getValue(L"SystemComponent");
		//if((uninstallSubkeyRegListStatus == 0) && (_tcscmp(uninstallSubkeyRegType, TEXT("REG_DWORD")) == 0) && (_tcscmp(uninstallSubkeyRegValue, TEXT("1")) == 0)){
		if(SystemComponentValue.name != L"" && SystemComponentValue.type == L"REG_DWORD" && SystemComponentValue.data == L"1"){
			if(debugLevel > 2){
				PrintDebugLog(L"continue. SystemComponentValue.name != \"\"\n");
			}
			continue;
		}
		//������r�̂��߃��W�X�g���L�[���擾
		currentSoftwareInfo.RegistoryKey = uninstallSubKeySubkey.name;
		//���������؁���
		//Publisher�Ƃ������W�X�g���l������A���^��REG_SZ�܂���REG_EXPAND_SZ�ł��̒l��""�łȂ��ꍇ�A���X�g�ɒǉ�����
		//�i������REG_MULTI_SZ��NG�j
		regValue PublisherValue = uninstallSubkeyRegList.getValue(L"Publisher");
		if(PublisherValue.name != L""){
			if(PublisherValue.type == L"REG_SZ" && PublisherValue.data != L""){
				currentSoftwareInfo.Vendor = PublisherValue.data;
				currentPatchInfo.Vendor = PublisherValue.data;
			}else if(PublisherValue.type == L"REG_EXPAND_SZ" && PublisherValue.data != L""){
				currentSoftwareInfo.Vendor = PublisherValue.data;
				currentPatchInfo.Vendor = PublisherValue.data;
			}
		}
		//���������؁���
		//DisplayVersion�Ƃ������W�X�g���l������A���^��REG_SZ�܂���REG_EXPAND_SZ�ł��̒l��""�łȂ��ꍇ�A���X�g�ɒǉ�����
		//�i������REG_MULTI_SZ��NG�j
		regValue DisplayVersionValue = uninstallSubkeyRegList.getValue(L"DisplayVersion");
		if(DisplayVersionValue.name != L""){
			if(DisplayVersionValue.type == L"REG_SZ" && DisplayVersionValue.data != L""){
				currentSoftwareInfo.Version = DisplayVersionValue.data;
				currentPatchInfo.Version = DisplayVersionValue.data;
			}else if(DisplayVersionValue.type == L"REG_EXPAND_SZ" && DisplayVersionValue.data != L""){
				currentSoftwareInfo.Version = DisplayVersionValue.data;
				currentPatchInfo.Version = DisplayVersionValue.data;
			}
		}
		//WindowsInstaller�Ƃ������W�X�g���l������A���^��REG_DWORD�ł��̒l��1�̏ꍇ�c
		//�iREG_QWORD��REG_DWORD�Ǝ����悤�Ȃ��̂����AREG_QWORD���ƒl��1�ł��_���j
		regValue WindowsInstallerValue = uninstallSubkeyRegList.getValue(L"WindowsInstaller");
		if(WindowsInstallerValue.name != L"" && WindowsInstallerValue.type == L"REG_DWORD" && WindowsInstallerValue.data == L"1"){
			if(debugLevel > 2){
				PrintDebugLog(L"hold. WindowsInstallerValue.data == \"1\"\n");
			}
			//������HKLM�z���łȂ��Ƃ��́A�������I��
			//���ԈႢ�B�S�Ă̏ꍇ�ɍs��
			//������HKLM�z���Ȃ炳��ɒ��ׂ�K�v������
			//�܂�GUID�̃t�H�[�}�b�g�ɍ��v���邩�ǂ����m�F
			//�iGUID�͖{��16�i���������e����Ȃ����A�v���O�����̒ǉ��ƍ폜�ł̕\���ł�0-9a-fA-F�ȊO�ł��A���t�@�x�b�g�ł���Βʂ�j
			//GUID�͂���Ȃ́�{EE936C7A-EA40-31D5-9B65-8E3E089C3828}
			if(uninstallSubKeySubkey.name.size() == 38
			&& uninstallSubKeySubkey.name.substr(0,1) == L"{"
			&& isAlphanumericString(uninstallSubKeySubkey.name.substr(1,8))
			&& uninstallSubKeySubkey.name.substr(9,1) == L"-"
			&& isAlphanumericString(uninstallSubKeySubkey.name.substr(10,4))
			&& uninstallSubKeySubkey.name.substr(14,1) == L"-"
			&& isAlphanumericString(uninstallSubKeySubkey.name.substr(15,4))
			&& uninstallSubKeySubkey.name.substr(19,1) == L"-"
			&& isAlphanumericString(uninstallSubKeySubkey.name.substr(20,4))
			&& uninstallSubKeySubkey.name.substr(24,1) == L"-"
			&& isAlphanumericString(uninstallSubKeySubkey.name.substr(25,12))
			&& uninstallSubKeySubkey.name.substr(37,1) == L"}"
			){
				//_tprintf(TEXT("uninstallSubKeySubkey.name:%s\n"), uninstallSubKeySubkey.name.c_str());
				wstring transformedUninstallSubKeySubkeyName;
				transformedUninstallSubKeySubkeyName=
					uninstallSubKeySubkey.name.substr(8,1)+
					uninstallSubKeySubkey.name.substr(7,1)+
					uninstallSubKeySubkey.name.substr(6,1)+
					uninstallSubKeySubkey.name.substr(5,1)+
					uninstallSubKeySubkey.name.substr(4,1)+
					uninstallSubKeySubkey.name.substr(3,1)+
					uninstallSubKeySubkey.name.substr(2,1)+
					uninstallSubKeySubkey.name.substr(1,1)+
					uninstallSubKeySubkey.name.substr(13,1)+
					uninstallSubKeySubkey.name.substr(12,1)+
					uninstallSubKeySubkey.name.substr(11,1)+
					uninstallSubKeySubkey.name.substr(10,1)+
					uninstallSubKeySubkey.name.substr(18,1)+
					uninstallSubKeySubkey.name.substr(17,1)+
					uninstallSubKeySubkey.name.substr(16,1)+
					uninstallSubKeySubkey.name.substr(15,1)+
					uninstallSubKeySubkey.name.substr(21,1)+
					uninstallSubKeySubkey.name.substr(20,1)+
					uninstallSubKeySubkey.name.substr(23,1)+
					uninstallSubKeySubkey.name.substr(22,1)+
					uninstallSubKeySubkey.name.substr(26,1)+
					uninstallSubKeySubkey.name.substr(25,1)+
					uninstallSubKeySubkey.name.substr(28,1)+
					uninstallSubKeySubkey.name.substr(27,1)+
					uninstallSubKeySubkey.name.substr(30,1)+
					uninstallSubKeySubkey.name.substr(29,1)+
					uninstallSubKeySubkey.name.substr(32,1)+
					uninstallSubKeySubkey.name.substr(31,1)+
					uninstallSubKeySubkey.name.substr(34,1)+
					uninstallSubKeySubkey.name.substr(33,1)+
					uninstallSubKeySubkey.name.substr(36,1)+
					uninstallSubKeySubkey.name.substr(35,1);
				wstring productsHKeyName;
				wstring productsSubKeyName;
				RegList productsRegList;
				//HKCU\Software\Microsoft\Installer\Products\�z���Ƀ��W�X�g���L�[������ꍇ
				//�i�L�[�����݂��邩��EnumKey�ŕԂ�l��2�łȂ����Ƃ�����΂悢�j
				//�i����Ǝ���HKCR\Installer\Products\�̗���������ꍇ�͂��ꂪ�D�悷��j
				productsHKeyName = L"HKEY_CURRENT_USER";
				productsSubKeyName = L"Software\\Microsoft\\Installer\\Products\\" + transformedUninstallSubKeySubkeyName;
				status = productsRegList.setHKeyAndSubKey(productsHKeyName, productsSubKeyName, scan64bit);
				if(status == 0){
/*
					//���������؁���
					//Publisher�Ƃ������W�X�g���l������A���^��REG_SZ�܂���REG_EXPAND_SZ�܂���REG_MULTI_SZ�ł��̒l��""�ł͂Ȃ��ꍇ�A��荞��
					//REG_DWORD��OK? QWORD��NG
					regValue PublisherValue = productsRegList.getValue(L"Publisher");
					if(PublisherValue.name != L""){
						if(PublisherValue.type == L"REG_SZ" && PublisherValue.data != L""){
							//_tprintf(TEXT("ProductName:%s\n"), ProductNameValue.data.c_str());
							currentsoftwareInfo.vendor = PublisherValue.data;
							continue;
						}else if(PublisherValue.type == L"REG_EXPAND_SZ" && PublisherValue.data != L""){
							//_tprintf(TEXT("ProductName:%s\n"), ProductNameValue.data.c_str());
							currentsoftwareInfo.vendor = PublisherValue.data;
							continue;
						}else if(PublisherValue.type == L"REG_MULTI_SZ" && PublisherValue.data != L""){
							//TODO REG_MULTI_SZ�̏ꍇ�A�v���O�����̒ǉ��ƍ폜�ł�1�s�ڂ̂ݕ\������邪������
							//_tprintf(TEXT("ProductName:%s\n"), ProductNameValue.data.c_str());
							currentsoftwareInfo.vendor = PublisherValue.data;
							continue;
						}else if(PublisherValue.type == L"REG_DWORD" && PublisherValue.data != L""){
							//TODO REG_DWORD�̏����i0�̓_���H�j���m�F�A�W�L0x1��1�ϊ�������
							//_tprintf(TEXT("ProductName:%s\n"), ProductNameValue.data.c_str());
							currentsoftwareInfo.vendor = PublisherValue.data;
							continue;
						//TODO REG_QWORD���m�F
						}
					}
*/
					//ProductName�Ƃ������W�X�g���l������A���^��REG_SZ�܂���REG_EXPAND_SZ�܂���REG_MULTI_SZ�ł��̒l��""�ł͂Ȃ��ꍇ�A��荞��
					//REG_DWORD��OK�������BQWORD��NG
					regValue ProductNameValue = productsRegList.getValue(L"ProductName");
					if(ProductNameValue.name != L""){
						wstring productNameValueStr = L"";
						if(ProductNameValue.type == L"REG_SZ" && ProductNameValue.data != L""){
							productNameValueStr = ProductNameValue.data;
						}else if(ProductNameValue.type == L"REG_EXPAND_SZ" && ProductNameValue.data != L""){
							productNameValueStr = ProductNameValue.data;
						}else if(ProductNameValue.type == L"REG_MULTI_SZ" && ProductNameValue.data != L""){
							//TODO REG_MULTI_SZ�̏ꍇ�A�v���O�����̒ǉ��ƍ폜�ł�1�s�ڂ̂ݕ\������邪������
							productNameValueStr = ProductNameValue.data;
						}else if(ProductNameValue.type == L"REG_DWORD" && ProductNameValue.data != L""){
							//TODO REG_DWORD�̏����i0�̓_���H�j���m�F�A�W�L0x1��1�ϊ�������
							productNameValueStr = ProductNameValue.data;
						//TODO REG_QWORD���m�F
						}
						if(productNameValueStr != L""){
							if(debugLevel > 2){
								PrintDebugLog(L"ProductNameValue.data != \"\"\n");
							}
							currentSoftwareInfo.Name = productNameValueStr;
							currentSoftwareInfos.push_back(currentSoftwareInfo);
							currentPatchInfo.ParentName = productNameValueStr;

							wstring productsSubKeyPatchKeyName;
							productsSubKeyPatchKeyName = productsSubKeyName + L"\\Patches";
							RegList productsPatchRegList;
							//HKCU\Software\Microsoft\Installer\Products\transformedUninstallSubKeySubkeyName\Patches�z���Ƀ��W�X�g���L�[������ꍇ
							status = productsPatchRegList.setHKeyAndSubKey(productsHKeyName, productsSubKeyPatchKeyName, scan64bit);
							if(status == 0){
								//Patches�Ƃ������W�X�g���l������A���^��REG_MULTI_SZ�iREG_SZ�AREG_EXPAND_SZ�s�j�ł��̒l��""�ł͂Ȃ��ꍇ�A��荞��
								regValue PatchesNameValue = productsPatchRegList.getValue(L"Patches");
								if(PatchesNameValue.name != L""){
									if(PatchesNameValue.type == L"REG_MULTI_SZ" && PatchesNameValue.data != L""){
										//1�s���ƂɎ��o��
										//https://stackoverflow.com/questions/36812132/splitting-stdwstring-into-stdvector
										wstringstream patchesNameStream(PatchesNameValue.data);
										wstring patchesNameValueLine;
										//MessageBox(NULL,PatchesNameValue.data.c_str(),L"test",MB_OK);
										while(getline(patchesNameStream, patchesNameValueLine)){
											//MessageBox(NULL,patchesNameValueLine.c_str(),L"test",MB_OK);
											wstring s1518PatchSubKeyKeyname;
											s1518PatchSubKeyKeyname = L"Software\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\" + transformedUninstallSubKeySubkeyName + L"\\Patches\\" + patchesNameValueLine;
											RegList s1518PatchRegList;
											//���̂悤�ȃ��W�X�g���L�[�����邩
											//HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\68AB67CA7DA7FFFFB744AA0000000010\Patches\68AB67CA7DA7FFFF5205A7C804101061
											status = s1518PatchRegList.setHKeyAndSubKey(L"HKEY_CURRENT_USER", s1518PatchSubKeyKeyname, scan64bit);
											if(status == 0){
												regValue DisplayNameValue = s1518PatchRegList.getValue(L"DisplayName");
												wstring displayNameValueStr = L"";
												//DisplayName�Ƃ������W�X�g���l������A���^��REG_SZ�ł��̒l��""�łȂ��ꍇ�A�p�b�`���X�g�ɒǉ�����
												//DisplayName�Ƃ������W�X�g���l������A���^��REG_DWORD�Ȃ炻�̒l�i10�i���j���p�b�`���X�g�ɒǉ�����
												if(DisplayNameValue.name != L""){
													if(DisplayNameValue.type == L"REG_SZ" && DisplayNameValue.data != L""){
														displayNameValueStr = DisplayNameValue.data;
													}else if(DisplayNameValue.type == L"REG_DWORD"){
														displayNameValueStr = DisplayNameValue.data;
													}
												}
												if(displayNameValueStr != L""){
													currentPatchInfo.Name = displayNameValueStr;
													currentPatchInfos.push_back(currentPatchInfo);
													if(debugLevel > 2){
														PrintDebugLog(L"Patch:DisplayNameValue.data ==");
														PrintDebugLog(displayNameValueStr);
														PrintDebugLog(L"\n");
													}
												//DisplayName�Ƃ������W�X�g���l������A�^��REG_SZ�����l��""�̏ꍇ
												//DisplayName�Ƃ������W�X�g���l�����邪�^��REG_SZ�ł�REG_DWORD�ł��Ȃ��ꍇ
												//DisplayName�Ƃ������W�X�g���l���Ȃ��ꍇ
												//��������v���O�����Ƌ@�\�ł͂����u�X�V�v���O�����v�ƕ\�������
												}else{
													currentPatchInfo.Name = L"Update Program";
													currentPatchInfos.push_back(currentPatchInfo);
													if(debugLevel > 2){
														PrintDebugLog(L"Patch:DisplayNameValue.data is not match.\n");
													}
												}
											}
										}
									}
								}
							}
							continue;
						}
					}
				}
				//HKCR\Installer\Products\�z���Ƀ��W�X�g���L�[������ꍇ
				productsHKeyName = L"HKEY_CLASSES_ROOT";
				productsSubKeyName = L"Installer\\Products\\" + transformedUninstallSubKeySubkeyName;
				//_tprintf(TEXT("transformedUninstallSubKeySubkeyName:%s\n"), transformedUninstallSubKeySubkeyName.c_str());
				status = productsRegList.setHKeyAndSubKey(productsHKeyName, productsSubKeyName, scan64bit);
				if(status == 0){
					//ProductName�Ƃ������W�X�g���l������A���^��REG_SZ�܂���REG_EXPAND_SZ�܂���REG_MULTI_SZ�ł��̒l��""�ł͂Ȃ��ꍇ�A��荞��
					//REG_DWORD��OK�������BQWORD��NG
					//�iWindows 7�Ŏ������Ƃ���AREG_EXPAND_SZ��GetStringValue�ł��擾�ł��A
					//�@�܂�REG_SZ��GetExpandedStringValue�Ŏ擾�ł��Ă��܂��̂�ElseIf reg.GetExpandedStringValue...�͈Ӗ����Ȃ������j
					//�iREG_MULTI_SZ�̏ꍇ�A�v���O�����̒ǉ��ƍ폜�ł�1�s�ڂ̂ݕ\�������j����������������
					regValue ProductNameValue = productsRegList.getValue(L"ProductName");
					if(ProductNameValue.name != L""){
						wstring productNameValueStr = L"";
						if(ProductNameValue.type == L"REG_SZ" && ProductNameValue.data != L""){
							productNameValueStr = ProductNameValue.data;
						}else if(ProductNameValue.type == L"REG_EXPAND_SZ" && ProductNameValue.data != L""){
							productNameValueStr = ProductNameValue.data;
						}else if(ProductNameValue.type == L"REG_MULTI_SZ" && ProductNameValue.data != L""){
							//TODO REG_MULTI_SZ�̏ꍇ�A�v���O�����̒ǉ��ƍ폜�ł�1�s�ڂ̂ݕ\������邪������
							productNameValueStr = ProductNameValue.data;
						}else if(ProductNameValue.type == L"REG_DWORD" && ProductNameValue.data != L""){
							//TODO REG_DWORD�̏����i0�̓_���H�j���m�F�A�W�L0x1��1�ϊ�������
							productNameValueStr = ProductNameValue.data;
						//TODO REG_QWORD���m�F
						}
						if(productNameValueStr != L""){
							if(debugLevel > 2){
								PrintDebugLog(L"ProductNameValue.data != \"\"\n");
							}
							currentSoftwareInfo.Name = productNameValueStr;
							currentSoftwareInfos.push_back(currentSoftwareInfo);
							currentPatchInfo.ParentName = productNameValueStr;

							wstring productsSubKeyPatchKeyName;
							productsSubKeyPatchKeyName = productsSubKeyName + L"\\Patches";
							RegList productsPatchRegList;
							//HKCU\Software\Microsoft\Installer\Products\transformedUninstallSubKeySubkeyName\Patches�z���Ƀ��W�X�g���L�[������ꍇ
							status = productsPatchRegList.setHKeyAndSubKey(productsHKeyName, productsSubKeyPatchKeyName, scan64bit);
							if(status == 0){
								//Patches�Ƃ������W�X�g���l������A���^��REG_MULTI_SZ�iREG_SZ�AREG_EXPAND_SZ�s�j�ł��̒l��""�ł͂Ȃ��ꍇ�A��荞��
								regValue PatchesNameValue = productsPatchRegList.getValue(L"Patches");
								if(PatchesNameValue.name != L""){
									if(PatchesNameValue.type == L"REG_MULTI_SZ" && PatchesNameValue.data != L""){
										//1�s���ƂɎ��o��
										//https://stackoverflow.com/questions/36812132/splitting-stdwstring-into-stdvector
										wstringstream patchesNameStream(PatchesNameValue.data);
										wstring patchesNameValueLine;
										//MessageBox(NULL,PatchesNameValue.data.c_str(),L"test",MB_OK);
										while(getline(patchesNameStream, patchesNameValueLine)){
											//MessageBox(NULL,patchesNameValueLine.c_str(),L"test",MB_OK);
											wstring s1518PatchSubKeyKeyname;
											s1518PatchSubKeyKeyname = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\" + transformedUninstallSubKeySubkeyName + L"\\Patches\\" + patchesNameValueLine;
											RegList s1518PatchRegList;
											//���̂悤�ȃ��W�X�g���L�[�����邩
											//HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\68AB67CA7DA7FFFFB744AA0000000010\Patches\68AB67CA7DA7FFFF5205A7C804101061
											//64bitOS�Ȃ�64bit�L�[�𒲂ׂ�@32bit�L�[�ɃA�N�Z�X���悤�Ƃ���ƕԂ�l2�ƂȂ�

											status = s1518PatchRegList.setHKeyAndSubKey(L"HKEY_LOCAL_MACHINE", s1518PatchSubKeyKeyname, IsWow64());
											//�e�X�g
											if(debugLevel > 2 && transformedUninstallSubKeySubkeyName == L"68AB67CA7DA7FFFFB744AA0000000010"){
												PrintDebugLog(L"PatchDebug: ");
												PrintDebugLog(s1518PatchSubKeyKeyname);
												PrintDebugLog(L"\n");
												s1518PatchRegList.setHKeyAndSubKey(L"HKEY_LOCAL_MACHINE", L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\68AB67CA7DA7FFFFB744AA0000000010\\Patches\\68AB67CA7DA7FFFF5205A7C804101061", true);
												s1518PatchRegList.setHKeyAndSubKey(L"HKEY_LOCAL_MACHINE", L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\68AB67CA7DA7FFFFB744AA0000000010\\Patches", true);
												s1518PatchRegList.setHKeyAndSubKey(L"HKEY_LOCAL_MACHINE", L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\68AB67CA7DA7FFFFB744AA0000000010", true);
												s1518PatchRegList.setHKeyAndSubKey(L"HKEY_LOCAL_MACHINE", L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products", true);
												s1518PatchRegList.setHKeyAndSubKey(L"HKEY_LOCAL_MACHINE", L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18", true);
												s1518PatchRegList.setHKeyAndSubKey(L"HKEY_LOCAL_MACHINE", L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData", true);
												s1518PatchRegList.setHKeyAndSubKey(L"HKEY_LOCAL_MACHINE", L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer", true);
												s1518PatchRegList.setHKeyAndSubKey(L"HKEY_LOCAL_MACHINE", L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\68AB67CA7DA7FFFFB744AA0000000010\\Patches\\68AB67CA7DA7FFFF5205A7C804101061", false);
												s1518PatchRegList.setHKeyAndSubKey(L"HKEY_LOCAL_MACHINE", L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\68AB67CA7DA7FFFFB744AA0000000010\\Patches", false);
												s1518PatchRegList.setHKeyAndSubKey(L"HKEY_LOCAL_MACHINE", L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\68AB67CA7DA7FFFFB744AA0000000010", false);
												s1518PatchRegList.setHKeyAndSubKey(L"HKEY_LOCAL_MACHINE", L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products", false);
												s1518PatchRegList.setHKeyAndSubKey(L"HKEY_LOCAL_MACHINE", L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18", false);
												s1518PatchRegList.setHKeyAndSubKey(L"HKEY_LOCAL_MACHINE", L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData", false);
												s1518PatchRegList.setHKeyAndSubKey(L"HKEY_LOCAL_MACHINE", L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer", false);
											}

											if(status == 0){
												regValue DisplayNameValue = s1518PatchRegList.getValue(L"DisplayName");
												wstring displayNameValueStr = L"";
												//DisplayName�Ƃ������W�X�g���l������A���^��REG_SZ�ł��̒l��""�łȂ��ꍇ�A�p�b�`���X�g�ɒǉ�����
												//DisplayName�Ƃ������W�X�g���l������A���^��REG_DWORD�Ȃ炻�̒l�i10�i���j���p�b�`���X�g�ɒǉ�����
												if(DisplayNameValue.name != L""){
													if(DisplayNameValue.type == L"REG_SZ" && DisplayNameValue.data != L""){
														displayNameValueStr = DisplayNameValue.data;
													}else if(DisplayNameValue.type == L"REG_DWORD"){
														displayNameValueStr = DisplayNameValue.data;
													}
												}
												if(displayNameValueStr != L""){
													currentPatchInfo.Name = displayNameValueStr;
													currentPatchInfos.push_back(currentPatchInfo);
													if(debugLevel > 2){
														PrintDebugLog(L"Patch:DisplayNameValue.data ==");
														PrintDebugLog(displayNameValueStr);
														PrintDebugLog(L"\n");
													}
												//DisplayName�Ƃ������W�X�g���l������A�^��REG_SZ�����l��""�̏ꍇ
												//DisplayName�Ƃ������W�X�g���l�����邪�^��REG_SZ�ł�REG_DWORD�ł��Ȃ��ꍇ
												//DisplayName�Ƃ������W�X�g���l���Ȃ��ꍇ
												//��������v���O�����Ƌ@�\�ł͂����u�X�V�v���O�����v�ƕ\�������
												}else{
													currentPatchInfo.Name = L"Update Program";
													currentPatchInfos.push_back(currentPatchInfo);
													if(debugLevel > 2){
														PrintDebugLog(L"Patch:DisplayNameValue.data is not match.\n");
													}
												}
											}
										}
									}
								}
							}
							continue;
						}
					}
				}
			}
			//continue;
		//WindowsInstaller�Ƃ������W�X�g���l���Ȃ����A�����Ă��^��REG_DWORD�ł��̒l��1�ȊO�̏ꍇ
		}else{
			if(debugLevel > 2){
				PrintDebugLog(L"hold. WindowsInstallerValue.data != \"1\"\n");
			}
			//UninstallString�Ƃ������W�X�g���l������A���^��REG_SZ�܂���REG_EXPAND_SZ�ł��̒l��""�ł͂Ȃ��ꍇ�A�������p��
			//�i������REG_MULTI_SZ��NG�j
			regValue UninstallStringValue = uninstallSubkeyRegList.getValue(L"UninstallString");
			//if(uninstallSubKeySubkey.name == L"{87434D51-51DB-4109-B68F-A829ECDCF380}"){
			//	_tprintf(TEXT("test:%s\n"), UninstallStringValue.data.c_str());
			//}
			if(UninstallStringValue.name != L""){
				if(UninstallStringValue.type == L"REG_SZ" && UninstallStringValue.data != L""){
					if(debugLevel > 2){
						PrintDebugLog(L"hold. UninstallStringValue.data != \"\"\n");
					}
					//�������p��
				}else if(UninstallStringValue.type == L"REG_EXPAND_SZ" && UninstallStringValue.data != L""){
					if(debugLevel > 2){
						PrintDebugLog(L"hold. UninstallStringValue.data != \"\"\n");
					}
					//�������p��
				}else{
					if(debugLevel > 2){
						PrintDebugLog(L"continue. UninstallStringValue.data is empty or == \"\"\n");
					}
					//��L�ȊO�͏������I��
					continue;
				}
			}else{
				if(debugLevel > 2){
					PrintDebugLog(L"continue. UninstallStringValue not found.\n");
				}
				//��L�ȊO�͏������I��
				continue;
			}
			//ParentKeyName�Ƃ������W�X�g���l������A���^��REG_DWORD�ł��̒l��0�łȂ��ꍇ�A�������I��
			//ParentKeyName�Ƃ������W�X�g���l������A���^��REG_SZ�܂���REG_EXPAND_SZ�܂���REG_MULTI_SZ�ł��̒l��""�łȂ��ꍇ�A�������I��
			//ParentKeyName�Ƃ������W�X�g���l������A���^��REG_BINARY�ŋ�łȂ��ꍇ�A�������I��
			//Vista�ȍ~��OS�ɂ����āAParentKeyName�Ƃ������W�X�g���l������A���^��REG_QWORD�ł��̒l��0�łȂ��ꍇ�A
			//�\�t�g�E�F�A���X�g�͏������I���A�p�b�`���X�g�͋t�ɏ������p��
			regValue ParentKeyNameValue = uninstallSubkeyRegList.getValue(L"ParentKeyName");
			bool parentKeyNameValueIsNotEmpty = false;
			if(ParentKeyNameValue.name != L""){
				if(ParentKeyNameValue.type == L"REG_DWORD" && ParentKeyNameValue.data != L"0"){
					parentKeyNameValueIsNotEmpty = true;
				}else if(ParentKeyNameValue.type == L"REG_SZ" && ParentKeyNameValue.data != L""){
					parentKeyNameValueIsNotEmpty = true;
				}else if(ParentKeyNameValue.type == L"REG_EXPAND_SZ" && ParentKeyNameValue.data != L""){
					parentKeyNameValueIsNotEmpty = true;
				}else if(ParentKeyNameValue.type == L"REG_MULTI_SZ" && ParentKeyNameValue.data != L""){
					parentKeyNameValueIsNotEmpty = true;
				}else if(ParentKeyNameValue.type == L"REG_BINARY" && ParentKeyNameValue.data != L""){
					parentKeyNameValueIsNotEmpty = true;
				}else if(ParentKeyNameValue.type == L"REG_QWORD" && ParentKeyNameValue.data != L"0"){
					parentKeyNameValueIsNotEmpty = true;
				}
				if(parentKeyNameValueIsNotEmpty){
					regValue ParentDisplayNameValue = uninstallSubkeyRegList.getValue(L"ParentDisplayName");
					//ParentDisplayName�Ƃ������W�X�g���l������A���^��REG_SZ�܂���REG_EXPAND_SZ�ł��̒l��""�łȂ��ꍇ�A�e�v���O�������Ƃ��ăp�b�`���X�g�ɒǉ�����
					if(ParentDisplayNameValue.name != L""){
						if(ParentDisplayNameValue.type == L"REG_SZ" && ParentDisplayNameValue.data != L""){
							currentPatchInfo.ParentName = ParentDisplayNameValue.data;
						}else if(ParentDisplayNameValue.type == L"REG_EXPAND_SZ" && ParentDisplayNameValue.data != L""){
							currentPatchInfo.ParentName = ParentDisplayNameValue.data;
						}
					}
					regValue DisplayNameValue = uninstallSubkeyRegList.getValue(L"DisplayName");
					//DisplayName�Ƃ������W�X�g���l������A���^��REG_SZ�܂���REG_EXPAND_SZ�ł��̒l��""�łȂ��ꍇ�A�p�b�`���X�g�ɒǉ�����
					if(DisplayNameValue.name != L""){
						if(DisplayNameValue.type == L"REG_SZ" && DisplayNameValue.data != L""){
							currentPatchInfo.Name = DisplayNameValue.data;
							currentPatchInfos.push_back(currentPatchInfo);
						}else if(DisplayNameValue.type == L"REG_EXPAND_SZ" && DisplayNameValue.data != L""){
							currentPatchInfo.Name = DisplayNameValue.data;
							currentPatchInfos.push_back(currentPatchInfo);
						}
					}
					if(debugLevel > 2){
						PrintDebugLog(L"continue. ParentKeyNameValue.data != \"\"\n");
					}
					continue;
				}
			}
			//�w�肳�ꂽ���W�X�g���L�[�z���̃T�u�L�[���T�|�[�g�Z�p���ԍ��iKB000000�j���̂��̂ł���ꍇ�A��{�I�ɂ͏������I�����邪
			//ParentKeyName�����݂��A0�܂���""�ł���ꍇ�̂ݏ������p��
			//�iWindows Search 4.0�Ŕ��o�j
			if((uninstallSubKeySubkey.name.size() == 8)
			&& (uninstallSubKeySubkey.name.substr(0,2) == L"KB")
			&& isNumericString(uninstallSubKeySubkey.name.substr(2,6))
			){
				//_tprintf(TEXT("test: %s\n"), uninstallSubKeySubkey.name.c_str());
				if(debugLevel > 2){
					PrintDebugLog(L"hold. uninstallSubKeySubkey.name is KB.\n");
				}
				//ParentKeyName�Ƃ������W�X�g���l������A���^��REG_DWORD�ł��̒l��0�̏ꍇ�A�������p��
				//ParentKeyName�Ƃ������W�X�g���l������A���^��REG_SZ�܂���REG_EXPAND_SZ�܂���REG_MULTI_SZ�ł��̒l��""�̏ꍇ�A�������p��
				//ParentKeyName�Ƃ������W�X�g���l������A���^��REG_BINARY�ŋ�̏ꍇ�A�������p��
				//Vista�ȍ~��OS�ɂ����āAParentKeyName�Ƃ������W�X�g���l������A���^��REG_QWORD�ł��̒l��0�̏ꍇ�A�������p��
				if(ParentKeyNameValue.name != L""){
					if(ParentKeyNameValue.type == L"REG_DWORD" && ParentKeyNameValue.data != L"0"){
						if(debugLevel > 2){
							PrintDebugLog(L"hold. ParentKeyNameValue.data !=\"\"\n");
						}
						//�������p��
					}else if(ParentKeyNameValue.type == L"REG_SZ" && ParentKeyNameValue.data != L""){
						if(debugLevel > 2){
							PrintDebugLog(L"hold. ParentKeyNameValue.data !=\"\"\n");
						}
						//�������p��
					}else if(ParentKeyNameValue.type == L"REG_EXPAND_SZ" && ParentKeyNameValue.data != L""){
						if(debugLevel > 2){
							PrintDebugLog(L"hold. ParentKeyNameValue.data !=\"\"\n");
						}
						//�������p��
					}else if(ParentKeyNameValue.type == L"REG_MULTI_SZ" && ParentKeyNameValue.data != L""){
						if(debugLevel > 2){
							PrintDebugLog(L"hold. ParentKeyNameValue.data !=\"\"\n");
						}
						//�������p��
					}else if(ParentKeyNameValue.type == L"REG_BINARY" && ParentKeyNameValue.data != L""){
						if(debugLevel > 2){
							PrintDebugLog(L"hold. ParentKeyNameValue.data !=\"\"\n");
						}
						//�������p��
					}else if(ParentKeyNameValue.type == L"REG_QWORD" && ParentKeyNameValue.data != L"0"){
						if(debugLevel > 2){
							PrintDebugLog(L"hold. ParentKeyNameValue.data !=\"\"\n");
						}
						//�������p��
					}else{
						if(debugLevel > 2){
							PrintDebugLog(L"continue");
						}
						//��L�ȊO�͏������I��
						continue;
					}
				}
			}
			//DisplayName_Localized�Ƃ������W�X�g���l������Ƃ��́ALocalizedString�����Ă�����擾����
			//��������������REG_SZ�ȊO�̊m�F�A�󕶎��񎞂̋����AProductName��Localized�Ή��L��
			//---------------------------
			regValue DisplayNameLocalizedValue = uninstallSubkeyRegList.getValue(L"DisplayName_Localized");
			if (DisplayNameLocalizedValue.name != L"") {
				if (DisplayNameLocalizedValue.type == L"REG_SZ" && DisplayNameLocalizedValue.data != L"")
				{
					try
					{
						wstring displayNameLocalizedValueData = getLocalizedString(DisplayNameLocalizedValue.data);
						//�����ŃG���[�����������ꍇ�͂��̉���if���ɓ���Ȃ��悤�ɂ��Acontinue���Ȃ��悤�ɂ���B�i2022/01/13�j
						if (displayNameLocalizedValueData != L"") {
							//_tprintf(TEXT("DisplayName_Localized:%s\n"), displayNameLocalizedValueData.c_str());
							currentSoftwareInfo.Name = displayNameLocalizedValueData;
							currentSoftwareInfos.push_back(currentSoftwareInfo);
							if (debugLevel > 2) {
								PrintDebugLog(L"continue. displayNameLocalizedValueData == ");
								PrintDebugLog(displayNameLocalizedValueData);
								PrintDebugLog(L"\n");
							}
							continue;
						}
					}
					catch (const std::exception&)
					{
						if (debugLevel > 2) {
							PrintDebugLog(L"Exception from get localize");
							PrintDebugLog(L"\n");
						}
					}
				}
			}
			//---------------------------
/*
			//���������؁���
			//Publisher�Ƃ������W�X�g���l������A���^��REG_SZ�܂���REG_EXPAND_SZ�ł��̒l��""�łȂ��ꍇ�A���X�g�ɒǉ�����
			//�i������REG_MULTI_SZ��NG�j
			regValue PublisherValue = uninstallSubkeyRegList.getValue(L"Publisher");
			if(PublisherValue.name != L""){
				if(PublisherValue.type == L"REG_SZ" && PublisherValue.data != L""){
					//_tprintf(TEXT("DisplayName:%s\n"), DisplayNameValue.data.c_str());
					currentsoftwareInfo.vendor = PublisherValue.data;
				}else if(PublisherValue.type == L"REG_EXPAND_SZ" && PublisherValue.data != L""){
					//_tprintf(TEXT("DisplayName:%s\n"), DisplayNameValue.data.c_str());
					currentsoftwareInfo.vendor = PublisherValue.data;
				}
			}
*/
			//DisplayName�Ƃ������W�X�g���l������A���^��REG_SZ�܂���REG_EXPAND_SZ�ł��̒l��""�łȂ��ꍇ�A���X�g�ɒǉ�����
			//�i������REG_MULTI_SZ��NG�j
			regValue DisplayNameValue = uninstallSubkeyRegList.getValue(L"DisplayName");
			if(DisplayNameValue.name != L""){
				if(DisplayNameValue.type == L"REG_SZ" && DisplayNameValue.data != L""){
					//_tprintf(TEXT("DisplayName:%s\n"), DisplayNameValue.data.c_str());
					currentSoftwareInfo.Name = DisplayNameValue.data;
					currentSoftwareInfos.push_back(currentSoftwareInfo);
					if(debugLevel > 2){
						PrintDebugLog(L"continue. DisplayNameValue.data ==");
						PrintDebugLog(DisplayNameValue.data);
						PrintDebugLog(L"\n");
					}
				}else if(DisplayNameValue.type == L"REG_EXPAND_SZ" && DisplayNameValue.data != L""){
					//_tprintf(TEXT("DisplayName:%s\n"), DisplayNameValue.data.c_str());
					currentSoftwareInfo.Name = DisplayNameValue.data;
					currentSoftwareInfos.push_back(currentSoftwareInfo);
					if(debugLevel > 2){
						PrintDebugLog(L"continue. DisplayNameValue.data == ");
						PrintDebugLog(DisplayNameValue.data);
						PrintDebugLog(L"\n");
					}
				}
			}
		}
		//_tprintf(TEXT("hkey:%s, subKey:%s\n"), uninstallHKeyName.c_str(), uninstallSubKeySubkey.fullPathName.c_str());
		//_tprintf(TEXT("hkey:%s, subKey:%s\n"), uninstallHKeyName, uninstallSubKeySubkeyFullPathName);
		//delete uninstallSubkeyRegList;

	}


    return 0;
}

//WMI�Ńp�b�`�����擾����֐��i�{�́j
long InventoryManager::getPatchFromWMI(){

	HRESULT hres;

	// Initialize COM. ------------------------------------------
    hres =  CoInitializeEx(0, COINIT_MULTITHREADED); 
    if (FAILED(hres)){
		if(debugLevel > 1){
			PrintDebugLog(L"Error:Failed to initialize COM library. Error code :");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%d"), hres);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
        return -11;               // Program has failed.
    }

    // Set general COM security levels --------------------------
    hres =  CoInitializeSecurity(
        NULL, 
        -1,                          // COM authentication
        NULL,                        // Authentication services
        NULL,                        // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
        RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
        NULL,                        // Authentication info
        EOAC_NONE,                   // Additional capabilities 
        NULL                         // Reserved
        );

	//RPC_E_TOO_LATE -2147417831
	//http://forums.codeguru.com/showthread.php?404520-Can-RPC_E_TOO_LATE-be-ignored
    if (hres != -2147417831 && FAILED(hres)){
		if(debugLevel > 1){
			PrintDebugLog(L"Error:Failed to initialize security. Error code :");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%d"), hres);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
        CoUninitialize();
        return -12;               // Program has failed.
    }
    
    // Obtain the initial locator to WMI -------------------------
    IWbemLocator *pLoc = NULL;

    hres = CoCreateInstance(
        MY_CLSID_WbemLocator,             
        0, 
        CLSCTX_INPROC_SERVER, 
        MY_IID_IWbemLocator, (LPVOID *) &pLoc);
 
    if (FAILED(hres)){
		if(debugLevel > 1){
			PrintDebugLog(L"Error:Failed to create IWbemLocator object. Error code :");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%d"), hres);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
        CoUninitialize();
        return -13;               // Program has failed.
    }

    // Connect to WMI through the IWbemLocator::ConnectServer method
    IWbemServices *pSvc = NULL;
    hres = pLoc->ConnectServer(
         _bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
         NULL,                    // User name. NULL = current user
         NULL,                    // User password. NULL = current
         0,                       // Locale. NULL indicates current
         NULL,                    // Security flags.
         0,                       // Authority (for example, Kerberos)
         0,                       // Context object 
         &pSvc                    // pointer to IWbemServices proxy
         );
    
    if (FAILED(hres)){
		if(debugLevel > 1){
			PrintDebugLog(L"Error:Could not connect. Error code :");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%d"), hres);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
        pLoc->Release();     
        CoUninitialize();
        return -14;               // Program has failed.
    }

    // Set security levels on the proxy -------------------------
    hres = CoSetProxyBlanket(
       pSvc,                        // Indicates the proxy to set
       RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
       RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
       NULL,                        // Server principal name 
       RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
       RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
       NULL,                        // client identity
       EOAC_NONE                    // proxy capabilities 
    );

    if (FAILED(hres)){
		if(debugLevel > 1){
			PrintDebugLog(L"Error:Could not set proxy blanket. Error code :");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%d"), hres);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
        pSvc->Release();
        pLoc->Release();     
        CoUninitialize();
        return -15;               // Program has failed.
    }

    // Use the IWbemServices pointer to make requests of WMI ----
    IEnumWbemClassObject* pEnumerator = NULL;
	hres = pSvc->ExecQuery(
        bstr_t("WQL"), 
        bstr_t("SELECT * FROM Win32_QuickFixEngineering"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, 
        NULL,
        &pEnumerator);
    
    if (FAILED(hres)){
		if(debugLevel > 1){
			PrintDebugLog(L"Error:Query for operating system name failed. Error code :");
			wchar_t stringBuffer[32];
			_stprintf(stringBuffer, TEXT("%d"), hres);
			PrintDebugLog(stringBuffer);
			PrintDebugLog(L"\n");
		}
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return -16;               // Program has failed.
    }

    // Get the data from the query in step 6 -------------------
    IWbemClassObject *pclsObj;
    ULONG uReturn = 0;

    while (pEnumerator){
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if(0 == uReturn){
            break;
        }
        VARIANT vtProp;

		//�l�̎擾
		//hr = pclsObj->Get(BSTR(L"Description"), 0, &vtProp, 0, 0);
		//if(!FAILED(hr)){
		//	if(vtProp.vt==VT_NULL || vtProp.vt==VT_EMPTY){
		//		//�������Ȃ�
		//	}else{
		//		_bstr_t bstr = vtProp;
		//		hotfixStr = bstr;
		//	}
		//}

		wstring hotfixStr = L"";
		hr = pclsObj->Get(BSTR(L"HotFixID"), 0, &vtProp, 0, 0);
		if(!FAILED(hr)){
			if(vtProp.vt==VT_NULL || vtProp.vt==VT_EMPTY){
				//�������Ȃ�
			}else{
				_bstr_t bstr = vtProp;
				//wstring hotfixID = bstr;
				//hotfixStr += L" (" + hotfixID + L")";
				hotfixStr = bstr;

			}
		}else{
			if(debugLevel > 1){
				PrintDebugLog(L"Error:pclsObj->Get failed.\n");
			}
		}

		//�ȑO��WindowsOS�Ŗ��p�b�`�̏ꍇ�AHotFixID��File 1�ƂȂ�͗l
		//https://stackoverflow.com/questions/10291296/what-is-the-hotfix-with-hotfixid-file-1
		//https://web.archive.org/web/20121013153025/http://support.microsoft.com/kb/831415
		//���̏ꍇ��ServicePackInEffect�̎擾�����݂�
		if(hotfixStr == L"File 1"){
			hotfixStr = L"";
			hr = pclsObj->Get(BSTR(L"ServicePackInEffect"), 0, &vtProp, 0, 0);
			if(!FAILED(hr)){
				if(vtProp.vt==VT_NULL || vtProp.vt==VT_EMPTY){
					//�������Ȃ�
				}else{
					_bstr_t bstr = vtProp;
					hotfixStr = bstr;
				}
			}else{
				if(debugLevel > 1){
					PrintDebugLog(L"Error:pclsObj->Get failed.\n");
				}
			}
		}

		if(hotfixStr != L""){
			patchInfo currentPatchInfo;
			currentPatchInfo.Name = hotfixStr;
			currentPatchInfo.ParentName = L"Microsoft Windows";
			currentPatchInfos.push_back(currentPatchInfo);
		}

		VariantClear(&vtProp);
        pclsObj->Release();
    }

	// Cleanup
    pSvc->Release();
    pLoc->Release();
    pEnumerator->Release();
    if(!pclsObj) pclsObj->Release();
    CoUninitialize();

	return 0;
}

//���e��S�o�͂���֐�
wstring InventoryManager::output(){

	//���P�[���������ݒ肵�Ă���
	//��������Ă����Ȃ��ƁASJIS�ŕۑ������Ƃ��ɓ��{�ꂪ??�ɂȂ�
	//�Q�l http://cx5software.com/article_vcpp_unicode/
	//setlocale�֐��̑������Ɂu""�v���w�肷��ƁAOS�̃f�t�H���g���P�[�����w�肳��܂��B�Ⴆ�΁A���{��ɐݒ肳��Ă���ꍇ�uJapanese_Japan.932�v���ݒ肳��܂��B���R�Ȃ���AOS�̐ݒ肪�ς��ƁA����ɍ��킹�ĕύX����܂��̂Œ��ӂ��K�v�ł��B
	_tsetlocale(LC_ALL, _T(""));
	//_tsetlocale(LC_ALL, L"Japanese");

	//Unicode��������R���\�[���ɕ\��������
	//_tsetlocale(LC_ALL, _tsetlocale(LC_CTYPE, L""));

	//�t�@�C���o�͂̃R���X�g���N�^
	//http://www.geocities.jp/ky_webid/cpp/library/033.html
	//http://sato-si.at.webry.info/200703/article_1.html

	//�\�t�g�E�F�A���X�g�̉�ʏo��
	//GUI�ł̓R�����g�A�E�g
/*
	vector<softwareInfo>::iterator softwareInfoIterator = softwareInfos.begin();  // �C�e���[�^�̃C���X�^���X��
	while(softwareInfoIterator != softwareInfos.end()){
		_tprintf(TEXT("%s\n"), (*softwareInfoIterator).name.c_str());  // *���Z�q�ŊԐڎQ��
		++softwareInfoIterator;                 // �C�e���[�^���P�i�߂�
	}
*/

	//��{�I�ȃC���x���g�����̎擾�������Ŏ��̂���߂�
	/*
	InventoryManager testInventoryManager;
	inventory currentHardwareInfo;
	currentHardwareInfo = testInventoryManager.getHardwareInfo();
	*/
	//SetDlgItemText(IDC_EDIT1, testInventory.ComputerName.c_str());
	//SetDlgItemText(IDC_EDIT2, testInventory.UserName.c_str());
	//SetDlgItemText(IDC_EDIT2, testInventory.MACAddress.c_str());
	//SetDlgItemText(IDC_EDIT3, testInventory.IPAddress.c_str());

	wstring returnString = L"";

	//�\�t�g�E�F�A���X�g�̃t�@�C���o��
	//http://cx5software.com/article_vcpp_unicode/

	//�t�H�[�}�b�g���ʏ�`���i�V���v����csv�j��SARMS�`�����ŕ���

	//�t�H�[�}�b�g��SARMS�`���̏ꍇ
	//return outputFileName;
	if(outputFileFormat == L"SARMS"){
		returnString = outputSARMS();
	//�t�H�[�}�b�g���ʏ�`���i�V���v����csv�j�̏ꍇ
	}else if(outputFileFormat == L"SIMPLE"){
		returnString = outputSIMPLE();
	//�t�H�[�}�b�g��AdvancedManager�`���̏ꍇ
	}else if(outputFileFormat == L"AdvancedManager"){
		returnString = outputAdvancedManager();
	}
	return returnString;
}

//SARMS�`���œ��e��S�o�͂���֐�
wstring InventoryManager::outputSARMS(){

		wstring returnOutputFileName = L"";
	
		//�A�v���P�[�V�������̕ۑ�
		FILE* fp = NULL;
		wstring outputFileName = outputPath + L"Ap_" + currentHardwareInfo.HardwareNoValue + L".csv";

		//�G���R�[�h��SJIS�ŕۑ�����i�I�������Ȃ��j
		_wfopen_s( &fp, outputFileName.c_str(), L"w" );

		//�^�C�g���s�̏o�́��Ȃ�
		vector<softwareInfo>::iterator softwareInfoIterator2 = currentSoftwareInfos.begin();  // �C�e���[�^�̃C���X�^���X��
		while(softwareInfoIterator2 != currentSoftwareInfos.end()){
			wstring tempField;

			tempField = Replace(currentHardwareInfo.HardwareNoValue, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			tempField= (*softwareInfoIterator2).Name.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			//_ftprintf(fp, TEXT("Software - Name:%s\t\n"), (*softwareInfoIterator).Name.c_str());  // *���Z�q�ŊԐڎQ��
			_ftprintf(fp, L"\"%s\",", tempField.c_str());  // *���Z�q�ŊԐڎQ��
			//_ftprintf(fp, L"%s,", (*softwareInfoIterator2).Name.c_str());  // *���Z�q�ŊԐڎQ��
			
			if(debugLevel > 0){

				tempField = (*softwareInfoIterator2).InstallDate.c_str();
				tempField = Replace(tempField, L"\"", L"\"\"");
				_ftprintf(fp, L"\"%s\",", tempField.c_str());

				tempField = (*softwareInfoIterator2).ScanRegistryRange.c_str();
				tempField = Replace(tempField, L"\"", L"\"\"");
				_ftprintf(fp, L"\"%s\",", tempField.c_str());

				tempField = (*softwareInfoIterator2).CompareName.c_str();
				tempField = Replace(tempField, L"\"", L"\"\"");
				_ftprintf(fp, L"\"%s\",", tempField.c_str());

				tempField = (*softwareInfoIterator2).NameCodePoint.c_str();
				tempField = Replace(tempField, L"\"", L"\"\"");
				_ftprintf(fp, L"\"%s\",", tempField.c_str());
			}

			_ftprintf(fp, L"\n");
			++softwareInfoIterator2;                 // �C�e���[�^���P�i�߂�
		}
		//_ftprintf(fp, L"\n");
		//_ftprintf(fp, L"Output by LIP (List Installed Programs) v0.1.");
	
		fclose(fp);
		returnOutputFileName += outputFileName.c_str();
		returnOutputFileName += L", ";

		//���̑��̃C���x���g�����̕ۑ�
		fp = NULL;
		outputFileName = outputPath + L"Inv_" + currentHardwareInfo.HardwareNoValue + L".csv";

		//�G���R�[�h��SJIS�ŕۑ�����i�I�������Ȃ��j
		_wfopen_s( &fp, outputFileName.c_str(), L"w" );

		//�^�C�g���s�̏o�́��Ȃ�
		wstring tempField;
		//�Ǘ��ԍ�
		tempField = Replace(currentHardwareInfo.HardwareNoValue, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//�g�D1
		tempField = Replace(currentHardwareInfo.Value1, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//�g�D2
		tempField = Replace(currentHardwareInfo.Value2, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//�g�D3
		tempField = Replace(currentHardwareInfo.Value3, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//�g�D4
		tempField = Replace(currentHardwareInfo.Value4, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//�g�D5
		tempField = Replace(currentHardwareInfo.Value5, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//IP�A�h���X
		tempField = Replace(currentHardwareInfo.IPAddress, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//MAC�A�h���X
		tempField = Replace(currentHardwareInfo.MACAddress, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//���[�U��
		tempField = Replace(currentHardwareInfo.UserName, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//�R���s���[�^��
		tempField = Replace(currentHardwareInfo.ComputerName, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//�@��
		tempField = Replace(currentHardwareInfo.SystemProductName, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//�x���_�[
		tempField = Replace(currentHardwareInfo.SystemManufacturer, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//OS
		tempField = Replace(currentHardwareInfo.OSName, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//���g�p��1���g�p�Җ�
		tempField = Replace(currentHardwareInfo.Value6, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//���g�p��2��CPU�^�C�v
		tempField = Replace(currentHardwareInfo.ProcessorName, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//���g�p��3
		_ftprintf(fp, L"\"\",");
		//�擾����
		//_ftprintf(fp, L"\"%04d/%02d/%02d %02d:%02d:%02d\",", timestamp.tm_year + 1900, timestamp.tm_mon + 1, timestamp.tm_mday, timestamp.tm_hour, timestamp.tm_min, timestamp.tm_sec);
		_ftprintf(fp, L"\"%s\",", currentHardwareInfo.Timestamp2.c_str());
		//�c�[���̃o�[�W����
		tempField = Replace(currentHardwareInfo.ToolVersion, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());

		fclose(fp);
		returnOutputFileName += outputFileName.c_str();

		return returnOutputFileName;
}


//SIMPLE�`���œ��e��S�o�͂���֐�
wstring InventoryManager::outputSIMPLE(){

		//wchar_t outputFileNameChars [100];
		//swprintf(outputFileNameChars, 100,  _T( "%s%s.csv" ), currentHardwareInfo.ComputerName.c_str(), currentHardwareInfo.Timestamp.c_str());
		//wstring outputFileName = outputFileNameChars;
		wstring outputFileName = outputPath + currentHardwareInfo.ComputerName + currentHardwareInfo.Timestamp + L".csv";

		FILE* fp = NULL;

		//���ʂɏ���������SJIS�ɂȂ邪�ASJIS���ƊۃA�[����������R�ɂȂ��Ă��܂�
		//�v���O�����Ƌ@�\�ƈ�v���������邽�߂ɂ́AUTF-8�ŕۑ����ׂ�
		if(outputFileEncode == L"SJIS"){
			_wfopen_s( &fp, outputFileName.c_str(), L"w" );
		}else if(outputFileEncode == L"UTF-8N"){
			_wfopen_s( &fp, outputFileName.c_str(), L"w,ccs=UTF-8" );
			//���ʂɂ���BOM�t����UTF-8�ƂȂ�̂ŁABOM������UTF-8�Ƃ������ꍇ��seek����
			//http://social.msdn.microsoft.com/Forums/ja-JP/965c68dc-3e86-49a2-9c65-b469beb1f69f/bomutf8
			fseek(fp, 0, SEEK_SET); 
		}else{
			_wfopen_s( &fp, outputFileName.c_str(), L"w,ccs=UTF-8" );
		}

		//�^�C�g���s�̏o��
		//�擾�����A���[�J�[�A�@��A�V���A���ԍ��ACPU�ǉ�
		_ftprintf(fp, L"\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\"\n", currentHardwareInfo.HardwareNoTitle.c_str(), currentHardwareInfo.Title1.c_str(), currentHardwareInfo.Title2.c_str(), currentHardwareInfo.Title3.c_str(), currentHardwareInfo.Title4.c_str(), currentHardwareInfo.Title5.c_str(), currentHardwareInfo.Title6.c_str(), L"�R���s���[�^��", L"IP�A�h���X", L"MAC�A�h���X", L"PC�x���_�[", L"PC�@��", L"�V���A���ԍ�", L"CPU�^�C�v", L"���O�C�����[�U", L"OS", L"�\�t�g�E�F�A��", L"�x���_�[��", L"�o�[�W����", L"�擾����");
		vector<softwareInfo>::iterator softwareInfoIterator2 = currentSoftwareInfos.begin();  // �C�e���[�^�̃C���X�^���X��
		while(softwareInfoIterator2 != currentSoftwareInfos.end()){
			//ofstreamString = (*softwareInfoIterator).name;
			//_tprintf(TEXT("Software - Name:%s\t\n"), ofstreamString.c_str());  // *���Z�q�ŊԐڎQ��
			//ofstream�ł�string��<<�ő��荞�߂Ȃ�
			//���Ƃ�����.c_str()�Ƃ���ƁA�|�C���^�̕����񂪎�荞�܂��
			//ofs << ofstreamString << endl;
			//ofs.put(ofstreamString.c_str());
			//_tprintf(TEXT("Software - Name:%s\t\n"), (*softwareInfoIterator2).name.c_str());  // *���Z�q�ŊԐڎQ��
			
			wstring tempField;

			//�Ǘ��ԍ�
			tempField = Replace(currentHardwareInfo.HardwareNoValue, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//����1�`6
			tempField = Replace(currentHardwareInfo.Value1, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			tempField = Replace(currentHardwareInfo.Value2, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			tempField = Replace(currentHardwareInfo.Value3, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			tempField = Replace(currentHardwareInfo.Value4, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			tempField = Replace(currentHardwareInfo.Value5, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			tempField = Replace(currentHardwareInfo.Value6, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//�R���s���[�^��
			tempField = Replace(currentHardwareInfo.ComputerName, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//IP�A�h���X
			tempField = Replace(currentHardwareInfo.IPAddress, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//MAC�A�h���X
			tempField = Replace(currentHardwareInfo.MACAddress, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//PC�x���_�[
			tempField = Replace(currentHardwareInfo.SystemManufacturer, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//PC�@��
			tempField = Replace(currentHardwareInfo.SystemProductName, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//�V���A���ԍ�
			tempField = Replace(currentHardwareInfo.IdentifyingNumber, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//CPU�^�C�v
			tempField = Replace(currentHardwareInfo.ProcessorName, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//���[�U��
			tempField = Replace(currentHardwareInfo.UserName, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//OS
			tempField = Replace(currentHardwareInfo.OSName, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//�\�t�g�E�F�A��
			tempField= (*softwareInfoIterator2).Name.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			//_ftprintf(fp, TEXT("Software - Name:%s\t\n"), (*softwareInfoIterator).name.c_str());  // *���Z�q�ŊԐڎQ��
			_ftprintf(fp, L"\"%s\",", tempField.c_str());  // *���Z�q�ŊԐڎQ��
			//_ftprintf(fp, L"%s,", (*softwareInfoIterator2).name.c_str());  // *���Z�q�ŊԐڎQ��
			//�\�t�g�E�F�A�x���_�[
			tempField = (*softwareInfoIterator2).Vendor.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//�\�t�g�E�F�A�o�[�W����
			tempField = (*softwareInfoIterator2).Version.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//�擾����
			//_ftprintf(fp, L"\"%04d/%02d/%02d %02d:%02d:%02d\",", timestamp.tm_year + 1900, timestamp.tm_mon + 1, timestamp.tm_mday, timestamp.tm_hour, timestamp.tm_min, timestamp.tm_sec);
			_ftprintf(fp, L"\"%s\",", currentHardwareInfo.Timestamp2.c_str());

			if(debugLevel > 0){

				tempField = (*softwareInfoIterator2).InstallDate.c_str();
				tempField = Replace(tempField, L"\"", L"\"\"");
				_ftprintf(fp, L"\"%s\",", tempField.c_str());

				tempField = (*softwareInfoIterator2).ScanRegistryRange.c_str();
				tempField = Replace(tempField, L"\"", L"\"\"");
				_ftprintf(fp, L"\"%s\",", tempField.c_str());

				tempField = (*softwareInfoIterator2).CompareName.c_str();
				tempField = Replace(tempField, L"\"", L"\"\"");
				_ftprintf(fp, L"\"%s\",", tempField.c_str());

				tempField = (*softwareInfoIterator2).NameCodePoint.c_str();
				tempField = Replace(tempField, L"\"", L"\"\"");
				_ftprintf(fp, L"\"%s\",", tempField.c_str());
			}

			_ftprintf(fp, L"\n");
			++softwareInfoIterator2;                 // �C�e���[�^���P�i�߂�
		}
		//_ftprintf(fp, L"\n");
		//_ftprintf(fp, L"Output by LIP (List Installed Programs) v0.1.");
		fclose(fp);
		return outputFileName;

}



//SIMPLE�`���œ��e��S�o�͂���֐�
wstring InventoryManager::outputAdvancedManager(){

		wstring returnString = L"�C���x���g������ۑ����܂����B\n�ۑ��t�@�C��: \n";
		wstring tempField;

		//%ALLUSERSPROFILE%�ւ̃f�[�^�ۑ�
		//XP�܂ł�C:\Documents and Settings\All Users
		//Vista�ȍ~��C:\ProgramData���w��
		//http://pasofaq.jp/windows/mycomputer/folderlist.htm
		//�ۑ��ꏊ���ǂ��ɂ��邩
		//���[�U�[���ƂɃf�[�^��ۑ�����̂ł����HKEY_LOCL_USER�܂���%LOCALAPPDATA%���悢
		//���W���̓��[�U�[���Ƃł͂Ȃ�PC���Ƃ̃f�[�^�Ȃ̂ŁA%ALLUSERSPROFILE%���K��
		//http://sygh.hatenadiary.jp/entry/2013/11/10/184200
		//�������A�u�������݃A�N�Z�X���́A�t�@�C���̍쐬�� (���L��) �݂̂ɕt�^�v����邽�߁A
		//���[�U�[���ƂɃt�@�C���𕪂��ĕۑ�����K�v����
		//http://sygh.hatenadiary.jp/entry/2013/11/10/184200
		//http://bbs.wankuma.com/index.cgi?mode=al2&namber=65112&KLOG=110

		//�t�H���_�쐬
		//wstring datPath = _wgetenv(L"ALLUSERSPROFILE");
		//datPath = datPath + L"\\LIP";
		//_wmkdir(datPath.c_str());

		//datPath = datPath + L"\\test.ini";
		//DeleteFile(datPath.c_str());

		//load();
		//save();

		//�\���̂����̂܂܃t�@�C���ɏ������ޏꍇ
		//https://jyn.jp/cpp-structure-file/
		//fwrite((char*)&currentHardwareInfo, sizeof(currentHardwareInfo), 1, fp);
		//�o�C�i���ŏ������߂邽�߈Í������l���Ȃ��Ă悢�o�[�W�����A�b�v�Ŏ擾���ڂ��ς��ƑΉ������G�ɂȂ邽�߁A�̗p������
		//FILE* fp = NULL;
		//_wfopen_s(&fp, datPath.c_str(), L"wb");
		//fclose(fp);

		//ini�t�@�C���Ƃ��ď������ޏꍇ
		//https://hack.jp/?p=296
		//WritePrivateProfileString(TEXT("APP"), TEXT("VALUE1"), TEXT("STRING1"), szINIFilePath);
		//WritePrivateProfileString(TEXT("APP"), TEXT("VALUE1"), TEXT("STRING1"), datPath.c_str());

		//fstream file;
		//file.open("./filename.dat", ios::binary | ios::in);
		//file.write((char*)&n, sizeof(n));
		

		//�\�t�g�E�F�A���̏����o��
		wstring outputFileName = outputPath + currentHardwareInfo.HardwareNoValue + L"_" + currentHardwareInfo.ComputerName + L"_SW_" + currentHardwareInfo.Timestamp + L".csv";

		//�ۑ��ꏊ��������Ȃ��ꍇ�̃G���[�������K�v
		FILE* fp = NULL;
		errno_t error;
		//UTF-8�ŕۑ�
		error = _wfopen_s( &fp, outputFileName.c_str(), L"w,ccs=UTF-8" );

		if(error != 0){
			returnString = L"�C���x���g�����̕ۑ��Ɏ��s���܂����B�t�@�C���p�X�ƃA�N�Z�X�����m�F���Ă��������B\n���s�t�@�C��: \n";
			returnString += outputFileName;
			return returnString;
		}

		//wstring tempField;

		//�w�b�_�����o��
		_ftprintf(fp, L"\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\"\n", currentHardwareInfo.HardwareNoTitle.c_str(), L"�C���x���g���[�c�[��ID", L"�V���A��", L"�R���s���[�^��", L"IP�A�h���X", L"MAC�A�h���X", L"�C���X�g�[������", L"Hotfix�敪", L"HotfixApp����", L"�x���_�[", L"�o�[�W����", L"���s����", L"�X�V�t���O");

		//int softwareCount = 0;
		//wchar_t softwareCountChars [100];
		vector<softwareInfo>::iterator softwareInfoIterator2 = currentSoftwareInfos.begin();  // �C�e���[�^�̃C���X�^���X��
		
		while(softwareInfoIterator2 != currentSoftwareInfos.end()){

			//�ۑ�INI�p�̃J�E���^
			//softwareCount += 1;
			//swprintf(softwareCountChars, 100, _T( "Software%d" ), softwareCount);

			//�n�[�h�E�F�A�Ǘ��ԍ�
			tempField = Replace(currentHardwareInfo.HardwareNoValue, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//�C���x���g���c�[��ID
			//tempField = Replace(VER_STR_INVENTORYTOOLID, L"\"", L"\"\"");
			//_ftprintf(fp, L"\"%s\",", tempField.c_str());
			_ftprintf(fp, L"\"\",");

			//�V���A��
			tempField = Replace(currentHardwareInfo.IdentifyingNumber, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//�R���s���[�^��
			tempField = Replace(currentHardwareInfo.ComputerName, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//IP�A�h���X
			tempField = Replace(currentHardwareInfo.IPAddress, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//MAC�A�h���X
			tempField = Replace(currentHardwareInfo.MACAddress, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//�\�t�g�E�F�A��
			tempField = (*softwareInfoIterator2).Name.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//WritePrivateProfileString(softwareCountChars, L"name", (*softwareInfoIterator2).name.c_str(), datPath.c_str());

			//Hotfix�敪
			_ftprintf(fp, L"\"�A�v���P�[�V����\",");

			//HotfixApp����
			_ftprintf(fp, L"\"\",");

			//�\�t�g�E�F�A�x���_�[
			tempField = (*softwareInfoIterator2).Vendor.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//WritePrivateProfileString(softwareCountChars, L"vendor", (*softwareInfoIterator2).vendor.c_str(), datPath.c_str());

			//�\�t�g�E�F�A�o�[�W����
			tempField = (*softwareInfoIterator2).Version.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//WritePrivateProfileString(softwareCountChars, L"version", (*softwareInfoIterator2).version.c_str(), datPath.c_str());

			//���s����
			_ftprintf(fp, L"\"%s\",", currentHardwareInfo.Timestamp2.c_str());

			//�X�V�t���O
			tempField = Replace((*softwareInfoIterator2).UpdateFlag, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			if(debugLevel > 0){

				tempField = (*softwareInfoIterator2).InstallDate.c_str();
				tempField = Replace(tempField, L"\"", L"\"\"");
				_ftprintf(fp, L"\"%s\",", tempField.c_str());

				tempField = (*softwareInfoIterator2).ScanRegistryRange.c_str();
				tempField = Replace(tempField, L"\"", L"\"\"");
				_ftprintf(fp, L"\"%s\",", tempField.c_str());

				tempField = (*softwareInfoIterator2).CompareName.c_str();
				tempField = Replace(tempField, L"\"", L"\"\"");
				_ftprintf(fp, L"\"%s\",", tempField.c_str());

				tempField = (*softwareInfoIterator2).NameCodePoint.c_str();
				tempField = Replace(tempField, L"\"", L"\"\"");
				_ftprintf(fp, L"\"%s\",", tempField.c_str());

				tempField = (*softwareInfoIterator2).RegistoryKey.c_str();
				tempField = Replace(tempField, L"\"", L"\"\"");
				_ftprintf(fp, L"\"%s\",", tempField.c_str());
			}

			_ftprintf(fp, L"\n");
			++softwareInfoIterator2;                 // �C�e���[�^���P�i�߂�
		}
/*		
		fclose(fp);


		returnOutputFileName += outputFileName.c_str();
		returnOutputFileName += L", ";

		//�\�t�g�E�F�A������INI�ɕۊ�
		//swprintf(softwareCountChars, 100, _T( "%d" ), softwareCount);
		//WritePrivateProfileString(L"Software", L"count", softwareCountChars, datPath.c_str());

		//�p�b�`���̏����o��
		outputFileName = outputPath + currentHardwareInfo.HardwareNoValue + L"_" + currentHardwareInfo.ComputerName + L"_HF_" + currentHardwareInfo.Timestamp + L".csv";

		fp = NULL;
		//UTF-8�ŕۑ�
		_wfopen_s( &fp, outputFileName.c_str(), L"w,ccs=UTF-8" );

		//�w�b�_�����o��
		_ftprintf(fp, L"\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\"\n", currentHardwareInfo.HardwareNoTitle.c_str(), L"�X�V�t���O", L"�X�V�v���O������", L"�\�t�g�E�F�A��", L"�x���_�[��", L"�o�[�W����", L"�擾����");
*/
		//int softwareCount = 0;
		//wchar_t softwareCountChars [100];
		vector<patchInfo>::iterator patchInfoIterator2 = currentPatchInfos.begin();  // �C�e���[�^�̃C���X�^���X��
		
		while(patchInfoIterator2 != currentPatchInfos.end()){

			//�n�[�h�E�F�A�Ǘ��ԍ�
			tempField = Replace(currentHardwareInfo.HardwareNoValue, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//�C���x���g���c�[��ID
			//tempField = Replace(VER_STR_INVENTORYTOOLID, L"\"", L"\"\"");
			//_ftprintf(fp, L"\"%s\",", tempField.c_str());
			_ftprintf(fp, L"\"\",");

			//�V���A��
			tempField = Replace(currentHardwareInfo.IdentifyingNumber, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//�R���s���[�^��
			tempField = Replace(currentHardwareInfo.ComputerName, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//IP�A�h���X
			tempField = Replace(currentHardwareInfo.IPAddress, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//MAC�A�h���X
			tempField = Replace(currentHardwareInfo.MACAddress, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//�X�V�v���O������
			tempField = (*patchInfoIterator2).Name.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//Hotfix�敪
			_ftprintf(fp, L"\"HOTFIX\",");

			//HotfixApp����
			tempField = (*patchInfoIterator2).ParentName.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//�\�t�g�E�F�A�x���_�[
			tempField = (*patchInfoIterator2).Vendor.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//�\�t�g�E�F�A�o�[�W����
			tempField = (*patchInfoIterator2).Version.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//���s����
			_ftprintf(fp, L"\"%s\",", currentHardwareInfo.Timestamp2.c_str());

			//�X�V�t���O
			tempField = Replace((*patchInfoIterator2).UpdateFlag, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			if(debugLevel > 0){

				tempField = (*patchInfoIterator2).InstallDate.c_str();
				tempField = Replace(tempField, L"\"", L"\"\"");
				_ftprintf(fp, L"\"%s\",", tempField.c_str());

				tempField = (*patchInfoIterator2).ScanRegistryRange.c_str();
				tempField = Replace(tempField, L"\"", L"\"\"");
				_ftprintf(fp, L"\"%s\",", tempField.c_str());

				tempField = (*patchInfoIterator2).CompareName.c_str();
				tempField = Replace(tempField, L"\"", L"\"\"");
				_ftprintf(fp, L"\"%s\",", tempField.c_str());

				tempField = (*patchInfoIterator2).RegistoryKey.c_str();
				tempField = Replace(tempField, L"\"", L"\"\"");
				_ftprintf(fp, L"\"%s\",", tempField.c_str());
			}

			_ftprintf(fp, L"\n");
			++patchInfoIterator2;                 // �C�e���[�^���P�i�߂�
		}
		
		fclose(fp);

		returnString += outputFileName.c_str();
		returnString += L"\n";

		//�n�[�h�E�F�A���̏����o��
		//swprintf(outputFileNameChars, 100,  _T( "%s_%s_HW_%s.csv" ), currentHardwareInfo.HardwareNoValue.c_str(), currentHardwareInfo.ComputerName.c_str(), currentHardwareInfo.Timestamp.c_str());
		//outputFileName = outputFileNameChars;
		outputFileName = outputPath + currentHardwareInfo.HardwareNoValue + L"_" + currentHardwareInfo.ComputerName + L"_HW_" + currentHardwareInfo.Timestamp + L".csv";

		fp = NULL;
		//UTF-8�ŕۑ�
		error = _wfopen_s( &fp, outputFileName.c_str(), L"w,ccs=UTF-8" );

		if(error != 0){
			returnString = L"�C���x���g�����̕ۑ��Ɏ��s���܂����B�t�@�C���p�X�ƃA�N�Z�X�����m�F���Ă��������B\n���s�t�@�C��: \n";
			returnString += outputFileName;
			return returnString;
		}


		wstring fileLine1 = L"";
		wstring fileLine2 = L"";
		wstring fileLine3 = L"";


		//�Ǘ��ԍ�
		fileLine1 += L"\"" + Replace(currentHardwareInfo.HardwareNoTitle, L"\"", L"\"\"") + L"\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.HardwareNoValue, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.HardwareNoValueUpdateFlag + L"\",";
		//_ftprintf(fp, L"\"%s\",\"%s\",\"%s\"\n", currentHardwareInfo.HardwareNoValue.c_str(), L"", tempField.c_str());
		//WritePrivateProfileString(L"Hardware", currentHardwareInfo.HardwareNoTitle.c_str(), currentHardwareInfo.HardwareNoValue.c_str(), datPath.c_str());

		//�C���x���g���c�[��ID
		fileLine1 += L"\"�C���x���g���[�c�[��ID\",";
		//fileLine2 += L"\"" + Replace(VER_STR_INVENTORYTOOLID, L"\"", L"\"\"") + L"\",";
		fileLine2 += L"\"\",";
		fileLine3 += L"\"\",";

		//����1�`6
		if(isInput1Valid){
			fileLine1 += L"\"" + Replace(currentHardwareInfo.Title1, L"\"", L"\"\"") + L"\",";
			fileLine2 += L"\"" + Replace(currentHardwareInfo.Value1, L"\"", L"\"\"") + L"\",";
			fileLine3 += L"\"" + currentHardwareInfo.Value1UpdateFlag + L"\",";
		}
		if(isInput2Valid){
			fileLine1 += L"\"" + Replace(currentHardwareInfo.Title2, L"\"", L"\"\"") + L"\",";
			fileLine2 += L"\"" + Replace(currentHardwareInfo.Value2, L"\"", L"\"\"") + L"\",";
			fileLine3 += L"\"" + currentHardwareInfo.Value2UpdateFlag + L"\",";
		}
		if(isInput3Valid){
			fileLine1 += L"\"" + Replace(currentHardwareInfo.Title3, L"\"", L"\"\"") + L"\",";
			fileLine2 += L"\"" + Replace(currentHardwareInfo.Value3, L"\"", L"\"\"") + L"\",";
			fileLine3 += L"\"" + currentHardwareInfo.Value3UpdateFlag + L"\",";
		}
		if(isInput4Valid){
			fileLine1 += L"\"" + Replace(currentHardwareInfo.Title4, L"\"", L"\"\"") + L"\",";
			fileLine2 += L"\"" + Replace(currentHardwareInfo.Value4, L"\"", L"\"\"") + L"\",";
			fileLine3 += L"\"" + currentHardwareInfo.Value4UpdateFlag + L"\",";
		}
		if(isInput5Valid){
			fileLine1 += L"\"" + Replace(currentHardwareInfo.Title5, L"\"", L"\"\"") + L"\",";
			fileLine2 += L"\"" + Replace(currentHardwareInfo.Value5, L"\"", L"\"\"") + L"\",";
			fileLine3 += L"\"" + currentHardwareInfo.Value5UpdateFlag + L"\",";
		}
		if(isInput6Valid){
			fileLine1 += L"\"" + Replace(currentHardwareInfo.Title6, L"\"", L"\"\"") + L"\",";
			fileLine2 += L"\"" + Replace(currentHardwareInfo.Value6, L"\"", L"\"\"") + L"\",";
			fileLine3 += L"\"" + currentHardwareInfo.Value6UpdateFlag + L"\",";
		}

		//���[�J�[��
		fileLine1 += L"\"���[�J�[��\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.SystemManufacturer, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.SystemManufacturerUpdateFlag + L"\",";
		//tempField = Replace(currentHardwareInfo.SystemManufacturer, L"\"", L"\"\"");
		//_ftprintf(fp, L"\"%s\",\"%s\",\"%s\"\n", L"���[�J�[��", L"", tempField.c_str());
		//WritePrivateProfileString(L"Hardware", L"SystemManufacturer", currentHardwareInfo.SystemManufacturer.c_str(), datPath.c_str());

		//�^��
		fileLine1 += L"\"�^��\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.SystemProductName, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.SystemProductNameUpdateFlag + L"\",";

		//�n�[�h�E�F�A���
		fileLine1 += L"\"�n�[�h�E�F�A���\",";
		fileLine2 += L"\"\",";
		fileLine3 += L"\"\",";

		//�V���A��
		fileLine1 += L"\"�V���A��\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.IdentifyingNumber, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.IdentifyingNumberUpdateFlag + L"\",";

		//�R���s���[�^��
		fileLine1 += L"\"�R���s���[�^��\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.ComputerName, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.ComputerNameUpdateFlag + L"\",";

		//CPU�^�C�v
		fileLine1 += L"\"CPU\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.ProcessorName, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.ProcessorNameUpdateFlag + L"\",";

		//CPU������
		fileLine1 += L"\"CPU������\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.NumberOfProcessors, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.NumberOfProcessorsUpdateFlag + L"\",";

		//CPU�R�A��
		fileLine1 += L"\"CPU�R�A��\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.NumberOfCores, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.NumberOfCoresUpdateFlag + L"\",";

		//CPU�_���v���Z�b�T��
		fileLine1 += L"\"CPU�_���v���Z�b�T��\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.NumberOfLogicalProcessors, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.NumberOfLogicalProcessorsUpdateFlag + L"\",";

		//CPU�N���b�N��
		fileLine1 += L"\"CPU�N���b�N��\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.ProcessorMaxClockSpeed, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.ProcessorMaxClockSpeedUpdateFlag + L"\",";

		//CPU�\�P�b�g��
		fileLine1 += L"\"CPU�\�P�b�g��\",";
		fileLine2 += L"\"\",";
		fileLine3 += L"\"\",";

		//�������e��
		fileLine1 += L"\"����������\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.TotalPhysicalMemory, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.TotalPhysicalMemoryUpdateFlag + L"\",";

		//�f�B�X�N�e��
		fileLine1 += L"\"�f�B�X�N�e��\",";
		fileLine2 += L"\"\",";
		fileLine3 += L"\"\",";

		//IP�A�h���X
		fileLine1 += L"\"IP�A�h���X\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.IPAddress, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.IPAddressUpdateFlag + L"\",";
		//tempField = Replace(currentHardwareInfo.IPAddress, L"\"", L"\"\"");
		//_ftprintf(fp, L"\"%s\",\"%s\",\"%s\"\n", L"IP�A�h���X", L"", tempField.c_str());
		//WritePrivateProfileString(L"Hardware", L"IPAddress", currentHardwareInfo.IPAddress.c_str(), datPath.c_str());

		//MAC�A�h���X
		fileLine1 += L"\"MAC�A�h���X\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.MACAddress, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.MACAddressUpdateFlag + L"\",";
		//tempField = Replace(currentHardwareInfo.MACAddress, L"\"", L"\"\"");
		//_ftprintf(fp, L"\"%s\",\"%s\",\"%s\"\n", L"MAC�A�h���X", L"", tempField.c_str());
		//WritePrivateProfileString(L"Hardware", L"MACAddress", currentHardwareInfo.MACAddress.c_str(), datPath.c_str());

		//OS
		fileLine1 += L"\"OS\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.OSName, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.OSNameUpdateFlag + L"\",";
		//tempField = Replace(currentHardwareInfo.OSName, L"\"", L"\"\"");
		//_ftprintf(fp, L"\"%s\",\"%s\",\"%s\"\n", L"OS", L"", tempField.c_str());
		//WritePrivateProfileString(L"Hardware", L"OSName", currentHardwareInfo.OSName.c_str(), datPath.c_str());

		//���O�C�����[�U
		fileLine1 += L"\"���O�C��ID\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.UserName, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.UserNameUpdateFlag + L"\",";
		//tempField = Replace(currentHardwareInfo.UserName, L"\"", L"\"\"");
		//_ftprintf(fp, L"\"%s\",\"%s\",\"%s\"\n", L"���O�C�����[�U", L"", tempField.c_str());
		//WritePrivateProfileString(L"Hardware", L"userName", currentHardwareInfo.UserName.c_str(), datPath.c_str());

		//���z����
		fileLine1 += L"\"���z��\",";
		fileLine2 += L"\"" + currentHardwareInfo.Virtualization + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.VirtualizationUpdateFlag + L"\",";
		//tempField = Replace(currentHardwareInfo.OSName, L"\"", L"\"\"");
		//_ftprintf(fp, L"\"%s\",\"%s\",\"%s\"\n", L"OS", L"", tempField.c_str());
		//WritePrivateProfileString(L"Hardware", L"OSName", currentHardwareInfo.OSName.c_str(), datPath.c_str());

		//���s����
		fileLine1 += L"\"���s����\",";
		fileLine2 += L"\"" + currentHardwareInfo.Timestamp2 + L"\",";
		fileLine3 += L"\"\",";
		//_ftprintf(fp, L"\"%s\",\"%s\",\"%04d/%02d/%02d %02d:%02d:%02d\"\n", L"���s����", L"", timestamp.tm_year + 1900, timestamp.tm_mon + 1, timestamp.tm_mday, timestamp.tm_hour, timestamp.tm_min, timestamp.tm_sec);

		//�X�V�t���O
		fileLine1 += L"\"�X�V�t���O\",";
		int indexUpdate = (int)fileLine3.find(L"UPDATE");
		if(indexUpdate != -1){
			fileLine2 += L"\"UPDATE\",";
		}else{
			fileLine2 += L"\"\",";
		}
		fileLine3 += L"\"\",";

		//�c�[���̃o�[�W����
		//_ftprintf(fp, VER_STR_INVENTORYTOOLID);

		_ftprintf(fp, L"%s\n", fileLine1.c_str());
		_ftprintf(fp, L"%s\n", fileLine2.c_str());
		//_ftprintf(fp, L"%s\n", fileLine3.c_str());


		fclose(fp);
		returnString += outputFileName.c_str();
		return returnString;

}

//�C���x���g����ۑ�����֐�
long InventoryManager::save(){

	//�t�H���_��������΍쐬
	wstring datPath = _wgetenv(L"ALLUSERSPROFILE");
	datPath = datPath + L"\\LIP";
	_wmkdir(datPath.c_str());

	//�t�H���_���̃t�@�C����S�폜�i�A�N�Z�X���̊֌W�ŉ\�ł���΁j
	//http://blog.systemjp.net/entry/20100318/p5
	//DeleteFile(datPath.c_str());

    HANDLE hFind;
    WIN32_FIND_DATA win32fd;//defined at Windwos.h
	wstring tempFilename;
	wstring searchPath = datPath + L"\\*.dat";
    hFind = FindFirstFile(searchPath.c_str(), &win32fd);

    if(hFind == INVALID_HANDLE_VALUE){
        //return -1;
		//�G���[�i�t�@�C����������Ȃ��ꍇ�܂ށj�͉������Ȃ�
    }
    do{
        if(win32fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY){
			//�w��t�H���_���̃T�u�t�H���_�͉������Ȃ�
        }
        else{
			//�w��t�H���_����dat�t�@�C���͍폜
            tempFilename = win32fd.cFileName;
			tempFilename = datPath + L"\\" + tempFilename;
			//MessageBox(NULL,tempFilename.c_str(),L"test",MB_OK);
			DeleteFile(tempFilename.c_str());
        }
    }while(FindNextFile(hFind, &win32fd));

    FindClose(hFind);

	//�f�[�^�̏�������
	//http://www.geocities.co.jp/SiliconValley/6071/technic/74.html
	//FILE *fp = fopen(datPath.c_str(), "wb");
	wstring savePath = datPath + L"\\" + currentHardwareInfo.Timestamp + L".dat";
	FILE* fp = NULL;
	_wfopen_s( &fp, savePath.c_str(), L"wb" );
	//MessageBox(NULL,datPath.c_str(),L"test",MB_OK);

	//fwriteWString(fp, L"test string.");

	vector<softwareInfo>::iterator softwareInfoIterator3 = currentSoftwareInfos.begin();  // �C�e���[�^�̃C���X�^���X��
	while(softwareInfoIterator3 != currentSoftwareInfos.end()){

		if((*softwareInfoIterator3).UpdateFlag != L"REMOVE"){
			//���W�X�g���̃T�u�L�[
			fwriteWString(fp, L"sw.RegistoryKey");
			fwriteWString(fp, (*softwareInfoIterator3).RegistoryKey);
			//�\�t�g�E�F�A��
			fwriteWString(fp, L"sw.Name");
			fwriteWString(fp, (*softwareInfoIterator3).Name);
			//�\�t�g�E�F�A�x���_�[
			fwriteWString(fp, L"sw.Vendor");
			fwriteWString(fp, (*softwareInfoIterator3).Vendor);
			//�\�t�g�E�F�A�o�[�W����
			fwriteWString(fp, L"sw.Version");
			fwriteWString(fp, (*softwareInfoIterator3).Version);
			
			//�ۊǂ��Ȃ�
			/*
			fwriteWString(fp, L"sw.InstallDate");
			fwriteWString(fp, (*softwareInfoIterator3).InstallDate);
			fwriteWString(fp, L"sw.ScanRegistryRange");
			fwriteWString(fp, (*softwareInfoIterator3).ScanRegistryRange);
			fwriteWString(fp, L"sw.CompareName");
			fwriteWString(fp, (*softwareInfoIterator3).CompareName);
			fwriteWString(fp, L"sw.NameCodePoint");
			fwriteWString(fp, (*softwareInfoIterator3).NameCodePoint);
			*/

			fwriteWString(fp, L"++sw");
		}
		//�C�e���[�^��1�i�߂�
		++softwareInfoIterator3;
	}

	vector<patchInfo>::iterator patchInfoIterator3 = currentPatchInfos.begin();  // �C�e���[�^�̃C���X�^���X��
	while(patchInfoIterator3 != currentPatchInfos.end()){

		if((*patchInfoIterator3).UpdateFlag != L"REMOVE"){
			//�X�V�v���O������
			fwriteWString(fp, L"hf.Name");
			fwriteWString(fp, (*patchInfoIterator3).Name);
			//�\�t�g�E�F�A��
			fwriteWString(fp, L"hf.ParentName");
			fwriteWString(fp, (*patchInfoIterator3).ParentName);
			//�\�t�g�E�F�A�x���_�[
			fwriteWString(fp, L"hf.Vendor");
			fwriteWString(fp, (*patchInfoIterator3).Vendor);
			//�\�t�g�E�F�A�o�[�W����
			fwriteWString(fp, L"hf.Version");
			fwriteWString(fp, (*patchInfoIterator3).Version);
			//��r���i�\�t�g�E�F�A���{�X�V�v���O�������j
			fwriteWString(fp, L"hf.CompareName");
			fwriteWString(fp, (*patchInfoIterator3).CompareName);
			
			fwriteWString(fp, L"++hf");
		}
		//�C�e���[�^��1�i�߂�
		++patchInfoIterator3;
	}

	//�Ǘ��ԍ�
	fwriteWString(fp, L"hw.HardwareNoTitle");
	fwriteWString(fp, currentHardwareInfo.HardwareNoTitle);
	fwriteWString(fp, L"hw.HardwareNoValue");
	fwriteWString(fp, currentHardwareInfo.HardwareNoValue);

	//����1�`6
	fwriteWString(fp, L"hw.Title1");
	fwriteWString(fp, currentHardwareInfo.Title1);
	fwriteWString(fp, L"hw.Value1");
	fwriteWString(fp, currentHardwareInfo.Value1);
	fwriteWString(fp, L"hw.Title2");
	fwriteWString(fp, currentHardwareInfo.Title2);
	fwriteWString(fp, L"hw.Value2");
	fwriteWString(fp, currentHardwareInfo.Value2);
	fwriteWString(fp, L"hw.Title3");
	fwriteWString(fp, currentHardwareInfo.Title3);
	fwriteWString(fp, L"hw.Value3");
	fwriteWString(fp, currentHardwareInfo.Value3);
	fwriteWString(fp, L"hw.Title4");
	fwriteWString(fp, currentHardwareInfo.Title4);
	fwriteWString(fp, L"hw.Value4");
	fwriteWString(fp, currentHardwareInfo.Value4);
	fwriteWString(fp, L"hw.Title5");
	fwriteWString(fp, currentHardwareInfo.Title5);
	fwriteWString(fp, L"hw.Value5");
	fwriteWString(fp, currentHardwareInfo.Value5);
	fwriteWString(fp, L"hw.Title6");
	fwriteWString(fp, currentHardwareInfo.Title6);
	fwriteWString(fp, L"hw.Value6");
	fwriteWString(fp, currentHardwareInfo.Value6);

	//�R���s���[�^��
	fwriteWString(fp, L"hw.ComputerName");
	fwriteWString(fp, currentHardwareInfo.ComputerName);

	//IP�A�h���X
	fwriteWString(fp, L"hw.IPAddress");
	fwriteWString(fp, currentHardwareInfo.IPAddress);

	//MAC�A�h���X
	fwriteWString(fp, L"hw.MACAddress");
	fwriteWString(fp, currentHardwareInfo.MACAddress);

	//���[�J�[��
	fwriteWString(fp, L"hw.SystemManufacturer");
	fwriteWString(fp, currentHardwareInfo.SystemManufacturer);

	//�^��
	fwriteWString(fp, L"hw.SystemProductName");
	fwriteWString(fp, currentHardwareInfo.SystemProductName);

	//�V���A��
	fwriteWString(fp, L"hw.IdentifyingNumber");
	fwriteWString(fp, currentHardwareInfo.IdentifyingNumber);

	//CPU�^�C�v
	fwriteWString(fp, L"hw.ProcessorName");
	fwriteWString(fp, currentHardwareInfo.ProcessorName);

	//CPU���g��
	fwriteWString(fp, L"hw.ProcessorMaxClockSpeed");
	fwriteWString(fp, currentHardwareInfo.ProcessorMaxClockSpeed);

	//CPU��
	fwriteWString(fp, L"hw.NumberOfProcessors");
	fwriteWString(fp, currentHardwareInfo.NumberOfProcessors);

	//�R�A��
	fwriteWString(fp, L"hw.NumberOfCores");
	fwriteWString(fp, currentHardwareInfo.NumberOfCores);

	//�_���v���Z�b�T��
	fwriteWString(fp, L"hw.NumberOfLogicalProcessors");
	fwriteWString(fp, currentHardwareInfo.NumberOfLogicalProcessors);

	//���O�C�����[�U
	fwriteWString(fp, L"hw.UserName");
	fwriteWString(fp, currentHardwareInfo.UserName);

	//OS
	fwriteWString(fp, L"hw.OSName");
	fwriteWString(fp, currentHardwareInfo.OSName);

	//���z����
	fwriteWString(fp, L"hw.Virtualization");
	fwriteWString(fp, currentHardwareInfo.Virtualization);

	//�������e��
	fwriteWString(fp, L"hw.TotalPhysicalMemory");
	fwriteWString(fp, currentHardwareInfo.TotalPhysicalMemory);

	//���s����
	fwriteWString(fp, L"hw.Timestamp");
	fwriteWString(fp, currentHardwareInfo.Timestamp);

	//�c�[���̃o�[�W����
	fwriteWString(fp, L"hw.ToolVersion");
	fwriteWString(fp, currentHardwareInfo.ToolVersion);

	fclose(fp);
	return 0;
}	

//�O��̃C���x���g�������[�h����֐�
long InventoryManager::load(){

	//�t�H���_��������΍쐬
	wstring datPath = _wgetenv(L"ALLUSERSPROFILE");
	datPath = datPath + L"\\LIP";
	_wmkdir(datPath.c_str());

	//�t�H���_���̃t�@�C���𑖍����āA�t�@�C�����̓����������_���O���ŐV�̂��̂��s�b�N�A�b�v
	//http://qiita.com/tkymx/items/f9190c16be84d4a48f8a
    HANDLE hFind;
    WIN32_FIND_DATA win32fd;//defined at Windwos.h

	wstring searchPath = datPath + L"\\*.dat";
	//MessageBox(NULL,searchPath.c_str(),L"search",MB_OK);
    hFind = FindFirstFile(searchPath.c_str(), &win32fd);
	wstring loadFilename = L"";
	wstring tempFilename;
	wstring nowFilename = currentHardwareInfo.Timestamp + L".dat";
	//MessageBox(NULL,nowFilename.c_str(),L"now",MB_OK);

	if(hFind == INVALID_HANDLE_VALUE){
		//MessageBox(NULL,L"-1",L"now",MB_OK);
        return -1;
    }

    do{
        if(win32fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY){
			//�������Ȃ�
        }
        else{
            tempFilename = win32fd.cFileName;
			//MessageBox(NULL,tempFilename.c_str(),L"temp",MB_OK);
			if(loadFilename < tempFilename && tempFilename < nowFilename){
				loadFilename = tempFilename;
				//MessageBox(NULL,loadFilename.c_str(),L"load",MB_OK);
			}
        }
    }while(FindNextFile(hFind, &win32fd));

    FindClose(hFind);
	//MessageBox(NULL,loadFilename.c_str(),L"last",MB_OK);

	wstring loadFilePath = datPath + L"\\" + loadFilename;

	//�f�[�^�̓ǂݍ���
	//http://www.geocities.co.jp/SiliconValley/6071/technic/74.html
	//FILE *fp = fopen(datPath.c_str(), "rb");
	FILE* fp = NULL;
	_wfopen_s( &fp, loadFilePath.c_str(), L"rb" );

	wstring readString;
	//long r;
	//r = freadWString(fp, readString);
	//if(r>0) MessageBox(NULL,readString.c_str(),L"test",MB_OK);

	//softwareInfo�\���̂�p��
	softwareInfo tempSoftwareInfo;
	patchInfo tempPatchInfo;

	//�\�t�g�E�F�A���X�g�̃t���O��S�āu�V�K�v�ɃZ�b�g
	vector<softwareInfo>::iterator softwareInfoIterator = currentSoftwareInfos.begin();
	while(softwareInfoIterator != currentSoftwareInfos.end()){
		(*softwareInfoIterator).UpdateFlag = L"ADD";
		//MessageBox(NULL,(*softwareInfoIterator).updateFlag.c_str(),L"test",MB_OK);
		++softwareInfoIterator;
	}
	vector<patchInfo>::iterator patchInfoIterator = currentPatchInfos.begin();
	while(patchInfoIterator != currentPatchInfos.end()){
		(*patchInfoIterator).UpdateFlag = L"ADD";
		//MessageBox(NULL,(*patchInfoIterator).updateFlag.c_str(),L"test",MB_OK);
		++patchInfoIterator;
	}

	while(freadWString(fp, readString) > 0){
		if(readString == L"sw.RegistoryKey"){
			freadWString(fp, readString);
			tempSoftwareInfo.RegistoryKey = readString;
		}else if(readString == L"sw.Name"){
			freadWString(fp, readString);
			tempSoftwareInfo.Name = readString;
			//MessageBox(NULL,readString.c_str(),L"test",MB_OK);
		}else if(readString == L"sw.Vendor"){
			freadWString(fp, readString);
			tempSoftwareInfo.Vendor = readString;
		}else if(readString == L"sw.Version"){
			freadWString(fp, readString);
			tempSoftwareInfo.Version = readString;
		}else if(readString == L"sw.InstallDate"){
			freadWString(fp, readString);
			tempSoftwareInfo.InstallDate = readString;
		}else if(readString == L"sw.ScanRegistryRange"){
			freadWString(fp, readString);
			tempSoftwareInfo.ScanRegistryRange = readString;
		}else if(readString == L"sw.CompareName"){
			freadWString(fp, readString);
			tempSoftwareInfo.CompareName = readString;
		}else if(readString == L"sw.NameCodePoint"){
			freadWString(fp, readString);
			tempSoftwareInfo.NameCodePoint = readString;
		}else if(readString == L"++sw"){
			//tempSoftwareInfo�i�̃��W�X�g���L�[�j�ƍ��v������̂��\�t�g�E�F�A���X�g�ɂ��邩����
			bool registoryKeyMatch = false;
			softwareInfoIterator = currentSoftwareInfos.begin();
			while(softwareInfoIterator != currentSoftwareInfos.end()){
				//���v������̂��������ꍇ
				if(tempSoftwareInfo.RegistoryKey == (*softwareInfoIterator).RegistoryKey){
					//MessageBox(NULL,tempSoftwareInfo.registoryKey.c_str(),L"test",MB_OK);
					registoryKeyMatch = true;
					//�\�t�g�E�F�A������v���Ă���΃t���O�͋�ɃZ�b�g
					if(tempSoftwareInfo.Name == (*softwareInfoIterator).Name){
						(*softwareInfoIterator).UpdateFlag = L"";
					//�\�t�g�E�F�A������v���Ă��Ȃ���΃\�t�g�E�F�A���̍X�V��������
					}else{
						(*softwareInfoIterator).UpdateFlag = L"UPDATE";
					}
				}
				//�C�e���[�^��1�i�߂�
				++softwareInfoIterator;
			}
			///tempSoftwareInfo�i�̃��W�X�g���L�[�j�ƍ��v������̂��\�t�g�E�F�A���X�g�Ɉꌏ�����������ꍇ���\�t�g�E�F�A���폜���ꂽ
			//2018/10/20 �폜�ς݂̃\�t�g�E�F�A��"REMOVE"�ƃt���O�𗧂Ă������Ń��X�g�ɏo���Ă������A
			//�v�]�ɂ�胊�X�g�ɏo���Ȃ��悤�ɕύX
			if(registoryKeyMatch == false){
				//tempSoftwareInfo.UpdateFlag = L"REMOVE";
				//currentSoftwareInfos.push_back(tempSoftwareInfo);				
			}

		}else if(readString == L"hf.Name"){
			freadWString(fp, readString);
			tempPatchInfo.Name = readString;
			//MessageBox(NULL,readString.c_str(),L"test",MB_OK);
		}else if(readString == L"hf.ParentName"){
			freadWString(fp, readString);
			tempPatchInfo.ParentName = readString;
		}else if(readString == L"hf.KB"){
			freadWString(fp, readString);
			tempPatchInfo.KB = readString;
		}else if(readString == L"hf.Vendor"){
			freadWString(fp, readString);
			tempPatchInfo.Vendor = readString;
		}else if(readString == L"hf.Version"){
			freadWString(fp, readString);
			tempPatchInfo.Version = readString;
		}else if(readString == L"hf.InstallDate"){
			freadWString(fp, readString);
			tempPatchInfo.InstallDate = readString;
		}else if(readString == L"hf.ScanRegistryRange"){
			freadWString(fp, readString);
			tempPatchInfo.ScanRegistryRange = readString;
		}else if(readString == L"hf.CompareName"){
			freadWString(fp, readString);
			tempPatchInfo.CompareName = readString;
		}else if(readString == L"hf.NameCodePoint"){
			freadWString(fp, readString);
			tempPatchInfo.NameCodePoint = readString;
		}else if(readString == L"hf.RegistoryKey"){
			freadWString(fp, readString);
			tempPatchInfo.RegistoryKey = readString;
		}else if(readString == L"++hf"){
			//tempPatchInfo�i�̖��O�j�ƍ��v������̂��\�t�g�E�F�A���X�g�ɂ��邩����
			bool registoryKeyMatch = false;
			patchInfoIterator = currentPatchInfos.begin();
			while(patchInfoIterator != currentPatchInfos.end()){
				//���v������̂��������ꍇ
				if(tempPatchInfo.CompareName == (*patchInfoIterator).CompareName){
					registoryKeyMatch = true;
					//�t���O�͋�ɃZ�b�g
					(*patchInfoIterator).UpdateFlag = L"";
				}
				//�C�e���[�^��1�i�߂�
				++patchInfoIterator;
			}
			///tempSoftwareInfo�i�̃��W�X�g���L�[�j�ƍ��v������̂��\�t�g�E�F�A���X�g�Ɉꌏ�����������ꍇ���\�t�g�E�F�A���폜���ꂽ
			//2018/10/20 �폜�ς݂̃\�t�g�E�F�A��"REMOVE"�ƃt���O�𗧂Ă������Ń��X�g�ɏo���Ă������A
			//�v�]�ɂ�胊�X�g�ɏo���Ȃ��悤�ɕύX
			if(registoryKeyMatch == false){
				//tempPatchInfo.UpdateFlag = L"REMOVE";
				//MessageBox(NULL,readString.c_str(),L"patch remove",MB_OK);
				//currentPatchInfos.push_back(tempPatchInfo);				
			}

		}else if(readString == L"hw.HardwareNoTitle"){
			freadWString(fp, readString);
		}else if(readString == L"hw.HardwareNoValue"){
			freadWString(fp, readString);
			//MessageBox(NULL,readString.c_str(),L"test0",MB_OK);
			if(currentHardwareInfo.HardwareNoValue != readString){
				currentHardwareInfo.HardwareNoValueUpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.HardwareNoValueUpdateFlag = L"";
			}
			previousHardwareInfo.HardwareNoValue = readString;
			//MessageBox(NULL,previousHardwareInfo.HardwareNoValue.c_str(),L"test1",MB_OK);
		}else if(readString == L"hw.Title1"){
			freadWString(fp, readString);
		}else if(readString == L"hw.Value1"){
			freadWString(fp, readString);
			if(currentHardwareInfo.Value1 != readString){
				currentHardwareInfo.Value1UpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.Value1UpdateFlag = L"";
			}
			previousHardwareInfo.Value1 = readString;
		}else if(readString == L"hw.Title2"){
			freadWString(fp, readString);
		}else if(readString == L"hw.Value2"){
			freadWString(fp, readString);
			if(currentHardwareInfo.Value2 != readString){
				currentHardwareInfo.Value2UpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.Value2UpdateFlag = L"";
			}
			previousHardwareInfo.Value2 = readString;
		}else if(readString == L"hw.Title3"){
			freadWString(fp, readString);
		}else if(readString == L"hw.Value3"){
			freadWString(fp, readString);
			if(currentHardwareInfo.Value3 != readString){
				currentHardwareInfo.Value3UpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.Value3UpdateFlag = L"";
			}
			previousHardwareInfo.Value3 = readString;
		}else if(readString == L"hw.Title4"){
			freadWString(fp, readString);
		}else if(readString == L"hw.Value4"){
			freadWString(fp, readString);
			if(currentHardwareInfo.Value4 != readString){
				currentHardwareInfo.Value4UpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.Value4UpdateFlag = L"";
			}
			previousHardwareInfo.Value4 = readString;
		}else if(readString == L"hw.Title5"){
			freadWString(fp, readString);
		}else if(readString == L"hw.Value5"){
			freadWString(fp, readString);
			if(currentHardwareInfo.Value5 != readString){
				currentHardwareInfo.Value5UpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.Value5UpdateFlag = L"";
			}
			previousHardwareInfo.Value5 = readString;
		}else if(readString == L"hw.Title6"){
			freadWString(fp, readString);
		}else if(readString == L"hw.Value6"){
			freadWString(fp, readString);
			if(currentHardwareInfo.Value6 != readString){
				currentHardwareInfo.Value6UpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.Value6UpdateFlag = L"";
			}
			previousHardwareInfo.Value6 = readString;
		}else if(readString == L"hw.ComputerName"){
			freadWString(fp, readString);
			if(currentHardwareInfo.ComputerName != readString){
				currentHardwareInfo.ComputerNameUpdateFlag = L"UPDATE";
				//wstring temp = currentHardwareInfo.ComputerName + L" != " + readString;
				//MessageBox(NULL,temp.c_str(),L"test",MB_OK);
			}else{
				currentHardwareInfo.ComputerNameUpdateFlag = L"";
			}
		}else if(readString == L"hw.IPAddress"){
			freadWString(fp, readString);
			if(currentHardwareInfo.IPAddress != readString){
				currentHardwareInfo.IPAddressUpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.IPAddressUpdateFlag = L"";
			}
		}else if(readString == L"hw.MACAddress"){
			freadWString(fp, readString);
			if(currentHardwareInfo.MACAddress != readString){
				currentHardwareInfo.MACAddressUpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.MACAddressUpdateFlag = L"";
			}
		}else if(readString == L"hw.SystemManufacturer"){
			freadWString(fp, readString);
			if(currentHardwareInfo.SystemManufacturer != readString){
				currentHardwareInfo.SystemManufacturerUpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.SystemManufacturerUpdateFlag = L"";
			}
		}else if(readString == L"hw.SystemProductName"){
			freadWString(fp, readString);
			if(currentHardwareInfo.SystemProductName != readString){
				currentHardwareInfo.SystemProductNameUpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.SystemProductNameUpdateFlag = L"";
			}
		}else if(readString == L"hw.IdentifyingNumber"){
			freadWString(fp, readString);
			if(currentHardwareInfo.IdentifyingNumber != readString){
				currentHardwareInfo.IdentifyingNumberUpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.IdentifyingNumberUpdateFlag = L"";
			}
		}else if(readString == L"hw.ProcessorName"){
			freadWString(fp, readString);
			if(currentHardwareInfo.ProcessorName != readString){
				currentHardwareInfo.ProcessorNameUpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.ProcessorNameUpdateFlag = L"";
			}
		}else if(readString == L"hw.ProcessorMaxClockSpeed"){
			freadWString(fp, readString);
			if(currentHardwareInfo.ProcessorMaxClockSpeed != readString){
				currentHardwareInfo.ProcessorMaxClockSpeedUpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.ProcessorMaxClockSpeedUpdateFlag = L"";
			}
		}else if(readString == L"hw.NumberOfProcessors"){
			freadWString(fp, readString);
			if(currentHardwareInfo.NumberOfProcessors != readString){
				currentHardwareInfo.NumberOfProcessorsUpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.NumberOfProcessorsUpdateFlag = L"";
			}
		}else if(readString == L"hw.NumberOfCores"){
			freadWString(fp, readString);
			if(currentHardwareInfo.NumberOfCores != readString){
				currentHardwareInfo.NumberOfCoresUpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.NumberOfCoresUpdateFlag = L"";
			}
		}else if(readString == L"hw.NumberOfLogicalProcessors"){
			freadWString(fp, readString);
			if(currentHardwareInfo.NumberOfLogicalProcessors != readString){
				currentHardwareInfo.NumberOfLogicalProcessorsUpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.NumberOfLogicalProcessorsUpdateFlag = L"";
			}
		}else if(readString == L"hw.UserName"){
			freadWString(fp, readString);
			if(currentHardwareInfo.UserName != readString){
				currentHardwareInfo.UserNameUpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.UserNameUpdateFlag = L"";
			}
		}else if(readString == L"hw.OSName"){
			//MessageBox(NULL,readString.c_str(),L"test1",MB_OK);
			freadWString(fp, readString);
			//MessageBox(NULL,readString.c_str(),L"test2",MB_OK);
			if(currentHardwareInfo.OSName != readString){
				currentHardwareInfo.OSNameUpdateFlag = L"UPDATE";
				//wstring temp = currentHardwareInfo.ComputerName + L" != " + readString;
				//MessageBox(NULL,temp.c_str(),L"test3",MB_OK);
			}else{
				currentHardwareInfo.OSNameUpdateFlag = L"";
			}
		}else if(readString == L"hw.Virtualzation"){
			freadWString(fp, readString);
			if(currentHardwareInfo.Virtualization != readString){
				currentHardwareInfo.VirtualizationUpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.VirtualizationUpdateFlag = L"";
			}
		}else if(readString == L"hw.TotalPhysicalMemory"){
			freadWString(fp, readString);
			if(currentHardwareInfo.TotalPhysicalMemory != readString){
				currentHardwareInfo.TotalPhysicalMemoryUpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.TotalPhysicalMemoryUpdateFlag = L"";
			}
		}else if(readString == L"hw.Timestamp"){
			freadWString(fp, readString);
			if(currentHardwareInfo.Timestamp != readString){
				currentHardwareInfo.TimestampUpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.TimestampUpdateFlag = L"";
			}
		}else if(readString == L"hw.ToolVersion"){
			freadWString(fp, readString);
			if(currentHardwareInfo.ToolVersion != readString){
				currentHardwareInfo.ToolVersionUpdateFlag = L"UPDATE";
			}else{
				currentHardwareInfo.ToolVersionUpdateFlag = L"";
			}
		}
	}

	fclose(fp);


	return 0;

}
//wstring���t�@�C���ɏ������ފ֐�
long InventoryManager::fwriteWString(FILE* fp, wstring string){

	//https://stackoverflow.com/questions/6975094/need-explanation-of-reading-and-writing-of-wchar-t-to-binary-file
	wchar_t* wc = (wchar_t*)string.c_str();
	unsigned int len = wcslen(wc) + 1;
	//�܂����ʕ�������������
	fwrite(&inventorySaveStringKey, sizeof(unsigned int), 1, fp);
	//������̒�������������
	fwrite(&len, sizeof(unsigned int), 1, fp);
	//�����ĕ��������������
	fwrite(wc, len, sizeof(wchar_t), fp);
	//fwrite(&testString, sizeof(testString) ,1 , fp);

	//MessageBox(NULL,wc,L"test",MB_OK);
	return 0;
}
//wstring���t�@�C������ǂݍ��ފ֐�
long InventoryManager::freadWString(FILE* fp, wstring &string){

	//https://stackoverflow.com/questions/6975094/need-explanation-of-reading-and-writing-of-wchar-t-to-binary-file
	unsigned int len;
	size_t r;
	//�܂����ʕ�����ǂݎ��
	r = fread(&len, sizeof(len), 1, fp);
	//�t�@�C���I�[���ǂݍ��݃G���[���I��
	if(r == 0){
		if(debugLevel > 1){
			PrintDebugLog(L"freadWString: fread return 0.\n");
		}
		return -1;
	}
	//���ʕ������Ⴄ���������t�@�C���ł͂Ȃ����I��
	if(len != inventorySaveStringKey){
		if(debugLevel > 1){
			PrintDebugLog(L"freadWString: inventorySaveStringKey unmatch.\n");
		}
		return -1;
	}
	//if(debugLevel > 1){
	//	PrintDebugLog(L"freadWString: inventorySaveStringKey match.\n");
	//}
	//������̒�����ǂݍ���
	r = fread(&len, sizeof(len), 1, fp);
	//�t�@�C���I�[���ǂݍ��݃G���[���I��
	if(r == 0){
		if(debugLevel > 1){
			PrintDebugLog(L"freadWString: fread return 0.\n");
		}
		return -1;
	}
	//�������ǂݍ���
	wchar_t* wc = new wchar_t[len];
	fread(wc, len, sizeof(wchar_t), fp);
	//�t�@�C���I�[���ǂݍ��݃G���[���I��
	if(r == 0){
		if(debugLevel > 1){
			PrintDebugLog(L"freadWString: fread return 0.\n");
		}
		return -1;
	}
	//MessageBox(NULL,wc,L"test2",MB_OK);
	string = wc;
	return 1;

}



//�O��擾���������擾����֐�
//�f�t�H���g���͒l�����[�h����֐�
long InventoryManager::loadDefaultSetting(){

	//�f�t�H���g���͒l�̃��[�h
	FILE* fp = NULL;
	//errno_t err_no;

	//�G���R�[�h�̓f�t�H���g�iSJIS�j
	//�t�@�C���ǂݍ��݃T���v���@http://pg-sample.sagami-ss.net/?eid=17
	//err_no = ;
	//�t�@�C���I�[�v���ɐ��������ꍇ
	if (_wfopen_s( &fp, defaultCSVPath.c_str(), L"r" ) == 0){
		//MessageBox(NULL,L"read", VER_STR_PRODUCTNAME, MB_OK);
		wchar_t identifyingNumber[64];
		wchar_t computerName[64];
		wchar_t ipAddress[64];
		wchar_t macAddress[64];
		wchar_t hardwareNo[1024];
		wchar_t input1[1024];
		wchar_t input2[1024];
		wchar_t input3[1024];
		wchar_t input4[1024];
		wchar_t input5[1024];
		wchar_t input6[1024];
		wchar_t buf[4096];
		wstring identifyingNumberStr;
		wstring computerNameStr;
		wstring ipAddressStr;
		wstring macAddressStr;
		wstring hardwareNoStr;
		wstring input1Str;
		wstring input2Str;
		wstring input3Str;
		wstring input4Str;
		wstring input5Str;
		wstring input6Str;
		wstring bufStr;
		int match;
		while (fgetws(buf, sizeof(buf), fp) != NULL) {
		//while (fwscanf_s(fp, L"%[^,],%[^,],%[^,],%[^,],%[^,],%[^,],%[^,],%[^,],%[^,],%[^,],%s", identifyingNumber, 64, computerName, 64, ipAddress, 64, macAddress, 64, hardwareNo, 1024, input1, 1024, input2, 1024, input3, 1024, input4, 1024, input5, 1024, input6, 1024) == 11){
		//while (fwscanf_s(fp, L"%[^,],%[^,],%[^,],%[^,],%s", identifyingNumber, 64, computerName, 64, ipAddress, 64, macAddress, 64, hardwareNo, 1024) == 3){
		//while (fwscanf_s(fp, L"%s", identifyingNumber, 64) == 3){
			//MessageBox(NULL,identifyingNumber, VER_STR_PRODUCTNAME, MB_OK);
			//MessageBox(NULL,hardwareNo, VER_STR_PRODUCTNAME, MB_OK);
			//MessageBox(NULL, buf, VER_STR_PRODUCTNAME, MB_OK);
			match = 0;
			bufStr = buf;
			//csv���̈��p���͂Ƃ肠��������
			bufStr = Replace(bufStr, L"\"", L"");
			//���g���Ȃ��J���}������scan�œǂݍ��߂Ȃ��̂ŁA����L�������ށ����Ƃŏ���
			bufStr = Replace(bufStr, L",", L"\a,\a");
			//MessageBox(NULL, bufStr.c_str(), VER_STR_PRODUCTNAME, MB_OK);
			if(swscanf_s(bufStr.c_str(), L"%[^,],%[^,],%[^,],%[^,],%[^,],%[^,],%[^,],%[^,],%[^,],%[^,],%[^,]", identifyingNumber, 64, computerName, 64, ipAddress, 64, macAddress, 64, hardwareNo, 1024, input1, 1024, input2, 1024, input3, 1024, input4, 1024, input5, 1024, input6, 1024) == 11){
				identifyingNumberStr = identifyingNumber;
				computerNameStr = computerName;
				ipAddressStr = ipAddress;
				macAddressStr = macAddress;
				hardwareNoStr = hardwareNo;
				input1Str = input1;
				input2Str = input2;
				input3Str = input3;
				input4Str = input4;
				input5Str = input5;
				input6Str = input6;
				
				identifyingNumberStr = Replace(identifyingNumberStr, L"\a", L"");
				computerNameStr = Replace(computerNameStr, L"\a", L"");
				ipAddressStr = Replace(ipAddressStr, L"\a", L"");
				macAddressStr = Replace(macAddressStr, L"\a", L"");
				macAddressStr = Replace(macAddressStr, L":", L"");
				hardwareNoStr = Replace(hardwareNoStr, L"\a", L"");
				input1Str = Replace(input1Str, L"\a", L"");
				input2Str = Replace(input2Str, L"\a", L"");
				input3Str = Replace(input3Str, L"\a", L"");
				input4Str = Replace(input4Str, L"\a", L"");
				input5Str = Replace(input5Str, L"\a", L"");
				input6Str = Replace(input6Str, L"\a", L"");
				//MessageBox(NULL,identifyingNumberStr.c_str(), VER_STR_PRODUCTNAME, MB_OK);
				//MessageBox(NULL,macAddressStr.c_str(), VER_STR_PRODUCTNAME, MB_OK);
				//MessageBox(NULL,input6Str.c_str(), VER_STR_PRODUCTNAME, MB_OK);

				//�f�t�H���g�l���Ȃ��Ƃ��̓t���O�͂��̂܂�
				if(identifyingNumberStr == L""){
					//match += 0;
				//�f�t�H���g�l������A���C���x���g���Ɠ������Ƃ��̓t���O�𗧂Ă�
				}else if(identifyingNumberStr == currentHardwareInfo.IdentifyingNumber){
					//MessageBox(NULL,identifyingNumberStr.c_str(), VER_STR_PRODUCTNAME, MB_OK);
					match += 1;
				//�f�t�H���g�l������A���C���x���g���ƈقȂ�Ƃ��̓t���O������
				}else{
					match = -1;
				}
				if(computerNameStr == L""){
					//match += 0;
				}else if(compareToIgnoreCase(computerNameStr, currentHardwareInfo.ComputerName)){
					//MessageBox(NULL,computerNameStr.c_str(), VER_STR_PRODUCTNAME, MB_OK);
					match += 1;
				}else{
					match = -1;
				}
				if(ipAddressStr == L""){
					//match += 0;
				}else if(ipAddressStr == currentHardwareInfo.IPAddress){
					//MessageBox(NULL,ipAddressStr.c_str(), VER_STR_PRODUCTNAME, MB_OK);
					match += 1;
				}else{
					match = -1;
				}
				if(macAddressStr == L""){
					//match += 0;
				}else if(compareToIgnoreCase(macAddressStr, currentHardwareInfo.MACAddress)){
					//MessageBox(NULL,macAddressStr.c_str(), VER_STR_PRODUCTNAME, MB_OK);
					match += 1;
				}else{
					match = -1;
				}

				if(match > 0){
					//MessageBox(NULL, hardwareNoStr.c_str(), VER_STR_PRODUCTNAME, MB_OK);
					defaultHardwareInfo.HardwareNoValue = hardwareNoStr;
					defaultHardwareInfo.Value1 = input1Str;
					defaultHardwareInfo.Value2 = input2Str;
					defaultHardwareInfo.Value3 = input3Str;
					defaultHardwareInfo.Value4 = input4Str;
					defaultHardwareInfo.Value5 = input5Str;
					defaultHardwareInfo.Value6 = input6Str;
					break;
				}

			}
		}
		fclose(fp);
	}
	

	return 0;
}
//�f�t�H���g���͒l�ݒ�t�@�C�������l�[������֐�
long InventoryManager::cleanupDefaultSetting(){

	wstring renameCSVPath = defaultCSVPath + L".bak";

	//���ɂ���default.csv.bak���폜���Ă���
	DeleteFile(renameCSVPath.c_str());

	//���l�[��
	if(_wrename(defaultCSVPath.c_str(), renameCSVPath.c_str()) == 0){
		return 0;
	}else{
		return -1;
	}
	
}
wstring InventoryManager::getPreviousHardwareNoValue(){
	//MessageBox(NULL,previousHardwareInfo.HardwareNoValue.c_str(),L"test2",MB_OK);
	if(defaultHardwareInfo.HardwareNoValue != L""){
		return defaultHardwareInfo.HardwareNoValue;
	}else{
		return previousHardwareInfo.HardwareNoValue;
	}
}
wstring InventoryManager::getPreviousValue1(){
	if(defaultHardwareInfo.Value1 != L""){
		return defaultHardwareInfo.Value1;
	}else{
		return previousHardwareInfo.Value1;
	}
}
wstring InventoryManager::getPreviousValue2(){
	if(defaultHardwareInfo.Value2 != L""){
		return defaultHardwareInfo.Value2;
	}else{
		return previousHardwareInfo.Value2;
	}
}
wstring InventoryManager::getPreviousValue3(){
	if(defaultHardwareInfo.Value3 != L""){
		return defaultHardwareInfo.Value3;
	}else{
		return previousHardwareInfo.Value3;
	}
}
wstring InventoryManager::getPreviousValue4(){
	if(defaultHardwareInfo.Value4 != L""){
		return defaultHardwareInfo.Value4;
	}else{
		return previousHardwareInfo.Value4;
	}
}
wstring InventoryManager::getPreviousValue5(){
	if(defaultHardwareInfo.Value5 != L""){
		return defaultHardwareInfo.Value5;
	}else{
		return previousHardwareInfo.Value5;
	}
}
wstring InventoryManager::getPreviousValue6(){
	if(defaultHardwareInfo.Value6 != L""){
		return defaultHardwareInfo.Value6;
	}else{
		return previousHardwareInfo.Value6;
	}
}


/*
//�\�t�g�E�F�A���̍\���̂�񋓂���֐�
softwareInfo InventoryManager::getNextSoftware(){
	softwareInfo blankSoftwareInfo;
	return blankSoftwareInfo;
}
*/
//main
int _tmain(int argc, _TCHAR* argv[]){
    //GUI�ł͂Ƃ肠�����R�����g�A�E�g
 /*
	//���P�[���������ݒ肵�Ă���
	//�Q�l http://cx5software.com/article_vcpp_unicode/
	//setlocale�֐��̑������Ɂu""�v���w�肷��ƁAOS�̃f�t�H���g���P�[�����w�肳��܂��B�Ⴆ�΁A���{��ɐݒ肳��Ă���ꍇ�uJapanese_Japan.932�v���ݒ肳��܂��B���R�Ȃ���AOS�̐ݒ肪�ς��ƁA����ɍ��킹�ĕύX����܂��̂Œ��ӂ��K�v�ł��B
	_tsetlocale(LC_ALL, _T(""));
	//_tsetlocale(LC_ALL, L"Japanese");

    //Unicode��������R���\�[���ɕ\��������
    //_tsetlocale(LC_ALL, _tsetlocale(LC_CTYPE, L""));

	wstring hKey = L"HKEY_LOCAL_MACHINE";
    wstring subKey = L"";
	BOOL displayHelp = FALSE;
	

	//�����̉��
    int i;
    for(i = 0; i < argc; ++i){
        //_tprintf(TEXT("argv[%d]=%s\n"), i, argv[i]);
		if((_tcscmp(StrToLower(argv[i]), L"/hkey") == 0) && (i + 1 < argc)){
			++i;
			hKey = argv[i];
			//StrToUpper(hKey);
		}else if((_tcscmp(StrToLower(argv[i]), L"/subkey") == 0) && (i + 1 < argc)){
			++i;
			subKey = argv[i];
		}else if((_tcscmp(StrToLower(argv[i]), L"/file") == 0) && (i + 1 < argc)){
			++i;
			outputFileName = argv[i];
		}else if((_tcscmp(StrToLower(argv[i]), L"/encode") == 0) && (i + 1 < argc)){
			++i;
			if(_tcscmp(StrToLower(argv[i]), L"utf-8") == 0){
				outputFileEncode = L"UTF-8";
			}else if(_tcscmp(StrToLower(argv[i]), L"utf8") == 0){
				outputFileEncode = L"UTF-8";
			}else if(_tcscmp(StrToLower(argv[i]), L"utf-8n") == 0){
				outputFileEncode = L"UTF-8N";
			}else if(_tcscmp(StrToLower(argv[i]), L"utf8n") == 0){
				outputFileEncode = L"UTF-8N";
			}else if(_tcscmp(StrToLower(argv[i]), L"sjis") == 0){
				outputFileEncode = L"SJIS";
			}else if(_tcscmp(StrToLower(argv[i]), L"shiftjis") == 0){
				outputFileEncode = L"SJIS";
			}
		}else if(_tcscmp(StrToLower(argv[i]), L"/debug") == 0){
			debugLevel = 1;
			++i;
		}else if(_tcscmp(StrToLower(argv[i]), L"/help") == 0){
			displayHelp = TRUE;
		}else if(_tcscmp(StrToLower(argv[i]), L"/h") == 0){
			displayHelp = TRUE;
		}else if(_tcscmp(StrToLower(argv[i]), L"/?") == 0){
			displayHelp = TRUE;
		}
    }

	_putws(L"LIP (List Installed Programs) v0.1");
	_putws(L"Copyright (C) 2011 Space Work");
	_putws(L"");

	if(displayHelp){
		//       123456789012345678901234567890123456789012345678901234567890
		_putws(L"�u�v���O�����Ƌ@�\�v�ɕ\�������\�t�g�E�F�A���̈ꗗ���o�͂��܂��B");
		_putws(L"");
		_putws(L"LIP [/file �t�@�C����] [/encode �G���R�[�h]");
		_putws(L"");
		_putws(L"  /file �t�@�C����");
		_putws(L"    �ۑ�����t�@�C�������w�肵�܂��B");
		_putws(L"    �w�肵�Ȃ��ꍇ�A�uSoftwareList.csv�v�ƂȂ�܂��B");
		_putws(L"  /encode �G���R�[�h");
		_putws(L"    �ۑ��t�@�C���̃G���R�[�h���w�肵�܂��B");
		_putws(L"    �uUTF-8�v�uUTF-8N�v�uSJIS�v�̂����ꂩ���w�肵�܂��B");
		_putws(L"    �w�肵�Ȃ��ꍇ�uUTF-8�v�ƂȂ�܂��B");
		_putws(L"");
		return 0;
	}
	if(debugLevel > 0){
		for(i = 0; i < argc; ++i){
			_tprintf(TEXT("argv[%d]= %s\n"), i, argv[i]);
		}
	}


	// /subkey���������Ƃ��́A�e�X�g�I�Ɏw�肳�ꂽHKEY�ASUBKEY�̏����擾���ăR���\�[���ɏo��
	if(subKey.size() != 0){
		RegList testRegList;
		_tprintf(TEXT("HKEY: %s\n"), hKey.c_str());
		_tprintf(TEXT("SUBKEY: %s\n"), hKey.c_str());
		testRegList.setHKeyAndSubKey(hKey, subKey, TRUE);
		testRegList.displayAll();
	}

	//���C�������@�\�t�g�E�F�A�ꗗ�̏o��
	InventoryManager testInventoryManager;
	testInventoryManager.listUp();
	testInventoryManager.displayAll();

	_putws(L"");
	_tprintf(TEXT("�\�t�g�E�F�A�̈ꗗ���u%s�v�ɕۑ����܂����B\n"), outputFileName.c_str());

	_putws(L"�I������ɂ͉����L�[�������Ă��������B");
	//getchar();
	//����getchar()���邾���ł̓G���^�[�L�[�����󂯕t���Ȃ�
	//http://stope.blog130.fc2.com/blog-entry-36.html
	//�ȉ��̌x�����o���̂�kbhit()��_kbhit()�ɕύX
	//warning C4996: 'kbhit': The POSIX name for this item is deprecated. Instead, use the ISO C++ conformant name: _kbhit. See online help for details.
	for (i = 0; ; i++) {
		if (_kbhit()!=0) {
			break;
		}
	}
*/

	return 0;
}

