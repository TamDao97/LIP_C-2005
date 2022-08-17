//これをしておかないと、mapを使うときに以下のエラーが出る
//Run-Time Check Failure #2 - Stack around the variable '_Lk' was corrupted.
//PlattformSDKのcrtにあるxtreeが読まれていることが原因らしい
//参考:http://unkar.org/r/tech/1204045410
#pragma warning(disable:4786)

//CUIアプリと違い、これを追加しないと以下のエラーが出る
//fatal error C1010: プリコンパイル ヘッダーを検索中に不明な EOF が見つかりました。'#include "stdafx.h"' をソースに追加しましたか?
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
//tramsformに必要
#include <algorithm>
#include <cctype>
#include <cstdio>
//ファイル書き出しに必要
#include <iostream>
#include <fstream>
//エンターキーを押さずに入力キーを受け付けるときに必要
#include <conio.h>
//ファイル名に日時を付けるときに必要
#include <time.h>
//ホスト名取得
#include <stdio.h>
#include <WinSock.h>
#pragma comment(lib, "wsock32.lib")
//ネットワークアダプタ取得
#include <Windows.h>
#include <Iphlpapi.h>
#include <Assert.h>
#pragma comment(lib, "iphlpapi.lib")

//WMI取得
#include <comdef.h>
//#include <comutil.h>
#include <Wbemidl.h>

//_mkdirに必要
#include <direct.h>

//WMIからCPU情報を取得するときSocketDesignationの重複除去カウントに必要
#include <set>

//ORACLE取得時のテキストファイル解析に必要
//VC2005には含まれておらずVC2010以降でないとダメ　boostを導入すると使えるがそこまでするかどうか
//#include <regex>


//UUID生成
//#include <rpc.h>
//#pragma comment(lib ,"rpcrt4.lib")
//https://stackoverflow.com/questions/24365331/how-can-i-generate-uuid-in-c-without-using-boost-library
//http://koolgeeks.seesaa.net/article/117666391.html

//これがないと名前にいちいち接頭辞"std::"をつけて 修飾しなければならない
//独自に名前空間を定義しないなら書いておく
//参考:http://www.amris.co.jp/cpp/c16.html
using namespace std;

//OS判別に追加
//http://pmakino.jp/tdiary/20090627.html
//GetVersionEx() で使われる定数 (VC++2005 未定義)
#define VER_SUITE_WH_SERVER 0x00008000
 
//GetSystemMetrics() で使われる定数 (VC++2005 未定義)
#define SM_TABLETPC     86
#define SM_MEDIACENTER  87
#define SM_STARTER      88
#define SM_SERVERR2     89
 
//GetProductInfo() で使われる定数 (VC++2005 未定義)
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
//定義漏れを追加
#define PRODUCT_MULTIPOINT_STANDARD_SERVER          0x0000004C
#define PRODUCT_MULTIPOINT_PREMIUM_SERVER           0x0000004D
#define PRODUCT_CORE                                0x00000065
#define PRODUCT_CORE_N                              0x00000062
#define PRODUCT_PROFESSIONAL_WMC                    0x00000067
//定義漏れを追加 2015-08-24
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

//_tolowerで使う
//#include <afx.h>

//WMI取得のため追加
#ifndef _WIN32_DCOM
#   define _WIN32_DCOM
#endif
//#pragma comment(lib, "wbemuuid.lib")
//↑でエラーが出る
//fatal error LNK1103: debugging information corrupt; recompile module Error executing link.exe.
//古いSDKでないとVC++2005には対応していない模様
//もう入手できないので、直接書き換える
//http://stackoverflow.com/questions/6568467/vc-compatibility-problem
GUID MY_CLSID_WbemLocator = {0x4590f811, 0x1d3a, 0x11d0, {0x89, 0x1f, 0x00, 0xaa, 0x00, 0x4b, 0x2e, 0x24}};
GUID MY_IID_IWbemLocator = {0xdc12a687, 0x737f, 0x11cf, {0x88, 0x4d, 0x00, 0xaa, 0x00, 0x4b, 0x2e, 0x24}};
 
_COM_SMARTPTR_TYPEDEF(IWbemLocator, __uuidof(IWbemLocator));
_COM_SMARTPTR_TYPEDEF(IWbemServices, __uuidof(IWbemServices));
_COM_SMARTPTR_TYPEDEF(IEnumWbemClassObject, __uuidof(IEnumWbemClassObject));
_COM_SMARTPTR_TYPEDEF(IWbemClassObject, __uuidof(IWbemClassObject));



//★グローバル変数

//OSが64bitか32bitか毎度調査せず調査した結果を保管
//-1:未調査
// 0:32bit
// 1:64bit
long isWow64Flag = -1;

//デバッグ（コンソールに詳しい情報を出力する）モードか否か
//0:デバッグモードではなく通常モード
//1:デバッグモード
long debugLevel = 0;
//保存ファイル名（あとで書き換わる）←やめ
//wstring outputFileName = L"InventoryManager.csv";
//デバッグログの保存ファイル名
wstring debugLogFileName = L"debug.log";
//デバッグログの保存ファイルのポインタ
FILE* debugFP = NULL;

//インベントリをPC内に保存する際のキーワード
unsigned int inventorySaveStringKey = 116; 

//SoftwarePicker.h用の出力設定
//保存ファイルのエンコード（デフォルトはUTF-8、その他UTF-8N、SJISも可）
wstring outputFileEncode = VER_STR_OUTPUTFILEENCODE;
//wstring outputFileEncode = L"SJIS";
//保存ファイルの形式（"SIMPLE","SARMS"）※SARMSのとき、エンコードはSJIS強制
wstring outputFileFormat = VER_STR_OUTPUTFILEFORMAT;
//wstring outputFileFormat = L"SIMPLE";
//MORE:保存ファイルがSIMPLE形式のとき、PCベンダー、PC機種、CPUタイプの列を追加する
//wstring outputFileOption = L"MORE";
wstring outputFileOption = VER_STR_OUTPUTFILEOPTION;

//動作モード（デフォルトはFALSE）
BOOL silentMode = FALSE;

//ウィンドウのラベル
wstring labelHardwareNo = VER_STR_LABEL_HARDWARE_NO;
wstring label1 = VER_STR_LABEL1;
wstring label2 = VER_STR_LABEL2;
wstring label3 = VER_STR_LABEL3;
wstring label4 = VER_STR_LABEL4;
wstring label5 = VER_STR_LABEL5;
wstring label6 = VER_STR_LABEL6;

//入力欄の有効・無効
bool isInputHardwareNoValid = true;
bool isInput1Valid = true;
bool isInput2Valid = true;
bool isInput3Valid = true;
bool isInput4Valid = true;
bool isInput5Valid = true;
bool isInput6Valid = true;

//iniファイルのパス
wstring iniPath = L".\\LIP.ini";
//iniファイルのデータのパス
wstring outputPath = L".\\";
wstring hardwareNoPrefix = L"";
unsigned int hardwareNoDigits = 0;
wstring iniProgram = L"";
wstring iniProgramArguments = L"";
unsigned int iniOracleCompatible = 0;

//default.csvファイルのパス
wstring defaultCSVPath = L".\\default.csv";



/*
参考

RegEnumValueのサンプル
http://d.hatena.ne.jp/s-kita/20120331/1333174926

RegQueryValueExのサンプル
http://d.hatena.ne.jp/s-kita/20120408/1333895565

RegEnumKeyExのサンプル
http://d.hatena.ne.jp/s-kita/20120401/1333290793

64bitOS判別のサンプル
http://msdn.microsoft.com/en-us/library/ms684139(VS.85).aspx

コマンドラインオプション解析のサンプル
http://masudahp.web.fc2.com/cl/tool/tool01.html
http://codezine.jp/article/detail/4086
http://www.opensource.apple.com/source/awk/awk-1.2/main.c
http://c-faq.com/misc/argv.html
http://sa.eei.eng.osaka-u.ac.jp/eeisa003/tani_prog/org_option.html

マルチバイト関連
http://cx5software.com/article_vcpp_unicode/
http://victreal.com/Junk/_T/

メモリリーク対策
http://www.aerith.net/cpp/safe-coding-j.html

レジストリの最大文字数等の情報
http://support.microsoft.com/kb/256986/ja

LocalizedString
http://msdn.microsoft.com/en-us/library/windows/desktop/dd374120(v=vs.85).aspx

LocalizedStringの読み込み
http://stackoverflow.com/questions/13862278/getting-string-from-a-dll-string-table
http://stackoverflow.com/questions/9324563/how-does-one-read-a-localized-name-from-the-registry

プログラムと機能の出力（VBS）
http://social.technet.microsoft.com/Forums/windows/en-US/e1599843-f5c8-45ca-9cdf-56d5e5ccf43b/addremove-programs-export-to-text-file?forum=itproxpsp


*/
/*
//最新版のWindows Platform SDK（V7.1）のMicrosoft SDKs\Windows\v7.1\Include\WinReg.hから移植
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

//デバッグログを出力
//http://program.station.ez-net.jp/special/vc/basic/function/stdarg.asp
//http://www.kde.cs.tut.ac.jp/~atsushi/?p=117
void PrintDebugLog(wstring text){
	/*
    FILE* fp = NULL;
	//BOM付きのUTF-8で出力
	//a+モードでファイル末尾に追加 http://msdn.microsoft.com/ja-jp/library/z5hh6ee9(v=vs.90).aspx
	LONG status;
	status = _wfopen_s( &fp, debugLogFileName.c_str(), L"a+,ccs=UTF-8" );
	//ファイルがちゃんと開けたか返り値をチェックしておかないと途中でエラー落ちする
	//いちいち書き込むのが問題か？バッファに溜めてから書き込んだほうがいいかどうか
	if(status == 0){
		_ftprintf(fp, L"%s", text.c_str());
		//閉じておかないと「他のプロセスで使用中」となってエラー落ち
		fclose(fp);
	}
	*/
	_ftprintf(debugFP, L"%s", text.c_str());
}

//文字列中の英大文字文字を小文字文字に変換
_TCHAR *StrToLower(_TCHAR *acString){
    for (_TCHAR* pc = acString; *pc; pc++) {
        *pc = tolower(*pc);
    }
    return (acString);
}

//文字列中の英小文字を大文字に変換
_TCHAR *StrToUpper(_TCHAR *acString){
    //_TCHAR acString[] = TEXT("aBcDe");
    for (_TCHAR* pc = acString; *pc; pc++) {
        *pc = _toupper(*pc);
    }
    //_tprintf(TEXT("%s\n"), acString);
	return (acString);
}

//文字列の大文字を小文字に変換
void toLowerCase(wstring& string){
	transform(string.begin (), string.end (), string.begin (), towlower);
}

//文字列の大文字小文字を無視した比較
bool compareToIgnoreCase(wstring string1, wstring string2){
	toLowerCase(string1);
	toLowerCase(string2);
	if(string1 == string2){
		return true;
	}else{
		return false;
	}
}


//文字列の置換
//http://www.geocities.jp/eneces_jupiter_jp/cpp1/010-055.html
//使い方
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


//半角カナを全角カナに置換
wstring HankakuKana2ZenkakuKana(wstring targetString){
	//半角、全角の参考 https://gist.github.com/punytan/379007
	wstring hankakuKana1 = L"｡｢｣､･ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝﾞﾟ";
	wstring zenkakuKana1 = L"。「」、・ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン゛゜";
	wstring hankakuKana2 = L"ｳﾞｶﾞｷﾞｸﾞｹﾞｺﾞｻﾞｼﾞｽﾞｾﾞｿﾞﾀﾞﾁﾞﾂﾞﾃﾞﾄﾞﾊﾞﾋﾞﾌﾞﾍﾞﾎﾞﾊﾟﾋﾟﾌﾟﾍﾟﾎﾟ";
	wstring zenkakuKana2 = L"ヴガギグゲゴザジズゼゾダヂヅデドバビブベボパピプペポ";
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
			//何もしない
		}else{
			//数値じゃなくなったら終了
			isNumericString = false;
			break;
		}
	}
	return isNumericString;
}
/*
//文字列中の最初に出てきた数値列を指定桁数でゼロ詰めする
wstring FirstNumericInStringZeroPadding(wstring targetString, wstring digits){
	//digitsが数字のみで構成された文字列かチェック
	//http://qiita.com/edo1z/items/da66e28e206d2b01157e
	//https://msdn.microsoft.com/en-us/library/fcc4ksh8(v=vs.80).aspx
	if(all_of(digits.cbegin(), digits.cend(), _istdigit)){
		//何もしない
	}else{
		//数字のみで構成された文字列ではない場合、"1"とする
		digits=L"1";
	}
	//sprintfSetting = L"%010d"
	wstring sprintfSetting = L"%0" + digits + L"d";

	int numericFirstIndex = -1;
	int numericLastIndex = -1;
	for(int i=0; i<(int)targetString.size(); ++i){
		//if(48 <= targetString[i].c_str()[0] && targetString[i].c_str()[0] <= 57){
		if(48 <= targetString[i] && targetString[i] <= 57){
			//初めて数値が出てきたらFirstIndexにセット
			if(numericFirstIndex == -1){
				numericFirstIndex = i;
			}
			//数値ならLastIndexにもセット
			numericLastIndex = i;
		}else{
			//数値じゃなくなったら終了
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
//文字列中の数値列を全て10桁でゼロ詰めする　エラー落ちするので直すこと！！
wstring NumericInStringZeroPadding(wstring targetString){
	int numericFirstIndex = -1;
	int numericLastIndex = -1;
	for(int i=(int)targetString.size()-1; i>=0; --i){
	//for(int i=0; i<(int)targetString.size(); ++i){
		//if(48 <= targetString[i].c_str()[0] && targetString[i].c_str()[0] <= 57){
		if(48 <= targetString[i] && targetString[i] <= 57){
			//初めて数値が出てきたらLastIndexにセット
			if(numericLastIndex == -1){
				numericLastIndex = i;
			}
			//数値ならFirstIndexにもセット
			numericFirstIndex = i;
		}else{
			//数値じゃなくなったら終了
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

//文字列が数字だけで構成されているかチェック
BOOL isNumericString(wstring targetString)
{
	for(int i=0; i<(int)targetString.size(); ++i){
		//if(48 <= targetString[i].c_str()[0] && targetString[i].c_str()[0] <= 57){
		if(48 <= targetString[i] && targetString[i] <= 57){
			//数字なのでOK
		}else{
			return FALSE;
		}
	}
	return TRUE;
}

//文字列が英数字だけで構成されているかチェック
BOOL isAlphanumericString(wstring targetString)
{
	for(int i=0; i<(int)targetString.size(); ++i){
		//if(48 <= targetString[i].c_str()[0] && targetString[i].c_str()[0] <= 57){
		if(48 <= targetString[i] && targetString[i] <= 57){
			//数字なのでOK
		}else if(65 <= targetString[i] && targetString[i] <= 90){
			//英字大文字なのでOK
		}else if(97 <= targetString[i] && targetString[i] <= 122){
			//英字小文字なのでOK
		}else{
			return FALSE;
		}
	}
	return TRUE;
}

//外部プログラムの実行
//http://tooljp.com/language/C-Languate/sample-code/CreateProcess-sample-code.html
//http://www.wabiapp.com/WabiSampleSource/windows/create_process.html
//コマンド、引数の渡し方
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
//外部プログラムを実行して結果を得る
//https://ameblo.jp/cpp-prg/entry-10232970547.html
wstring RunProcess2(wstring programString, wstring argumentsString){

	if(debugLevel > 1){
		PrintDebugLog(L"RunProcess programString=");
		PrintDebugLog(programString);
		PrintDebugLog(L"\n");
	}
	//パイプ作成
	SECURITY_ATTRIBUTES sa={0}; 
	sa.nLength = sizeof( SECURITY_ATTRIBUTES );
	sa.lpSecurityDescriptor = NULL;
	sa.bInheritHandle = TRUE;
	HANDLE hReadPipe, hWritePipe ;
	CreatePipe( &hReadPipe, &hWritePipe, &sa, 50000 );
	//起動情報設定
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

		//出力を保存
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
		//終了処理
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

//64bitOSか否かの判定
BOOL IsWow64(){
	//グローバル変数に結果が残っていればそれを使用
	if(isWow64Flag == 1){
		return TRUE;
	}else if(isWow64Flag == 0){
		return FALSE;
	//グローバル変数に結果が残っていない場合
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


//WMIの取得
wstring getWMI(wstring WMIQueryString, wstring WMIObjectString){

	//シリアル番号の取得
	//http://www.programmershare.com/1579747/ OS名取得までコード書いてあるが誤字等多く修正が手間すぎる
	//http://togarasi.wordpress.com/2009/12/25/c-%E3%81%A7-wmi-%E3%82%AF%E3%83%A9%E3%82%B9%E3%82%92%E4%BD%BF%E3%81%86/　完成度高いがエラー
	//http://msdn.microsoft.com/en-us/library/aa390423(v=vs.85).aspx

	//「サイドバイサイド構成が正しくないため、アプリケーションを開始できませんでした」エラーが表示される
	//スタティックリンクすることで解決
	//http://d.hatena.ne.jp/soappp/20110717/1310914195
	//ロジェクトのプロパティで、[C/C++]-[コード生成]-[ランタイムライブラリ]で、[マルチスレッド DLL (/MD)] の代わりに [マルチスレッド (/MT)] を選ぶ

	//_CrtDbgReportW云々のエラーが出る
	//http://dixq.net/forum/viewtopic.php?f=3&t=1228
	//デバッグ環境なら/MTではなく/MTdとすることで治る
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
		//エラーコードの判定もしておく
		//https://www.codeproject.com/Articles/19433/WMI-Sample-Get-IP-address
		if(!FAILED(hr)){ // && vtProp.bstrVal){
			//bstrがNULLのとき、wstringに変換するとエラー落ちするので判定を挟む
			//http://www.sutosoft.com/room/archives/000355.html
			//↑ではなく↓を採用
			//https://stackoverflow.com/questions/22525890/c-getting-wmi-array-data-from-the-local-computer
			if(vtProp.vt==VT_NULL || vtProp.vt==VT_EMPTY){
				//何もしない
			}else if(vtProp.vt & VT_ARRAY){
				//配列のときの処理は現時点で必要ないので何もしない　※これがないとエラー落ちする
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

//WMIからCPUに関する情報の取得
//CPUが刺さるソケット数NumberOfSocketsは取得方法調査中
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
	//Win32_Processorから情報を取得する
	//Win32_ComputerSystemからもNumberOfCores、NumberOfLogicalProcessorsが取れるが、古いWindowsだと値が無く
	//またWin32_Processorと同じ情報量しかないため使えない
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

        //値の取得
		hr = pclsObj->Get(BSTR(L"Name"), 0, &vtProp, 0, 0);
		if(!FAILED(hr)){
			if(vtProp.vt==VT_NULL || vtProp.vt==VT_EMPTY){
				//何もしない
			}else{
				_bstr_t bstr = vtProp;
				currentProcessorName = bstr;
				if(sProcessorName.str() != currentProcessorName){
					if(sProcessorName.str() != L""){
						sProcessorName << L";";
					}
					sProcessorName << currentProcessorName;
				}else{
					//同じなら追加しない
				}
			}
		}

		hr = pclsObj->Get(BSTR(L"MaxClockSpeed"), 0, &vtProp, 0, 0);
		if(!FAILED(hr)){
			if(vtProp.vt==VT_NULL || vtProp.vt==VT_EMPTY){
				//何もしない
			}else{
				_bstr_t bstr = vtProp;
				currentProcessorMaxClockSpeed = bstr;
				if(sProcessorMaxClockSpeed.str() != currentProcessorMaxClockSpeed){
					if(sProcessorMaxClockSpeed.str() != L""){
						sProcessorMaxClockSpeed << L";";
					}
					sProcessorMaxClockSpeed << currentProcessorMaxClockSpeed;
				}else{
					//同じなら追加しない
				}
			}
		}

		//CPU（プロセッサ、石）のカウント
		//https://community.spiceworks.com/topic/126129-how-to-determine-physical-processor-count-with-windows-os
		hr = pclsObj->Get(BSTR(L"NumberOfCores"), 0, &vtProp, 0, 0);
		//NumberOfCoresの取得が成功する＝hotfix済みのWMI
		if(!FAILED(hr)){
			//エントリの数がCPUの数
			iHotfixAfterNumberOfProcessors += 1;
			if(vtProp.vt==VT_NULL || vtProp.vt==VT_EMPTY){
				//
			}else{
				//NumberOfCoresの合計がコアの合計
				iHotfixAfterNumberOfCores += vtProp.intVal;
			}
			hr = pclsObj->Get(BSTR(L"NumberOfLogicalProcessors"), 0, &vtProp, 0, 0);
			if(!FAILED(hr)){
				if(vtProp.vt==VT_NULL || vtProp.vt==VT_EMPTY){
					//
				}else{
					//NumberOfLogicalProcessorsの合計が論理プロセッサの合計
					iHotfixAfterNumberOfLogicalProcessors += vtProp.intVal;
				}
			}			
		//NumberOfCoresの取得が失敗する＝hotfix前のWMI
		}else{
			hr = pclsObj->Get(BSTR(L"SocketDesignation"), 0, &vtProp, 0, 0);
			if(!FAILED(hr)){
				if(vtProp.vt==VT_NULL || vtProp.vt==VT_EMPTY){
					//
				}else{
					//SocketDesignationの重複を除去した合計がCPU数
					socketDesignationSet.insert(vtProp.bstrVal);
				}
			}			
			//エントリの数が論理プロセッサの数
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


//OS名の取得
wstring getOSName(){
	//メイン処理
	//OS名をエディション、バージョン含めて取得する
	//骨格の参考 http://www.gesource.jp/programming/bcb/77.html
	//分岐の参考 http://stackoverflow.com/questions/1268178/how-to-check-in-delphi-the-os-version-windows-7-or-server-2008-r2
	//95の参考 http://chokuto.ifdef.jp/advanced/version.html
	//NTの参考 http://www.delphimaster.net/view/7-33832/all
	//XPの参考 https://handle2.uc3m.es/nestk/trunk/deps/openni/Source/OpenNI/Win32/XnOSWin32.cpp
	//あとで気がついた詳しくてC++のサンプル http://pmakino.jp/tdiary/20090627.html
	//ここも詳しい http://ht-deko.minim.ne.jp/tech002.html
	wstring OSName = L"";

	// GetVersionEx による基礎判別
	//http://pmakino.jp/tdiary/20090627.html
	OSVERSIONINFOEX info;
	ZeroMemory(&info, sizeof(OSVERSIONINFOEX)); // 必要あるのか?←と書いてあった
	info.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);
	BOOL bOsVersionInfoEx;
	if ((bOsVersionInfoEx = GetVersionEx((OSVERSIONINFO*)&info)) == FALSE) {
		// Windows NT 4.0 SP5 以前と Windows 9x
		info.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
		if (!GetVersionEx((OSVERSIONINFO *)&info)) return FALSE;
	}

	// Windows XP 以降では GetNativeSystemInfo を、それ以外では GetSystemInfo を使う←このセクションの必要性不明
	//http://pmakino.jp/tdiary/20090627.html
	SYSTEM_INFO sysinfo;
	ZeroMemory(&sysinfo, sizeof(SYSTEM_INFO)); // 必要あるのか?←と書いてあった
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

	// Vista 以降では GetProductInfo が使える
	DWORD type;
	//DWORD dwordCSDVersion;
	if(info.dwMajorVersion >= 6){
		PGPI pGPI = (PGPI)GetProcAddress(GetModuleHandle(TEXT("kernel32.dll")), "GetProductInfo");
		pGPI(info.dwMajorVersion, info.dwMinorVersion, info.wServicePackMajor, info.wServicePackMinor, &type);
		
		//Technical Previewの判定のためにinfo.szCSDVersionを数値に変換
		//↑間違い　必要なのはinfo.dwBuildNumberなので元から数値になっている
		//dwordCSDVersion = wcstod(info.szCSDVersion, _T('\0'));
	}

	//Windows8.1でGetVersionEx()がWindows8に偽装した値を返す件

	//対策1 shell32.dllは正しい値を返すのでshell32.dllに問い合わせる
	//http://www.all.undo.jp/asr/Ver35-4/10.html
	//DLLVERSIONINFO dvi;
	//if(info.dwMajorVersion >= 6){
	//	DLLGETVERSIONINFOPROC pDGV = (DLLGETVERSIONINFOPROC)GetProcAddress(GetModuleHandle(TEXT("shell32.dll")), "DllGetVersion");
	//	dvi.cbSize = sizeof(dvi);
	//	pDGV(&dvi);
	//}
	//↑遅延読み込みの形で書こうとしたがエラー落ちするので諦めて対策2を採用

	//対策2 Windows8.1に対応していることを示すマニフェストを埋め込む（MSが推奨する方法）
	//http://msdn.microsoft.com/en-us/library/windows/desktop/dn302074(v=vs.85).aspx
	//http://www.inasoft.org/talk/h201310a.html
	//マニフェストの埋め込み方法
	//http://www.g-ishihara.com/vc_wi_01.htm
	//マニフェストのexe名は、マニフェストと実際のexe名が違っていても問題ないっぽい

	//検証用に値を表示する
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

	//ここからOSのバージョン、エディション判断
	if(info.dwPlatformId == VER_PLATFORM_WIN32s){
		OSName += L"Microsoft Win32s";
	// Windows 9x系
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
		//このセクション最初に移動
/*
		//2000以降で詳しい判別を行うためにOSVERSIONINFOEXを取る
		//OSVERSIONINFOEXは2000以降でないと呼び出しエラーになる
		//http://chokuto.ifdef.jp/advanced/version.html
		OSVERSIONINFOEX infoex;
		ZeroMemory(&infoex, sizeof(OSVERSIONINFOEX));
		infoex.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);
		
		//64bitOSか判別する
		SYSTEM_INFO sysinfo = {0};
		GetSystemInfo(&sysinfo);
*/
		if(info.dwMajorVersion == 3){
			OSName += L"Windows NT";

			//NTのエディションはレジストリを見ないといけない
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
		//wikipediaとかmicrosoftサイトとか書き方がバラバラすぎる
		//以下のサポートライフサイクルのページの書き方とする
		//http://support.microsoft.com/lifecycle/?LN=en-us&p1=3194&x=2&y=14
		//http://support.microsoft.com/lifecycle/?LN=en-us&p1=3188&x=12&y=11
		//以下は判別方法が見つからなかったので未対応
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
		//2000のエディション
		//http://support.microsoft.com/lifecycle/search/?sort=PN&alpha=Windows+2000&Filter=FilterNO
		//ただ、Windows 2000 Professional EditionはWindows 2000 Professionalとする
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
		//XPのエディション
		//http://support.microsoft.com/lifecycle/?c2=1173
		//★★未実装★★Kエディションの取り方 http://support.microsoft.com/kb/922474
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
			//2003のエディション
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
		//Vistaまたは2008
		}else if(info.dwMajorVersion == 6 && info.dwMinorVersion == 0){
			//Vistaのエディション
			//http://support.microsoft.com/lifecycle/?c2=11732
			//韓国版としてKエディション、KNエディションがあるらしいがライフサイクル表にはない。また判別資料も見つからない。
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
				//VistaにはHome Preminum Nはない  Home BasicとBusinessだけ
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
			//2008のエディション
			//http://support.microsoft.com/lifecycle/search/?sort=PN&alpha=Windows+Server+2008&Filter=FilterNO
			//↑だと足りない
			//http://support.microsoft.com/lifecycle/search/default.aspx?sort=PN&alpha=Server+2008&Filter=FilterNO
			//ライフサイクル表では64と32の区別がないが区別を付ける
			//★★未実装★★ Windows Storage Server 2008 Basicの判別方法が見つからない…typeでもbasicっぽいものがない
			//以下のシリーズがある
			//Windows Storage Server 2008 Basic
			//Windows Storage Server 2008 Basic 32-bit
			//Windows Storage Server 2008 Basic Embedded
			//Windows Storage Server 2008 Basic Embedded 32-bit
			//以下も未実装
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
					//製品名にEntrepriseが付かないがtypeにはEntrepriseが付く
					//参考 http://www.microsoft.com/en-us/download/details.aspx?id=12794
					OSName += L" for Itanium-based Systems";
				}else if(type == PRODUCT_SERVER_FOR_SMALLBUSINESS){
					//多分64bit版しかない
					//参考 http://en.wikipedia.org/wiki/Windows_Server_Essentials
					//Windows Small Business Server 2008 is only available for the x86-64 (64-bit) architecture
					OSName += L" for Windows Essential Server Solutions";
				}else if(type == PRODUCT_SERVER_FOR_SMALLBUSINESS_V){
					//ライフサイクル表は
					//Windows Server 2008 for Windows Essential Server Solutions without Hyper-V
					//http://msdn.microsoft.com/en-us/library/windows/desktop/ms724358(v=vs.85).aspx は語順が逆で
					//Windows Server 2008 without Hyper-V for Windows Essential Server Solutions
					//ライフサイクル表を採用
					OSName += L" for Windows Essential Server Solutions without Hyper-V";
				}else if(type == PRODUCT_SERVER_FOUNDATION){
					//64bit版しかない
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
					//64bit版しかない
					//参考 http://en.wikipedia.org/wiki/Windows_Server_Essentials
					//Windows Small Business Server 2008 is only available for the x86-64 (64-bit) architecture
					OSName = L"Windows Small Business Server 2008 Premium";
				}else if(type == PRODUCT_SMALLBUSINESS_SERVER){
					//64bit版しかない
					//参考 http://en.wikipedia.org/wiki/Windows_Server_Essentials
					//Windows Small Business Server 2008 is only available for the x86-64 (64-bit) architecture
					OSName = L"Windows Small Business Server 2008 Standard";
				//http://en.wikipedia.org/wiki/Windows_Essential_Business_Server_2008
				//http://msdn.microsoft.com/en-us/library/windows/desktop/ms724358(v=vs.85).aspx
				//http://support.microsoft.com/lifecycle/search/default.aspx?sort=PN&alpha=Essential+Business+Server&Filter=FilterNO
				}else if(type == PRODUCT_MEDIUMBUSINESS_SERVER_MANAGEMENT || type == PRODUCT_MEDIUMBUSINESS_SERVER_MESSAGING || type == PRODUCT_MEDIUMBUSINESS_SERVER_SECURITY){
					//ライフサイクル表には無印とStandardがあるが、どう区別付けるかわけわかめ
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
					//typeの説明が「Server Hyper Core V」とおかしいが傾向から「without Hyper-V」と判断
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
		//7または2008R2(MultiPoint Server 2010/2011、Home Server 2011含む)
		}else if(info.dwMajorVersion == 6 && info.dwMinorVersion == 1){
			//7のエディション
			//http://support.microsoft.com/lifecycle/search/?sort=PN&alpha=Windows+7&Filter=FilterNO
			//日本語版wikipediaではNエディションはHome Premium、Professional、Enterprise、Ultimateとなっている
			//http://ja.wikipedia.org/wiki/Microsoft_Windows_7
			//MSのページではStarter, Home Premium, Professional, Enterprise, and Ultimateとなっている（Starterがプラス）
			//http://windows.microsoft.com/en-us/windows7/products/what-is-windows-7-n-edition
			//ライフサイクル表ではHome PremiumにNエディションがない…
			//http://support.microsoft.com/lifecycle/search/?sort=PN&alpha=Windows+7&Filter=FilterNO
			//typeのページではPRODUCT_HOME_BASIC_Nがある…
			//http://msdn.microsoft.com/en-us/library/windows/desktop/ms724358(v=vs.85).aspx
			//韓国版としてKNエディションがあるらしいがライフサイクル表にはない。未実装。
			//http://en.wikipedia.org/wiki/Windows_7_editions#Special-purpose_editions
			//ライフサイクル表では64と32の区別がないが区別を付ける
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
					//ライフサイクル表にはないが…
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

			//Windows Server 2008 R2のエディション
			//ライフサイクル表 http://support.microsoft.com/lifecycle/search/default.aspx?sort=PN&alpha=Server+2008&Filter=FilterNO
			//wikipedia http://ja.wikipedia.org/wiki/Microsoft_Windows_Server_2008_R2
			//MSページ https://www.microsoft.com/ja-jp/server-cloud/local/windows-server/2008/r2/editions/features.aspx
			//MSページとwikipediaにはFoundationがある ライフサイクル表にはない
			//Windows Server から派生したMultiPoint Serverなるものがある(#｀Д´)ﾉﾉ┻┻;
			//http://social.msdn.microsoft.com/Forums/vstudio/en-US/25e6c227-3586-47f1-aadb-298e462ddfa0/how-to-differentiate-between-windows-server-2008-r2-and-windows-multipoint-server-2011?forum=vcgeneral
			//2008R2は64bit版しかない

			//MultiPoint Server 2010/2011のエディション
			//Windows MultiPoint Server 2010
			//Windows MultiPoint Server 2010 Academic←アカデミックを判断する方法が見つからない 未実装
			//Windows MultiPoint Server 2011 Premium
			//Windows MultiPoint Server 2011 Standard
			//ライフサイクル表 http://support.microsoft.com/lifecycle/search/default.aspx?sort=PN&alpha=multipoint&Filter=FilterNO
			//wikipedia http://en.wikipedia.org/wiki/Windows_MultiPoint_Server
			//MSページ http://www.microsoft.com/ja-jp/education/multipoint.aspx

			//Home Server 2011
			}else{
				OSName += L"Windows Server 2008 R2";
				if(type == PRODUCT_DATACENTER_SERVER || type == PRODUCT_DATACENTER_SERVER_CORE){
					OSName += L" Datacenter";
				}else if(type == PRODUCT_ENTERPRISE_SERVER || type == PRODUCT_ENTERPRISE_SERVER_CORE){
					OSName += L" Enterprise";
				}else if(type == PRODUCT_ENTERPRISE_SERVER_IA64){
					//製品名にEntrepriseが付かないがtypeにはEntrepriseが付く
					//参考 http://www.microsoft.com/en-us/download/details.aspx?id=12794
					OSName += L" for Itanium-based Systems";
				}else if(type == PRODUCT_SERVER_FOUNDATION){
					//MSページとwikipediaにはFoundationがある ライフサイクル表にはない
					OSName += L" Foundation";
				}else if(type == PRODUCT_STANDARD_SERVER || type == PRODUCT_STANDARD_SERVER_CORE){
					OSName += L" Standard";
				}else if(type == PRODUCT_STORAGE_STANDARD_SERVER || type == PRODUCT_STORAGE_STANDARD_SERVER_CORE){
					//2008R2のStorage ServerにはStandardがなく、無印がある
					//typeがこれで合っているかは載っているサイトなかったので実機で検証するしかなさそう
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
		//8または2012
		}else if(info.dwMajorVersion == 6 && info.dwMinorVersion == 2){
			//8のエディション
			//エディションは、無印、Pro、Enterpriseの三種類 それぞれにNエディションがある Windows RTはとりあえず考慮しない
			//ライフサイクル表 http://support.microsoft.com/lifecycle/search/?sort=PN&alpha=Windows+8&Filter=FilterNO
			//wikipwdia http://ja.wikipedia.org/wiki/Microsoft_Windows_8
			//韓国版としてKNエディションがあるらしいがライフサイクル表にはない。未実装。
			//http://support.microsoft.com/kb/2703761/ja
			//typeの表には「Windows 8 China」があるが、検索してもよくわからないのでエディション設けない
			//http://msdn.microsoft.com/en-us/library/windows/desktop/ms724358(v=vs.85).aspx
			//ライフサイクル表では64と32の区別がないが区別を付ける
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
					//無印
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_CORE_N){
					//無印のNエディション
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
				//ProにMedia Center Packを適用するとtypeがPRODUCT_PROFESSIONAL_WMCに変わるらしい(#｀Д´)ﾉﾉ┻┻;
				//http://blog.uskanda.com/2012/11/05/windows-8-pro-media-center-product-inf/
				//Media Center Packは有償（一時期のキャンペーン中は無料）なので、別エディションとして扱う
				//MSページもそういうエディションぽく書いている http://windows.microsoft.com/ja-jp/windows-8/feature-packs
				//Pro NにMedia Center Pack入れられたらどうやって判断するかは未検証
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
			//Windows Server 2012のエディション
			//エディションはDatacenter、Standard、Essentials、Foundation
			//ライフサイクル表 http://support.microsoft.com/lifecycle/search/default.aspx?sort=PN&alpha=Server+2012&Filter=FilterNO
			//wikipedia http://ja.wikipedia.org/wiki/Microsoft_Windows_Server_2012
			//MSページ http://download.microsoft.com/download/B/F/4/BF474812-BE9E-41CE-9F5F-6C6E2F0B5B22/WS2012_Licensing-Pricing_Datasheet_ja.pdf
			//2012は64bit版しかない
			//MultiPoint Server 2012のエディション
			//Windows MultiPoint Server 2012 Premium
			//Windows MultiPoint Server 2012 Standard
			//ライフサイクル表 http://support.microsoft.com/lifecycle/search/default.aspx?sort=PN&alpha=multipoint&Filter=FilterNO
			//wikipedia http://en.wikipedia.org/wiki/Windows_MultiPoint_Server
			//MSページ http://www.microsoft.com/ja-jp/education/multipoint.aspx
			}else{
				OSName += L"Windows Server 2012";
				if(type == PRODUCT_DATACENTER_SERVER || type == PRODUCT_DATACENTER_SERVER_CORE){
					OSName += L" Datacenter";
				}else if(type == PRODUCT_SERVER_FOUNDATION){
					OSName += L" Foundation";
				}else if(type == PRODUCT_STANDARD_SERVER || type == PRODUCT_STANDARD_SERVER_CORE){
					OSName += L" Standard";
				//Windows Server 2012 Essentialsはtypeが何と出るのか丁度合致するものがなく不明
				//仮置きしておく
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
		//8.1または2012R2
		//マニフェスト書かないと8に偽装される
		//http://msdn.microsoft.com/en-us/library/windows/desktop/dn302074(v=vs.85).aspx
		//http://www.inasoft.org/talk/h201310a.html
		//マニフェストの埋め込み方法
		//http://www.g-ishihara.com/vc_wi_01.htm
		//マニフェストのexe名は、マニフェストと実際のexe名が違っていても問題ないっぽい
		}else if(info.dwMajorVersion == 6 && info.dwMinorVersion == 3){
			//8.1のエディション
			//エディションは8と同じく、無印、Pro、Enterpriseの三種類 それぞれにNエディションがある
			//ライフサイクル表 http://support.microsoft.com/lifecycle/search/?sort=PN&alpha=Windows+8&Filter=FilterNO
			//wikipwdia http://ja.wikipedia.org/wiki/Microsoft_Windows_8#Windows_8.1
			//韓国版としてKNエディションがあるらしいがライフサイクル表にはない。未実装。
			//http://support.microsoft.com/kb/2835517
			//ライフサイクル表では64と32の区別がないが区別を付ける
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
					//無印
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_CORE_N){
					//無印のNエディション
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
				//ProにMedia Center Packを適用するとtypeがPRODUCT_PROFESSIONAL_WMCに変わるらしい(#｀Д´)ﾉﾉ┻┻;
				//http://blog.uskanda.com/2012/11/05/windows-8-pro-media-center-product-inf/
				//Media Center Packは有償（一時期のキャンペーン中は無料）なので、別エディションとして扱う
				//MSページもそういうエディションぽく書いている http://windows.microsoft.com/ja-jp/windows-8/feature-packs
				//Pro NにMedia Center Pack入れられたらどうやって判断するかは未検証
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
			//2012R2のエディション
			//ライフサイクル表 http://support.microsoft.com/lifecycle/search/default.aspx?sort=PN&alpha=Server+2012&Filter=FilterNO
			//wikipedia http://ja.wikipedia.org/wiki/Microsoft_Windows_Server_2012
			//MSページ http://download.microsoft.com/download/B/F/4/BF474812-BE9E-41CE-9F5F-6C6E2F0B5B22/WS2012_Licensing-Pricing_Datasheet_ja.pdf
			//エディションは2012と同じくDatacenter、Standard、Essentials、Foundation
			//2012は64bit版しかない
			}else{
				OSName += L"Windows Server 2012 R2";
				if(type == PRODUCT_DATACENTER_SERVER || type == PRODUCT_DATACENTER_SERVER_CORE){
					OSName += L" Datacenter";
				}else if(type == PRODUCT_SERVER_FOUNDATION){
					OSName += L" Foundation";
				}else if(type == PRODUCT_STANDARD_SERVER || type == PRODUCT_STANDARD_SERVER_CORE){
					OSName += L" Standard";
				//Windows Server 2012 R2 Essentialsはtypeが何と出るのか丁度合致するものがなく不明
				//仮置きしておく
				}else if(type == PRODUCT_SERVER_FOR_SMALLBUSINESS){
					OSName += L" Essentials";
				}else{
					OSName += L" (Unknown Edition)";
				}
			}
		//Windows10 Technical Previewの途中まで
		//GUID←該当verのWindowsに同梱されているwscript.exeをリソースエディタで開いて確認
		//{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}
		//http://social.msdn.microsoft.com/Forums/azure/en-US/07cbfc3a-bced-45b7-80d2-a9d32a7c95d4/supportedos-manifest-for-windows-10?forum=windowsgeneraldevelopmentissues
		//マニフェスト書かないと8.1に偽装される
		//http://blogs.msdn.com/b/chuckw/archive/2013/09/10/manifest-madness.aspx
		//マニフェストの埋め込み方法
		//http://www.g-ishihara.com/vc_wi_01.htm
		//マニフェストのexe名は、マニフェストと実際のexe名が違っていても問題ないっぽい
		}else if(info.dwMajorVersion == 6 && info.dwMinorVersion == 4){
			//Windows10 Technical Preview build 9841から9879まで
			//http://it.srad.jp/story/14/11/23/0357207/
			//エディションはTechnical Previewの段階では無印（Pro相当）とEnterpriseの2種類
			//ライフサイクル表では64と32の区別がないが区別を付ける
			if(info.wProductType == VER_NT_WORKSTATION){
				//dwordCSDVersion
				OSName += L"Windows 10 Technical Preview";
				//Technical Preview
				//Technical PreviewのBuild 10240が製品版になった
				//http://www.tenforums.com/windows-insider/1946-download-windows-10-insider-preview-iso-file.html
				//http://tattu1902.blog136.fc2.com/blog-entry-99.html
				if(type == PRODUCT_ENTERPRISE){
					//Technical PreviewのEnterprise版はforが付く
					//http://www.microsoft.com/en-us/evalcenter/evaluate-windows-technical-preview-for-enterprise
					OSName += L" for Enterprise";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_PROFESSIONAL){
					//Technical Previewの無印版は内部的にはPro相当
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
		//Windows10（Technical Previewの途中から正式版まで）または次期サーバOS
		//GUID←該当verのWindowsに同梱されているwscript.exeをリソースエディタで開いて確認
		//{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}
		//https://msdn.microsoft.com/en-us/library/windows/desktop/dn481241(v=vs.85).aspx
		//マニフェスト書かないとVer6.2＝Windows8に偽装される
		//https://msdn.microsoft.com/en-us/library/windows/desktop/ms724833(v=vs.85).aspx
		//マニフェストの埋め込み方法
		//http://www.g-ishihara.com/vc_wi_01.htm
		//マニフェストのexe名は、マニフェストと実際のexe名が違っていても問題ないっぽい
		}else if(info.dwMajorVersion == 10 && info.dwMinorVersion == 0){
			//-----------------
			if (info.dwBuildNumber >= 22000) {
				OSName += L"Windows 11";
				///		Windows11のエディション
					//	エディションは7種類※ヨーロッパの反トラスト法に対応するため、MediaPlayer12を外したNがつくものもある
					/*Windows 11 Home
						Windows 11 Education
						Windows 11 Pro
						Windows 11 Pro Education
						Windows 11 Pro for Workstations
						Windows 11 IoT Enterprise
						Windows 11 Enterprise*/
						///	正式版で確認する事項：
						//	・Nエディションがあるものがある。
						//	ライフサイクル表 https://docs.microsoft.com/ja-jp/lifecycle/faq/windows
				if (info.wProductType == VER_NT_WORKSTATION) {
					//	dwordCSDVersion
					/*OSName += L"Windows 11";*/
					//		Windows11のBuildNumberは22000から始まっている。
					//		http://www.tenforums.com/windows-insider/1946-download-windows-10-insider-preview-iso-file.html
					//		http://tattu1902.blog136.fc2.com/blog-entry-99.html
					{
						if (type == PRODUCT_CORE) {
							//	Home相当は無印
							OSName += L" Home";
							if (sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64) {
								OSName += L" (64bit)";
							}
							else {
								OSName += L" (32bit)";
							}
						}
						else if (type == PRODUCT_CORE_N) {
							///		Home相当は無印
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
							///	Windows 10 Enterprise Eとは何かよく分からないが一応追加
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
			
			//Windows10のエディション
			//エディションは6種類
			//正式版で確認する事項：
			//・それぞれにNエディションがあるか→ある模様
			//・Pro相当ではない無印があるか→無印ではなくHome
			//・with Bingがあるか→未確認
			//・韓国版としてKNエディションがあるか http://support.microsoft.com/kb/2835517
			//ライフサイクル表 http://support.microsoft.com/lifecycle/search/?sort=PN&alpha=Windows+10&Filter=FilterNO
			//wikipwdia https://en.wikipedia.org/wiki/Windows_10_editions
			//ライフサイクル表では64と32の区別がないが区別を付ける
			if(info.wProductType == VER_NT_WORKSTATION){
				//dwordCSDVersion
				OSName += L"Windows 10";
				//Technical PreviewのBuild 10240が製品版になったが、試用版の細かいエディションを確認していないのでBuild Numberで判断するのはやめる
				//http://www.tenforums.com/windows-insider/1946-download-windows-10-insider-preview-iso-file.html
				//http://tattu1902.blog136.fc2.com/blog-entry-99.html
				//if(info.dwBuildNumber < 10240){
				if(type == PRODUCT_CORE){
					//Home相当は無印
					//OSName += L" Home";
					if(sysinfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64){
						OSName += L" (64bit)";
					}else{
						OSName += L" (32bit)";
					}
				}else if(type == PRODUCT_CORE_N){
					//Home相当は無印
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
					//Windows 10 Enterprise Eとは何かよく分からないが一応追加
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
			//Server 2016以降のエディション
			//ライフサイクル表 http://support.microsoft.com/lifecycle/search/default.aspx?sort=PN&alpha=Server+2012&Filter=FilterNO
			//wikipedia http://ja.wikipedia.org/wiki/Microsoft_Windows_Server_2012
			//MSページ http://download.microsoft.com/download/B/F/4/BF474812-BE9E-41CE-9F5F-6C6E2F0B5B22/WS2012_Licensing-Pricing_Datasheet_ja.pdf
			//エディションは2012と同じくDatacenter、Standard、Essentials、Foundation
			}else{
				//Server 2016と2019の判定はレジストリが必要
				HKEY hKey;
				TCHAR szReleaseId[9] = TEXT("");
				DWORD dwBufLen = sizeof(TCHAR) * sizeof(szReleaseId);
				LONG lRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion"), 0, KEY_QUERY_VALUE, &hKey);
				if (lRet == ERROR_SUCCESS) {
					lRet = RegQueryValueEx(hKey, TEXT("ReleaseId"), NULL, NULL, (LPBYTE)szReleaseId, &dwBufLen);
					RegCloseKey(hKey);
				}
				int ReleaseId = _wtoi(szReleaseId);
				
				//ReleaseIdで判定
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
				//Windows Server 2012 R2 Essentialsはtypeが何と出るのか丁度合致するものがなく不明
				//仮置きしておく
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

//LocalizedStringの取得
wstring getLocalizedString(wstring inputString){
	//HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\DW WLAN Card Utility
	//@C:\Windows\system32\bcmwlrc.dll,-4001
	//HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{3E29EE6C-963A-4aae-86C1-DC237C4A49FC}
	//@C:\Program Files (x86)\Intel\Intel(R) Rapid Storage Technology\Uninstall\Setup.exe,-2018

	//LocalizedStringは以下の形式で渡される
	//"@<PE-path>,-<stringID>[;<comment>]"
	//http://msdn.microsoft.com/en-us/library/windows/desktop/dd374120(v=vs.85).aspx
	//末尾の","はないことがある
	//まずは区切り文字の位置を調べる
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
	//int index2 = (int)inputString.find(L",-", 0); //元々
	int index2 = (int)inputString.find(L",-", 0); //debuglogからコピペ
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
	//区切り文字「;」がない場合（＝stringIDのあとにcommentがない場合）
	if(index3 == string::npos){
		stringID = inputString.substr(index2 + 2, inputString.size() - index2);
	//区切り文字「;」がある場合（＝stringIDのあとにcommentがある場合）
	}else{
		stringID = inputString.substr(index2 + 2, index3 - index2 - 1);
	}

	//_tprintf(TEXT("path: %s\n"), path.c_str());
	//_tprintf(TEXT("stringID: %s\n"), stringID.c_str());

	//ここから文字列のロード
	//LoadLibraryExもあるが、とりあえずLoadLibraryで
	//LoadLibraryExなら、パス指定でファイル検索順を変えられる
	//http://msdn.microsoft.com/ja-jp/library/cc429243.aspx
	//どちらがより適切かは未検証

	//64bitOSに32bitアプリケーションでアクセスしたときsystem32が開けない問題あり
	//この場合、Sysnativeにアクセスすることで回避できる
	//http://yamori-jp.blogspot.jp/2011/04/64bitwindows32bitsystem32.html
	//http://msdn.microsoft.com/ja-jp/library/aa384187(v=vs.85).aspx
	//詳細な挙動は未検証

	//64bitOSで
	if(IsWow64()){
		//比較用に小文字にする文字列を用意
		wstring pathLower(path.c_str());
		transform(pathLower.begin (), pathLower.end (), pathLower.begin (), tolower);
		//pathにsystem32が含まれているかチェック
		int indexSystem32 = (int)pathLower.find(L"\\system32\\", 0);
		//あればSysnativeに置換
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
	//バッファを超える長さの文字列は切り捨てられるのでバッファは大きめに取っておく
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
	//ソートできるように比較演算子を定義
	//http://d.hatena.ne.jp/minus9d/20130501/1367415668
	//こうしておくと
	//sort(data_array.begin(), data_array.end());のようにできる
	//最後のconstを忘れると"instantiated from here"というエラーが出てコンパイルできないので注意
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
	//ソートできるように比較演算子を定義
	//http://d.hatena.ne.jp/minus9d/20130501/1367415668
	//こうしておくと
	//sort(data_array.begin(), data_array.end());のようにできる
	//最後のconstを忘れると"instantiated from here"というエラーが出てコンパイルできないので注意
	//bool operator<( const data_t& right ) const {
		//return num == right.num ? str < right.str : num < right.num;
	//	return name < right.name;
	//}
	//http://homepage2.nifty.com/well/sort.html ←こっちを採用
	//softwareInfo::nameCompare
	//並び替えのルール（詳細は未検証）
	//英数は大文字小文字無視
	//ひらがなカタカナの違いは無視
	//漢字はUTF-8ではなくJIS順
	static bool nameCompare(const softwareInfo& left, const softwareInfo& right){
		//比較用文字列を用意
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
	//ソートできるように比較演算子を定義
	//http://d.hatena.ne.jp/minus9d/20130501/1367415668
	static bool nameCompare(const patchInfo& left, const patchInfo& right){
		//比較用文字列を用意
		wstring compareLeft = left.CompareName.c_str();
		wstring compareRight = right.CompareName.c_str();
		return compareLeft < compareRight;
	}
};


//指定されたキー直下の値とサブキーを取得するクラス
class RegList{
	private:
		//キー　　：値の名前
		//値　　　：値の中身
		//map<_TCHAR*, _TCHAR*> regValues;
		//キー　　：値の名前
		//値　　　：値のデータ型
		//map<_TCHAR*, _TCHAR*> regTypes;
		//map<_TCHAR*, _TCHAR* > data;
		//サブキーの配列
		//_TCHAR** subKeys;←宣言時にはエラーにならないが使うときexeがエラー落ちする
		//vector<_TCHAR*> subKeys;
		unsigned int subKeysIndex;
		vector<regValue> regValues;
		vector<subKey> subKeys;

    //入力
    //data[TEXT("Environment")][TEXT("ProductName")] = TEXT("Name");
	public:
		//メンバ関数の宣言
		//必要情報（主キーとサブキー）を受け取ってレジストリをスキャンする関数
		long setHKeyAndSubKey(wstring hKey, wstring subKey, BOOL scan64bit);
		//取得したデータを標準出力に出す関数
		void displayAll();
		//指定された名前の値の中身と型を返す関数
		//名前は大文字小文字無視
		regValue getValue(wstring name);
		//見つかったサブキーを列挙する関数（呼ばれる度に1つずつ参照渡しして最後まで出たら戻り値が-1）
		subKey getNextSubKey();
};

//必要情報（主キーとサブキー）を受け取ってレジストリをスキャンする関数
long RegList::setHKeyAndSubKey(wstring hKeyChar, wstring subKeyChar, BOOL scan64bit){
	subKeysIndex = 0;
	//clearしないと再度setHKeyAndSubKeyを呼び出したときに以前のデータがあっておかしくなる
	regValues.clear();
	subKeys.clear();

	HKEY hKey;
	
	//大文字にするのはエラー落ちしたことがあるのでとりあえず停止　原因追求はまだ
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
	//↓のようにすると、「間接操作のレベルが異なります」というエラーが出る
	//ULONG_PTR hKey = HKEY_LOCAL_MACHINE;
	//ただのLONGでもない
	//よく分かっていないので対策不明
	//LONG hKey = HKEY_LOCAL_MACHINE;

    //Status = RegOpenKeyEx(hKey, subKeyChar, 0, KEY_READ  | (IsWow64() ? KEY_WOW64_64KEY : 0), &phkResult);
	if(IsWow64() == TRUE){
		//64bitOSで64bitキーをスキャンする場合
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
		//64bitOSで32bitキーをスキャンする場合
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
		//32bitOSで64bitキーをスキャンする場合
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
		//newで動的に確保しないと上書きしてしまう
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

			//今から取得する値についてregValue構造体を用意する
			regValue currentRegValue;
			//currentRegValue.name = szValueNameBuffer;
			//tcharをwstringに代入
			currentRegValue.name = szValueName;
			//レジストリのキー名は大文字小文字無視なので小文字にした名前も用意
			//tolowerで比較すると、日本語などで文字化けする　towlowerを使う
			currentRegValue.lowerName = szValueName;
			transform(currentRegValue.lowerName.begin (), currentRegValue.lowerName.end (), currentRegValue.lowerName.begin (), towlower);

            switch(dwType){

            case REG_BINARY:
                //バイナリ値は、Hexダンプする
                //puts("REG_BINARY");
				//regTypes[szValueNameBuffer] = TEXT("REG_BINARY");
				//regValues[szValueNameBuffer] = TEXT("");
				currentRegValue.type = L"REG_BINARY";
				//currentRegValue.dataに空文字列をセットすると文字列を追加したタイミングでアクセス違反の例外となる
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
                //ExpandEnvironmentStringsで環境変数文字列を展開してから表示する
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
				//REG_MULTI_SZはヌル文字（\0）で区切られ、終端はヌル文字が連続（\0\0）している。英語ではdouble-null-terminated list
				//REG_SZの方法をそのまま適用すると、最初のヌル文字までしか取得できない
				//_stprintf(regDataCharBuffer, TEXT("%s"), lpData);
				//currentRegValue.data = regDataCharBuffer;

				//サイズを指定してもうまくいかず（詳細未検証）
				//swprintf(regDataCharBuffer, dwDataSize, TEXT("%s"), lpData);
				
				//wstringの初期化にそのまま渡してもうまくいかず（詳細未検証）
				//wstring regDataStrBuffer(char(lpData), 0, dwDataSize);
				//currentRegValue.data = regDataStrBuffer;

				//ヌル文字を改行に置換してから処理
				//これにより
				//36 00 38 00 41 00 42 00 36 00 37 00 43 00 41 00 37 00 44 00 41 00 37 00 46 00 46 00 46 00 46 00 35 00 32 00 30 00 35 00 41 00 37 00 43 00 38 00 30 00 34 00 31 00 30 00 31 00 30 00 36 00 31 00 00 00 31 00 32 00 33 00 33 00 00 00 00 00 
				//↓
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

			//レジストリの値データをデバッグログとして保存しようとすると時間がかかりすぎるのでレベル格上げ→戻す
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


	//指定キー配下のサブキー名を取得

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

			//subKey構造体を用意する
			subKey currentSubKey;

 			_TCHAR *subKeyNameBuffer;
			//newで動的に確保しないと上書きしてしまう
			//http://www.geocities.jp/ky_webid/cpp/language/012.html
			subKeyNameBuffer = new _TCHAR[MAX_PATH];

			_TCHAR *subKeyFullPathNameBuffer;
			subKeyFullPathNameBuffer = new _TCHAR[MAX_PATH];

			FILETIME ftLocalTime;
            SYSTEMTIME LastWriteTime;
			//newで動的に確保しないと上書きしてしまう
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

			//レジストリの列挙データをデバッグログとして保存しようとすると時間がかかりすぎるのでレベル格上げ→戻す
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
			//列挙し終わるとRegEnumKeyExはERROR_NO_MORE_ITEMSを返す
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

//取得したデータを標準出力に出す関数
void RegList::displayAll(){
	//値リストの出力
	vector<regValue>::iterator regValuesIterator = regValues.begin();  // イテレータのインスタンス化
	while(regValuesIterator != regValues.end()){
		_tprintf(TEXT("Value - Name:%s\tType:%s\tData:%s\n"), (*regValuesIterator).name.c_str(), (*regValuesIterator).type.c_str(), (*regValuesIterator).data.c_str());  // *演算子で間接参照
		++regValuesIterator;                 // イテレータを１つ進める
	}

	//サブキー名の出力
	vector<subKey>::iterator subKeysIterator = subKeys.begin();  // イテレータのインスタンス化
	while(subKeysIterator != subKeys.end()){
		_tprintf(TEXT("Subkey - Name:%s\tFullPathName:%s\tLastWriteTime:%s\n"), (*subKeysIterator).name.c_str(), (*subKeysIterator).fullPathName.c_str(), (*subKeysIterator).lastWriteTime.c_str());  // *演算子で間接参照
		++subKeysIterator;                 // イテレータを１つ進める
	}
}

//指定された名前の値の中身と型を返す関数
regValue RegList::getValue(wstring name){
//戻り値を複数返したいので、参照渡しを使っている
//http://dixq.net/forum/viewtopic.php?f=3&t=3058
//http://www.geocities.jp/ky_webid/cpp/language/015.html

	//小文字にする
	//tolowerで比較すると、日本語などで文字化けする　towlowerを使う
	transform(name.begin (), name.end (), name.begin (), towlower);

	vector<regValue>::iterator regValuesIterator = regValues.begin();  // イテレータのインスタンス化
	while(regValuesIterator != regValues.end()){
		//_tprintf(TEXT("Name:%s\tType:%s\tData:%s\n"), (*regValuesIterator).name, (*regValuesIterator).type, (*regValuesIterator).data);  // *演算子で間接参照
		//if(_tcscmp(name, (*regValuesIterator).name.c_str()) == 0){
		if(name == (*regValuesIterator).lowerName){
			//*value = (*regValuesIterator).data.c_str();
			//*type = (*regValuesIterator).type;
			//_tprintf(TEXT("value: %s\n"), value);
			//合致するものがあれば0を返して終了
			return (*regValuesIterator);
		}
		// イテレータを１つ進める
		++regValuesIterator;
	}

/*

	//for ( vector<regValue>::iterator position = a.begin();
	//	(position = find_if(position, a.end(), &age5)) != a.end(); // ココで検索
	//	++position ) {
	//	cout << *position << endl;
	//}

	//_tprintf(TEXT("name: [%s]\n"), name);
	map<_TCHAR*, _TCHAR*>::iterator regValuesIterator = regValues.begin();
	while(regValuesIterator != regValues.end()){
		//引数として指定された値の名前と、regValuesに保管されている値の名前とを比較
		//文字配列なので.findは使えない
		if(_tcscmp(name, regValuesIterator->first) == 0){
			*value = regValuesIterator->second;
			//_tprintf(TEXT("value: %s\n"), value);
			//型リストから該当するものを調査
			map<_TCHAR*, _TCHAR*>::iterator regTypesIterator = regTypes.find(regValuesIterator->first);
			if (regTypesIterator != regTypes.end()) {
				*type = regTypesIterator->second;
				//_tprintf(TEXT("type: %s\n"), type);
			}else{
				//値があるのに型が保存されていない場合は-2を返す
				return -2;
			}
			//合致するものがあれば0を返して終了
			return 0;
		}
		++regValuesIterator;
	}
*/
	//値が見つからない場合は-1を返す
	//return -1;
	regValue regValue;
	return regValue;
}

//見つかったサブキーを列挙する関数（呼ばれる度に1つずつ参照渡しして最後まで出たら戻り値が-1）
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

//ソフトウェア名などのリストを取得するクラス
class InventoryManager{
	private:
		//キー　　：値の名前
		//値　　　：値の中身
		//map<_TCHAR*, _TCHAR*> regValues;
		//キー　　：値の名前
		//値　　　：値のデータ型
		//map<_TCHAR*, _TCHAR*> regTypes;
		//map<_TCHAR*, _TCHAR* > data;
		//サブキーの配列
		//_TCHAR** subKeys;←宣言時にはエラーにならないが使うときexeがエラー落ちする
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


    //入力
    //data[TEXT("Environment")][TEXT("ProductName")] = TEXT("Name");
	public:
		//コンストラクタ
		InventoryManager(){

			if(debugLevel > 1){
				PrintDebugLog(L"InventoryManager::constructor\n");
			}

			//現在日時の取得（出力ファイルに書き込む）
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

		//メンバ関数の宣言
		//入力された情報を保管する関数
		long setInputInfo(wstring editHardwareNo, wstring edit1, wstring edit2, wstring edit3, wstring edit4, wstring edit5, wstring edit6);
		//ハードウェアの必要情報（コンピュータ名、ユーザ名、IP、MAC等）を取得する関数
		long getHardwareInfo();
		//ソフトウェア名を調べる関数（ラッパ）
		long listUp();
		//指定されたレジストリからソフトウェア名を調べる関数（本体）
		long scanRegistry(BOOL scan64bit, BOOL scanHKCU);
		//WMIからパッチ情報を取得する関数（本体）
		long getPatchFromWMI();
		//内容を全出力する関数
		wstring output();
		wstring outputSARMS();
		wstring outputSIMPLE();
		wstring outputAdvancedManager();

		//ソフトウェア名の構造体を列挙する関数
		//softwareInfo getNextSoftware();
		//インベントリを保存する関数
		long save();
		//前回のインベントリをロードする関数
		long load();
		//wstringをファイルに書き込む関数
		long fwriteWString(FILE* fp, wstring string);
		//wstringをファイルから読み込む関数
		long freadWString(FILE* fp, wstring &string);
		//前回取得した情報を取得する関数
		//default.csvから初期値をロードする関数
		long loadDefaultSetting();
		//default.csvをリネームする関数
		long cleanupDefaultSetting();
		wstring getPreviousHardwareNoValue();
		wstring getPreviousValue1();
		wstring getPreviousValue2();
		wstring getPreviousValue3();
		wstring getPreviousValue4();
		wstring getPreviousValue5();
		wstring getPreviousValue6();

};

//入力された情報を保管する関数
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

//ハードウェアの必要情報（コンピュータ名、ユーザ名、IP、MAC等）を取得する関数
long InventoryManager::getHardwareInfo(){

	if(debugLevel > 1){
		PrintDebugLog(L"InventoryManager::getHardwareInfo() ->\n");
	}

	//コンピュータ名の取得
	//IPアドレスはこの方法で取ると取れないことがあるので別の方法で
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
			/*IPアドレスはこの方法で取ると取れないことがあるので別の方法で　ログも化ける（未対策）
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

	//MACアドレス、IPアドレスの取得
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
		
		//MACアドレス取得　ログ出力時のためにここで取得　
		tempMACAddress = L"";
		for (UINT i = 0; i < pAdapter->AddressLength; i++) {
			//MACアドレスをハイフン区切りしない
			//if (i == (pAdapter->AddressLength - 1)){
			//	swprintf(MACAddressPart, 10, L"%.2X",(int)pAdapter->Address[i]);
			//}
			//else{
			//	swprintf(MACAddressPart, 10, L"%.2X-",(int)pAdapter->Address[i]);
			//}
			swprintf(MACAddressPart, 10, L"%.2X",(int)pAdapter->Address[i]);
			tempMACAddress += MACAddressPart;

		}
		//IPアドレス取得
		//tempIPAddress = pAdapter->IpAddressList.IpAddress.String;
		//charをまずtcharに変換してからwstringに代入
		mbstowcs(wIPAddress, pAdapter->IpAddressList.IpAddress.String, sizeof(pAdapter->IpAddressList.IpAddress.String));
		tempIPAddress = wIPAddress;

		//その他取得
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

		//pAdapter->Indexが一番小さいものがPrimary Adapter！
		if(pAdapter->Index < tempAdapterIndex){
			if(debugLevel > 1){
				PrintDebugLog(L" Index < tempIndex\n");
			}

			//無線LANで0.0.0.0になる暫定対策
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

	//現在のユーザ名取得
	//http://okwave.jp/qa/q2597801.html
	//http://jehupc.exblog.jp/13175145/

	//_TCHAR wuserName[1024];などと長さによって文字化けしたり空文字列になったりする
	//524、64でもダメ
	//MyDialog.cppから呼び出すとなぜか成功する
	//DWORDに初期値を設定すると大丈夫という記事を見つける
	//http://www.geocities.jp/midorinopage/Tips/getusername.html
	//確かにこれで大丈夫になった…
	_TCHAR wuserName[1024];
	DWORD dwUserSize = 1024; // 取得したユーザ名の文字列の長さ
	if ( ! GetUserName(wuserName,&dwUserSize) ){
		//return -1;
	}
	//MessageBox(NULL,wuserName,_T("ユーザ名"),MB_OK);
	currentHardwareInfo.UserName = wuserName;

	//OSの取得
	currentHardwareInfo.OSName = getOSName();

	//機種、ベンダーの取得　×レジストリからWMIに切り替え
	
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

	//仮想環境か否か
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

	//CPU名の取得　レジストリからWMIに切り替え
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

	//CPU関連情報の取得
	getCPUInfo(
		currentHardwareInfo.ProcessorName,
		currentHardwareInfo.ProcessorMaxClockSpeed,
		currentHardwareInfo.NumberOfProcessors,
		currentHardwareInfo.NumberOfCores,
		currentHardwareInfo.NumberOfLogicalProcessors
	);

	//シリアル番号の取得
	currentHardwareInfo.IdentifyingNumber = getWMI(L"SELECT * FROM Win32_ComputerSystemProduct", L"IdentifyingNumber");

	//メモリ容量の取得
	currentHardwareInfo.TotalPhysicalMemory = getWMI(L"SELECT * FROM Win32_ComputerSystem", L"TotalPhysicalMemory");

	//wstring test = getWMI(L"SELECT * FROM Win32_ComputerSystem", L"Model");
	//MessageBox(NULL,test.c_str(),L"test",MB_OK);

	return 0;

}

//ソフトウェア名を調べる関数（ラッパ）
long InventoryManager::listUp(){

	if(debugLevel > 1){
		PrintDebugLog(L"InventoryManager::listUp() ->\n");
	}
	//clearしないと再度listUpを呼び出したときに以前のデータがあっておかしくなる
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
	//64bitOSなら64bitキーも調べる
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
		//HKCUは32bitでも64bitでも変わらないので調べない
		//status = scanRegistry(TRUE, TRUE);
		//if(status != 0){
		//	if(debugLevel > 0){
		//		fwprintf_s(stderr, TEXT("scanRegistry(scan64bit=TRUE, scanHKCU=TRUE) error. status:%d\n"),status);
		//	}
		//}
	}


	//WMIからWindowsパッチ情報の取得
	//WMIから取得すると遅い
	//Windows Update Agent API(WUI)経由で取る方法もあるが、一致するものが取れない？
	//結局WMIがおすすめらしい
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

	//softwareInfosのソート
	//ソート用フィールド(compareName)の準備
	vector<softwareInfo>::iterator softwareInfoIterator = currentSoftwareInfos.begin();  // イテレータのインスタンス化
	while(softwareInfoIterator != currentSoftwareInfos.end()){
		//比較用文字列を用意
		wstring compareName = (*softwareInfoIterator).Name.c_str();

		//一文字ずつ切り出してコードポイントを付けておく
		//（制御コードが今後混じってきたときに確認するため）
		//一文字ずつ切り出す方法 http://d.hatena.ne.jp/minus9d/20120517/1337265190
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

		//wstring compareName = L"にほんご日本語123１２３カタカナ";
		//比較用文字列を小文字にする
		//tolowerで比較すると、日本語などで文字化けする　towlowerを使う
		transform(compareName.begin (), compareName.end (), compareName.begin (), towlower);

		//プログラムと機能のソートルールは以下のとおり
		//(1)ひらがな、カタカナ、半角カタカナは区別しない
		//(2)英字の大文字、小文字は区別しない
		//(3)ざっくり言うと数字→英語→かな→漢字の順に並ぶ
		//(4)数字が入っていると、そこは数値としてソートする
		//   ※2000までは普通に文字順だったのにXP以降このルールになったらしい
		//     エクスプローラのファイルをファイル名でソートしたときと同様
		//     1つのファイル名の中に複数の数値部分がある場合は、最初のものが対象となる
		//     参考 http://support.microsoft.com/default.aspx?kbid=319827
		//          http://www.atmarkit.co.jp/fwin2k/win2ktips/342xpsort/xpsort.html
		//     これを実現するための関数としてStrCmpLogicalWが提供されているが、XP以降でしか使えないので独自実装する
		//     http://msdn.microsoft.com/en-us/library/bb759947%28VS.85%29.aspx
		//一方で、ソート順は以下のとおり
		//半角数字→半角英字→全角ひらがな→全角漢字→半角カナ→全角数字→全角英字
		//半角カナ対策としてLCMAP_FULLWIDTHすると、英字の前にひらがなと漢字が来てしまうのでNG
		//逆にLCMAP_HALFWIDTHすると、ひらがなの前に漢字が来てしまってこれもNG
		//LCMAP_FULLWIDTH、LCMAP_HALFWIDTHは悪影響が大きいので、半角カナは独自実装で全角カナに変換する
		//※LCMAP_FULLWIDTHで「EPSONﾌﾟﾘﾝﾀﾄﾞﾗｲﾊﾞ･ﾕｰﾃｨﾘﾃｨ」が「ｅｐｓｏｎプリンタドライバ・ユーティリティｄｅｌ」と、末尾に「ｄｅｌ」がなぜか付いてしまう問題もある。
		//  生成されたcsvをバイナリエディタで見ても、制御コードのDELは付いていないようなので原因不明

		//★★未実装★★
		//第二ソートは何でやっているか、インストール日？それともバージョン？未確認

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
		++softwareInfoIterator;                 // イテレータを１つ進める
	}
	sort(currentSoftwareInfos.begin(), currentSoftwareInfos.end(), softwareInfo::nameCompare);

	//Oracle対応
	//Oracleがプログラムと機能に現れなくてもiniで設定されていたら取得する
	if(iniOracleCompatible > 0){
		//Oracleのコマンドを入力し、oracle.invとして一度保存

		//_wgetenvでは存在しない名称の環境変数を入れようとするとエラー落ちするので安全な_wgetenv_sを使うように変更
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

				//以下を試してもSJISの日本語が文字化けする
				//_tsetlocale(LC_ALL, L".ACP");
				//_tsetlocale(LC_ALL, L"japanese");
				//wcout.imbue(std::locale("Japanese"));
				//wcout.imbue(std::locale(""));

				//以下を書くとSJISの日本語を文字化けせずに読み込ませられる
				//https://nullnull.hatenablog.com/entry/20120629/1340935277
				//必要最小限以外はコメントアウト
				//ios_base::sync_with_stdio(false);
				locale default_loc("");
				locale::global(default_loc);
				//locale ctype_default(locale::classic(), default_loc, locale::ctype); //※
				//wcout.imbue(ctype_default);
				//wcin.imbue(ctype_default);

				wifstream ifs(oracle_output_file.c_str());
				//ファイルが存在するなら中身を取得
				if(ifs.is_open()){
					//wstring str((istreambuf_iterator<TCHAR>(ifs)), std::istreambuf_iterator<TCHAR>());
					wstring line;
					bool oracle_products_found = false;
					int sprit_pos;
					wstring oracle_products_name;
					wstring oracle_products_ver;
					while (getline(ifs, line)){
						//製品名列挙後にフラグをfalseに
						if(line.find(L"の製品がインストールされています。") != std::string::npos) {
							oracle_products_found = false;
						}else if(line.find(L"products installed in this Oracle Home.") != std::string::npos) {
							oracle_products_found = false;
						}
						//製品名取得
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
							//softwareInfo構造体を用意
							softwareInfo currentSoftwareInfo;
							currentSoftwareInfo.Name = oracle_products_name;
							currentSoftwareInfo.Vendor = L"Oracle Corporation";
							currentSoftwareInfo.Version = oracle_products_ver;
							currentSoftwareInfos.push_back(currentSoftwareInfo);
						}
						//製品名列挙直前にフラグをtrueに
						if(line.find(L"インストールされた最上位製品") != std::string::npos) {
							oracle_products_found = true;
						}else if(line.find(L"Installed Top-level Products") != std::string::npos) {
							oracle_products_found = true;
						}
												
					}
				}
				//一時ファイル削除　ただし、iniで2が指定されていれば一時ファイルを残す
				if(iniOracleCompatible != 2){
					DeleteFile(oracle_output_file.c_str());
				}
			}
		}

	}


	//patchInfosのソート
	//ソート用フィールド(compareName)の準備
	vector<patchInfo>::iterator patchInfoIterator = currentPatchInfos.begin();  // イテレータのインスタンス化
	while(patchInfoIterator != currentPatchInfos.end()){
		//比較用文字列を用意
		wstring compareName = (*patchInfoIterator).ParentName + (*patchInfoIterator).Name;

		//一文字ずつ切り出してコードポイントを付けておく
		//（制御コードが今後混じってきたときに確認するため）
		//一文字ずつ切り出す方法 http://d.hatena.ne.jp/minus9d/20120517/1337265190
		wchar_t buffer [100];
		for(int i = 0; i < (int)compareName.size(); ++i){
			swprintf(buffer, 100, L"%c:U+%lX ", compareName[i], compareName[i]);
			(*patchInfoIterator).NameCodePoint += buffer;
		}

		//wstring compareName = L"にほんご日本語123１２３カタカナ";
		//比較用文字列を小文字にする
		//tolowerで比較すると、日本語などで文字化けする　towlowerを使う
		transform(compareName.begin (), compareName.end (), compareName.begin (), towlower);

		//プログラムと機能のソートルールは以下のとおり
		//(1)ひらがな、カタカナ、半角カタカナは区別しない
		//(2)英字の大文字、小文字は区別しない
		//(3)ざっくり言うと数字→英語→かな→漢字の順に並ぶ
		//(4)数字が入っていると、そこは数値としてソートする
		//   ※2000までは普通に文字順だったのにXP以降このルールになったらしい
		//     エクスプローラのファイルをファイル名でソートしたときと同様
		//     1つのファイル名の中に複数の数値部分がある場合は、最初のものが対象となる
		//     参考 http://support.microsoft.com/default.aspx?kbid=319827
		//          http://www.atmarkit.co.jp/fwin2k/win2ktips/342xpsort/xpsort.html
		//     これを実現するための関数としてStrCmpLogicalWが提供されているが、XP以降でしか使えないので独自実装する
		//     http://msdn.microsoft.com/en-us/library/bb759947%28VS.85%29.aspx
		//一方で、ソート順は以下のとおり
		//半角数字→半角英字→全角ひらがな→全角漢字→半角カナ→全角数字→全角英字
		//半角カナ対策としてLCMAP_FULLWIDTHすると、英字の前にひらがなと漢字が来てしまうのでNG
		//逆にLCMAP_HALFWIDTHすると、ひらがなの前に漢字が来てしまってこれもNG
		//LCMAP_FULLWIDTH、LCMAP_HALFWIDTHは悪影響が大きいので、半角カナは独自実装で全角カナに変換する
		//※LCMAP_FULLWIDTHで「EPSONﾌﾟﾘﾝﾀﾄﾞﾗｲﾊﾞ･ﾕｰﾃｨﾘﾃｨ」が「ｅｐｓｏｎプリンタドライバ・ユーティリティｄｅｌ」と、末尾に「ｄｅｌ」がなぜか付いてしまう問題もある。
		//  生成されたcsvをバイナリエディタで見ても、制御コードのDELは付いていないようなので原因不明

		//★★未実装★★
		//第二ソートは何でやっているか、インストール日？それともバージョン？未確認

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
		++patchInfoIterator;                 // イテレータを１つ進める
	}
	sort(currentPatchInfos.begin(), currentPatchInfos.end(), patchInfo::nameCompare);


	return 0;
}

//指定されたレジストリからソフトウェア名を調べる関数（本体）
long InventoryManager::scanRegistry(BOOL scan64bit, BOOL scanHKCU){

	//_tprintf(TEXT("scanRegistry\n"));
	//Uninstallキーの取得
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

	//Uninstallキー配下のサブキーごとに処理
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
		//今から取得する値についてsoftwareInfo構造体を用意する
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
		//サブキー名が空なら終了
		if(uninstallSubKeySubkey.name == L""){
			if(debugLevel > 2){
				PrintDebugLog(L"break. uninstallSubKeySubkey.name == \"\"\n");
			}
			break;
		}
		//_tprintf(TEXT("subKey: %s\n"), uninstallSubKeySubkeyFullPathName);
		//RegList uninstallSubkeyRegList = new RegList;
		uninstallSubkeyRegList.setHKeyAndSubKey(uninstallHKeyName, uninstallSubKeySubkey.fullPathName, scan64bit);
		//#SystemComponentというレジストリ値があり、かつ型がREG_DWORDでその値が1の場合、処理を終了
		regValue SystemComponentValue = uninstallSubkeyRegList.getValue(L"SystemComponent");
		//if((uninstallSubkeyRegListStatus == 0) && (_tcscmp(uninstallSubkeyRegType, TEXT("REG_DWORD")) == 0) && (_tcscmp(uninstallSubkeyRegValue, TEXT("1")) == 0)){
		if(SystemComponentValue.name != L"" && SystemComponentValue.type == L"REG_DWORD" && SystemComponentValue.data == L"1"){
			if(debugLevel > 2){
				PrintDebugLog(L"continue. SystemComponentValue.name != \"\"\n");
			}
			continue;
		}
		//差分比較のためレジストリキーを取得
		currentSoftwareInfo.RegistoryKey = uninstallSubKeySubkey.name;
		//★★未検証★★
		//Publisherというレジストリ値があり、かつ型がREG_SZまたはREG_EXPAND_SZでその値が""でない場合、リストに追加する
		//（ここもREG_MULTI_SZはNG）
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
		//★★未検証★★
		//DisplayVersionというレジストリ値があり、かつ型がREG_SZまたはREG_EXPAND_SZでその値が""でない場合、リストに追加する
		//（ここもREG_MULTI_SZはNG）
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
		//WindowsInstallerというレジストリ値があり、かつ型がREG_DWORDでその値が1の場合…
		//（REG_QWORDもREG_DWORDと似たようなものだが、REG_QWORDだと値が1でもダメ）
		regValue WindowsInstallerValue = uninstallSubkeyRegList.getValue(L"WindowsInstaller");
		if(WindowsInstallerValue.name != L"" && WindowsInstallerValue.type == L"REG_DWORD" && WindowsInstallerValue.data == L"1"){
			if(debugLevel > 2){
				PrintDebugLog(L"hold. WindowsInstallerValue.data == \"1\"\n");
			}
			//処理がHKLM配下でないときは、処理を終了
			//↑間違い。全ての場合に行う
			//処理がHKLM配下ならさらに調べる必要がある
			//まずGUIDのフォーマットに合致するかどうか確認
			//（GUIDは本来16進数しか許容されないが、プログラムの追加と削除での表示では0-9a-fA-F以外でもアルファベットであれば通る）
			//GUIDはこんなの→{EE936C7A-EA40-31D5-9B65-8E3E089C3828}
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
				//HKCU\Software\Microsoft\Installer\Products\配下にレジストリキーがある場合
				//（キーが存在するかはEnumKeyで返り値が2でないことを見ればよい）
				//（これと次のHKCR\Installer\Products\の両方がある場合はこれが優先する）
				productsHKeyName = L"HKEY_CURRENT_USER";
				productsSubKeyName = L"Software\\Microsoft\\Installer\\Products\\" + transformedUninstallSubKeySubkeyName;
				status = productsRegList.setHKeyAndSubKey(productsHKeyName, productsSubKeyName, scan64bit);
				if(status == 0){
/*
					//★★未検証★★
					//Publisherというレジストリ値があり、かつ型がREG_SZまたはREG_EXPAND_SZまたはREG_MULTI_SZでその値が""ではない場合、取り込む
					//REG_DWORDもOK? QWORDはNG
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
							//TODO REG_MULTI_SZの場合、プログラムの追加と削除では1行目のみ表示されるが未処理
							//_tprintf(TEXT("ProductName:%s\n"), ProductNameValue.data.c_str());
							currentsoftwareInfo.vendor = PublisherValue.data;
							continue;
						}else if(PublisherValue.type == L"REG_DWORD" && PublisherValue.data != L""){
							//TODO REG_DWORDの条件（0はダメ？）未確認、標記0x1→1変換未処理
							//_tprintf(TEXT("ProductName:%s\n"), ProductNameValue.data.c_str());
							currentsoftwareInfo.vendor = PublisherValue.data;
							continue;
						//TODO REG_QWORD未確認
						}
					}
*/
					//ProductNameというレジストリ値があり、かつ型がREG_SZまたはREG_EXPAND_SZまたはREG_MULTI_SZでその値が""ではない場合、取り込む
					//REG_DWORDもOKだった。QWORDはNG
					regValue ProductNameValue = productsRegList.getValue(L"ProductName");
					if(ProductNameValue.name != L""){
						wstring productNameValueStr = L"";
						if(ProductNameValue.type == L"REG_SZ" && ProductNameValue.data != L""){
							productNameValueStr = ProductNameValue.data;
						}else if(ProductNameValue.type == L"REG_EXPAND_SZ" && ProductNameValue.data != L""){
							productNameValueStr = ProductNameValue.data;
						}else if(ProductNameValue.type == L"REG_MULTI_SZ" && ProductNameValue.data != L""){
							//TODO REG_MULTI_SZの場合、プログラムの追加と削除では1行目のみ表示されるが未処理
							productNameValueStr = ProductNameValue.data;
						}else if(ProductNameValue.type == L"REG_DWORD" && ProductNameValue.data != L""){
							//TODO REG_DWORDの条件（0はダメ？）未確認、標記0x1→1変換未処理
							productNameValueStr = ProductNameValue.data;
						//TODO REG_QWORD未確認
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
							//HKCU\Software\Microsoft\Installer\Products\transformedUninstallSubKeySubkeyName\Patches配下にレジストリキーがある場合
							status = productsPatchRegList.setHKeyAndSubKey(productsHKeyName, productsSubKeyPatchKeyName, scan64bit);
							if(status == 0){
								//Patchesというレジストリ値があり、かつ型がREG_MULTI_SZ（REG_SZ、REG_EXPAND_SZ不可）でその値が""ではない場合、取り込む
								regValue PatchesNameValue = productsPatchRegList.getValue(L"Patches");
								if(PatchesNameValue.name != L""){
									if(PatchesNameValue.type == L"REG_MULTI_SZ" && PatchesNameValue.data != L""){
										//1行ごとに取り出す
										//https://stackoverflow.com/questions/36812132/splitting-stdwstring-into-stdvector
										wstringstream patchesNameStream(PatchesNameValue.data);
										wstring patchesNameValueLine;
										//MessageBox(NULL,PatchesNameValue.data.c_str(),L"test",MB_OK);
										while(getline(patchesNameStream, patchesNameValueLine)){
											//MessageBox(NULL,patchesNameValueLine.c_str(),L"test",MB_OK);
											wstring s1518PatchSubKeyKeyname;
											s1518PatchSubKeyKeyname = L"Software\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\" + transformedUninstallSubKeySubkeyName + L"\\Patches\\" + patchesNameValueLine;
											RegList s1518PatchRegList;
											//↓のようなレジストリキーがあるか
											//HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\68AB67CA7DA7FFFFB744AA0000000010\Patches\68AB67CA7DA7FFFF5205A7C804101061
											status = s1518PatchRegList.setHKeyAndSubKey(L"HKEY_CURRENT_USER", s1518PatchSubKeyKeyname, scan64bit);
											if(status == 0){
												regValue DisplayNameValue = s1518PatchRegList.getValue(L"DisplayName");
												wstring displayNameValueStr = L"";
												//DisplayNameというレジストリ値があり、かつ型がREG_SZでその値が""でない場合、パッチリストに追加する
												//DisplayNameというレジストリ値があり、かつ型がREG_DWORDならその値（10進数）をパッチリストに追加する
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
												//DisplayNameというレジストリ値があり、型がREG_SZだが値が""の場合
												//DisplayNameというレジストリ値があるが型がREG_SZでもREG_DWORDでもない場合
												//DisplayNameというレジストリ値がない場合
												//いずれもプログラムと機能ではただ「更新プログラム」と表示される
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
				//HKCR\Installer\Products\配下にレジストリキーがある場合
				productsHKeyName = L"HKEY_CLASSES_ROOT";
				productsSubKeyName = L"Installer\\Products\\" + transformedUninstallSubKeySubkeyName;
				//_tprintf(TEXT("transformedUninstallSubKeySubkeyName:%s\n"), transformedUninstallSubKeySubkeyName.c_str());
				status = productsRegList.setHKeyAndSubKey(productsHKeyName, productsSubKeyName, scan64bit);
				if(status == 0){
					//ProductNameというレジストリ値があり、かつ型がREG_SZまたはREG_EXPAND_SZまたはREG_MULTI_SZでその値が""ではない場合、取り込む
					//REG_DWORDもOKだった。QWORDはNG
					//（Windows 7で試したところ、REG_EXPAND_SZはGetStringValueでも取得でき、
					//　またREG_SZもGetExpandedStringValueで取得できてしまうのでElseIf reg.GetExpandedStringValue...は意味がなさそう）
					//（REG_MULTI_SZの場合、プログラムの追加と削除では1行目のみ表示される）←★★未実装★★
					regValue ProductNameValue = productsRegList.getValue(L"ProductName");
					if(ProductNameValue.name != L""){
						wstring productNameValueStr = L"";
						if(ProductNameValue.type == L"REG_SZ" && ProductNameValue.data != L""){
							productNameValueStr = ProductNameValue.data;
						}else if(ProductNameValue.type == L"REG_EXPAND_SZ" && ProductNameValue.data != L""){
							productNameValueStr = ProductNameValue.data;
						}else if(ProductNameValue.type == L"REG_MULTI_SZ" && ProductNameValue.data != L""){
							//TODO REG_MULTI_SZの場合、プログラムの追加と削除では1行目のみ表示されるが未処理
							productNameValueStr = ProductNameValue.data;
						}else if(ProductNameValue.type == L"REG_DWORD" && ProductNameValue.data != L""){
							//TODO REG_DWORDの条件（0はダメ？）未確認、標記0x1→1変換未処理
							productNameValueStr = ProductNameValue.data;
						//TODO REG_QWORD未確認
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
							//HKCU\Software\Microsoft\Installer\Products\transformedUninstallSubKeySubkeyName\Patches配下にレジストリキーがある場合
							status = productsPatchRegList.setHKeyAndSubKey(productsHKeyName, productsSubKeyPatchKeyName, scan64bit);
							if(status == 0){
								//Patchesというレジストリ値があり、かつ型がREG_MULTI_SZ（REG_SZ、REG_EXPAND_SZ不可）でその値が""ではない場合、取り込む
								regValue PatchesNameValue = productsPatchRegList.getValue(L"Patches");
								if(PatchesNameValue.name != L""){
									if(PatchesNameValue.type == L"REG_MULTI_SZ" && PatchesNameValue.data != L""){
										//1行ごとに取り出す
										//https://stackoverflow.com/questions/36812132/splitting-stdwstring-into-stdvector
										wstringstream patchesNameStream(PatchesNameValue.data);
										wstring patchesNameValueLine;
										//MessageBox(NULL,PatchesNameValue.data.c_str(),L"test",MB_OK);
										while(getline(patchesNameStream, patchesNameValueLine)){
											//MessageBox(NULL,patchesNameValueLine.c_str(),L"test",MB_OK);
											wstring s1518PatchSubKeyKeyname;
											s1518PatchSubKeyKeyname = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\" + transformedUninstallSubKeySubkeyName + L"\\Patches\\" + patchesNameValueLine;
											RegList s1518PatchRegList;
											//↓のようなレジストリキーがあるか
											//HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\68AB67CA7DA7FFFFB744AA0000000010\Patches\68AB67CA7DA7FFFF5205A7C804101061
											//64bitOSなら64bitキーを調べる　32bitキーにアクセスしようとすると返り値2となる

											status = s1518PatchRegList.setHKeyAndSubKey(L"HKEY_LOCAL_MACHINE", s1518PatchSubKeyKeyname, IsWow64());
											//テスト
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
												//DisplayNameというレジストリ値があり、かつ型がREG_SZでその値が""でない場合、パッチリストに追加する
												//DisplayNameというレジストリ値があり、かつ型がREG_DWORDならその値（10進数）をパッチリストに追加する
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
												//DisplayNameというレジストリ値があり、型がREG_SZだが値が""の場合
												//DisplayNameというレジストリ値があるが型がREG_SZでもREG_DWORDでもない場合
												//DisplayNameというレジストリ値がない場合
												//いずれもプログラムと機能ではただ「更新プログラム」と表示される
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
		//WindowsInstallerというレジストリ値がないか、あっても型がREG_DWORDでその値が1以外の場合
		}else{
			if(debugLevel > 2){
				PrintDebugLog(L"hold. WindowsInstallerValue.data != \"1\"\n");
			}
			//UninstallStringというレジストリ値があり、かつ型がREG_SZまたはREG_EXPAND_SZでその値が""ではない場合、処理を継続
			//（ここはREG_MULTI_SZはNG）
			regValue UninstallStringValue = uninstallSubkeyRegList.getValue(L"UninstallString");
			//if(uninstallSubKeySubkey.name == L"{87434D51-51DB-4109-B68F-A829ECDCF380}"){
			//	_tprintf(TEXT("test:%s\n"), UninstallStringValue.data.c_str());
			//}
			if(UninstallStringValue.name != L""){
				if(UninstallStringValue.type == L"REG_SZ" && UninstallStringValue.data != L""){
					if(debugLevel > 2){
						PrintDebugLog(L"hold. UninstallStringValue.data != \"\"\n");
					}
					//処理を継続
				}else if(UninstallStringValue.type == L"REG_EXPAND_SZ" && UninstallStringValue.data != L""){
					if(debugLevel > 2){
						PrintDebugLog(L"hold. UninstallStringValue.data != \"\"\n");
					}
					//処理を継続
				}else{
					if(debugLevel > 2){
						PrintDebugLog(L"continue. UninstallStringValue.data is empty or == \"\"\n");
					}
					//上記以外は処理を終了
					continue;
				}
			}else{
				if(debugLevel > 2){
					PrintDebugLog(L"continue. UninstallStringValue not found.\n");
				}
				//上記以外は処理を終了
				continue;
			}
			//ParentKeyNameというレジストリ値があり、かつ型がREG_DWORDでその値が0でない場合、処理を終了
			//ParentKeyNameというレジストリ値があり、かつ型がREG_SZまたはREG_EXPAND_SZまたはREG_MULTI_SZでその値が""でない場合、処理を終了
			//ParentKeyNameというレジストリ値があり、かつ型がREG_BINARYで空でない場合、処理を終了
			//Vista以降のOSにおいて、ParentKeyNameというレジストリ値があり、かつ型がREG_QWORDでその値が0でない場合、
			//ソフトウェアリストは処理を終了、パッチリストは逆に処理を継続
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
					//ParentDisplayNameというレジストリ値があり、かつ型がREG_SZまたはREG_EXPAND_SZでその値が""でない場合、親プログラム名としてパッチリストに追加する
					if(ParentDisplayNameValue.name != L""){
						if(ParentDisplayNameValue.type == L"REG_SZ" && ParentDisplayNameValue.data != L""){
							currentPatchInfo.ParentName = ParentDisplayNameValue.data;
						}else if(ParentDisplayNameValue.type == L"REG_EXPAND_SZ" && ParentDisplayNameValue.data != L""){
							currentPatchInfo.ParentName = ParentDisplayNameValue.data;
						}
					}
					regValue DisplayNameValue = uninstallSubkeyRegList.getValue(L"DisplayName");
					//DisplayNameというレジストリ値があり、かつ型がREG_SZまたはREG_EXPAND_SZでその値が""でない場合、パッチリストに追加する
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
			//指定されたレジストリキー配下のサブキーがサポート技術情報番号（KB000000）そのものである場合、基本的には処理を終了するが
			//ParentKeyNameが存在し、0または""である場合のみ処理を継続
			//（Windows Search 4.0で発覚）
			if((uninstallSubKeySubkey.name.size() == 8)
			&& (uninstallSubKeySubkey.name.substr(0,2) == L"KB")
			&& isNumericString(uninstallSubKeySubkey.name.substr(2,6))
			){
				//_tprintf(TEXT("test: %s\n"), uninstallSubKeySubkey.name.c_str());
				if(debugLevel > 2){
					PrintDebugLog(L"hold. uninstallSubKeySubkey.name is KB.\n");
				}
				//ParentKeyNameというレジストリ値があり、かつ型がREG_DWORDでその値が0の場合、処理を継続
				//ParentKeyNameというレジストリ値があり、かつ型がREG_SZまたはREG_EXPAND_SZまたはREG_MULTI_SZでその値が""の場合、処理を継続
				//ParentKeyNameというレジストリ値があり、かつ型がREG_BINARYで空の場合、処理を継続
				//Vista以降のOSにおいて、ParentKeyNameというレジストリ値があり、かつ型がREG_QWORDでその値が0の場合、処理を継続
				if(ParentKeyNameValue.name != L""){
					if(ParentKeyNameValue.type == L"REG_DWORD" && ParentKeyNameValue.data != L"0"){
						if(debugLevel > 2){
							PrintDebugLog(L"hold. ParentKeyNameValue.data !=\"\"\n");
						}
						//処理を継続
					}else if(ParentKeyNameValue.type == L"REG_SZ" && ParentKeyNameValue.data != L""){
						if(debugLevel > 2){
							PrintDebugLog(L"hold. ParentKeyNameValue.data !=\"\"\n");
						}
						//処理を継続
					}else if(ParentKeyNameValue.type == L"REG_EXPAND_SZ" && ParentKeyNameValue.data != L""){
						if(debugLevel > 2){
							PrintDebugLog(L"hold. ParentKeyNameValue.data !=\"\"\n");
						}
						//処理を継続
					}else if(ParentKeyNameValue.type == L"REG_MULTI_SZ" && ParentKeyNameValue.data != L""){
						if(debugLevel > 2){
							PrintDebugLog(L"hold. ParentKeyNameValue.data !=\"\"\n");
						}
						//処理を継続
					}else if(ParentKeyNameValue.type == L"REG_BINARY" && ParentKeyNameValue.data != L""){
						if(debugLevel > 2){
							PrintDebugLog(L"hold. ParentKeyNameValue.data !=\"\"\n");
						}
						//処理を継続
					}else if(ParentKeyNameValue.type == L"REG_QWORD" && ParentKeyNameValue.data != L"0"){
						if(debugLevel > 2){
							PrintDebugLog(L"hold. ParentKeyNameValue.data !=\"\"\n");
						}
						//処理を継続
					}else{
						if(debugLevel > 2){
							PrintDebugLog(L"continue");
						}
						//上記以外は処理を終了
						continue;
					}
				}
			}
			//DisplayName_Localizedというレジストリ値があるときは、LocalizedString化してそれを取得する
			//★★未実装★★REG_SZ以外の確認、空文字列時の挙動、ProductNameのLocalized対応有無
			//---------------------------
			regValue DisplayNameLocalizedValue = uninstallSubkeyRegList.getValue(L"DisplayName_Localized");
			if (DisplayNameLocalizedValue.name != L"") {
				if (DisplayNameLocalizedValue.type == L"REG_SZ" && DisplayNameLocalizedValue.data != L"")
				{
					try
					{
						wstring displayNameLocalizedValueData = getLocalizedString(DisplayNameLocalizedValue.data);
						//ここでエラーが発生した場合はこの下のif分に入らないようにし、continueしないようにする。（2022/01/13）
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
			//★★未検証★★
			//Publisherというレジストリ値があり、かつ型がREG_SZまたはREG_EXPAND_SZでその値が""でない場合、リストに追加する
			//（ここもREG_MULTI_SZはNG）
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
			//DisplayNameというレジストリ値があり、かつ型がREG_SZまたはREG_EXPAND_SZでその値が""でない場合、リストに追加する
			//（ここもREG_MULTI_SZはNG）
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

//WMIでパッチ情報を取得する関数（本体）
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

		//値の取得
		//hr = pclsObj->Get(BSTR(L"Description"), 0, &vtProp, 0, 0);
		//if(!FAILED(hr)){
		//	if(vtProp.vt==VT_NULL || vtProp.vt==VT_EMPTY){
		//		//何もしない
		//	}else{
		//		_bstr_t bstr = vtProp;
		//		hotfixStr = bstr;
		//	}
		//}

		wstring hotfixStr = L"";
		hr = pclsObj->Get(BSTR(L"HotFixID"), 0, &vtProp, 0, 0);
		if(!FAILED(hr)){
			if(vtProp.vt==VT_NULL || vtProp.vt==VT_EMPTY){
				//何もしない
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

		//以前のWindowsOSで未パッチの場合、HotFixIDがFile 1となる模様
		//https://stackoverflow.com/questions/10291296/what-is-the-hotfix-with-hotfixid-file-1
		//https://web.archive.org/web/20121013153025/http://support.microsoft.com/kb/831415
		//この場合はServicePackInEffectの取得を試みる
		if(hotfixStr == L"File 1"){
			hotfixStr = L"";
			hr = pclsObj->Get(BSTR(L"ServicePackInEffect"), 0, &vtProp, 0, 0);
			if(!FAILED(hr)){
				if(vtProp.vt==VT_NULL || vtProp.vt==VT_EMPTY){
					//何もしない
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

//内容を全出力する関数
wstring InventoryManager::output(){

	//ロケールを自動設定しておく
	//これをしておかないと、SJISで保存したときに日本語が??になる
	//参考 http://cx5software.com/article_vcpp_unicode/
	//setlocale関数の第二引数に「""」を指定すると、OSのデフォルトロケールが指定されます。例えば、日本語に設定されている場合「Japanese_Japan.932」が設定されます。当然ながら、OSの設定が変わると、それに合わせて変更されますので注意が必要です。
	_tsetlocale(LC_ALL, _T(""));
	//_tsetlocale(LC_ALL, L"Japanese");

	//Unicode文字列をコンソールに表示させる
	//_tsetlocale(LC_ALL, _tsetlocale(LC_CTYPE, L""));

	//ファイル出力のコンストラクタ
	//http://www.geocities.jp/ky_webid/cpp/library/033.html
	//http://sato-si.at.webry.info/200703/article_1.html

	//ソフトウェアリストの画面出力
	//GUI版はコメントアウト
/*
	vector<softwareInfo>::iterator softwareInfoIterator = softwareInfos.begin();  // イテレータのインスタンス化
	while(softwareInfoIterator != softwareInfos.end()){
		_tprintf(TEXT("%s\n"), (*softwareInfoIterator).name.c_str());  // *演算子で間接参照
		++softwareInfoIterator;                 // イテレータを１つ進める
	}
*/

	//基本的なインベントリ情報の取得→ここで取るのをやめる
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

	//ソフトウェアリストのファイル出力
	//http://cx5software.com/article_vcpp_unicode/

	//フォーマットが通常形式（シンプルなcsv）かSARMS形式かで分岐

	//フォーマットがSARMS形式の場合
	//return outputFileName;
	if(outputFileFormat == L"SARMS"){
		returnString = outputSARMS();
	//フォーマットが通常形式（シンプルなcsv）の場合
	}else if(outputFileFormat == L"SIMPLE"){
		returnString = outputSIMPLE();
	//フォーマットがAdvancedManager形式の場合
	}else if(outputFileFormat == L"AdvancedManager"){
		returnString = outputAdvancedManager();
	}
	return returnString;
}

//SARMS形式で内容を全出力する関数
wstring InventoryManager::outputSARMS(){

		wstring returnOutputFileName = L"";
	
		//アプリケーション情報の保存
		FILE* fp = NULL;
		wstring outputFileName = outputPath + L"Ap_" + currentHardwareInfo.HardwareNoValue + L".csv";

		//エンコードはSJISで保存する（選択させない）
		_wfopen_s( &fp, outputFileName.c_str(), L"w" );

		//タイトル行の出力→なし
		vector<softwareInfo>::iterator softwareInfoIterator2 = currentSoftwareInfos.begin();  // イテレータのインスタンス化
		while(softwareInfoIterator2 != currentSoftwareInfos.end()){
			wstring tempField;

			tempField = Replace(currentHardwareInfo.HardwareNoValue, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			tempField= (*softwareInfoIterator2).Name.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			//_ftprintf(fp, TEXT("Software - Name:%s\t\n"), (*softwareInfoIterator).Name.c_str());  // *演算子で間接参照
			_ftprintf(fp, L"\"%s\",", tempField.c_str());  // *演算子で間接参照
			//_ftprintf(fp, L"%s,", (*softwareInfoIterator2).Name.c_str());  // *演算子で間接参照
			
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
			++softwareInfoIterator2;                 // イテレータを１つ進める
		}
		//_ftprintf(fp, L"\n");
		//_ftprintf(fp, L"Output by LIP (List Installed Programs) v0.1.");
	
		fclose(fp);
		returnOutputFileName += outputFileName.c_str();
		returnOutputFileName += L", ";

		//その他のインベントリ情報の保存
		fp = NULL;
		outputFileName = outputPath + L"Inv_" + currentHardwareInfo.HardwareNoValue + L".csv";

		//エンコードはSJISで保存する（選択させない）
		_wfopen_s( &fp, outputFileName.c_str(), L"w" );

		//タイトル行の出力→なし
		wstring tempField;
		//管理番号
		tempField = Replace(currentHardwareInfo.HardwareNoValue, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//組織1
		tempField = Replace(currentHardwareInfo.Value1, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//組織2
		tempField = Replace(currentHardwareInfo.Value2, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//組織3
		tempField = Replace(currentHardwareInfo.Value3, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//組織4
		tempField = Replace(currentHardwareInfo.Value4, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//組織5
		tempField = Replace(currentHardwareInfo.Value5, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//IPアドレス
		tempField = Replace(currentHardwareInfo.IPAddress, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//MACアドレス
		tempField = Replace(currentHardwareInfo.MACAddress, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//ユーザ名
		tempField = Replace(currentHardwareInfo.UserName, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//コンピュータ名
		tempField = Replace(currentHardwareInfo.ComputerName, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//機種
		tempField = Replace(currentHardwareInfo.SystemProductName, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//ベンダー
		tempField = Replace(currentHardwareInfo.SystemManufacturer, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//OS
		tempField = Replace(currentHardwareInfo.OSName, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//未使用欄1→使用者名
		tempField = Replace(currentHardwareInfo.Value6, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//未使用欄2→CPUタイプ
		tempField = Replace(currentHardwareInfo.ProcessorName, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());
		//未使用欄3
		_ftprintf(fp, L"\"\",");
		//取得日時
		//_ftprintf(fp, L"\"%04d/%02d/%02d %02d:%02d:%02d\",", timestamp.tm_year + 1900, timestamp.tm_mon + 1, timestamp.tm_mday, timestamp.tm_hour, timestamp.tm_min, timestamp.tm_sec);
		_ftprintf(fp, L"\"%s\",", currentHardwareInfo.Timestamp2.c_str());
		//ツールのバージョン
		tempField = Replace(currentHardwareInfo.ToolVersion, L"\"", L"\"\"");
		_ftprintf(fp, L"\"%s\",", tempField.c_str());

		fclose(fp);
		returnOutputFileName += outputFileName.c_str();

		return returnOutputFileName;
}


//SIMPLE形式で内容を全出力する関数
wstring InventoryManager::outputSIMPLE(){

		//wchar_t outputFileNameChars [100];
		//swprintf(outputFileNameChars, 100,  _T( "%s%s.csv" ), currentHardwareInfo.ComputerName.c_str(), currentHardwareInfo.Timestamp.c_str());
		//wstring outputFileName = outputFileNameChars;
		wstring outputFileName = outputPath + currentHardwareInfo.ComputerName + currentHardwareInfo.Timestamp + L".csv";

		FILE* fp = NULL;

		//普通に書きだすとSJISになるが、SJISだと丸アールがただのRになってしまう
		//プログラムと機能と一致する情報を取るためには、UTF-8で保存すべき
		if(outputFileEncode == L"SJIS"){
			_wfopen_s( &fp, outputFileName.c_str(), L"w" );
		}else if(outputFileEncode == L"UTF-8N"){
			_wfopen_s( &fp, outputFileName.c_str(), L"w,ccs=UTF-8" );
			//普通にやるとBOM付きのUTF-8となるので、BOM無しのUTF-8としたい場合はseekする
			//http://social.msdn.microsoft.com/Forums/ja-JP/965c68dc-3e86-49a2-9c65-b469beb1f69f/bomutf8
			fseek(fp, 0, SEEK_SET); 
		}else{
			_wfopen_s( &fp, outputFileName.c_str(), L"w,ccs=UTF-8" );
		}

		//タイトル行の出力
		//取得日時、メーカー、機種、シリアル番号、CPU追加
		_ftprintf(fp, L"\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\"\n", currentHardwareInfo.HardwareNoTitle.c_str(), currentHardwareInfo.Title1.c_str(), currentHardwareInfo.Title2.c_str(), currentHardwareInfo.Title3.c_str(), currentHardwareInfo.Title4.c_str(), currentHardwareInfo.Title5.c_str(), currentHardwareInfo.Title6.c_str(), L"コンピュータ名", L"IPアドレス", L"MACアドレス", L"PCベンダー", L"PC機種", L"シリアル番号", L"CPUタイプ", L"ログインユーザ", L"OS", L"ソフトウェア名", L"ベンダー名", L"バージョン", L"取得日時");
		vector<softwareInfo>::iterator softwareInfoIterator2 = currentSoftwareInfos.begin();  // イテレータのインスタンス化
		while(softwareInfoIterator2 != currentSoftwareInfos.end()){
			//ofstreamString = (*softwareInfoIterator).name;
			//_tprintf(TEXT("Software - Name:%s\t\n"), ofstreamString.c_str());  // *演算子で間接参照
			//ofstreamではstringを<<で送り込めない
			//かといって.c_str()とすると、ポインタの文字列が取り込まれる
			//ofs << ofstreamString << endl;
			//ofs.put(ofstreamString.c_str());
			//_tprintf(TEXT("Software - Name:%s\t\n"), (*softwareInfoIterator2).name.c_str());  // *演算子で間接参照
			
			wstring tempField;

			//管理番号
			tempField = Replace(currentHardwareInfo.HardwareNoValue, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//入力1～6
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
			//コンピュータ名
			tempField = Replace(currentHardwareInfo.ComputerName, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//IPアドレス
			tempField = Replace(currentHardwareInfo.IPAddress, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//MACアドレス
			tempField = Replace(currentHardwareInfo.MACAddress, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//PCベンダー
			tempField = Replace(currentHardwareInfo.SystemManufacturer, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//PC機種
			tempField = Replace(currentHardwareInfo.SystemProductName, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//シリアル番号
			tempField = Replace(currentHardwareInfo.IdentifyingNumber, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//CPUタイプ
			tempField = Replace(currentHardwareInfo.ProcessorName, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//ユーザ名
			tempField = Replace(currentHardwareInfo.UserName, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//OS
			tempField = Replace(currentHardwareInfo.OSName, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//ソフトウェア名
			tempField= (*softwareInfoIterator2).Name.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			//_ftprintf(fp, TEXT("Software - Name:%s\t\n"), (*softwareInfoIterator).name.c_str());  // *演算子で間接参照
			_ftprintf(fp, L"\"%s\",", tempField.c_str());  // *演算子で間接参照
			//_ftprintf(fp, L"%s,", (*softwareInfoIterator2).name.c_str());  // *演算子で間接参照
			//ソフトウェアベンダー
			tempField = (*softwareInfoIterator2).Vendor.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//ソフトウェアバージョン
			tempField = (*softwareInfoIterator2).Version.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//取得日時
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
			++softwareInfoIterator2;                 // イテレータを１つ進める
		}
		//_ftprintf(fp, L"\n");
		//_ftprintf(fp, L"Output by LIP (List Installed Programs) v0.1.");
		fclose(fp);
		return outputFileName;

}



//SIMPLE形式で内容を全出力する関数
wstring InventoryManager::outputAdvancedManager(){

		wstring returnString = L"インベントリ情報を保存しました。\n保存ファイル: \n";
		wstring tempField;

		//%ALLUSERSPROFILE%へのデータ保存
		//XPまではC:\Documents and Settings\All Users
		//Vista以降はC:\ProgramDataを指す
		//http://pasofaq.jp/windows/mycomputer/folderlist.htm
		//保存場所をどこにするか
		//ユーザーごとにデータを保存するのであればHKEY_LOCL_USERまたは%LOCALAPPDATA%がよい
		//収集情報はユーザーごとではなくPCごとのデータなので、%ALLUSERSPROFILE%が適当
		//http://sygh.hatenadiary.jp/entry/2013/11/10/184200
		//ただし、「書き込みアクセス権は、ファイルの作成者 (所有者) のみに付与」されるため、
		//ユーザーごとにファイルを分けて保存する必要あり
		//http://sygh.hatenadiary.jp/entry/2013/11/10/184200
		//http://bbs.wankuma.com/index.cgi?mode=al2&namber=65112&KLOG=110

		//フォルダ作成
		//wstring datPath = _wgetenv(L"ALLUSERSPROFILE");
		//datPath = datPath + L"\\LIP";
		//_wmkdir(datPath.c_str());

		//datPath = datPath + L"\\test.ini";
		//DeleteFile(datPath.c_str());

		//load();
		//save();

		//構造体をそのままファイルに書き込む場合
		//https://jyn.jp/cpp-structure-file/
		//fwrite((char*)&currentHardwareInfo, sizeof(currentHardwareInfo), 1, fp);
		//バイナリで書き込めるため暗号化を考えなくてよいバージョンアップで取得項目が変わると対応が複雑になるため、採用見送り
		//FILE* fp = NULL;
		//_wfopen_s(&fp, datPath.c_str(), L"wb");
		//fclose(fp);

		//iniファイルとして書き込む場合
		//https://hack.jp/?p=296
		//WritePrivateProfileString(TEXT("APP"), TEXT("VALUE1"), TEXT("STRING1"), szINIFilePath);
		//WritePrivateProfileString(TEXT("APP"), TEXT("VALUE1"), TEXT("STRING1"), datPath.c_str());

		//fstream file;
		//file.open("./filename.dat", ios::binary | ios::in);
		//file.write((char*)&n, sizeof(n));
		

		//ソフトウェア情報の書き出し
		wstring outputFileName = outputPath + currentHardwareInfo.HardwareNoValue + L"_" + currentHardwareInfo.ComputerName + L"_SW_" + currentHardwareInfo.Timestamp + L".csv";

		//保存場所が見つからない場合のエラー処理が必要
		FILE* fp = NULL;
		errno_t error;
		//UTF-8で保存
		error = _wfopen_s( &fp, outputFileName.c_str(), L"w,ccs=UTF-8" );

		if(error != 0){
			returnString = L"インベントリ情報の保存に失敗しました。ファイルパスとアクセス権を確認してください。\n失敗ファイル: \n";
			returnString += outputFileName;
			return returnString;
		}

		//wstring tempField;

		//ヘッダ書き出し
		_ftprintf(fp, L"\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\"\n", currentHardwareInfo.HardwareNoTitle.c_str(), L"インベントリーツールID", L"シリアル", L"コンピュータ名", L"IPアドレス", L"MACアドレス", L"インストール名称", L"Hotfix区分", L"HotfixApp名称", L"ベンダー", L"バージョン", L"実行日時", L"更新フラグ");

		//int softwareCount = 0;
		//wchar_t softwareCountChars [100];
		vector<softwareInfo>::iterator softwareInfoIterator2 = currentSoftwareInfos.begin();  // イテレータのインスタンス化
		
		while(softwareInfoIterator2 != currentSoftwareInfos.end()){

			//保存INI用のカウンタ
			//softwareCount += 1;
			//swprintf(softwareCountChars, 100, _T( "Software%d" ), softwareCount);

			//ハードウェア管理番号
			tempField = Replace(currentHardwareInfo.HardwareNoValue, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//インベントリツールID
			//tempField = Replace(VER_STR_INVENTORYTOOLID, L"\"", L"\"\"");
			//_ftprintf(fp, L"\"%s\",", tempField.c_str());
			_ftprintf(fp, L"\"\",");

			//シリアル
			tempField = Replace(currentHardwareInfo.IdentifyingNumber, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//コンピュータ名
			tempField = Replace(currentHardwareInfo.ComputerName, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//IPアドレス
			tempField = Replace(currentHardwareInfo.IPAddress, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//MACアドレス
			tempField = Replace(currentHardwareInfo.MACAddress, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//ソフトウェア名
			tempField = (*softwareInfoIterator2).Name.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//WritePrivateProfileString(softwareCountChars, L"name", (*softwareInfoIterator2).name.c_str(), datPath.c_str());

			//Hotfix区分
			_ftprintf(fp, L"\"アプリケーション\",");

			//HotfixApp名称
			_ftprintf(fp, L"\"\",");

			//ソフトウェアベンダー
			tempField = (*softwareInfoIterator2).Vendor.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//WritePrivateProfileString(softwareCountChars, L"vendor", (*softwareInfoIterator2).vendor.c_str(), datPath.c_str());

			//ソフトウェアバージョン
			tempField = (*softwareInfoIterator2).Version.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());
			//WritePrivateProfileString(softwareCountChars, L"version", (*softwareInfoIterator2).version.c_str(), datPath.c_str());

			//実行日時
			_ftprintf(fp, L"\"%s\",", currentHardwareInfo.Timestamp2.c_str());

			//更新フラグ
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
			++softwareInfoIterator2;                 // イテレータを１つ進める
		}
/*		
		fclose(fp);


		returnOutputFileName += outputFileName.c_str();
		returnOutputFileName += L", ";

		//ソフトウェア件数をINIに保管
		//swprintf(softwareCountChars, 100, _T( "%d" ), softwareCount);
		//WritePrivateProfileString(L"Software", L"count", softwareCountChars, datPath.c_str());

		//パッチ情報の書き出し
		outputFileName = outputPath + currentHardwareInfo.HardwareNoValue + L"_" + currentHardwareInfo.ComputerName + L"_HF_" + currentHardwareInfo.Timestamp + L".csv";

		fp = NULL;
		//UTF-8で保存
		_wfopen_s( &fp, outputFileName.c_str(), L"w,ccs=UTF-8" );

		//ヘッダ書き出し
		_ftprintf(fp, L"\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\"\n", currentHardwareInfo.HardwareNoTitle.c_str(), L"更新フラグ", L"更新プログラム名", L"ソフトウェア名", L"ベンダー名", L"バージョン", L"取得日時");
*/
		//int softwareCount = 0;
		//wchar_t softwareCountChars [100];
		vector<patchInfo>::iterator patchInfoIterator2 = currentPatchInfos.begin();  // イテレータのインスタンス化
		
		while(patchInfoIterator2 != currentPatchInfos.end()){

			//ハードウェア管理番号
			tempField = Replace(currentHardwareInfo.HardwareNoValue, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//インベントリツールID
			//tempField = Replace(VER_STR_INVENTORYTOOLID, L"\"", L"\"\"");
			//_ftprintf(fp, L"\"%s\",", tempField.c_str());
			_ftprintf(fp, L"\"\",");

			//シリアル
			tempField = Replace(currentHardwareInfo.IdentifyingNumber, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//コンピュータ名
			tempField = Replace(currentHardwareInfo.ComputerName, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//IPアドレス
			tempField = Replace(currentHardwareInfo.IPAddress, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//MACアドレス
			tempField = Replace(currentHardwareInfo.MACAddress, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//更新プログラム名
			tempField = (*patchInfoIterator2).Name.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//Hotfix区分
			_ftprintf(fp, L"\"HOTFIX\",");

			//HotfixApp名称
			tempField = (*patchInfoIterator2).ParentName.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//ソフトウェアベンダー
			tempField = (*patchInfoIterator2).Vendor.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//ソフトウェアバージョン
			tempField = (*patchInfoIterator2).Version.c_str();
			tempField = Replace(tempField, L"\"", L"\"\"");
			_ftprintf(fp, L"\"%s\",", tempField.c_str());

			//実行日時
			_ftprintf(fp, L"\"%s\",", currentHardwareInfo.Timestamp2.c_str());

			//更新フラグ
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
			++patchInfoIterator2;                 // イテレータを１つ進める
		}
		
		fclose(fp);

		returnString += outputFileName.c_str();
		returnString += L"\n";

		//ハードウェア情報の書き出し
		//swprintf(outputFileNameChars, 100,  _T( "%s_%s_HW_%s.csv" ), currentHardwareInfo.HardwareNoValue.c_str(), currentHardwareInfo.ComputerName.c_str(), currentHardwareInfo.Timestamp.c_str());
		//outputFileName = outputFileNameChars;
		outputFileName = outputPath + currentHardwareInfo.HardwareNoValue + L"_" + currentHardwareInfo.ComputerName + L"_HW_" + currentHardwareInfo.Timestamp + L".csv";

		fp = NULL;
		//UTF-8で保存
		error = _wfopen_s( &fp, outputFileName.c_str(), L"w,ccs=UTF-8" );

		if(error != 0){
			returnString = L"インベントリ情報の保存に失敗しました。ファイルパスとアクセス権を確認してください。\n失敗ファイル: \n";
			returnString += outputFileName;
			return returnString;
		}


		wstring fileLine1 = L"";
		wstring fileLine2 = L"";
		wstring fileLine3 = L"";


		//管理番号
		fileLine1 += L"\"" + Replace(currentHardwareInfo.HardwareNoTitle, L"\"", L"\"\"") + L"\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.HardwareNoValue, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.HardwareNoValueUpdateFlag + L"\",";
		//_ftprintf(fp, L"\"%s\",\"%s\",\"%s\"\n", currentHardwareInfo.HardwareNoValue.c_str(), L"", tempField.c_str());
		//WritePrivateProfileString(L"Hardware", currentHardwareInfo.HardwareNoTitle.c_str(), currentHardwareInfo.HardwareNoValue.c_str(), datPath.c_str());

		//インベントリツールID
		fileLine1 += L"\"インベントリーツールID\",";
		//fileLine2 += L"\"" + Replace(VER_STR_INVENTORYTOOLID, L"\"", L"\"\"") + L"\",";
		fileLine2 += L"\"\",";
		fileLine3 += L"\"\",";

		//入力1～6
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

		//メーカー名
		fileLine1 += L"\"メーカー名\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.SystemManufacturer, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.SystemManufacturerUpdateFlag + L"\",";
		//tempField = Replace(currentHardwareInfo.SystemManufacturer, L"\"", L"\"\"");
		//_ftprintf(fp, L"\"%s\",\"%s\",\"%s\"\n", L"メーカー名", L"", tempField.c_str());
		//WritePrivateProfileString(L"Hardware", L"SystemManufacturer", currentHardwareInfo.SystemManufacturer.c_str(), datPath.c_str());

		//型番
		fileLine1 += L"\"型番\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.SystemProductName, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.SystemProductNameUpdateFlag + L"\",";

		//ハードウェア種別
		fileLine1 += L"\"ハードウェア種別\",";
		fileLine2 += L"\"\",";
		fileLine3 += L"\"\",";

		//シリアル
		fileLine1 += L"\"シリアル\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.IdentifyingNumber, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.IdentifyingNumberUpdateFlag + L"\",";

		//コンピュータ名
		fileLine1 += L"\"コンピュータ名\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.ComputerName, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.ComputerNameUpdateFlag + L"\",";

		//CPUタイプ
		fileLine1 += L"\"CPU\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.ProcessorName, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.ProcessorNameUpdateFlag + L"\",";

		//CPU物理数
		fileLine1 += L"\"CPU物理数\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.NumberOfProcessors, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.NumberOfProcessorsUpdateFlag + L"\",";

		//CPUコア数
		fileLine1 += L"\"CPUコア数\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.NumberOfCores, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.NumberOfCoresUpdateFlag + L"\",";

		//CPU論理プロセッサ数
		fileLine1 += L"\"CPU論理プロセッサ数\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.NumberOfLogicalProcessors, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.NumberOfLogicalProcessorsUpdateFlag + L"\",";

		//CPUクロック数
		fileLine1 += L"\"CPUクロック数\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.ProcessorMaxClockSpeed, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.ProcessorMaxClockSpeedUpdateFlag + L"\",";

		//CPUソケット数
		fileLine1 += L"\"CPUソケット数\",";
		fileLine2 += L"\"\",";
		fileLine3 += L"\"\",";

		//メモリ容量
		fileLine1 += L"\"実装メモリ\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.TotalPhysicalMemory, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.TotalPhysicalMemoryUpdateFlag + L"\",";

		//ディスク容量
		fileLine1 += L"\"ディスク容量\",";
		fileLine2 += L"\"\",";
		fileLine3 += L"\"\",";

		//IPアドレス
		fileLine1 += L"\"IPアドレス\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.IPAddress, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.IPAddressUpdateFlag + L"\",";
		//tempField = Replace(currentHardwareInfo.IPAddress, L"\"", L"\"\"");
		//_ftprintf(fp, L"\"%s\",\"%s\",\"%s\"\n", L"IPアドレス", L"", tempField.c_str());
		//WritePrivateProfileString(L"Hardware", L"IPAddress", currentHardwareInfo.IPAddress.c_str(), datPath.c_str());

		//MACアドレス
		fileLine1 += L"\"MACアドレス\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.MACAddress, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.MACAddressUpdateFlag + L"\",";
		//tempField = Replace(currentHardwareInfo.MACAddress, L"\"", L"\"\"");
		//_ftprintf(fp, L"\"%s\",\"%s\",\"%s\"\n", L"MACアドレス", L"", tempField.c_str());
		//WritePrivateProfileString(L"Hardware", L"MACAddress", currentHardwareInfo.MACAddress.c_str(), datPath.c_str());

		//OS
		fileLine1 += L"\"OS\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.OSName, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.OSNameUpdateFlag + L"\",";
		//tempField = Replace(currentHardwareInfo.OSName, L"\"", L"\"\"");
		//_ftprintf(fp, L"\"%s\",\"%s\",\"%s\"\n", L"OS", L"", tempField.c_str());
		//WritePrivateProfileString(L"Hardware", L"OSName", currentHardwareInfo.OSName.c_str(), datPath.c_str());

		//ログインユーザ
		fileLine1 += L"\"ログインID\",";
		fileLine2 += L"\"" + Replace(currentHardwareInfo.UserName, L"\"", L"\"\"") + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.UserNameUpdateFlag + L"\",";
		//tempField = Replace(currentHardwareInfo.UserName, L"\"", L"\"\"");
		//_ftprintf(fp, L"\"%s\",\"%s\",\"%s\"\n", L"ログインユーザ", L"", tempField.c_str());
		//WritePrivateProfileString(L"Hardware", L"userName", currentHardwareInfo.UserName.c_str(), datPath.c_str());

		//仮想環境か
		fileLine1 += L"\"仮想環境\",";
		fileLine2 += L"\"" + currentHardwareInfo.Virtualization + L"\",";
		fileLine3 += L"\"" + currentHardwareInfo.VirtualizationUpdateFlag + L"\",";
		//tempField = Replace(currentHardwareInfo.OSName, L"\"", L"\"\"");
		//_ftprintf(fp, L"\"%s\",\"%s\",\"%s\"\n", L"OS", L"", tempField.c_str());
		//WritePrivateProfileString(L"Hardware", L"OSName", currentHardwareInfo.OSName.c_str(), datPath.c_str());

		//実行日時
		fileLine1 += L"\"実行日時\",";
		fileLine2 += L"\"" + currentHardwareInfo.Timestamp2 + L"\",";
		fileLine3 += L"\"\",";
		//_ftprintf(fp, L"\"%s\",\"%s\",\"%04d/%02d/%02d %02d:%02d:%02d\"\n", L"実行日時", L"", timestamp.tm_year + 1900, timestamp.tm_mon + 1, timestamp.tm_mday, timestamp.tm_hour, timestamp.tm_min, timestamp.tm_sec);

		//更新フラグ
		fileLine1 += L"\"更新フラグ\",";
		int indexUpdate = (int)fileLine3.find(L"UPDATE");
		if(indexUpdate != -1){
			fileLine2 += L"\"UPDATE\",";
		}else{
			fileLine2 += L"\"\",";
		}
		fileLine3 += L"\"\",";

		//ツールのバージョン
		//_ftprintf(fp, VER_STR_INVENTORYTOOLID);

		_ftprintf(fp, L"%s\n", fileLine1.c_str());
		_ftprintf(fp, L"%s\n", fileLine2.c_str());
		//_ftprintf(fp, L"%s\n", fileLine3.c_str());


		fclose(fp);
		returnString += outputFileName.c_str();
		return returnString;

}

//インベントリを保存する関数
long InventoryManager::save(){

	//フォルダが無ければ作成
	wstring datPath = _wgetenv(L"ALLUSERSPROFILE");
	datPath = datPath + L"\\LIP";
	_wmkdir(datPath.c_str());

	//フォルダ内のファイルを全削除（アクセス権の関係で可能であれば）
	//http://blog.systemjp.net/entry/20100318/p5
	//DeleteFile(datPath.c_str());

    HANDLE hFind;
    WIN32_FIND_DATA win32fd;//defined at Windwos.h
	wstring tempFilename;
	wstring searchPath = datPath + L"\\*.dat";
    hFind = FindFirstFile(searchPath.c_str(), &win32fd);

    if(hFind == INVALID_HANDLE_VALUE){
        //return -1;
		//エラー（ファイルが見つからない場合含む）は何もしない
    }
    do{
        if(win32fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY){
			//指定フォルダ内のサブフォルダは何もしない
        }
        else{
			//指定フォルダ内のdatファイルは削除
            tempFilename = win32fd.cFileName;
			tempFilename = datPath + L"\\" + tempFilename;
			//MessageBox(NULL,tempFilename.c_str(),L"test",MB_OK);
			DeleteFile(tempFilename.c_str());
        }
    }while(FindNextFile(hFind, &win32fd));

    FindClose(hFind);

	//データの書き込み
	//http://www.geocities.co.jp/SiliconValley/6071/technic/74.html
	//FILE *fp = fopen(datPath.c_str(), "wb");
	wstring savePath = datPath + L"\\" + currentHardwareInfo.Timestamp + L".dat";
	FILE* fp = NULL;
	_wfopen_s( &fp, savePath.c_str(), L"wb" );
	//MessageBox(NULL,datPath.c_str(),L"test",MB_OK);

	//fwriteWString(fp, L"test string.");

	vector<softwareInfo>::iterator softwareInfoIterator3 = currentSoftwareInfos.begin();  // イテレータのインスタンス化
	while(softwareInfoIterator3 != currentSoftwareInfos.end()){

		if((*softwareInfoIterator3).UpdateFlag != L"REMOVE"){
			//レジストリのサブキー
			fwriteWString(fp, L"sw.RegistoryKey");
			fwriteWString(fp, (*softwareInfoIterator3).RegistoryKey);
			//ソフトウェア名
			fwriteWString(fp, L"sw.Name");
			fwriteWString(fp, (*softwareInfoIterator3).Name);
			//ソフトウェアベンダー
			fwriteWString(fp, L"sw.Vendor");
			fwriteWString(fp, (*softwareInfoIterator3).Vendor);
			//ソフトウェアバージョン
			fwriteWString(fp, L"sw.Version");
			fwriteWString(fp, (*softwareInfoIterator3).Version);
			
			//保管しない
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
		//イテレータを1つ進める
		++softwareInfoIterator3;
	}

	vector<patchInfo>::iterator patchInfoIterator3 = currentPatchInfos.begin();  // イテレータのインスタンス化
	while(patchInfoIterator3 != currentPatchInfos.end()){

		if((*patchInfoIterator3).UpdateFlag != L"REMOVE"){
			//更新プログラム名
			fwriteWString(fp, L"hf.Name");
			fwriteWString(fp, (*patchInfoIterator3).Name);
			//ソフトウェア名
			fwriteWString(fp, L"hf.ParentName");
			fwriteWString(fp, (*patchInfoIterator3).ParentName);
			//ソフトウェアベンダー
			fwriteWString(fp, L"hf.Vendor");
			fwriteWString(fp, (*patchInfoIterator3).Vendor);
			//ソフトウェアバージョン
			fwriteWString(fp, L"hf.Version");
			fwriteWString(fp, (*patchInfoIterator3).Version);
			//比較名（ソフトウェア名＋更新プログラム名）
			fwriteWString(fp, L"hf.CompareName");
			fwriteWString(fp, (*patchInfoIterator3).CompareName);
			
			fwriteWString(fp, L"++hf");
		}
		//イテレータを1つ進める
		++patchInfoIterator3;
	}

	//管理番号
	fwriteWString(fp, L"hw.HardwareNoTitle");
	fwriteWString(fp, currentHardwareInfo.HardwareNoTitle);
	fwriteWString(fp, L"hw.HardwareNoValue");
	fwriteWString(fp, currentHardwareInfo.HardwareNoValue);

	//入力1～6
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

	//コンピュータ名
	fwriteWString(fp, L"hw.ComputerName");
	fwriteWString(fp, currentHardwareInfo.ComputerName);

	//IPアドレス
	fwriteWString(fp, L"hw.IPAddress");
	fwriteWString(fp, currentHardwareInfo.IPAddress);

	//MACアドレス
	fwriteWString(fp, L"hw.MACAddress");
	fwriteWString(fp, currentHardwareInfo.MACAddress);

	//メーカー名
	fwriteWString(fp, L"hw.SystemManufacturer");
	fwriteWString(fp, currentHardwareInfo.SystemManufacturer);

	//型番
	fwriteWString(fp, L"hw.SystemProductName");
	fwriteWString(fp, currentHardwareInfo.SystemProductName);

	//シリアル
	fwriteWString(fp, L"hw.IdentifyingNumber");
	fwriteWString(fp, currentHardwareInfo.IdentifyingNumber);

	//CPUタイプ
	fwriteWString(fp, L"hw.ProcessorName");
	fwriteWString(fp, currentHardwareInfo.ProcessorName);

	//CPU周波数
	fwriteWString(fp, L"hw.ProcessorMaxClockSpeed");
	fwriteWString(fp, currentHardwareInfo.ProcessorMaxClockSpeed);

	//CPU数
	fwriteWString(fp, L"hw.NumberOfProcessors");
	fwriteWString(fp, currentHardwareInfo.NumberOfProcessors);

	//コア数
	fwriteWString(fp, L"hw.NumberOfCores");
	fwriteWString(fp, currentHardwareInfo.NumberOfCores);

	//論理プロセッサ数
	fwriteWString(fp, L"hw.NumberOfLogicalProcessors");
	fwriteWString(fp, currentHardwareInfo.NumberOfLogicalProcessors);

	//ログインユーザ
	fwriteWString(fp, L"hw.UserName");
	fwriteWString(fp, currentHardwareInfo.UserName);

	//OS
	fwriteWString(fp, L"hw.OSName");
	fwriteWString(fp, currentHardwareInfo.OSName);

	//仮想環境か
	fwriteWString(fp, L"hw.Virtualization");
	fwriteWString(fp, currentHardwareInfo.Virtualization);

	//メモリ容量
	fwriteWString(fp, L"hw.TotalPhysicalMemory");
	fwriteWString(fp, currentHardwareInfo.TotalPhysicalMemory);

	//実行日時
	fwriteWString(fp, L"hw.Timestamp");
	fwriteWString(fp, currentHardwareInfo.Timestamp);

	//ツールのバージョン
	fwriteWString(fp, L"hw.ToolVersion");
	fwriteWString(fp, currentHardwareInfo.ToolVersion);

	fclose(fp);
	return 0;
}	

//前回のインベントリをロードする関数
long InventoryManager::load(){

	//フォルダが無ければ作成
	wstring datPath = _wgetenv(L"ALLUSERSPROFILE");
	datPath = datPath + L"\\LIP";
	_wmkdir(datPath.c_str());

	//フォルダ内のファイルを走査して、ファイル名の日時が現時点より前かつ最新のものをピックアップ
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
			//何もしない
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

	//データの読み込み
	//http://www.geocities.co.jp/SiliconValley/6071/technic/74.html
	//FILE *fp = fopen(datPath.c_str(), "rb");
	FILE* fp = NULL;
	_wfopen_s( &fp, loadFilePath.c_str(), L"rb" );

	wstring readString;
	//long r;
	//r = freadWString(fp, readString);
	//if(r>0) MessageBox(NULL,readString.c_str(),L"test",MB_OK);

	//softwareInfo構造体を用意
	softwareInfo tempSoftwareInfo;
	patchInfo tempPatchInfo;

	//ソフトウェアリストのフラグを全て「新規」にセット
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
			//tempSoftwareInfo（のレジストリキー）と合致するものがソフトウェアリストにあるか検索
			bool registoryKeyMatch = false;
			softwareInfoIterator = currentSoftwareInfos.begin();
			while(softwareInfoIterator != currentSoftwareInfos.end()){
				//合致するものがあった場合
				if(tempSoftwareInfo.RegistoryKey == (*softwareInfoIterator).RegistoryKey){
					//MessageBox(NULL,tempSoftwareInfo.registoryKey.c_str(),L"test",MB_OK);
					registoryKeyMatch = true;
					//ソフトウェア名も一致していればフラグは空にセット
					if(tempSoftwareInfo.Name == (*softwareInfoIterator).Name){
						(*softwareInfoIterator).UpdateFlag = L"";
					//ソフトウェア名が一致していなければソフトウェア名の更新があった
					}else{
						(*softwareInfoIterator).UpdateFlag = L"UPDATE";
					}
				}
				//イテレータを1つ進める
				++softwareInfoIterator;
			}
			///tempSoftwareInfo（のレジストリキー）と合致するものがソフトウェアリストに一件も無かった場合→ソフトウェアが削除された
			//2018/10/20 削除済みのソフトウェアも"REMOVE"とフラグを立てたうえでリストに出していたが、
			//要望によりリストに出さないように変更
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
			//tempPatchInfo（の名前）と合致するものがソフトウェアリストにあるか検索
			bool registoryKeyMatch = false;
			patchInfoIterator = currentPatchInfos.begin();
			while(patchInfoIterator != currentPatchInfos.end()){
				//合致するものがあった場合
				if(tempPatchInfo.CompareName == (*patchInfoIterator).CompareName){
					registoryKeyMatch = true;
					//フラグは空にセット
					(*patchInfoIterator).UpdateFlag = L"";
				}
				//イテレータを1つ進める
				++patchInfoIterator;
			}
			///tempSoftwareInfo（のレジストリキー）と合致するものがソフトウェアリストに一件も無かった場合→ソフトウェアが削除された
			//2018/10/20 削除済みのソフトウェアも"REMOVE"とフラグを立てたうえでリストに出していたが、
			//要望によりリストに出さないように変更
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
//wstringをファイルに書き込む関数
long InventoryManager::fwriteWString(FILE* fp, wstring string){

	//https://stackoverflow.com/questions/6975094/need-explanation-of-reading-and-writing-of-wchar-t-to-binary-file
	wchar_t* wc = (wchar_t*)string.c_str();
	unsigned int len = wcslen(wc) + 1;
	//まず識別符号を書き込み
	fwrite(&inventorySaveStringKey, sizeof(unsigned int), 1, fp);
	//文字列の長さを書き込み
	fwrite(&len, sizeof(unsigned int), 1, fp);
	//続いて文字列を書き込み
	fwrite(wc, len, sizeof(wchar_t), fp);
	//fwrite(&testString, sizeof(testString) ,1 , fp);

	//MessageBox(NULL,wc,L"test",MB_OK);
	return 0;
}
//wstringをファイルから読み込む関数
long InventoryManager::freadWString(FILE* fp, wstring &string){

	//https://stackoverflow.com/questions/6975094/need-explanation-of-reading-and-writing-of-wchar-t-to-binary-file
	unsigned int len;
	size_t r;
	//まず識別符号を読み取り
	r = fread(&len, sizeof(len), 1, fp);
	//ファイル終端か読み込みエラー→終了
	if(r == 0){
		if(debugLevel > 1){
			PrintDebugLog(L"freadWString: fread return 0.\n");
		}
		return -1;
	}
	//識別符号が違う→正しいファイルではない→終了
	if(len != inventorySaveStringKey){
		if(debugLevel > 1){
			PrintDebugLog(L"freadWString: inventorySaveStringKey unmatch.\n");
		}
		return -1;
	}
	//if(debugLevel > 1){
	//	PrintDebugLog(L"freadWString: inventorySaveStringKey match.\n");
	//}
	//文字列の長さを読み込み
	r = fread(&len, sizeof(len), 1, fp);
	//ファイル終端か読み込みエラー→終了
	if(r == 0){
		if(debugLevel > 1){
			PrintDebugLog(L"freadWString: fread return 0.\n");
		}
		return -1;
	}
	//文字列を読み込み
	wchar_t* wc = new wchar_t[len];
	fread(wc, len, sizeof(wchar_t), fp);
	//ファイル終端か読み込みエラー→終了
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



//前回取得した情報を取得する関数
//デフォルト入力値をロードする関数
long InventoryManager::loadDefaultSetting(){

	//デフォルト入力値のロード
	FILE* fp = NULL;
	//errno_t err_no;

	//エンコードはデフォルト（SJIS）
	//ファイル読み込みサンプル　http://pg-sample.sagami-ss.net/?eid=17
	//err_no = ;
	//ファイルオープンに成功した場合
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
			//csv中の引用符はとりあえず除去
			bufStr = Replace(bufStr, L"\"", L"");
			//中身がないカンマ続きはscanで読み込めないので、制御記号を挟む→あとで除去
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

				//デフォルト値がないときはフラグはそのまま
				if(identifyingNumberStr == L""){
					//match += 0;
				//デフォルト値があり、かつインベントリと等しいときはフラグを立てる
				}else if(identifyingNumberStr == currentHardwareInfo.IdentifyingNumber){
					//MessageBox(NULL,identifyingNumberStr.c_str(), VER_STR_PRODUCTNAME, MB_OK);
					match += 1;
				//デフォルト値があり、かつインベントリと異なるときはフラグを消す
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
//デフォルト入力値設定ファイルをリネームする関数
long InventoryManager::cleanupDefaultSetting(){

	wstring renameCSVPath = defaultCSVPath + L".bak";

	//既にあるdefault.csv.bakを削除しておく
	DeleteFile(renameCSVPath.c_str());

	//リネーム
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
//ソフトウェア名の構造体を列挙する関数
softwareInfo InventoryManager::getNextSoftware(){
	softwareInfo blankSoftwareInfo;
	return blankSoftwareInfo;
}
*/
//main
int _tmain(int argc, _TCHAR* argv[]){
    //GUI版はとりあえずコメントアウト
 /*
	//ロケールを自動設定しておく
	//参考 http://cx5software.com/article_vcpp_unicode/
	//setlocale関数の第二引数に「""」を指定すると、OSのデフォルトロケールが指定されます。例えば、日本語に設定されている場合「Japanese_Japan.932」が設定されます。当然ながら、OSの設定が変わると、それに合わせて変更されますので注意が必要です。
	_tsetlocale(LC_ALL, _T(""));
	//_tsetlocale(LC_ALL, L"Japanese");

    //Unicode文字列をコンソールに表示させる
    //_tsetlocale(LC_ALL, _tsetlocale(LC_CTYPE, L""));

	wstring hKey = L"HKEY_LOCAL_MACHINE";
    wstring subKey = L"";
	BOOL displayHelp = FALSE;
	

	//引数の解析
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
		_putws(L"「プログラムと機能」に表示されるソフトウェア名の一覧を出力します。");
		_putws(L"");
		_putws(L"LIP [/file ファイル名] [/encode エンコード]");
		_putws(L"");
		_putws(L"  /file ファイル名");
		_putws(L"    保存するファイル名を指定します。");
		_putws(L"    指定しない場合、「SoftwareList.csv」となります。");
		_putws(L"  /encode エンコード");
		_putws(L"    保存ファイルのエンコードを指定します。");
		_putws(L"    「UTF-8」「UTF-8N」「SJIS」のいずれかを指定します。");
		_putws(L"    指定しない場合「UTF-8」となります。");
		_putws(L"");
		return 0;
	}
	if(debugLevel > 0){
		for(i = 0; i < argc; ++i){
			_tprintf(TEXT("argv[%d]= %s\n"), i, argv[i]);
		}
	}


	// /subkeyがあったときは、テスト的に指定されたHKEY、SUBKEYの情報を取得してコンソールに出す
	if(subKey.size() != 0){
		RegList testRegList;
		_tprintf(TEXT("HKEY: %s\n"), hKey.c_str());
		_tprintf(TEXT("SUBKEY: %s\n"), hKey.c_str());
		testRegList.setHKeyAndSubKey(hKey, subKey, TRUE);
		testRegList.displayAll();
	}

	//メイン処理　ソフトウェア一覧の出力
	InventoryManager testInventoryManager;
	testInventoryManager.listUp();
	testInventoryManager.displayAll();

	_putws(L"");
	_tprintf(TEXT("ソフトウェアの一覧を「%s」に保存しました。\n"), outputFileName.c_str());

	_putws(L"終了するには何かキーを押してください。");
	//getchar();
	//ただgetchar()するだけではエンターキーしか受け付けない
	//http://stope.blog130.fc2.com/blog-entry-36.html
	//以下の警告が出たのでkbhit()を_kbhit()に変更
	//warning C4996: 'kbhit': The POSIX name for this item is deprecated. Instead, use the ISO C++ conformant name: _kbhit. See online help for details.
	for (i = 0; ; i++) {
		if (_kbhit()!=0) {
			break;
		}
	}
*/

	return 0;
}

