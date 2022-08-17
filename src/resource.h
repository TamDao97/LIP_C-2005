//{{NO_DEPENDENCIES}}
// Microsoft Visual C++ generated include file.
// Used by Resource.rc
//

// include the Resource IDs defined by Win32++
//エラー	1	fatal error RC1015: cannot open include file 'default_resource.h'.	d:\tool\Dropbox\work\2013\Yoji\project\LIP016\src\resource.h	7	
#include "default_resource.h"
//#include "resource.h"

//Resource.rcに埋め込むファイル情報の定義
//FileDescriptionを付けるとなぜか管理者権限が必要になる
//→理由が判明。install setup update などの特定キーワードが入っているとダメらしい
//http://stackoverflow.com/questions/4133337/how-do-i-avoid-uac-when-my-exe-file-name-contains-the-word-update
//http://blogs.wankuma.com/yo/archive/2007/10/04/99879.aspx
//http://blogs.wankuma.com/jitta/archive/2006/09/26/39667.aspx
//#define VER_STR_FILEDESCRIPTION  "List installed programs"
#define VER_STR_FILEDESCRIPTION  L"Output Add/Remove Programs"
#define VER_STR_COMMENTS         L"Output Add/Remove Programs"
#define VER_STR_INTERNALNAME     L"LIP\0"
#define VER_STR_ORIGINALFILENAME L"LIP.exe\0"
//#define VER_STR_LEGALCOPYRIGHT   L"Copyright (C) 2014 Cross Beat Co., Ltd.\0"
//#define VER_STR_COMPANYNAME      L"Cross Beat Co., Ltd.\0"
#define VER_STR_LEGALCOPYRIGHT   L"Copyright (C) 2014-2019 FIT Inc.\0"
#define VER_STR_COMPANYNAME      L"FIT Inc.\0"
#define VER_STR_PRODUCTNAME      L"LIP"
#define VER_FILEVERSION               1,8,3,2
#define VER_STR_FILEVERSION         L"1.8.3.2"
#define VER_PRODUCTVERSION            1,8,3,2
#define VER_STR_PRODUCTVERSION      L"1.8.3.2"
#define VER_STR_INVENTORYTOOLID L"LIP 1.8.3.2"


//タイトル
#define VER_STR_WINDOWCAPTION  L"LIP試用版"
//#define VER_STR_WINDOWCAPTION  L"LIP"

//SoftwarePicker.hに埋め込む有効期限の定義　画面にも埋め込む
#define VER_INT_PERIOD_YEAR  2021
#define VER_INT_PERIOD_MONTH   4
#define VER_INT_PERIOD_DAY     30
//#define VER_INT_PERIOD_YEAR  2018
//#define VER_INT_PERIOD_MONTH   3
//#define VER_INT_PERIOD_DAY     31

//画面に埋め込む提供お客様の定義
//#define VER_STR_CUSTOMER  L"for 三菱電機インフォメーションネットワーク"
//#define VER_STR_CUSTOMER  L"for 沖縄県農協電算センター"
#define VER_STR_CUSTOMER  L"広成建設株式会社"
//#define VER_STR_CUSTOMER  L"for 宮崎県"
//#define VER_STR_CUSTOMER  L"for 神戸市"
//#define VER_STR_CUSTOMER  L"for NGK"
//#define VER_STR_CUSTOMER  L"for IIJ"
//#define VER_STR_CUSTOMER  L""

//収集項目のタイトル
#define VER_STR_LABEL_HARDWARE_NO L"ハードウェア管理番号"
#define VER_STR_LABEL1 L"入力項目1"
#define VER_STR_LABEL2 L"入力項目2"
#define VER_STR_LABEL3 L"入力項目3"
#define VER_STR_LABEL4 L"入力項目4"
#define VER_STR_LABEL5 L"入力項目5"
#define VER_STR_LABEL6 L"入力項目6"

//SoftwarePicker.h用の出力設定
//保存ファイルのエンコード（デフォルトはUTF-8、その他UTF-8N、SJISも可）
#define VER_STR_OUTPUTFILEENCODE L"UTF-8"
//wstring outputFileEncode = L"SJIS";
//保存ファイルの形式（"SIMPLE","SARMS","AdvancedManager"）※SARMSのとき、エンコードはSJIS強制
//#define VER_STR_OUTPUTFILEFORMAT L"SARMS"
//#define VER_STR_OUTPUTFILEFORMAT L"SIMPLE"
#define VER_STR_OUTPUTFILEFORMAT L"AdvancedManager"
//MORE :保存ファイルがSIMPLE形式のとき、PCベンダー、PC機種、CPUタイプの列を追加する
//MORE2:保存ファイルがSIMPLE形式のとき、PCベンダー、PC機種、シリアル番号、CPUタイプの列を追加する
//MORE3:保存ファイルがSIMPLE形式のとき、取得日時、PCベンダー、PC機種、シリアル番号、CPUタイプの列を追加する
//#define VER_STR_OUTPUTFILEOPTION L""
#define VER_STR_OUTPUTFILEOPTION L"MORE3"



/*
SXS_MANIFEST_RESOURCE_ID=1
SXS_MANIFEST=foo.manifest
SXS_ASSEMBLY_NAME=Microsoft.Windows.Foo
SXS_ASSEMBLY_VERSION=1.0	
SXS_ASSEMBLY_LANGUAGE_INDEPENDENT=1
SXS_MANIFEST_IN_RESOURCES=1
*/


//The resource ID for MENU, ICON, ToolBar Bitmap, Accelerator,
//  and Window Caption

//Resource IDs for the dialog
#define IDD_DIALOG1                     101
//#define IDC_RADIO1						110
//#define IDC_RADIO2						111
//#define IDC_RADIO3						112
//#define IDC_CHECK1						113
//#define IDC_CHECK2						114
//#define IDC_CHECK3						115
#define IDC_EDIT_HARDWARE_NO						110
#define IDC_EDIT1						120
#define IDC_EDIT2						130
#define IDC_EDIT3						140
#define IDC_EDIT4						150
#define IDC_EDIT5						160
#define IDC_EDIT6						170
#define IDC_STATIC_HARDWARE_NO                     210
#define IDC_STATIC1                     220
#define IDC_STATIC2                     230
#define IDC_STATIC3                     240
#define IDC_STATIC4                     250
#define IDC_STATIC5                     260
#define IDC_STATIC6                     270
#define IDC_STATIC_OK                     280
//#define IDC_LIST1						121
//#define IDC_BUTTON1						200
//#define IDC_RICHEDIT1					123
//#define IDC_STATIC1                     130
//#define IDC_STATIC2                     131
//#define IDC_STATIC3                     132
//#define IDC_HOTKEY1                     140
//#define IDB_BITMAP1                     150

// Next default values for new objects
// 
#ifdef APSTUDIO_INVOKED
#ifndef APSTUDIO_READONLY_SYMBOLS
#define _APS_NO_MFC                     1
#define _APS_NEXT_RESOURCE_VALUE        129
#define _APS_NEXT_COMMAND_VALUE         32771
#define _APS_NEXT_CONTROL_VALUE         1000
#define _APS_NEXT_SYMED_VALUE           130
#endif
#endif

