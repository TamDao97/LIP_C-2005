//{{NO_DEPENDENCIES}}
// Microsoft Visual C++ generated include file.
// Used by Resource.rc
//

// include the Resource IDs defined by Win32++
//�G���[	1	fatal error RC1015: cannot open include file 'default_resource.h'.	d:\tool\Dropbox\work\2013\Yoji\project\LIP016\src\resource.h	7	
#include "default_resource.h"
//#include "resource.h"

//Resource.rc�ɖ��ߍ��ރt�@�C�����̒�`
//FileDescription��t����ƂȂ����Ǘ��Ҍ������K�v�ɂȂ�
//�����R�������Binstall setup update �Ȃǂ̓���L�[���[�h�������Ă���ƃ_���炵��
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


//�^�C�g��
#define VER_STR_WINDOWCAPTION  L"LIP���p��"
//#define VER_STR_WINDOWCAPTION  L"LIP"

//SoftwarePicker.h�ɖ��ߍ��ޗL�������̒�`�@��ʂɂ����ߍ���
#define VER_INT_PERIOD_YEAR  2021
#define VER_INT_PERIOD_MONTH   4
#define VER_INT_PERIOD_DAY     30
//#define VER_INT_PERIOD_YEAR  2018
//#define VER_INT_PERIOD_MONTH   3
//#define VER_INT_PERIOD_DAY     31

//��ʂɖ��ߍ��ޒ񋟂��q�l�̒�`
//#define VER_STR_CUSTOMER  L"for �O�H�d�@�C���t�H���[�V�����l�b�g���[�N"
//#define VER_STR_CUSTOMER  L"for ���ꌧ�_���d�Z�Z���^�["
#define VER_STR_CUSTOMER  L"�L�����݊������"
//#define VER_STR_CUSTOMER  L"for �{�茧"
//#define VER_STR_CUSTOMER  L"for �_�ˎs"
//#define VER_STR_CUSTOMER  L"for NGK"
//#define VER_STR_CUSTOMER  L"for IIJ"
//#define VER_STR_CUSTOMER  L""

//���W���ڂ̃^�C�g��
#define VER_STR_LABEL_HARDWARE_NO L"�n�[�h�E�F�A�Ǘ��ԍ�"
#define VER_STR_LABEL1 L"���͍���1"
#define VER_STR_LABEL2 L"���͍���2"
#define VER_STR_LABEL3 L"���͍���3"
#define VER_STR_LABEL4 L"���͍���4"
#define VER_STR_LABEL5 L"���͍���5"
#define VER_STR_LABEL6 L"���͍���6"

//SoftwarePicker.h�p�̏o�͐ݒ�
//�ۑ��t�@�C���̃G���R�[�h�i�f�t�H���g��UTF-8�A���̑�UTF-8N�ASJIS���j
#define VER_STR_OUTPUTFILEENCODE L"UTF-8"
//wstring outputFileEncode = L"SJIS";
//�ۑ��t�@�C���̌`���i"SIMPLE","SARMS","AdvancedManager"�j��SARMS�̂Ƃ��A�G���R�[�h��SJIS����
//#define VER_STR_OUTPUTFILEFORMAT L"SARMS"
//#define VER_STR_OUTPUTFILEFORMAT L"SIMPLE"
#define VER_STR_OUTPUTFILEFORMAT L"AdvancedManager"
//MORE :�ۑ��t�@�C����SIMPLE�`���̂Ƃ��APC�x���_�[�APC�@��ACPU�^�C�v�̗��ǉ�����
//MORE2:�ۑ��t�@�C����SIMPLE�`���̂Ƃ��APC�x���_�[�APC�@��A�V���A���ԍ��ACPU�^�C�v�̗��ǉ�����
//MORE3:�ۑ��t�@�C����SIMPLE�`���̂Ƃ��A�擾�����APC�x���_�[�APC�@��A�V���A���ԍ��ACPU�^�C�v�̗��ǉ�����
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

