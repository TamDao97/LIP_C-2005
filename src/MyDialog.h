///////////////////////////////////////
// MyDialog.h

#ifndef MYDIALOG_H
#define MYDIALOG_H



// Declaration of the CMyDialog class
class CMyDialog : public CDialog
{
public:
	CMyDialog(UINT nResID, CWnd* pParent = NULL);
	virtual ~CMyDialog();

protected:
	virtual BOOL OnInitDialog() sealed;
	virtual INT_PTR DialogProc(UINT uMsg, WPARAM wParam, LPARAM lParam);
	virtual BOOL OnCommand(WPARAM wParam, LPARAM lParam);
	virtual void OnOK();
	virtual void OnDraw(CDC* pDC);	

private:
	void OnButton();
	//void OnRadio1();
	//void OnRadio2();
	//void OnRadio3();
	//void OnCheck1();
	//void OnCheck2();
	//void OnCheck3();

	HMODULE m_hInstRichEdit;
};

#endif //MYDIALOG_H
