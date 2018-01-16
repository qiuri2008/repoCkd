#if !defined(AFX_PROP2_H__396E8C12_D168_44C0_965A_7A6518DA1B66__INCLUDED_)
#define AFX_PROP2_H__396E8C12_D168_44C0_965A_7A6518DA1B66__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Prop2.h : header file
//
#include "SortListCtrl.h"
#include "Excel2003.h"
/////////////////////////////////////////////////////////////////////////////
// CPageCkd dialog

class CPageCkd : public CPropertyPage
{
	DECLARE_DYNCREATE(CPageCkd)

// Construction
public:
	CPageCkd();
	~CPageCkd();
	void SaveShippingMark(CExcel &Excel);
	void SaveAccordingDate(CExcel &Excel);
	void SaveBoxList(CExcel &Excel);
	void PlayVoice(void);
	//void Speak(CString m_sText);
	//void MessageBox1(CString strText,CString strCaption = "",UINT nType=MB_OK);
	unsigned int m_iSaveMode;
	BOOL m_bOnlyOnce;

// Dialog Data
	//{{AFX_DATA(CPageCkd)
	enum { IDD = IDD_PAGE_CKD};
	CSortListCtrl	m_ListMain;
	CString	m_csReceive;
	CString	m_csPartNo;
	CString	m_csInbox;
	CString	m_csOutBox;
	CString	m_csQty;
	CString	m_csUndoInbox;
	UINT	m_unRatio;
	CString	m_strZCS;
	//}}AFX_DATA


// Overrides
	// ClassWizard generate virtual function overrides
	//{{AFX_VIRTUAL(CPageCkd)
	public:
	virtual BOOL OnSetActive();
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CPageCkd)
	afx_msg void OnButton1();
	afx_msg void OnButton2();
	afx_msg void OnBtnSave();
	afx_msg void OnBtnLoadBaseList();
	afx_msg void OnBtnReload();
	afx_msg void OnBtnCheck();
	afx_msg void OnSetfocusEditUndo();
	afx_msg void OnTimer(UINT nIDEvent);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
private:
	bool m_listInit;

public:
	void ReloadData(BOOL bActive = FALSE);
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PROP2_H__396E8C12_D168_44C0_965A_7A6518DA1B66__INCLUDED_)
