#if !defined(AFX_PROP3_H__9D59055C_9BE5_4475_8746_D5C3CEAA6449__INCLUDED_)
#define AFX_PROP3_H__9D59055C_9BE5_4475_8746_D5C3CEAA6449__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Prop3.h : header file
//
#include "SortListCtrl.h"
/////////////////////////////////////////////////////////////////////////////
// CPageList dialog

class CPageList : public CPropertyPage
{
	DECLARE_DYNCREATE(CPageList)

// Construction
public:
	CPageList();
	~CPageList();

// Dialog Data
	//{{AFX_DATA(CPageList)
	enum { IDD = IDD_PAGE_LIST };
	CSortListCtrl	m_ListMain;
	CString	m_csReceive;
	//}}AFX_DATA


// Overrides
	// ClassWizard generate virtual function overrides
	//{{AFX_VIRTUAL(CPageList)
	public:
	virtual BOOL OnSetActive();
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual void CalcWindowRect(LPRECT lpClientRect, UINT nAdjustType = adjustBorder);
	//}}AFX_VIRTUAL

// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CPageList)
	afx_msg void OnBtnImportBaseList();
	afx_msg void OnBtnSaveBaseList();
	afx_msg void OnBtnSaveBom();
	afx_msg void OnBtnLoad();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

private :
	bool m_listInit;

public :
	//void MessageBox1(CString strText,CString strCaption = "",UINT nType=MB_OK);

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PROP3_H__9D59055C_9BE5_4475_8746_D5C3CEAA6449__INCLUDED_)
