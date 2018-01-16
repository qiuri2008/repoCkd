#if !defined(AFX_PROP1_H__E4A943D0_8116_4CF7_8E6D_3BC6FB33CE61__INCLUDED_)
#define AFX_PROP1_H__E4A943D0_8116_4CF7_8E6D_3BC6FB33CE61__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Prop1.h : header file
//

#include "excel2003.h"
/////////////////////////////////////////////////////////////////////////////
// CPageShipMark dialog

class CPageShipMark : public CPropertyPage
{
	DECLARE_DYNCREATE(CPageShipMark)

// Construction
public:
	CPageShipMark();
	~CPageShipMark();
	void SaveShippingMark1(CExcel &Excel);

// Dialog Data
	//{{AFX_DATA(CPageShipMark)
	enum { IDD = IDD_PAGE_SHIPMARK };
	CString	m_cs11;
	CString	m_cs12;
	CString	m_cs13;
	CString	m_cs14;
	CString	m_cs21;
	CString	m_cs22;
	CString	m_cs23;
	CString	m_cs24;
	CString	m_cs31;
	CString	m_cs32;
	CString	m_cs33;
	CString	m_cs34;
	int		m_nRadio1;
	BOOL	m_bSpeak;
	//}}AFX_DATA


// Overrides
	// ClassWizard generate virtual function overrides
	//{{AFX_VIRTUAL(CPageShipMark)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CPageShipMark)
	afx_msg void OnBtnLoadFile();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PROP1_H__E4A943D0_8116_4CF7_8E6D_3BC6FB33CE61__INCLUDED_)
