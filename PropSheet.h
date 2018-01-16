#if !defined(AFX_PROPSHEET_H__FAE94575_0F47_483F_A019_FC2902D55ADF__INCLUDED_)
#define AFX_PROPSHEET_H__FAE94575_0F47_483F_A019_FC2902D55ADF__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// PropSheet.h : header file
//
#include "PageShipMark.h"
#include "PageCkd.h"
#include "PageList.h"
#include "PageCheck.h"
#include "AutoWeight.h"
/////////////////////////////////////////////////////////////////////////////
// CPropSheet

class CPropSheet : public CPropertySheet
{
	DECLARE_DYNAMIC(CPropSheet)

// Construction
public:
	CPropSheet();
	CPropSheet(UINT nIDCaption, CWnd* pParentWnd = NULL, UINT iSelectPage = 0);
	CPropSheet(LPCTSTR pszCaption, CWnd* pParentWnd = NULL, UINT iSelectPage = 0);

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CPropSheet)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CPropSheet();

	// Generated message map functions
protected:
	//{{AFX_MSG(CPropSheet)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

private:
	CPageShipMark	m_pageShipMark;
	CPageCkd		m_pageCkd;
	CPageList		m_pageList;
	CPageCheck		m_pageCheck;
	CAutoWeight		m_pageWeight;
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PROPSHEET_H__FAE94575_0F47_483F_A019_FC2902D55ADF__INCLUDED_)
