// CKDÂ¼Èë.h : main header file for the CKDÂ¼Èë application
//

#if !defined(AFX_CKD_H__386A5C4E_BBC1_4702_8523_6092752E305F__INCLUDED_)
#define AFX_CKD_H__386A5C4E_BBC1_4702_8523_6092752E305F__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CCKDApp:
// See CKDÂ¼Èë.cpp for the implementation of this class
//

class CCKDApp : public CWinApp
{
public:
	CCKDApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CCKDApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CCKDApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_CKD_H__386A5C4E_BBC1_4702_8523_6092752E305F__INCLUDED_)
