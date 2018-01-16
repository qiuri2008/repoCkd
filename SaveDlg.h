#if !defined(AFX_SAVEDLG_H__4A026467_A962_42B8_9BE1_B06C8F968DEF__INCLUDED_)
#define AFX_SAVEDLG_H__4A026467_A962_42B8_9BE1_B06C8F968DEF__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// SaveDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CSaveDlg dialog

class CSaveDlg : public CDialog
{
// Construction
public:
	CSaveDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CSaveDlg)
	enum { IDD = IDD_DLG_SAVE };
	BOOL	m_bBoxList;
	BOOL	m_bDayCount;
	BOOL	m_bShipMark;
	BOOL	m_bSaveDayOnly;
	int		m_iSaveMode;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSaveDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CSaveDlg)
		// NOTE: the ClassWizard will add member functions here
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
public:
//	afx_msg void OnBnClickedOk();
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_SAVEDLG_H__4A026467_A962_42B8_9BE1_B06C8F968DEF__INCLUDED_)
