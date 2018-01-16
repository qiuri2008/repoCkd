// CKDÂ¼ÈëDlg.h : header file
//

#if !defined(AFX_CKDDLG_H__AB1D0B68_9F0F_4774_8803_454C8A65B450__INCLUDED_)
#define AFX_CKDDLG_H__AB1D0B68_9F0F_4774_8803_454C8A65B450__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
#include "PropSheet.h"
/////////////////////////////////////////////////////////////////////////////
// CCKDDlg dialog

class CCKDDlg : public CDialog
{
// Construction
public:
	CCKDDlg(CWnd* pParent = NULL);	// standard constructor
	CPropSheet m_vSheet;
	void InitConfigFile(BOOL read);

// Dialog Data
	//{{AFX_DATA(CCKDDlg)
	enum { IDD = IDD_CKD_DIALOG };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CCKDDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CCKDDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_CKDDLG_H__AB1D0B68_9F0F_4774_8803_454C8A65B450__INCLUDED_)
