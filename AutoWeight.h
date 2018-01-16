#if !defined(AFX_AUTOWEIGHT_H__7B935E5F_CF5C_458C_99F7_1D1D09740822__INCLUDED_)
#define AFX_AUTOWEIGHT_H__7B935E5F_CF5C_458C_99F7_1D1D09740822__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// AutoWeight.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CAutoWeight dialog

class CAutoWeight : public CPropertyPage
{
	DECLARE_DYNCREATE(CAutoWeight)

// Construction
public:
	CAutoWeight();
	~CAutoWeight();

// Dialog Data
	//{{AFX_DATA(CAutoWeight)
	enum { IDD = IDD_PAGE_WEIGHT };
	CSortListCtrl	m_ListMain;
	CString	m_csPallet;
	CString	m_csOutBox;
	CString	m_csReceive;
	CString	m_strZCS;
	CString	m_srtPalletWeight;
	CString	m_strOutBoxWeight;
	//}}AFX_DATA


// Overrides
	// ClassWizard generate virtual function overrides
	//{{AFX_VIRTUAL(CAutoWeight)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual BOOL OnSetActive();
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CAutoWeight)
	afx_msg void OnBtnLoadWeightList();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg void OnBtnSave();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

private:
	bool m_listInit;
	double m_dPalletWeight;

public :
	double GetOutBoxWeight(CString strOutbox,CString strPallet="", bool bAddFlag=false);
	void UpdateList(void);
	double PalletWeight(CString csPallet);
	void SaveInvoiceDate(CExcel &Excel) ;

	void ReloadData(void);
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_AUTOWEIGHT_H__7B935E5F_CF5C_458C_99F7_1D1D09740822__INCLUDED_)
