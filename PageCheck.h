#pragma once
#include "afxcmn.h"
#include "SortListCtrl.h"
#include "afxwin.h"

// CPageCheck 对话框

class CPageCheck : public CPropertyPage
{
	DECLARE_DYNAMIC(CPageCheck)

public:
	CPageCheck();
	virtual ~CPageCheck();

// 对话框数据
	enum { IDD = IDD_PAGE_CHECK };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	CString m_csReceive;
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	CSortListCtrl m_ListMain;
	virtual BOOL OnSetActive();
	bool m_listInit;
	int m_iRadio;
	void CheckProvider(void);
	void LibraryLogin(void);
	void LibraryLogout(void);
	CString m_csKitking;
	CString m_csProvider;
	void CheckUpdateData(void);
	afx_msg void OnBnClickedRadioCheck();
	afx_msg void OnBnClickedRadioLogin();
	afx_msg void OnBnClickedRadioLogout();
	void ReloadData(void);
	CString m_strZCS;
	BOOL m_flgTaojian;
	CString m_csQty;
	afx_msg void OnBnClickedBtnLoad();
	afx_msg void OnBnClickedBtnLogin();
	afx_msg void OnBnClickedBtnLogout();
};
