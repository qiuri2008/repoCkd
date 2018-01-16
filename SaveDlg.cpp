// SaveDlg.cpp : implementation file
//

#include "stdafx.h"
#include "CKD录入.h"
#include "SaveDlg.h"
#include ".\savedlg.h"
#include "Global.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CSaveDlg dialog


CSaveDlg::CSaveDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CSaveDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CSaveDlg)
	m_bBoxList = TRUE;
	m_bDayCount = TRUE;
	m_bShipMark = TRUE;
	m_bSaveDayOnly = FALSE;
	m_iSaveMode = g_iSaveMode;
	//}}AFX_DATA_INIT
}


void CSaveDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CSaveDlg)
	DDX_Check(pDX, IDC_CHECK1, m_bBoxList);
	DDX_Check(pDX, IDC_CHECK2, m_bDayCount);
	DDX_Check(pDX, IDC_CHECK3, m_bShipMark);
	DDX_Check(pDX, IDC_CHECK4, m_bSaveDayOnly);
	DDX_Radio(pDX, IDC_RADIO1, m_iSaveMode);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CSaveDlg, CDialog)
	//{{AFX_MSG_MAP(CSaveDlg)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
//	ON_BN_CLICKED(IDOK, OnBnClickedOk)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CSaveDlg message handlers
  
//void CSaveDlg::OnBnClickedOk()
//{
//	// TODO: 在此添加控件通知处理程序代码
//	OnOK();
//}
