// CKD录入Dlg.cpp : implementation file
//

#include "stdafx.h"
#include "CKD录入.h"
#include "CKD录入Dlg.h"

#define _GLOBAL_C_

#include "Global.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
		// No message handlers
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CCKDDlg dialog

CCKDDlg::CCKDDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CCKDDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CCKDDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CCKDDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CCKDDlg)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CCKDDlg, CDialog)
	//{{AFX_MSG_MAP(CCKDDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


void CCKDDlg::InitConfigFile(BOOL read)
{
	TCHAR lpFileName[MAX_PATH];
	GetModuleFileName(AfxGetInstanceHandle(),lpFileName,MAX_PATH);

	CString strFileName = lpFileName;
	int nIndex = strFileName.ReverseFind ('\\');
	
	CString strPath;
	CString cs;

	if (nIndex > 0)
		strPath = strFileName.Left (nIndex);
	else
		strPath = "";
	strPath += "\\Config.ini";

	HANDLE h;
	LPWIN32_FIND_DATA pFD=new WIN32_FIND_DATA;
	BOOL bFound=FALSE;
	if(pFD)
	{
		h=FindFirstFile(strPath,pFD);
		bFound=(h!=INVALID_HANDLE_VALUE);
		if(bFound)
		{
			FindClose(h);
		}
		delete pFD;
	}
	if(!bFound)
	{
		MessageBox("找不到配置文件");
		return ;
	}

	if(read)
	{
		g_bUseDefine = GetPrivateProfileInt(_T("SECTION 1"), _T("自定义显示"), 0,strPath) ;
		g_iDisplayWith = GetPrivateProfileInt(_T("SECTION 1"), _T("显示宽度"), 0,strPath) ;
		g_iDisplayHigh = GetPrivateProfileInt(_T("SECTION 1"), _T("显示高度"), 0,strPath) ;
		g_iSaveMode = GetPrivateProfileInt(_T("SECTION 1"), _T("唛头保存模式"), 0,strPath) ;
	}
	else
	{
		cs.Format("%d",g_iSaveMode);
		WritePrivateProfileString(_T("SECTION 1"),_T("唛头保存模式"),cs,strPath);
	}
	cs.Format("%d-%d-%d",g_iDisplayWith,g_iDisplayHigh,g_iSaveMode);
	//MessageBox(cs);
}


/////////////////////////////////////////////////////////////////////////////
// CCKDDlg message handlers

BOOL CCKDDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	
	// TODO: Add extra initialization here
	InitConfigFile(true);

	{
		int cx = GetSystemMetrics( SM_CXSCREEN ); 
		int cy = GetSystemMetrics( SM_CYSCREEN );
		DEVMODE DevMode;  //屏幕信息结构体
		EnumDisplaySettings(NULL,ENUM_CURRENT_SETTINGS,&DevMode);
		if(DevMode.dmPelsWidth*10 / cx != 10 || DevMode.dmPelsHeight*10 / cy != 10  )    //实际分辨率与缩放后分辨率不同，则存在缩放
		{
			;//MoveWindow(cx/5,cy*3/20,cx/10*12/2,cy/10*7);
		}


		if(g_bUseDefine)
			MoveWindow(cx/5,cy*3/20,g_iDisplayWith,g_iDisplayHigh);
		else
		{
			if(DevMode.dmPelsWidth == 1920)
				MoveWindow(cx/5,cy*3/20,DevMode.dmPelsWidth*10/(24),DevMode.dmPelsHeight*10/18);
			else
				MoveWindow(cx/5,cy*3/20,DevMode.dmPelsWidth*10/(20),DevMode.dmPelsHeight*10/17);
		}

		#ifdef _DEBUG
		TRACE("------------------>Dlg Debug: %d-%d-%d\n", cx,DevMode.dmPelsWidth,cx*cx*10/(24));
		#endif
	}


	if (::CoInitialize(NULL)!=0) 
	{ 
		//AfxMessageBox("初始化COM支持库失败!"); 
		//exit(1); 
	} 

	m_vSheet.Create(this,WS_CHILD | WS_VISIBLE,WS_EX_CONTROLPARENT);
	RECT rect;
	m_vSheet.GetWindowRect(&rect);
	int width = rect.right - rect.left;
	int height = rect.bottom - rect.top;

	m_vSheet.SetWindowPos(NULL,0,0,width,height,SWP_NOZORDER | SWP_NOACTIVATE);
	m_vSheet.SetActivePage(0);

	//======================================================
	//前端显示
	//SetWindowPos(&wndTopMost,0,0,0,0,SWP_NOSIZE|SWP_NOMOVE);
	//======================================================

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CCKDDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CCKDDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CCKDDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}
