#include "stdafx.h"
#include "CKD录入.h"
#include "PageCkd.h"
#include <bitset>
#include "Global.h"

#include "excel2003.h"
#include "user.h"
#include "PropSheet.h"

BOOL g_bUseDefine = true;
unsigned int g_iDisplayWith = 0;
unsigned int g_iDisplayHigh = 0;
unsigned int g_iSaveMode = 0;

void Speak(CString m_sText)
{
	tts.Create(SP_CHINESE);//SP_ENGLISH//SP_CHINESE
	g_ttsCreate = TRUE;
	tts.SetVolume( 100 );
	tts.SetRate( 0 );
	MultiByteToWideChar( CP_ACP, //把char转换成wchar_t
						  0,
						  m_sText,
						  m_sText.GetLength() * sizeof(char),
						  g_wcRead,
						  m_sText.GetLength() * sizeof( wchar_t ) );
	//tts.SetVolume ( 100 );		//设置音量  0-100
	tts.SetRate( 0 );			//设置语速  -10-10	
	tts.Speak( g_wcRead );
	for(int i=0;i<m_sText.GetLength();i++)
		g_wcRead[i] = 0;
	//memset(g_wcRead,0,m_sText.GetLength()-1);  //用这句有问题
}

void MessageBox1(CString strText,CString strCaption,UINT nType)
{
//	UpdateData(TRUE);
	if(1)//g_flgSpeak)
	{
		if(strText.GetLength() >= SPEAK_BUFFSIZE-1)
			strText = strText.Left(SPEAK_BUFFSIZE-1);
		Speak(strText);
	}
	else
		;//MessageBoxA(strText,strCaption,nType);

}

BOOL IsFileExist(CString strFn, BOOL bDir)
{
    HANDLE h;
	LPWIN32_FIND_DATA pFD=new WIN32_FIND_DATA;
	BOOL bFound=FALSE;
	if(pFD)
	{
		h=FindFirstFile(strFn,pFD);
		bFound=(h!=INVALID_HANDLE_VALUE);
		if(bFound)
		{
			if(bDir)
				bFound= (pFD->dwFileAttributes&FILE_ATTRIBUTE_DIRECTORY)!=NULL;
			FindClose(h);
		}
		delete pFD;
	}
	return bFound;
}


void InitConfigFile(BOOL read)
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


