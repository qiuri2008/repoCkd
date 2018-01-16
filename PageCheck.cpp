// PageCheck.cpp : 实现文件
//

#include "stdafx.h"
#include "CKD录入.h"
#include "PageCheck.h"
#include ".\pagecheck.h"
#include "excel2003.h"
#include "Global.h"
#include <bitset>
#include "PropSheet.h"



using std::bitset;
bitset<2> g_bitCheck;


// CPageCheck 对话框

IMPLEMENT_DYNAMIC(CPageCheck, CPropertyPage)
CPageCheck::CPageCheck()
	: CPropertyPage(CPageCheck::IDD)
	, m_csReceive(_T(""))
	, m_iRadio(0)
	, m_csKitking(_T(""))
	, m_csProvider(_T(""))
	, m_strZCS(_T(""))
	, m_csQty(_T(""))
{
	m_listInit = false;
}

CPageCheck::~CPageCheck()
{
}

void CPageCheck::DoDataExchange(CDataExchange* pDX)
{
	CPropertyPage::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT_CHECK_RECEIVE, m_csReceive);
	DDX_Control(pDX, IDC_LIST1, m_ListMain);
	DDX_Radio(pDX, IDC_RADIO_CHECK, m_iRadio);
	DDX_Text(pDX, IDC_EDIT_KITKING, m_csKitking);
	DDX_Text(pDX, IDC_EDIT_PROVIDER, m_csProvider);
	DDX_Text(pDX, IDC_EDIT_ZCS, m_strZCS);
	DDX_Text(pDX, IDC_EDIT_QTY, m_csQty);
}


BEGIN_MESSAGE_MAP(CPageCheck, CPropertyPage)
//	ON_BN_CLICKED(IDC_BUTTON1, OnBnClickedButton1)
//	ON_BN_CLICKED(IDC_BUTTON2, OnBnClickedButton2)
ON_BN_CLICKED(IDC_RADIO_CHECK, OnBnClickedRadioCheck)
ON_BN_CLICKED(IDC_RADIO_LOGIN, OnBnClickedRadioLogin)
ON_BN_CLICKED(IDC_RADIO_LOGOUT, OnBnClickedRadioLogout)
ON_BN_CLICKED(IDC_BTN_LOAD, OnBnClickedBtnLoad)
ON_BN_CLICKED(IDC_BTN_LOGIN, OnBnClickedBtnLogin)
ON_BN_CLICKED(IDC_BTN_LOGOUT, OnBnClickedBtnLogout)
END_MESSAGE_MAP()


// CPageCheck 消息处理程序

BOOL CPageCheck::OnSetActive()
{
	// TODO: 在此添加专用代码和/或调用基类
	DWORD dwStyle;
	if(!m_listInit)
	{
		dwStyle=m_ListMain.GetExtendedStyle();
		dwStyle|=LVS_EX_FULLROWSELECT;
		dwStyle|=LVS_EX_GRIDLINES;

		(void)m_ListMain.SetExtendedStyle( dwStyle );///整行选择模式//LVS_EX_FULLROWSELECT
		m_ListMain.SetHeadings("序号,50;部品号,110;供应商号,110;数量,50"); ///设置列头信息 
		m_ListMain.LoadColumnInfo(); 

		CString strFileName = CExcel::GetAppPath() + "\\基准库\\" + "BaseProvider" + ".ckd";
		CString strPartNo;
		
		if(IsFileExist(strFileName,FALSE)==TRUE)
		{
			mapCheck.RemoveAll();
			m_ListMain.DeleteAllItems();
			//序列化文件
			CFile file;
			file.Open(strFileName, CFile::modeReadWrite);
			CArchive ar(&file, CArchive::load);
			mapCheck.Serialize(ar);
			ar.Close();
			file.Close();

			CheckUpdateData();
		
		}
		else
		{
			;//MessageBox("序号清单未载入!", "载入", MB_ICONWARNING);
		}


		m_listInit = true;
		m_flgTaojian = FALSE;
	}

	CEdit* pEdit = (CEdit*)GetDlgItem(IDC_EDIT_CHECK_RECEIVE);
	pEdit->SetFocus();

	return CPropertyPage::OnSetActive();
}

BOOL CPageCheck::PreTranslateMessage(MSG* pMsg)
{
	// TODO: 在此添加专用代码和/或调用基类
	if(pMsg->message == WM_KEYDOWN)
	{
		if(pMsg->wParam == VK_RETURN)
		{
			UINT nID = GetFocus()->GetDlgCtrlID();
			switch(nID)
			{
			case IDC_EDIT_ZCS:
				ReloadData();
				return true;
				break;
			case IDC_EDIT_CHECK_RECEIVE:
				UpdateData(true);
				m_csReceive.MakeUpper();
				if(m_csReceive.Find('(') !=-1 && m_csReceive.Find(')') != -1 && m_csReceive.Find('-') != -1
					&& (m_csReceive.SpanExcluding("()").GetLength() == 13||m_csReceive.SpanExcluding("()").GetLength() == 14||m_csReceive.SpanExcluding("()").GetLength() == 15)
					)
				{
					m_csKitking = m_csReceive.SpanExcluding("()");
					m_csQty = m_csReceive.Mid(m_csReceive.Find('(')+1,m_csReceive.Find(')')-m_csReceive.Find('(')-1);
					g_bitCheck.set(STYLE_PART);
					m_csReceive.Empty();
				}
				else
				{
					m_csProvider = m_csReceive;//.SpanExcluding(" ");
					g_bitCheck.set(STYLE_PROVIDER);
					m_csReceive.Empty();
				}
				UpdateData(FALSE);
				
				if(g_bitCheck.count() == STYLE_END_CHECK)
				{
					
					g_bitCheck.reset();
					if(m_iRadio == 0)
						CheckProvider();
					/*
					else if(m_iRadio == 1)
						LibraryLogin();
					else if(m_iRadio == 2)
						LibraryLogout();
					*/

					m_csKitking.Empty();
					m_csProvider.Empty();
				}
				
				return true;
				break;

			default:
				return true;
				break;
			}
		}
	}

	return CPropertyPage::PreTranslateMessage(pMsg);
}

void CPageCheck::CheckProvider(void)
{
	CStringList *csTempList = NULL;
	BOOL bFlagFind = FALSE;
	if(mapCheck.Lookup(m_csKitking, (CObject *&)csTempList))
	{
		POSITION listPos = csTempList->GetHeadPosition();
		while(listPos)
		{
			if(m_csProvider.Find(csTempList->GetNext(listPos)) != -1)
			{
				bFlagFind = TRUE;
				break;
			}
		}

		if(bFlagFind)
		{
			if(m_flgTaojian = FALSE)
			{
				MessageBox1("匹配成功!","录入成功!",MB_OK);
				return;
			}
			g_ssk.Clear();
			if(mapInBox.Lookup(m_csKitking,g_ssk))
			{
				g_ssk.iQty += atoi(m_csQty);
			}
			else
				g_ssk.iQty = atoi(m_csQty);
			CTime tm = CTime::GetCurrentTime();  
			CString strTime = tm.Format("%m-%d %X"); 
			g_ssk.strOrder = strTime;
			g_ssk.strPartNo = m_csKitking;
			g_ssk.strUserCode = m_csProvider;	
			mapInBox.SetAt(g_ssk.strPartNo,g_ssk);
			CLayers cBase,cValue;
			CString csTotal,strKey,cs;
			if(mapBaseList.Lookup(g_ssk.strPartNo,cBase))
			{
				if(g_ssk.iQty == cBase.iQty)
				{
					csTotal.Format("%d这颗物料已经齐套",g_ssk.iQty);
					Speak(csTotal);
				}
				else
				{
					csTotal.Format("%d",g_ssk.iQty);
					Speak(csTotal);
				}
			}
			else
			{
				csTotal.Format("%d",g_ssk.iQty);
				Speak(csTotal);
			}
			m_LayerList.RemoveAll();
			m_ListMain.DeleteAllItems();
			g_iOrder = 0;
			POSITION pos = mapInBox.GetStartPosition();
			while(pos)
			{
				mapInBox.GetNextAssoc(pos, strKey, cValue);
				//生成保存清单容器
				CLayers* pLayersItem = new CLayers(cValue);
				m_LayerList.AddTail(pLayersItem);
				g_iOrder++;
				cs.Format("%d",pLayersItem->iQty);
				m_ListMain.AddItem(pLayersItem->strOrder,pLayersItem->strPartNo,pLayersItem->strUserCode,cs);
			}
			m_ListMain.Sort(0,true);
			if(g_iOrder)
				m_ListMain.EnsureVisible(g_iOrder-1,TRUE);


			//保存清单
			CFile file;
			CString strFileName = CExcel::GetAppPath() + "\\Record List\\" + m_strZCS + ".ckd";
			file.Open(strFileName, CFile::modeCreate|CFile::modeReadWrite);
			CArchive ar(&file, CArchive::store);
			m_LayerList.Serialize(ar);
			ar.Close();
			file.Close();
		}
		else
			MessageBox1("供应商号不匹配!","录入!",MB_OK);
	}
	else
		MessageBox1("部品号未录入!","录入成功!",MB_OK);
}

void CPageCheck::LibraryLogin(void)
{
	CStringList *csTempList = NULL;
	if(mapCheck.Lookup(m_csKitking, (CObject *&)csTempList))
	{
		if(csTempList->Find(m_csProvider) == NULL)
		{
			csTempList->AddTail(m_csProvider);
			mapCheck.SetAt(m_csKitking,(CObject *&)csTempList);
			CheckUpdateData();
			MessageBox1("录入成功","录入成功!",MB_OK);
		}
		else
			MessageBox1("供应商号已经存在","未成功!",MB_OK);
	}
	else
	{
		CStringList *csList = new CStringList;
		csList->AddTail(m_csProvider);
		mapCheck.SetAt(m_csKitking,(CObject *&)csList);
		//delete csList;  //勿删除实例
		CheckUpdateData();
		MessageBox1("录入成功","录入成功!",MB_OK);
	}

	CString strFileName = CExcel::GetAppPath() + "\\基准库\\" + "BaseProvider" + ".ckd";

	//序列化文件
	CFile file;
	file.Open(strFileName, CFile::modeCreate|CFile::modeReadWrite);
	CArchive ar(&file, CArchive::store);
	mapCheck.Serialize(ar);
	ar.Close();
	file.Close();
}

void CPageCheck::LibraryLogout(void)
{
	CStringList *csTempList = NULL;
	if(mapCheck.Lookup(m_csKitking, (CObject *&)csTempList))
	{
		if(csTempList->Find(m_csProvider))
		{
			csTempList->RemoveAt(csTempList->Find(m_csProvider));
			if(csTempList->GetCount())
				mapCheck.SetAt(m_csKitking,(CObject *&)csTempList);
			else
				mapCheck.RemoveKey(m_csKitking);
			MessageBox1("撤销成功!","撤销成功!",MB_OK);
			CheckUpdateData();

			CString strFileName = CExcel::GetAppPath() + "\\基准库\\" + "BaseProvider" + ".ckd";

			//序列化文件
			CFile file;
			file.Open(strFileName, CFile::modeReadWrite);
			CArchive ar(&file, CArchive::store);
			mapCheck.Serialize(ar);
			ar.Close();
			file.Close();
		}
		else
			MessageBox1("未找到要撤销的供应商号","撤销!",MB_OK);
	}
	else
		MessageBox1("未找到要撤销的部品号","撤销!",MB_OK);
}

void CPageCheck::CheckUpdateData(void)
{
	g_iOrder = 0;
	m_ListMain.DeleteAllItems();
	POSITION pos = mapCheck.GetStartPosition();
	CString strPartNo;
	CStringList* listProvider;
	while(pos)
	{
		mapCheck.GetNextAssoc(pos, strPartNo, (CObject *&)listProvider);
		POSITION listPos = listProvider->GetHeadPosition();
		while(listPos)
		{
			g_csTemp.Format("%d",++g_iOrder);
			m_ListMain.AddItem(g_csTemp,strPartNo,listProvider->GetNext(listPos)," ");
		}

	}
	//m_ListMain.Sort(0,true);
	if(g_iOrder)
		m_ListMain.EnsureVisible(g_iOrder-1,TRUE);
}

void CPageCheck::OnBnClickedRadioCheck()
{
	// TODO: 在此添加控件通知处理程序代码
	g_bitCheck.reset();
	m_csProvider.Empty();
	m_csKitking.Empty();
	m_csReceive.Empty();
	m_iRadio = 0;
	UpdateData(FALSE);

}

void CPageCheck::OnBnClickedRadioLogin()
{
	// TODO: 在此添加控件通知处理程序代码
	g_bitCheck.reset();
	m_csProvider.Empty();
	m_csKitking.Empty();
	m_csReceive.Empty();
	m_iRadio = 1;
	UpdateData(FALSE);
}

void CPageCheck::OnBnClickedRadioLogout()
{
	// TODO: 在此添加控件通知处理程序代码
	g_bitCheck.reset();
	m_csProvider.Empty();
	m_csKitking.Empty();
	m_csReceive.Empty();
	m_iRadio = 2;
	UpdateData(FALSE);
}


void CPageCheck::ReloadData(void)
{
	UpdateData(TRUE);
	m_strZCS.MakeUpper();
	mapInBox.RemoveAll();
	mapOutBox.RemoveAll();
	m_LayerList.RemoveAll();
	mapPartCnt.RemoveAll();
	CString cs;
	CString strFileName = CExcel::GetAppPath() + "\\Record List\\" + m_strZCS + ".ckd";
	if(IsFileExist(strFileName,FALSE)==TRUE)
	{
		//==============================================================================================
		//载入录入清单
		m_ListMain.DeleteAllItems();
		g_iOrder = 0;

		//序列化文件
		CFile file;
		file.Open(strFileName, CFile::modeReadWrite);
		CArchive ar(&file, CArchive::load);
		m_LayerList.Serialize(ar);
		ar.Close();
		file.Close();

		POSITION pos = m_LayerList.GetHeadPosition();
		while (pos != NULL)
		{
			CLayers* pLayer = m_LayerList.GetNext(pos);
			int count = 0;
			g_iOrder++;
			mapInBox.SetAt(pLayer->strPartNo, *pLayer);
			cs.Format("%d",pLayer->iQty);
			m_ListMain.AddItem(pLayer->strOrder,pLayer->strPartNo,pLayer->strUserCode,cs);
		}
		m_flgTaojian = TRUE;
		m_ListMain.Sort(0,true);
		if(g_iOrder)
			m_ListMain.EnsureVisible(g_iOrder-1,TRUE);
		//MessageBox1("套件载入成功!","载入成功", MB_ICONINFORMATION);
		//return true;
		//==================================================================================================

		//==================================================================================================
		//载入基础清单
		strFileName = CExcel::GetAppPath() + "\\Record List\\" + m_strZCS + "_BaseList" + ".ckd";
		if(IsFileExist(strFileName,FALSE)==TRUE)
		{
			mapBaseList.RemoveAll();
			m_LayerList.RemoveAll();
			//序列化文件
			CFile file;
			file.Open(strFileName, CFile::modeReadWrite);
			CArchive ar(&file, CArchive::load);
			m_LayerList.Serialize(ar);
			ar.Close();
			file.Close();

			pos = m_LayerList.GetHeadPosition();
			while (pos != NULL)
			{
				CLayers* pLayer = m_LayerList.GetNext(pos);
				mapBaseList.SetAt(pLayer->strPartNo, *pLayer);	
			}
			MessageBox1("套件载入成功,基础清单载入成功!", "载入成功", MB_ICONINFORMATION);
		}
		else
		{
			MessageBox1("套件载入成功,基础清单未载入!", "载入", MB_ICONINFORMATION);
		}
		CEdit* pEdit = (CEdit*)GetDlgItem(IDC_EDIT_ZCS);
		pEdit->EnableWindow(FALSE);
		
	}
	else if(m_strZCS.IsEmpty())
	{
		MessageBox("请输入有效的套件号!","错误",MB_ICONERROR);
		m_flgTaojian = FALSE;
	}
	else
	{
		m_flgTaojian = TRUE;
		mapBaseList.RemoveAll();
		m_LayerList.RemoveAll();
		m_ListMain.DeleteAllItems();
		g_iOrder = 0;
		MessageBox("新套件号开始录入!","提示",MB_OK);
		CEdit* pEdit = (CEdit*)GetDlgItem(IDC_EDIT_ZCS);
		pEdit->EnableWindow(FALSE);
	}

	CEdit* pEdit = (CEdit*)GetDlgItem(IDC_EDIT_CHECK_RECEIVE);
	pEdit->SetFocus();
}


void CPageCheck::OnBnClickedBtnLoad()
{
	if(!m_flgTaojian)
	{
		MessageBox1("没有指定套件,请先指定套件号!","错误", MB_ICONERROR);
		return ;
	}

	CFileDialog dlg(true,"*.xls","",OFN_HIDEREADONLY,"Excel文件(*.xls)|*.xls");
	//===========================================================
	//打开特定目录
	CString strFileName = CExcel::GetAppPath() + "\\Base List";
	dlg.m_ofn.lpstrInitialDir = strFileName;
	//===========================================================
	if(dlg.DoModal()==IDOK)
		g_strOpenFile=dlg.GetPathName();

	if(g_strOpenFile.Find(".xls")<0)        //没有读取则直接返回
		return;

	unsigned int i,j;
	CString strCell;
	CExcel Excel;
	Excel.AddNewFile(g_strOpenFile);				// 打开一个文件
	Excel.SetVisible(false);						// 设置可见
	Excel.SelectSheet(1);				        // 激活工作簿1
	m_usedRow = Excel.GetUsedRowCount();
	m_usedCol = Excel.GetUsedColCount();
	Excel.SelectSheet(1);				        // 重新设置RANGE
	m_icolProductCnt = m_icolPart = m_irowPart = m_icolDetail = m_icolUserCode = 0;
	m_flgCheck = false;
	
	//////////////////////////////////////
	//获取部品号 位号 描述 封装列位置
	for(i=1;i<=m_usedRow;i++)
	{
		for(j=1;j<=m_usedCol;j++)
		{
			strCell = Excel.GetCell(i,j).bstrVal;
			strCell = Excel.DeleteBlackSpace(strCell);  //删除空格
			char *pCell = (LPTSTR)(LPCTSTR)strCell;
			if(strcmp(pCell,TOULIAO_LIST_PRODUCT_CNT) == 0)
			{
				m_icolProductCnt = j;
			}
			else if(strcmp(pCell,TOULIAO_LIST_PART) == 0)
			{
				m_icolPart = j;
				m_irowPart = i;
			}
			else if(strcmp(pCell,TOULIAO_LIST_DETAIL) == 0)
			{
				m_icolDetail = j;
			}
			else if(strcmp(pCell,TOULIAO_LIST_USER) == 0)
			{
				m_icolUserCode = j;
			}
			else if(strcmp(pCell,TOULIAO_LIST_NO) == 0)
			{
				m_icolNo = j;
			}

			if(m_icolProductCnt && m_icolPart && m_irowPart && m_icolDetail && m_icolUserCode && m_icolNo)
			{
				m_flgCheck = true;
				break;
			}
		}

		if(m_icolProductCnt && m_icolPart && m_irowPart && m_icolDetail && m_icolUserCode && m_icolNo)
		{
				m_flgCheck = true;
				break;
		}
	}

	/////////////////////////////////////////////////////
	//检查是否为标准K3清单
	if(!m_flgCheck)  
	{
		Excel.Save(true);
		MessageBox("请导入正确的基础清单!!!");
		return ;
	}

	CLayers ssk;
	mapBaseList.RemoveAll();
	/////////////////////////////////////////////////////	
	for(i=m_irowPart+1;i<=m_usedRow;i++)
	{
		strCell = Excel.GetCell(i,m_icolPart).bstrVal; 
		strCell = Excel.DeleteBlackSpace(strCell);
		if(strCell.GetLength() == 14||strCell.GetLength() == 15)
		{
			///////////////////////////////////////////////////
			//写入哈希表
			strCell.MakeUpper();
			ssk.strPartNo = strCell;

			strCell = Excel.GetCell(i,m_icolUserCode).bstrVal;
			strCell.MakeUpper();
			ssk.strUserCode = strCell;

			strCell = Excel.GetCell(i,m_icolDetail).bstrVal;
			strCell.MakeUpper();
			ssk.strDetail = strCell;

			strCell = Excel.GetCell(i,m_icolNo).bstrVal;
			strCell.MakeUpper();
			ssk.strNo = strCell;

			ssk.iQty = Excel.GetCellValue(i,m_icolProductCnt);
			mapBaseList.SetAt(ssk.strPartNo,ssk);
		}
	}
	
	m_ListMain.DeleteAllItems();
	m_LayerList.RemoveAll();
	POSITION pos = mapBaseList.GetStartPosition();
	CString strKey;
	CString cs;
	g_iOrder = 0;
    while(pos)
    {
		g_iOrder++;
        mapBaseList.GetNextAssoc(pos, strKey, ssk);
		cs.Format("%d",ssk.iQty);
		m_ListMain.AddItem(ssk.strNo,strKey,ssk.strUserCode,ssk.strDetail,cs);	

		//生成保存清单容器
		CLayers* pLayersItem = new CLayers(ssk);
		m_LayerList.AddTail(pLayersItem);
    }
	m_ListMain.Sort(0,true);
	m_ListMain.EnsureVisible(g_iOrder-1,TRUE);

	//保存清单
	CFile file;
	strFileName = CExcel::GetAppPath() + "\\Record List\\" + m_strZCS + "_BaseList" + ".ckd";
	file.Open(strFileName, CFile::modeCreate|CFile::modeReadWrite);
	CArchive ar(&file, CArchive::store);
	m_LayerList.Serialize(ar);
	ar.Close();
	file.Close();

	MessageBox1("基础清单载入成功!", "载入成功", MB_ICONINFORMATION);	
}
void CPageCheck::OnBnClickedBtnLogin()
{
	UpdateData(TRUE);
	if(m_iRadio !=1 )
	{
		MessageBox1("请正确选择扫描方式","录入!",MB_OK);
		return;
	}
	if(m_csKitking.IsEmpty())
	{
		MessageBox1("请先填写部品号","录入!",MB_OK);
		return;
	}
	else if(m_csProvider.IsEmpty())
	{
		MessageBox1("请先填写供应商号","录入!",MB_OK);
		return;
	}

	if(m_csKitking.Find('-') != -1
		&& (m_csKitking.GetLength() == 13||m_csKitking.GetLength() == 14||m_csKitking.GetLength() == 15))
		;
	else
	{
		MessageBox1("请正确填写部品号","录入!",MB_OK);
		return;
	}

	m_csKitking.MakeUpper();
	m_csProvider.MakeUpper();

	CStringList *csTempList = NULL;
	if(mapCheck.Lookup(m_csKitking, (CObject *&)csTempList))
	{
		if(csTempList->Find(m_csProvider) == NULL)
		{
			csTempList->AddTail(m_csProvider);
			mapCheck.SetAt(m_csKitking,(CObject *&)csTempList);
			CheckUpdateData();
			m_csKitking.Empty(); m_csProvider.Empty(); UpdateData(FALSE);
			MessageBox1("录入成功","录入成功!",MB_OK);
		}
		else
			MessageBox1("供应商号已经存在","未成功!",MB_OK);
	}
	else
	{
		CStringList *csList = new CStringList;
		csList->AddTail(m_csProvider);
		mapCheck.SetAt(m_csKitking,(CObject *&)csList);
		//delete csList;  //勿删除实例
		CheckUpdateData();
		m_csKitking.Empty(); m_csProvider.Empty(); UpdateData(FALSE);
		MessageBox1("录入成功","录入成功!",MB_OK);
	}

	CString strFileName = CExcel::GetAppPath() + "\\基准库\\" + "BaseProvider" + ".ckd";

	//序列化文件
	CFile file;
	file.Open(strFileName, CFile::modeCreate|CFile::modeReadWrite);
	CArchive ar(&file, CArchive::store);
	mapCheck.Serialize(ar);
	ar.Close();
	file.Close();
}

void CPageCheck::OnBnClickedBtnLogout()
{
	CStringList *csTempList = NULL;
	UpdateData(TRUE);
	if(m_iRadio !=2 )
	{
		MessageBox1("请正确选择扫描方式","录入!",MB_OK);
		return;
	}
	if(m_csKitking.IsEmpty())
	{
		MessageBox1("请填写部品号","录入!",MB_OK);
		return;
	}
	else if(m_csProvider.IsEmpty())
	{
		MessageBox1("请填写供应商号","录入!",MB_OK);
		return;
	}

	if(m_csKitking.Find('-') != -1
		&& (m_csKitking.GetLength() == 13||m_csKitking.GetLength() == 14||m_csKitking.GetLength() == 15))
		;
	else
	{
		MessageBox1("请正确填写部品号","录入!",MB_OK);
		return;
	}

	m_csKitking.MakeUpper();
	m_csProvider.MakeUpper();

	if(mapCheck.Lookup(m_csKitking, (CObject *&)csTempList))
	{
		if(csTempList->Find(m_csProvider))
		{
			csTempList->RemoveAt(csTempList->Find(m_csProvider));
			if(csTempList->GetCount())
				mapCheck.SetAt(m_csKitking,(CObject *&)csTempList);
			else
				mapCheck.RemoveKey(m_csKitking);
			MessageBox1("撤销成功!","撤销成功!",MB_OK);
			CheckUpdateData();

			CString strFileName = CExcel::GetAppPath() + "\\基准库\\" + "BaseProvider" + ".ckd";

			//序列化文件
			CFile file;
			file.Open(strFileName, CFile::modeReadWrite);
			CArchive ar(&file, CArchive::store);
			mapCheck.Serialize(ar);
			ar.Close();
			file.Close();
		}
		else
			MessageBox1("未找到要撤销的供应商号","撤销!",MB_OK);
	}
	else
		MessageBox1("未找到要撤销的部品号","撤销!",MB_OK);
}
