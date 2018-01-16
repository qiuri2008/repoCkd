// AutoWeight.cpp : implementation file
//

#include "stdafx.h"
#include "CKD录入.h"
#include "PageCkd.h"
#include <bitset>
#include "Global.h"
#include <afxtempl.h>
#include "excel2003.h"
#include "user.h"
#include "PropSheet.h"
#include "SaveDlg.h"
#include "AutoWeight.h"
#include ".\autoweight.h"


#define TEST_AUTOWEIGHT		0   //打开 会忽略未称 重量

using std::bitset;
bitset<3> g_bitPallet;

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAutoWeight property page

IMPLEMENT_DYNCREATE(CAutoWeight, CPropertyPage)

CAutoWeight::CAutoWeight() : CPropertyPage(CAutoWeight::IDD)
{
	//{{AFX_DATA_INIT(CAutoWeight)
	m_csPallet = _T("");
	m_csOutBox = _T("");
	m_csReceive = _T("");
	m_strZCS = _T("");
	m_srtPalletWeight = _T("托盘重量");
	m_strOutBoxWeight = _T("");
	//}}AFX_DATA_INIT
	m_listInit = false;
	m_dPalletWeight = 0;
}

CAutoWeight::~CAutoWeight()
{
}

void CAutoWeight::DoDataExchange(CDataExchange* pDX)
{
	CPropertyPage::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAutoWeight)
	DDX_Control(pDX, IDC_LIST2, m_ListMain);
	DDX_Text(pDX, IDC_EDIT1_BIGBOX, m_csPallet);
	DDX_Text(pDX, IDC_EDIT1_OUTBOX, m_csOutBox);
	DDX_Text(pDX, IDC_EDIT1_RECEIVE, m_csReceive);
	DDX_Text(pDX, IDC_EDIT1_ZCS, m_strZCS);
	DDX_Text(pDX, IDC_STATIC_1, m_srtPalletWeight);
	DDX_Text(pDX, IDC_EDIT_OUTBOX_WEIGHT, m_strOutBoxWeight);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CAutoWeight, CPropertyPage)
	//{{AFX_MSG_MAP(CAutoWeight)
	ON_BN_CLICKED(IDC_BTN_LOAD, OnBtnLoadWeightList)
	ON_WM_CTLCOLOR()
	ON_BN_CLICKED(IDC_BTN1_SAVE, OnBtnSave)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CAutoWeight message handlers

BOOL CAutoWeight::OnSetActive() 
{
	// TODO: Add your specialized code here and/or call the base class
	CPropSheet* pParent = (CPropSheet*) GetParent();
	CPageShipMark* prop1 = (CPageShipMark*)pParent->GetPage(pParent->GetPageCount()-1);
	
	g_flgSpeak = prop1->m_bSpeak;

	DWORD dwStyle;
	if(!m_listInit)
	{
		dwStyle=m_ListMain.GetExtendedStyle();
		dwStyle|=LVS_EX_FULLROWSELECT;
		dwStyle|=LVS_EX_GRIDLINES;

		(void)m_ListMain.SetExtendedStyle( dwStyle );///整行选择模式//LVS_EX_FULLROWSELECT
		m_ListMain.SetHeadings("序号,100;托盘,60;外箱,60;内箱,60;重量,80"); ///设置列头信息 
		m_ListMain.LoadColumnInfo(); 
		m_listInit = true;
	}
	//=================================================
	//=== 不同模块交替使用时，要重新加载数据===========
	else if(m_flgTaojian == TRUE && !m_strZCS.IsEmpty())
		ReloadData();
	//=================================================
	
	return CPropertyPage::OnSetActive();
}



BOOL CAutoWeight::PreTranslateMessage(MSG* pMsg) 
{
	// TODO: Add your specialized code here and/or call the base class
	CString cs, cs1;
	if(pMsg->message == WM_KEYDOWN) 
	{   
		if(pMsg->wParam == VK_RETURN) 
		{   
			UINT nID = GetFocus()->GetDlgCtrlID(); 
			switch(nID)
			{   
				case IDC_EDIT1_ZCS:
					ReloadData();					
					return true;
					break;
				case IDC_EDIT1_RECEIVE:
					
					UpdateData(TRUE);
					m_csReceive.MakeUpper();
					if(g_ttsCreate)
						tts.Pause();

					if(!m_flgTaojian)
					{
						MessageBox1("没有指定套件,请先指定套件号!","错误", MB_ICONERROR);
						m_csReceive.Empty();
						UpdateData(false);
						return true;
					}

					//  录入-托盘
					if(m_csReceive.Find('<') != -1 && m_csReceive.Find('>') != -1)
					{
						CString cs1,csTemp;
						csTemp = m_csReceive.Mid(m_csReceive.Find('<')+1,m_csReceive.Find('>')-m_csReceive.Find('<')-1);
						if(!m_csPallet.IsEmpty() && m_csPallet != csTemp)
						{
							//MessageBox1("托盘已变更为" + csTemp,"托盘变更提示",MB_ICONINFORMATION);
							//if(MessageBox("托盘已变更为" + csTemp + ",请确认?","托盘变更提示",MB_OKCANCEL) == IDOK)
							//{
								m_csPallet = csTemp;
								
								g_bitPallet.set(STYLE_PALLET);
								m_csReceive.Empty();
								m_dPalletWeight = PalletWeight(m_csPallet);
								m_srtPalletWeight.Format("%.2f",m_dPalletWeight);
								UpdateData(false);
								MessageBox1("托盘确认变更为" + csTemp + "," + m_srtPalletWeight + "公斤开始录入","托盘变更",MB_ICONINFORMATION);
								return true;
							//}
							//else
							//{
								//MessageBox1("取消变更托盘," + m_csPallet + "继续录入","托盘变更",MB_ICONINFORMATION);
								//return true;
							//}
						}
						else
							m_csPallet = csTemp;
						g_bitPallet.set(STYLE_PALLET);
						m_csReceive.Empty();
						m_dPalletWeight = PalletWeight(m_csPallet);
						m_srtPalletWeight.Format("%.2f",m_dPalletWeight);
						if(m_dPalletWeight)
							MessageBox1("托盘" + m_csPallet + "," + m_srtPalletWeight + "公斤继续录入","托盘",MB_ICONINFORMATION);
						else
							MessageBox1("托盘" + m_csPallet + "开始录入","托盘",MB_ICONINFORMATION);
						UpdateData(false);

					}
					//  录入-外箱
					else if(m_csReceive.Find('{') != -1 && m_csReceive.Find('}') != -1)
					{
						CString csPallet, csWeight;
						int count = 0;
						m_csOutBox = m_csReceive.Mid(m_csReceive.Find('{')+1,m_csReceive.Find('}')-m_csReceive.Find('{')-1);
						g_bitPallet.set(STYLE_PALLET_OUTBOX);
						if(!mapOutBox.Lookup(m_csOutBox,count))
						{
							MessageBox1(m_csOutBox+"不存于"+m_strZCS + "的套件中","外箱检查",MB_ICONINFORMATION);
							return true;
						}
						if(g_bitPallet.count() == STYLE_PALLET_END)
						{
							double dOutBoxWeight = 0;
							if(mapOutboxPallet.Lookup(m_csOutBox,csPallet))
							{
								if(csPallet == m_csPallet)
								{
									mapOutboxPallet.RemoveKey(m_csOutBox);
									//GetOutBoxWeight  将托盘号更新到 mapInBox
									dOutBoxWeight = GetOutBoxWeight(m_csOutBox,csPallet,false);
									m_dPalletWeight -= dOutBoxWeight;
									csWeight.Format("删除%g公斤",dOutBoxWeight);
								}
								else
									MessageBox1("不能删除外箱" + m_csOutBox + "存在于" + csPallet +"托盘中","外箱检查",MB_ICONINFORMATION);
								
							}
							else
							{
								mapOutboxPallet.SetAt(m_csOutBox, m_csPallet);
								//GetOutBoxWeight  将托盘号更新到 mapInBox
								dOutBoxWeight = GetOutBoxWeight(m_csOutBox,m_csPallet,true);
								m_dPalletWeight += dOutBoxWeight;
								csWeight.Format("添加%g公斤",dOutBoxWeight);
							}
							m_csReceive.Empty();
							m_srtPalletWeight.Format("%.2fKg",m_dPalletWeight);
							m_strOutBoxWeight.Format("%g",dOutBoxWeight);
							UpdateData(false);
							MessageBox1(csWeight, "外箱录入", MB_ICONINFORMATION);
							UpdateList();
						}
						else if(!g_bitPallet[STYLE_PALLET])
						{
							g_bitPallet.reset();
							m_csPallet.Empty(); m_csOutBox.Empty();
							MessageBox1("托盘未录入,请先录入托盘!","托盘未录入",MB_ICONERROR);
						}
					
					}
					return true;
					break;
				default:
					{
						CEdit* pEdit = (CEdit*)GetDlgItem(IDC_EDIT1_RECEIVE);
						pEdit->SetFocus();
					}
					return true;
					break;
				}
			}
		}
	
	return CPropertyPage::PreTranslateMessage(pMsg);  
}

void CAutoWeight::UpdateList()
{
	CString strKey,cs1;
	m_ListMain.DeleteAllItems();
	m_LayerList.RemoveAll();
	g_iReOrder = 0;
	POSITION pos = mapInBox.GetStartPosition();
	CLayers cValue;

	while(pos)
    {	
		mapInBox.GetNextAssoc(pos, strKey, cValue);
		if(!cValue.strPallet.IsEmpty())
		{
			g_iReOrder++;
			cs1.Format("%g", cValue.iQty*cValue.iUnitWeight*cValue.iRadio);
			m_ListMain.AddItem(cValue.strOrder, cValue.strPallet,cValue.strOutBox, strKey, cs1);
		}
		//生成保存清单容器
		CLayers* pLayersItem = new CLayers(cValue);
		m_LayerList.AddTail(pLayersItem);
	}
	m_ListMain.Sort(0,true);
	m_ListMain.EnsureVisible(g_iReOrder-1,TRUE);

	CPropSheet* pParent = (CPropSheet*) GetParent();
	CPageShipMark* prop1 = (CPageShipMark*)pParent->GetPage(pParent->GetPageCount()-1);

	//保存内箱总清单
	CFile file;
	CString strFileName = CExcel::GetAppPath() + "\\Record List\\" + m_strZCS + ".ckd";
	file.Open(strFileName, CFile::modeCreate|CFile::modeReadWrite);
	CArchive ar(&file, CArchive::store);
	ar<<prop1->m_cs11<<prop1->m_cs12<<prop1->m_cs13<<prop1->m_cs14
	  <<prop1->m_cs21<<prop1->m_cs22<<prop1->m_cs23<<prop1->m_cs24
	  <<prop1->m_cs31<<prop1->m_cs32<<prop1->m_cs33<<prop1->m_cs34;
	m_LayerList.Serialize(ar);
	ar.Close();
	file.Close();

	//保存“外箱-托盘”清单
	if(mapOutboxPallet.GetSize() == 0)
		return;
	strFileName = CExcel::GetAppPath() + "\\Record List\\" + m_strZCS + "_PalletList" + ".ckd";
	file.Open(strFileName, CFile::modeCreate|CFile::modeReadWrite);
	CArchive ar1(&file, CArchive::store);
	mapOutboxPallet.Serialize(ar1);
	ar1.Close();
	file.Close();
}

double CAutoWeight::GetOutBoxWeight(CString strOutbox,CString strPallet, bool bAddFlag)
{
	POSITION pos = mapInBox.GetStartPosition();
	CString strinbox;
	CLayers cValue;
	double outBoxWeight = 0;

	/////////////////////////////////////////////////////////////////
	//检索某个外箱的所有内箱 并计算整个外箱的重量
	while(pos)
	{
		mapInBox.GetNextAssoc(pos, strinbox, cValue);
		if(strOutbox == cValue.strOutBox)
		{
			outBoxWeight += cValue.iQty*cValue.iUnitWeight*cValue.iRadio;
			if(cValue.strPallet == "" && bAddFlag)
			{
				cValue.strPallet = strPallet;
				CTime tm = CTime::GetCurrentTime();  
				CString strTime = tm.Format("%m-%d %X"); 
				cValue.strOrder = strTime;
				mapInBox.SetAt(strinbox, cValue);
			}
			else if(cValue.strPallet != "" && bAddFlag == false)
			{
				cValue.strPallet = "";
				mapInBox.SetAt(strinbox, cValue);
			}
		}
	}
	return outBoxWeight;
}

double CAutoWeight::PalletWeight(CString csPallet)
{
	double palletWeight = 0;
	CString strOutbox,strPallet;
	/*
	POSITION pos = mapOutboxPallet.GetStartPosition();
	while(pos)
    {
		mapOutboxPallet.GetNextAssoc(pos, strOutbox, strPallet);
		if(csPallet == strPallet)
			palletWeight += GetOutBoxWeight(strOutbox);
	}
	*/
	
		CLayers cValue;
		CString strinbox;
		POSITION pos = mapInBox.GetStartPosition();
		while(pos)
		{
			mapInBox.GetNextAssoc(pos, strinbox, cValue);
			if(csPallet == cValue.strPallet)
				palletWeight += cValue.iQty*cValue.iUnitWeight*cValue.iRadio;
		}
		
	return palletWeight;
}



void CAutoWeight::OnBtnLoadWeightList() 
{
	// TODO: Add your control notification handler code here
	CFileDialog dlg(true,"*.xls","",OFN_HIDEREADONLY,"Excel文件(*.xls)|*.xls");
	//===========================================================
	//打开特定目录
	CString strFileName = CExcel::GetAppPath() + "\\托盘清单";
	dlg.m_ofn.lpstrInitialDir = strFileName;
	//===========================================================
	if(dlg.DoModal()==IDOK)
		g_strOpenFile=dlg.GetPathName();

	if(g_strOpenFile.Find(".xls")<0)				//没有读取则直接返回
		return;

	unsigned int i,j;
	CString strCell, cs1, cs2;
	CExcel Excel;
	Excel.AddNewFile(g_strOpenFile);				// 打开一个文件
	Excel.SetVisible(false);						// 设置可见
	Excel.SelectSheet(1);							// 激活工作簿1
	m_usedRow = Excel.GetUsedRowCount();
	m_usedCol = Excel.GetUsedColCount();
	Excel.SelectSheet(1);							// 重新设置RANGE
	m_icolProductCnt = m_icolPart = m_irowPart = m_icolDetail = m_icolUserCode = m_icolWeight = m_icolRadio = m_icolUnitPrice =0;
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

			if(strcmp(pCell,TOULIAO_LIST_PART) == 0)
			{
				m_icolPart = j;
				m_irowPart = i;
			}
			else if(strcmp(pCell,TOULIAO_LIST_DETAIL) == 0)
			{
				m_icolDetail = j;
			}
			else if(strcmp(pCell,"序号") == 0)
			{
				m_icolNo = j;
			}
			else if(strcmp(pCell,"单位重量") == 0)
			{
				m_icolWeight = j;
			}
			else if(strcmp(pCell,"加乘系数") == 0)
			{
				m_icolRadio = j;
			}
			else if(strcmp(pCell,"单价") == 0)
			{
				m_icolUnitPrice = j;
			}

			if(m_icolPart && m_irowPart && m_icolDetail && m_icolNo && m_icolWeight && m_icolRadio && m_icolUnitPrice)
			{
				m_flgCheck = true;
				break;
			}
		}

		if(m_icolPart && m_irowPart && m_icolDetail && m_icolNo && m_icolWeight && m_icolRadio && m_icolUnitPrice)
		{
				m_flgCheck = true;
				break;
		}
	}

	/////////////////////////////////////////////////////
	//检查
	if(!m_flgCheck)  
	{
		Excel.Save(true);
		MessageBox("请导入正确的重量清单!!!", "错误", MB_ICONERROR);
		return ;
	}

	CLayers ssk;
	mapWeightList.RemoveAll();
		
	/////////////////////////////////////////////////////	
	for(i=m_irowPart+1;i<=m_usedRow;i++)
	{
		strCell = Excel.GetCell(i,m_icolPart).bstrVal; 
		strCell = Excel.DeleteBlackSpace(strCell);
		if(strCell.GetLength() >= 13)
		{
			///////////////////////////////////////////////////
			//写入哈希表
			strCell.MakeUpper();
			ssk.strPartNo = strCell;

			strCell = Excel.GetCell(i,m_icolDetail).bstrVal;
			strCell.MakeUpper();
			ssk.strDetail = strCell;

			strCell = Excel.GetCell(i,m_icolNo ).bstrVal;
			strCell.MakeUpper();
			ssk.strNo = strCell;
			//文件清单中单位是g，要转换为Kg
			ssk.iUnitWeight = Excel.GetCellValueFloat(i,m_icolWeight)/1000;
			ssk.iRadio = Excel.GetCellValueFloat(i,m_icolRadio);
			ssk.iUnitPrice = Excel.GetCellValueFloat(i,m_icolUnitPrice);

			mapWeightList.SetAt(ssk.strPartNo,ssk);
		}
	}
	
	m_ListMain.DeleteAllItems();
	m_LayerList.RemoveAll();
	POSITION pos = mapWeightList.GetStartPosition();
	CString strKey;
	g_iOrder = 0;
    while(pos)
    {
		g_iOrder++;
        mapWeightList.GetNextAssoc(pos, strKey, ssk);
		cs1.Format("%g", ssk.iUnitWeight);
		cs2.Format("%g", ssk.iRadio);
		m_ListMain.AddItem(ssk.strNo,strKey,cs1,cs2,"");	

		//生成保存清单容器
		CLayers* pLayersItem = new CLayers(ssk);
		m_LayerList.AddTail(pLayersItem);
    }
	m_ListMain.Sort(0,true);
	m_ListMain.EnsureVisible(g_iOrder-1,TRUE);

	//保存清单
	CFile file;
	strFileName = CExcel::GetAppPath() + "\\托盘清单\\" + "_WeightList" + ".ckd";
	file.Open(strFileName, CFile::modeCreate|CFile::modeReadWrite);
	CArchive ar(&file, CArchive::store);
	m_LayerList.Serialize(ar);
	ar.Close();
	file.Close();

	MessageBox1("重量清单载入成功!", "载入成功", MB_ICONINFORMATION);
	
}

HBRUSH CAutoWeight::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor) 
{
	HBRUSH hbr = CPropertyPage::OnCtlColor(pDC, pWnd, nCtlColor);
	
	// TODO: Change any attributes of the DC here
	CFont *f; 
    f = new CFont; 

	switch(pWnd->GetDlgCtrlID())
	{
	case IDC_EDIT_OUTBOX_WEIGHT:
		//f->CreateFont(14,0,0,0,FW_SEMIBOLD,FALSE,FALSE,0, 
		//ANSI_CHARSET,OUT_DEFAULT_PRECIS,
		//CLIP_DEFAULT_PRECIS,DEFAULT_QUALITY,
		//DEFAULT_PITCH&FF_SWISS,"Arial");
		//pDC->SetBkMode(TRANSPARENT);
		pDC->SetTextColor(RGB(255,0, 0));
		//pDC->SelectObject(f);//设置字体
		break;
	case IDC_STATIC_LOOKUP:
		f->CreateFont(16,0,0,0,FW_SEMIBOLD,FALSE,FALSE,0, 
		ANSI_CHARSET,OUT_DEFAULT_PRECIS,
		CLIP_DEFAULT_PRECIS,DEFAULT_QUALITY,
		DEFAULT_PITCH&FF_SWISS,"Arial");
		
		pDC->SetBkMode(TRANSPARENT);
		pDC->SetTextColor(RGB(0,0, 255));
		pDC->SelectObject(f);//设置字体
		break;
	case IDC_STATIC_1:
		f->CreateFont(26,0,0,0,FW_SEMIBOLD,FALSE,FALSE,0, 
		ANSI_CHARSET,OUT_DEFAULT_PRECIS,
		CLIP_DEFAULT_PRECIS,DEFAULT_QUALITY,
		DEFAULT_PITCH&FF_SWISS,"Arial");
		pDC->SetBkMode(TRANSPARENT);
		pDC->SetTextColor(RGB(255,0, 0));
		pDC->SelectObject(f);//设置字体
	default :
		break;
	}

	
	// TODO: Return a different brush if the default is not desired
	return hbr;
}

template <typename T, typename U> void BubbleSort(T& collection, U element, int count, bool ascend = true)
{
	for (int i = 0; i < count-1; i++)
	for (int j = 0; j < count-1-i; j++)
		if (ascend)
		{
		// 升序
			if (collection[j] > collection[j+1])
			{
				U temp = collection[j];
				collection[j] = collection[j+1];
				collection[j+1] = temp;
			}
		}
		else
		{
		// 降序
			if (collection[j] < collection[j+1])
			{
			U temp = collection[j];
			collection[j] = collection[j+1];
			collection[j+1] = temp;
			}
		}
}


void CAutoWeight::SaveInvoiceDate(CExcel &Excel) 
{
	CString strCell;
	Excel.SelectSheet(2);				        // 激活工作簿1
	Excel.ActiveSheet().SetName("发票");

	Excel.SetColWidth(1,14);
	Excel.SetCell(1,1,"PO No.");
	Excel.SetCell(1,2,"Item Code");
	Excel.SetCell(1,3,"Item Description");
	Excel.SetCell(1,4,"Kitking Code");
	Excel.SetCell(1,5,"Name");
	Excel.SetCell(1,6,"Quantity");
	Excel.SetCell(1,7,"Unit Price");
	Excel.SetCell(1,8,"Total Amount");

	POSITION pos = mapInBox.GetStartPosition();
	CString strKey;
	CLayers cValue,cBase;
	int i = 2;


	CArrayClayer NoArray;
	//================================================================
	//按料号排序
	while(pos)
	{
		mapInBox.GetNextAssoc(pos, strKey, cValue);
		if(cValue.strPallet != "")			//只有放进托盘的内箱才做发票
			NoArray.Add(cValue);
	}
	g_OrderMode = ORDER_NO;
	BubbleSort(NoArray, NoArray[0], NoArray.GetSize(), true);
	//================================================================
	unsigned int iBorder = 2;	
    for(int j=0; j<NoArray.GetSize();)
    {
		//===========================================
		//检索相同部品号
		unsigned int count = 1;
		cValue = NoArray[j];
		cValue.strOrder = cValue.strOrder.Left(5);
		for(int k=j+1; k<NoArray.GetSize(); k++)
		{
			if(cValue.strNo == NoArray[k].strNo)
			{
				cValue.iQty += NoArray[k].iQty;
				count++;
			}
			else
				break;
		}
		j = j+count;
		Excel.SetCell(i,2,cValue.strUserCode);
		Excel.SetCell(i,3,cValue.strDetail);
		Excel.SetCell(i,4,cValue.strPartNo);
		Excel.SetCell(i,6,(long)cValue.iQty);
		Excel.SetCell(i,7,cValue.iUnitPrice,6);
		Excel.SetCell(i,8,cValue.iUnitPrice*cValue.iQty,6);
		i++;
		
    }

	for(i=1;i<=8;i++)
		Excel.SetColAutoFit(i);	

}



void CAutoWeight::OnBtnSave() 
{
	// TODO: Add your control notification handler code here
	if(!m_flgTaojian)
	{
		MessageBox1("没有指定套件,请先指定套件号!","错误", MB_ICONERROR);
		return ;
	}
	if(mapInBox.GetCount() == 0)
	{
		MessageBox1("未录入任何数据,请先录入!","保存错误",MB_ICONERROR);
		return ;
	}

	CFileDialog dlg(FALSE, "xls", "", OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, "Excel文件(*.xls)|*.xls" );  
	//===========================================================
	//保存到指定目录
	CString strFileName = CExcel::GetAppPath() + "\\托盘清单";
	dlg.m_ofn.lpstrInitialDir = strFileName;
	//===========================================================
	if(dlg.DoModal()==IDOK)
		g_strSaveFile=dlg.GetPathName();

	if(g_strSaveFile.Find(".xls")<0)        //没有保存则直接返回
		return;

	CExcel Excel;
	Excel.AddNewFile();				            // 新建一个文件
	Excel.SetVisible(FALSE);					// 设置不可见

	Excel.SelectSheet(1);
	Excel.ActiveSheet().SetName("箱单");

	Excel.SetColWidth(1,14);
	Excel.SetCell(1,1,"SUPPLIER NAME");
	Excel.SetCell(1,2,"P/T NO.");
	CRange range(Excel.GetRange(1,3,1,4));
	range.Merge();
	range = "CTN NO.";
	Excel.SelectSheet(1);
	Excel.SetCell(1,5,"P/O NO.");
	Excel.SetCell(1,6,"CODE NO.");
	Excel.SetCell(1,7,"DESCRIPTION");
	Excel.SetCell(1,8,"KITKING CODE");
	Excel.SetCell(1,9,"Unit");
	Excel.SetCell(1,10,"QTY");
	Excel.SetCell(1,11,"N/wet(kg)");
	Excel.SetCell(1,12,"G/wet(kg)");

	POSITION pos = mapInBox.GetStartPosition();

	CString strKey,csTemp;
	CLayers cValue;
	int i = 2;
	int count = 0;
	//double dPalletWeight = 0;

	CArrayClayer palletArray;
	mapPalletCnt.RemoveAll();
	//================================================================
	//按托盘号排序
	while(pos)
	{
		mapInBox.GetNextAssoc(pos, strKey, cValue);
		if(!cValue.strPallet.IsEmpty())
		{
			palletArray.Add(cValue);

			if(mapPalletCnt.Lookup(cValue.strPallet,count))
				mapPalletCnt.SetAt(cValue.strPallet, count+1);
			else
				mapPalletCnt.SetAt(cValue.strPallet, 1);
		}
	}
	g_OrderMode = ORDER_PALLET;
	BubbleSort(palletArray, palletArray[0], palletArray.GetSize(), true);

	//================================================================

    for(int j=0; j<palletArray.GetSize(); j++)
    {
		cValue = palletArray[j];
		Excel.SetCell(i,2, cValue.strPallet);
		Excel.SetCell(i,3, cValue.strOutBox);
		if(cValue.strInBox.Find(INBOX_COMMON) != -1)
			Excel.SetCell(i,4,INBOX_COMMON);
		else
			Excel.SetCell(i,4,cValue.strInBox);

		Excel.SetCell(i,6,cValue.strUserCode);
		Excel.SetCell(i,7,cValue.strDetail);
		Excel.SetCell(i,8,cValue.strPartNo);

		Excel.SetCell(i,10,(long)cValue.iQty);
		Excel.SetCell(i,11,cValue.iQty*cValue.iUnitWeight*cValue.iRadio,2);	
		
		if(csTemp.IsEmpty())
			csTemp = cValue.strPallet;
		//==============================================================
		//框选第N个托盘，并结算当前托盘重量
		if(j == palletArray.GetSize()-1)
		{
			if(mapPalletCnt.Lookup(csTemp,count))
			{
				Excel.SetCell(i,12,(long)PalletWeight(cValue.strPallet));
				CRange range(Excel.GetRange(i-count+1,1,i,12));
				range.Border();
				Excel.SelectSheet(1);
			}
		}
		//=============================================================

		//=============================================================
		//框选第N-1个托盘，并计算当前托盘重量
		if(csTemp != cValue.strPallet)
		{
			if(mapPalletCnt.Lookup(csTemp,count))
			{
				Excel.SetCell(i-1,12,(long)PalletWeight(csTemp));
				CRange range(Excel.GetRange(i-count,1,i-1,12));
				range.Border();	
				Excel.SelectSheet(1);
			}
			csTemp = cValue.strPallet;
		}
		//=============================================================
	    
		i++;
    }

	for(i=1;i<=12;i++)
		Excel.SetColAutoFit(i);	

	SaveInvoiceDate(Excel);

	Excel.SaveAs(g_strSaveFile);
		MessageBox1("保存成功!", "保存成功", MB_ICONINFORMATION);
	
}

void CAutoWeight::ReloadData(void)
{
	CString cs, cs1;
	UpdateData(TRUE);
	m_strZCS.MakeUpper();
	mapInBox.RemoveAll();
	mapOutBox.RemoveAll();
	m_LayerList.RemoveAll();
	mapPartCnt.RemoveAll();
	mapWeightList.RemoveAll();
	mapOutboxPallet.RemoveAll();
	mapPalletCnt.RemoveAll();
	CString strFileName = CExcel::GetAppPath() + "\\Record List\\" + m_strZCS + ".ckd";
	if(IsFileExist(strFileName,FALSE)==TRUE)
	{
		//==============================================================================================
		//载入录入清单 复活 mapInBox mapOutBox
		m_ListMain.DeleteAllItems();
		g_iOrder = 0;

		CPropSheet* pParent = (CPropSheet*) GetParent();
		CPageShipMark* prop1 = (CPageShipMark*)pParent->GetPage(pParent->GetPageCount()-1);
		//序列化文件
		CFile file;
		file.Open(strFileName, CFile::modeReadWrite);
		CArchive ar(&file, CArchive::load);
		ar>>prop1->m_cs11>>prop1->m_cs12>>prop1->m_cs13>>prop1->m_cs14
			>>prop1->m_cs21>>prop1->m_cs22>>prop1->m_cs23>>prop1->m_cs24
			>>prop1->m_cs31>>prop1->m_cs32>>prop1->m_cs33>>prop1->m_cs34;
		m_LayerList.Serialize(ar);
		ar.Close();
		file.Close();

		POSITION pos = m_LayerList.GetHeadPosition();
		while (pos != NULL)
		{
			CLayers* pLayer = m_LayerList.GetNext(pos);
			int count = 0;
			g_iOrder++;
			//pLayer->strPallet = "";
			mapInBox.SetAt(pLayer->strInBox, *pLayer);

			if(mapOutBox.Lookup(pLayer->strOutBox,count))
				mapOutBox.SetAt(pLayer->strOutBox, count+1);
			else
				mapOutBox.SetAt(pLayer->strOutBox, 1);
			cs.Format("%d",pLayer->iQty);
			m_ListMain.AddItem(pLayer->strOrder,pLayer->strPallet,pLayer->strOutBox,pLayer->strInBox,cs);
		}

		m_ListMain.Sort(0,true);
		m_ListMain.EnsureVisible(g_iOrder-1,TRUE);
		//MessageBox1("套件载入成功!");
		//return true;
		//==================================================================================================

		//==================================================================================================
		//载入重量清单  复活 mapWeightList
		//郇工测量的重量清单，部品号作为关键键，重量单位为“g”
		strFileName = CExcel::GetAppPath() + "\\托盘清单\\"  + "_WeightList" + ".ckd";
		if(IsFileExist(strFileName,FALSE)==TRUE)
		{
			mapWeightList.RemoveAll();
			m_LayerList.RemoveAll();
			m_ListMain.DeleteAllItems();
			g_iOrder = 0;
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
				g_iOrder++;
				CLayers* pLayer = m_LayerList.GetNext(pos);
				mapWeightList.SetAt(pLayer->strPartNo, *pLayer);
				cs.Format("%g", pLayer->iUnitWeight);
				cs1.Format("%g", pLayer->iRadio);	
				m_ListMain.AddItem(pLayer->strNo,pLayer->strPartNo,cs,cs1,"");	
			}
			m_ListMain.Sort(0,true);
			m_ListMain.EnsureVisible(g_iOrder-1,TRUE);
			//MessageBox1("套件载入成功,重量清单载入成功!", "载入成功", MB_ICONINFORMATION);
		}
		else
		{
			MessageBox1("套件载入成功,重量清单未载入!", "载入", MB_ICONINFORMATION);
			m_strZCS.Empty();
			UpdateData(false);
			return ;
		}
		//==================================================================================================
		m_ListMain.DeleteAllItems();
		pos = mapInBox.GetStartPosition();
		CString strInBox;
		CLayers cV, cBase;
		g_iOrder = 0;
		CMapInBox mapLossWeight;
		BOOL bLossFlag = false;
		while(pos)
		{
			// 完善记录表 mapInBox 缺失的 [单位重量] [加乘系数]以及[单价]
			mapInBox.GetNextAssoc(pos, strInBox, cV);
			if(mapWeightList.Lookup(cV.strPartNo, cBase))
			{
				cV.iUnitWeight = cBase.iUnitWeight;
				cV.iRadio = cBase.iRadio;
				cV.iUnitPrice = cBase.iUnitPrice;
				mapInBox.SetAt(strInBox,cV);
			}
			else
			{
				#if TEST_AUTOWEIGHT
				cV.iUnitWeight = 0.0;//cBase.iUnitWeight;
				cV.iRadio = 1;//cBase.iRadio;
				cV.iUnitPrice = 0;
				mapInBox.SetAt(strInBox,cV);
				#else
				m_strZCS.Empty();
				UpdateData(false);
				bLossFlag = true;

				//============================================
				//如果有重量缺失情况，表格只做显示缺失项用
				//所以在第一次发现缺失重量时，即清空表格
				static BOOL bClearList = true;
				if(bClearList)
				{
					m_ListMain.DeleteAllItems();
					g_iOrder = 0;
					bClearList = false;
				}
				//============================================
				//============================================
				//显示缺失重量的项目
				if(!mapLossWeight.Lookup(cV.strPartNo, cBase))
				{
					g_iOrder++;
					m_ListMain.AddItem(cV.strNo,cV.strPartNo,"","","缺失");
					mapLossWeight.SetAt(cV.strPartNo,cV);
				}
				//============================================
				#endif

			}
			if(!cV.strPallet.IsEmpty()
				#if !TEST_AUTOWEIGHT
				&& bLossFlag == false
				#endif
				)
			{
				g_iOrder++;
				cs1.Format("%g",cV.iQty*cV.iUnitWeight*cV.iRadio);
				m_ListMain.AddItem(cV.strOrder,cV.strPallet,cV.strOutBox,cV.strInBox,cs1);	
			}
			m_ListMain.Sort(0,true);
			m_ListMain.EnsureVisible(g_iOrder-1,TRUE);
		}

		#if !TEST_AUTOWEIGHT
		if(bLossFlag)
		{
			MessageBox("重量缺失，缺失项详见列表，请补齐后重新载入清单！","重量缺失", MB_ICONERROR);
			return ;
		}
		#endif
		//==================================================================================================
		//载入托盘清单
		strFileName = CExcel::GetAppPath() + "\\Record List\\" + m_strZCS + "_PalletList" + ".ckd";
		if(IsFileExist(strFileName,FALSE)==TRUE)
		{
			CString cs1,cs2,strOutbox,strPallet;
			mapOutboxPallet.RemoveAll();
			g_iOrder = 0;
			//序列化文件
			CFile file;
			file.Open(strFileName, CFile::modeReadWrite);
			CArchive ar(&file, CArchive::load);
			mapOutboxPallet.Serialize(ar);
			ar.Close();
			file.Close();

			MessageBox1("套件载入成功,重量清单载入成功,托盘清单载入成功!", "托盘清单", MB_ICONINFORMATION);
		}
		else
		{
			pos = mapInBox.GetStartPosition();
			while(pos)
			{
				mapInBox.GetNextAssoc(pos, strInBox, cV);
				cV.strPallet.Empty();
				mapInBox.SetAt(strInBox,cV);
			}
			m_ListMain.DeleteAllItems();
			UpdateList();
			MessageBox1("新套件开始录入,重量清单载入成功!", "托盘清单", MB_ICONINFORMATION);
		}
		//==================================================================================================


		m_flgTaojian = TRUE;
	}
	else if(m_strZCS.IsEmpty())
	{
		MessageBox("请输入有效的套件号!","错误",MB_ICONERROR);
		m_flgTaojian = FALSE;
	}
	else
	{
		m_flgTaojian = FALSE;
		MessageBox1("不是有效套件,请输入有效套件!","提示",MB_OK);
	}

	CEdit* pEdit = (CEdit*)GetDlgItem(IDC_EDIT1_RECEIVE);
	pEdit->SetFocus();
}
