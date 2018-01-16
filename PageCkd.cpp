// Prop2.cpp : implementation file
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

#include <windows.h>
#include <mmsystem.h>
#include ".\pageckd.h"
#pragma comment(lib, "WINMM.LIB")


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

using std::bitset;
bitset<4> g_bitCheck;

/////////////////////////////////////////////////////////////////////////////
// CPageCkd property page

IMPLEMENT_DYNCREATE(CPageCkd, CPropertyPage)

CPageCkd::CPageCkd() : CPropertyPage(CPageCkd::IDD)
{
	//{{AFX_DATA_INIT(CPageCkd)
	m_csReceive = _T("");
	m_csPartNo = _T("");
	m_csInbox = _T("");
	m_csOutBox = _T("");
	m_csQty = _T("");
	m_csUndoInbox = _T("内/外箱号");
	m_unRatio = 1;
	m_strZCS = _T("");
	//}}AFX_DATA_INIT
	m_listInit = false;
}

CPageCkd::~CPageCkd()
{
}

void CPageCkd::DoDataExchange(CDataExchange* pDX)
{
	CPropertyPage::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CPageCkd)
	DDX_Control(pDX, IDC_LIST1, m_ListMain);
	DDX_Text(pDX, IDC_EDIT_RECEIVE, m_csReceive);
	DDX_Text(pDX, IDC_EDIT_PART, m_csPartNo);
	DDX_Text(pDX, IDC_EDIT_INBOX, m_csInbox);
	DDX_Text(pDX, IDC_EDIT_OUTBOX, m_csOutBox);
	DDX_Text(pDX, IDC_EDIT_QTY, m_csQty);
	DDX_Text(pDX, IDC_EDIT_UNDO, m_csUndoInbox);
	DDX_Text(pDX, IDC_EDIT_RATIO, m_unRatio);
	DDX_Text(pDX, IDC_EDIT_ZCS, m_strZCS);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CPageCkd, CPropertyPage)
	//{{AFX_MSG_MAP(CPageCkd)
	ON_BN_CLICKED(IDC_BUTTON1, OnButton1)
	ON_BN_CLICKED(IDC_BUTTON2, OnButton2)
	ON_BN_CLICKED(IDC_BTN_SAVE, OnBtnSave)
	ON_BN_CLICKED(IDC_BTN_SHIPPINGMARK, OnBtnLoadBaseList)
	ON_BN_CLICKED(IDC_BTN_RELOAD, OnBtnReload)
	ON_BN_CLICKED(IDC_BTN_CHECK, OnBtnCheck)
	ON_EN_SETFOCUS(IDC_EDIT_UNDO, OnSetfocusEditUndo)
	ON_WM_TIMER()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CPageCkd message handlers

BOOL CPageCkd::OnSetActive() 
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
		m_ListMain.SetHeadings("日期,60;部品号,110;外箱号,60;内箱号,60;数量,60"); ///设置列头信息 
		m_ListMain.LoadColumnInfo(); 


		m_listInit = true;
		m_flgTaojian = FALSE;
	}
	//=================================================
	//=== 不同模块交替使用时，要重新加载数据===========
	//else if(m_flgTaojian == TRUE && !m_strZCS.IsEmpty())
		//ReloadData();
	//=================================================

	return CPropertyPage::OnSetActive();
}

BOOL CPageCkd::PreTranslateMessage(MSG* pMsg) 
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
				case IDC_EDIT_RECEIVE: 
					UpdateData(TRUE);
					m_csReceive.MakeUpper();
					if(g_ttsCreate)
						tts.Pause();
					//m_csReceive.Replace(" ","");
					if(!m_flgTaojian)
					{
						MessageBox1("没有指定套件,请先指定套件号!","错误", MB_ICONERROR);
						m_csReceive.Empty();
						UpdateData(false);
						return true;
					}
					
					//  录入-内箱号
					if(m_csReceive.Find('[') != -1 && m_csReceive.Find(']') != -1)
					{
						//=====================================
						//1秒内连续输入两次内箱号，触发语音校验
						SetTimer(1,1500,NULL);
						if(++g_iSpeakCheckCnt >= 2)
						{
							g_iSpeakCheckCnt = 0;
							KillTimer(1);

							g_ssk.Clear();
							g_bitCheck.reset();
							m_csPartNo.Empty(); m_csOutBox.Empty(); m_csQty.Empty();m_csInbox.Empty();
							m_csReceive.Empty();
							UpdateData(FALSE);
							OnBtnCheck();
							return true;
						}
						//=======================================

						CString csTemp;  
						csTemp = m_csReceive.Mid(m_csReceive.Find('[')+1,m_csReceive.Find(']')-m_csReceive.Find('[')-1);
						if(m_strZCS.IsEmpty())
						{
							MessageBox1("未指定套件号,请先指定套件号!","错误", MB_ICONERROR);
							return true;
						}
						else if(!m_csInbox.IsEmpty() && m_csInbox != csTemp && !m_csPartNo.IsEmpty())
						{
							cs1 = "内箱" + csTemp + "开始录入!" + "内箱" + m_csInbox + "将取消";
							cs1.Replace("IN","");
							MessageBox1(cs1,"提示-内箱号变更",MB_OK);
						}
						else
						{
							cs1.Format("乘%d",m_unRatio);
							if(g_flgSpeak)
								Speak(cs1);
							else
								PlayVoice();
						}

						m_csInbox = csTemp;

						g_bitCheck.set(STYLE_INBOX);
						g_ssk.Clear();
						m_csPartNo.Empty(); m_csOutBox.Empty(); m_csQty.Empty();
					}
					//  录入-部品号-数量
					else if(m_csReceive.Find('(') != -1 && m_csReceive.Find(')') != -1)
					{
						if(g_bitCheck[STYLE_INBOX])
						{
							m_csPartNo = m_csReceive.SpanExcluding("()");
							m_csQty = m_csReceive.Mid(m_csReceive.Find('(')+1,m_csReceive.Find(')')-m_csReceive.Find('(')-1);
							g_bitCheck.set(STYLE_PARTNO);
							g_bitCheck.set(STYLE_QTY);		
							g_ssk.iQty += (atoi(m_csQty) * m_unRatio);
							m_csPartNo = CExcel::DeleteBlackSpace(m_csPartNo);
							UpdateData(FALSE);
							if(g_ssk.strPartNo.IsEmpty() || g_ssk.strPartNo == m_csPartNo)
							{
								int iRecordTotal = 0;	//已经录入的总数量
								CLayers cBase;
								CString csTotal;		//单次录入数量
								g_ssk.strPartNo = m_csPartNo;
								if(mapPartCnt.Lookup(g_ssk.strPartNo,iRecordTotal))
									iRecordTotal += g_ssk.iQty;
								else
									iRecordTotal = g_ssk.iQty;
								if(mapBaseList.Lookup(g_ssk.strPartNo,cBase))
								{
									if(iRecordTotal == cBase.iQty)
									{
										csTotal.Format("%d这颗物料已经齐套",g_ssk.iQty);
										if(g_flgSpeak)
										Speak(csTotal);
									}
									else
									{
										csTotal.Format("%d",g_ssk.iQty);
										if(g_flgSpeak)
										Speak(csTotal);
									}
								}
								else
								{
									csTotal.Format("%d",g_ssk.iQty);
									if(g_flgSpeak)
									Speak(csTotal);
								}
							}
							else
							{
								g_ssk.Clear();
								g_bitCheck.reset();
								m_csPartNo.Empty(); m_csOutBox.Empty(); m_csInbox.Empty(); m_csQty.Empty();
								UpdateData(FALSE);
								MessageBox1("内箱号存在不同部品号,请重新开始录入!","错误", MB_ICONERROR);
							}
						}
						else
						{
							MessageBox1("内箱号不存在,请先录入内箱号!","内箱号不存在!",MB_ICONERROR);
						}
					}
					//  录入-系数
					else if(m_csReceive.Find('<') != -1 && m_csReceive.Find('>') != -1)
					{
						CString cs;
						cs = m_csReceive.Mid(m_csReceive.Find('<')+1,m_csReceive.Find('>')-m_csReceive.Find('<')-1);
						m_unRatio = atoi(cs);
						UpdateData(false);
						cs.Format("乘%d",m_unRatio);
						if(g_flgSpeak)
							Speak(cs);
						else
							PlayVoice();
					}
					//  录入-外箱号
					else if(m_csReceive.Find('{') != -1 && m_csReceive.Find('}') != -1)
					{
						m_csOutBox = m_csReceive.Mid(m_csReceive.Find('{')+1,m_csReceive.Find('}')-m_csReceive.Find('{')-1);
						g_bitCheck.set(STYLE_OUTBOX);

						if(g_bitCheck.count() == STYLE_END)
						{
							g_iOrder++;
							CTime tm = CTime::GetCurrentTime();  
							CString strTime = tm.Format("%m-%d %X"); 
							g_ssk.strOrder = strTime;
							g_ssk.strOutBox = m_csOutBox;
							g_ssk.strInBox = m_csInbox;	
							

							//======================================================
							//检查[CKD]特殊内箱
							if(m_csInbox == INBOX_COMMON)
							{
								g_ssk.strInBox = m_csInbox + strTime;
								POSITION pos = mapInBox.GetStartPosition();
								CString strKey, strInboxCKD;
								CLayers cValue;							
								while(pos)
								{
									mapInBox.GetNextAssoc(pos, strKey, cValue);
									if(cValue.strOutBox == m_csOutBox)
									{
										strInboxCKD = cValue.strInBox;
										mapInBox.RemoveKey(strInboxCKD);
										break;
									}
								}
							}
							//================================================================

							mapInBox.SetAt(g_ssk.strInBox,g_ssk);
							g_ssk.Clear();
							g_bitCheck.reset();
							m_csPartNo.Empty(); m_csOutBox.Empty(); m_csInbox.Empty(); m_csQty.Empty();
							UpdateData(false);
							OnBtnReload();
							if(g_flgSpeak)
								Speak("录入成功");
							else
								PlaySound(CExcel::GetAppPath() + "\\Voice\\NOKIA_clip.wav", NULL, SND_FILENAME | SND_ASYNC);
						}
						else if(!g_bitCheck[STYLE_INBOX])
						{
							g_bitCheck.reset();
							m_csPartNo.Empty(); m_csOutBox.Empty(); m_csInbox.Empty(); m_csQty.Empty();
							MessageBox1("内箱号未录入,请先录入内箱号!","内箱号未录入",MB_ICONERROR);
						}
						else if(!g_bitCheck[STYLE_PARTNO])
						{
							m_csOutBox.Empty();
							MessageBox1("部品号未录入,请重新录入部品号!","部品号未录入",MB_ICONERROR);
						}
						
					}
					m_csReceive.Empty();
					UpdateData(false);
					//MessageBox(m_csReceive);
					return true;
					break;
				case IDC_EDIT_RATIO:
					{
						UpdateData(TRUE);
						CString cs;
						cs.Format("乘%d",m_unRatio);
						UpdateData(false);
						if(g_flgSpeak)
						Speak(cs);
						else
						PlayVoice();
						UpdateData(FALSE);
						CEdit* pEdit = (CEdit*)GetDlgItem(IDC_EDIT_RECEIVE);
						pEdit->SetFocus();
					}
					return true;
					break;
				case IDC_EDIT_UNDO:
					{
						int u32OutBox = 0;
						UpdateData(TRUE);
						m_csUndoInbox.MakeUpper();
						if(m_csUndoInbox.Find("IN") != -1)
						{
							if(mapInBox.RemoveKey(m_csUndoInbox))
							{
								OnBtnReload();
								m_csUndoInbox = "内箱" + m_csUndoInbox + "撤销成功!";
								m_csUndoInbox.Replace("IN","");
								MessageBox1(m_csUndoInbox,"撤销成功!",MB_OK);
							}
							else
								MessageBox1("未找到要撤销的内箱号","撤销错误!",MB_ICONERROR);
						}
						else if(m_csUndoInbox.Find('C') != -1)
						{
							if(mapOutBox.Lookup(m_csUndoInbox,u32OutBox))
							{
								POSITION pos = mapInBox.GetStartPosition();
								CString strKey;
								CLayers cValue;
								while(pos)
								{
									mapInBox.GetNextAssoc(pos, strKey, cValue);
									if(cValue.strOutBox == m_csUndoInbox)
										mapInBox.RemoveKey(strKey);
								}
								OnBtnReload();
								m_csUndoInbox = "外箱" + m_csUndoInbox + "撤销成功!";
								m_csUndoInbox.Replace("C","");
								MessageBox1(m_csUndoInbox,"撤销成功!",MB_OK);
							}
							else
								MessageBox1("未找到要撤销的外箱号","撤销错误!",MB_ICONERROR);

						}
						m_csUndoInbox.Empty();
						UpdateData(FALSE);
						CEdit* pEdit = (CEdit*)GetDlgItem(IDC_EDIT_RECEIVE);
						pEdit->SetFocus();
					}
					return true;
					break;
				case IDC_EDIT_ZCS:
					ReloadData();
					return true;
					break;
				default:
					{
						CEdit* pEdit = (CEdit*)GetDlgItem(IDC_EDIT_RECEIVE);
						pEdit->SetFocus();
					}
					return true;
					break;
			}
		}
	}
	
	return CPropertyPage::PreTranslateMessage(pMsg);
}



//============================
//基础为内箱清单
//1、补齐内箱清单缺失项
//2、重新生成外箱清单
//3、重新生成部品号-数量清单
//===========================

void CPageCkd::OnBtnReload() 
{
	// TODO: Add your control notification handler code here
	if(mapInBox.GetCount() == 0)
	{
		MessageBox1("未录入数据,请先录入!","未录入数据",MB_ICONERROR);
		return ;
	}

	m_ListMain.DeleteAllItems();
	g_iReOrder = 0;
	POSITION pos = mapInBox.GetStartPosition();
	CString strKey;
	CLayers cValue,cBase;
	CString cs;
	int iFindCnt = 0;
	m_LayerList.RemoveAll();
	mapOutBox.RemoveAll();
	mapPartCnt.RemoveAll();
    while(pos)
    {
		g_iReOrder++;
		mapInBox.GetNextAssoc(pos, strKey, cValue);
		cs.Format("%d",cValue.iQty);
		if(strKey.Find(INBOX_COMMON) != -1)
			m_ListMain.AddItem(cValue.strOrder,cValue.strPartNo,cValue.strOutBox,INBOX_COMMON,cs);
		else
			m_ListMain.AddItem(cValue.strOrder,cValue.strPartNo,cValue.strOutBox,strKey,cs);

		//每次录入新记录，补齐缺失项
		if(mapBaseList.Lookup(cValue.strPartNo, cBase))
		{
			cValue.strUserCode = cBase.strUserCode;
			cValue.strDetail = cBase.strDetail;
			cValue.strNo = cBase.strNo;
			cValue.strFactoryNo = cBase.strFactoryNo;
			mapInBox.SetAt(strKey,cValue);
		}
		else if(!mapBaseList.IsEmpty() && iFindCnt<=5)
		{
			cs = "内箱" + cValue.strInBox + "中的物料在基础清单中未找到";
			cs.Replace("IN","");
			if(g_flgSpeak)
				Speak(cs);
			iFindCnt++;
		}
		
		//生成保存清单容器
		CLayers* pLayersItem = new CLayers(cValue);
		m_LayerList.AddTail(pLayersItem);
		//生成外箱清单
		int count = 0;
		if(mapOutBox.Lookup(cValue.strOutBox,count))
			mapOutBox.SetAt(cValue.strOutBox, count+1);
		else
			mapOutBox.SetAt(cValue.strOutBox, 1);
		// 生成 "部品号-总数量" 的映射表
		int iTemp = 0;
		if(mapPartCnt.Lookup(cValue.strPartNo, iTemp))
		{
			iTemp += cValue.iQty;
			mapPartCnt.SetAt(cValue.strPartNo,iTemp);
		}
		else
			mapPartCnt.SetAt(cValue.strPartNo,cValue.iQty);
    }

	m_ListMain.Sort(0,true);
	m_ListMain.EnsureVisible(g_iReOrder-1,TRUE);


	CPropSheet* pParent = (CPropSheet*) GetParent();
	CPageShipMark* prop1 = (CPageShipMark*)pParent->GetPage(pParent->GetPageCount()-1);
	//保存清单
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
	
}

template <typename T, typename U> void BubbleSort1(T& collection, U element, int count, bool ascend = true)
{
	for (int i = 0; i < count-1; i++)
	for (int j = 0; j < count-1-i; j++)
		if (ascend)
		{
		// 升序
			if (collection[j] - collection[j+1])
			{
				U temp = collection[j];
				collection[j] = collection[j+1];
				collection[j+1] = temp;
			}
		}
		else
		{
		// 降序
			if (collection[j] + collection[j+1])
			{
			U temp = collection[j];
			collection[j] = collection[j+1];
			collection[j+1] = temp;
			}
		}
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

void CPageCkd::SaveAccordingDate(CExcel &Excel) 
{
	CString strCell;
	Excel.SelectSheet(2);				        // 激活工作簿1
	Excel.ActiveSheet().SetName("汇总");

	Excel.SetColWidth(1,14);
	Excel.SetCell(1,1,"料号");
	Excel.SetCell(1,2,"部品号");
	Excel.SetCell(1,3,"时间");
	Excel.SetCell(1,4,"单日数量");
	Excel.SetCell(1,5,"总计");
	Excel.SetCell(1,6,"需求数量");
	Excel.SetCell(1,7,"状态");

	POSITION pos = mapInBox.GetStartPosition();
	CString strKey;
	CLayers cValue,cValue1,cBase;
	int i = 2;


	CArrayClayer NoArray;
	CArrayClayer DateArray;
	//================================================================
	//按料号排序
	while(pos)
	{
		mapInBox.GetNextAssoc(pos, strKey, cValue);
		NoArray.Add(cValue);
	}
	g_OrderMode = ORDER_NO;
	BubbleSort(NoArray, NoArray[0], NoArray.GetSize(), true);
	//================================================================
	unsigned int iBorder = 2;	
    for(int j=0; j<NoArray.GetSize();)
    {
		//===========================================
		//检索相同部品号 放进DateArray数组
		unsigned int count = 1;
		cValue = NoArray[j];
		cValue.strOrder = cValue.strOrder.Left(5);
		DateArray.Add(cValue);
		for(int k=j+1; k<NoArray.GetSize(); k++)
		{
			if(cValue.strNo == NoArray[k].strNo)
			{
				cValue1 = NoArray[k];
				cValue1.strOrder = NoArray[k].strOrder.Left(5);
				DateArray.Add(cValue1);
				count++;
			}
			else
				break;
		}
		j = j+count;
		//==============================================
		
		//==============================================
		//相同料号按日期排序
		g_OrderMode = ORDER_DATE;
		BubbleSort(DateArray, DateArray[0], DateArray.GetSize(), true);
		unsigned int iTotal = 0;
		Excel.SelectSheet(2);  
		iBorder = i;
		for(int w=0; w<DateArray.GetSize();)
		{
			cValue = DateArray[w];
			count = 1;
			//========================================
			//合并相同日期的单元
			
			for(int y=w+1; y<DateArray.GetSize(); y++)
			{
				if(cValue.strOrder == DateArray[y].strOrder)
				{
					cValue.iQty += DateArray[y].iQty;
					count++;
				}
				else
					break;
			}
			
			//========================================
			w = w + count;

			Excel.SetCell(i,3, cValue.strOrder);
			Excel.SetCell(i,4,(long)cValue.iQty);
			iTotal += cValue.iQty;
			i++;
		}

		CRange range(Excel.GetRange(iBorder,1,i-1,7));
		range.Border();	
		//range.Merge();
		Excel.SetCell(1,1, cValue.strNo);
		Excel.SetCell(1,2,cValue.strPartNo);
		Excel.SetCell(1,5,(long)iTotal);
		if(mapBaseList.Lookup(cValue.strPartNo, cBase))
		{
			Excel.SetCell(1,6,(long)cBase.iQty);
		}
		if(iTotal == cBase.iQty)
			Excel.SetCell(1,7,"齐套");
		else
			Excel.SetCell(1,7,"未齐套");

		
		
		DateArray.RemoveAll();
		
    }
	for(i=1;i<=11;i++)
		Excel.SetColAutoFit(i);	

}

void CPageCkd::SaveBoxList(CExcel &Excel)
{
	Excel.SelectSheet(1);
	Excel.ActiveSheet().SetName("箱单");

	Excel.SetColWidth(1,14);
	Excel.SetCell(1,1,"No.");
	Excel.SetCell(1,2,"料号");
	Excel.SetCell(1,3,"客户品号");
	Excel.SetCell(1,4,"描述");
	Excel.SetCell(1,5,"部品号");
	Excel.SetCell(1,6,"发货数量");
	Excel.SetCell(1,7,"托盘号");
	Excel.SetCell(1,8,"外箱箱号");
	Excel.SetCell(1,9,"内箱号");
	Excel.SetCell(1,10,"内箱数量");
	Excel.SetCell(1,11,"数量-内箱");
	Excel.SetCell(1,12,"总数量");

	POSITION pos = mapInBox.GetStartPosition();
	CString strKey;
	CLayers cValue;
	int i = 2;


	CArrayClayer dateArray;
	//================================================================
	//按日期排序
	while(pos)
	{
		mapInBox.GetNextAssoc(pos, strKey, cValue);
		dateArray.Add(cValue);
	}
	g_OrderMode = ORDER_DATE;
	BubbleSort(dateArray, dateArray[0], dateArray.GetSize(), true);
	//================================================================

    for(int j=0; j<dateArray.GetSize(); j++)
    {
		cValue = dateArray[j];
		Excel.SetCell(i,1, cValue.strOrder);
		Excel.SetCell(i,2, cValue.strNo);
		Excel.SetCell(i,3,cValue.strUserCode);
		Excel.SetCell(i,4,cValue.strDetail);
		Excel.SetCell(i,5,cValue.strPartNo);
		
		Excel.SetCell(i,8,cValue.strOutBox);
		if(cValue.strInBox.Find(INBOX_COMMON) != -1)
			Excel.SetCell(i,9,INBOX_COMMON);
		else
			Excel.SetCell(i,9,cValue.strInBox);
		Excel.SetCell(i,10,"1");
		Excel.SetCell(i,11,(long)cValue.iQty);	
		Excel.SetCell(i,12,(long)cValue.iQty);

		i++;
    }

	for(i=1;i<=11;i++)
		Excel.SetColAutoFit(i);	
}

void CPageCkd::OnBtnSave()
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
	CSaveDlg saveDlg;
	if(IDOK == saveDlg.DoModal())
	{
		CFileDialog dlg(FALSE, "xls", "", OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, "Excel文件(*.xls)|*.xls" );  
		//===========================================================
		//保存到指定目录
		CString strFileName = CExcel::GetAppPath() + "\\唛头文件";
		dlg.m_ofn.lpstrInitialDir = strFileName;
		//===========================================================
		if(dlg.DoModal()==IDOK)
			g_strSaveFile=dlg.GetPathName();

		if(g_strSaveFile.Find(".xls")<0)        //没有保存则直接返回
			return;

		CString strCell;
		CExcel Excel;
		Excel.AddNewFile();				            // 新建一个文件
		Excel.SetVisible(FALSE);					// 设置不可见

		m_iSaveMode = g_iSaveMode = saveDlg.m_iSaveMode;
		m_bOnlyOnce = saveDlg.m_bSaveDayOnly;
		InitConfigFile(false);

		if(saveDlg.m_bBoxList)
			SaveBoxList(Excel);

		if(saveDlg.m_bDayCount)
			SaveAccordingDate(Excel);

		if(saveDlg.m_bShipMark)
			SaveShippingMark(Excel);

		Excel.SaveAs(g_strSaveFile);
		MessageBox1("保存成功!", "保存成功", MB_ICONINFORMATION);
	}
}

void CPageCkd::SaveShippingMark(CExcel &Excel) 
{
	CString strCell;
	Excel.SelectSheet(3);				        // 激活工作簿1
	Excel.ActiveSheet().SetName("唛头");

	CPropSheet* pParent = (CPropSheet*) GetParent();//首先获得属性页父窗口指针
	CPageShipMark* prop1 = (CPageShipMark*)pParent->GetPage(pParent->GetPageCount()-1);

	//////////////////////////////////////////////////////////////////
	//设置表头
	//Excel.InsertRow(1);
	for(int j=1;j<=SHIPPING_WIDTH;j++)
		Excel.SetColWidth(j,1);

	CString strKey;
	int count = 0;
	int i =2;

	//=========================================================
	//检索所有外箱  并按外箱排序
	CStringArray arOutBox;
	CStringArray arOutBox1;
	
	POSITION pos = mapOutBox.GetStartPosition();
	while(pos)
    {
		mapOutBox.GetNextAssoc(pos, strKey, count);
		if(strKey.GetLength()<=4)
			arOutBox.Add(strKey);
		else
			arOutBox1.Add(strKey);
	}
	if(arOutBox.GetSize()>0)
		BubbleSort(arOutBox, arOutBox[0], arOutBox.GetSize(), true);

	//==========================================================
	//外箱标签打印设定为4位，如C002 C200.
	//多于C999, 如C1000会排在C999的前面.多于5位的单独排列
	if(arOutBox1.GetSize()>0)
	{
		BubbleSort(arOutBox1, arOutBox1[0], arOutBox1.GetSize(), true);
		for(int w=0; w<arOutBox1.GetSize(); w++)
			arOutBox.Add(arOutBox1[w]);
	}
	//=========================================================

	//==========================================================

    for(int k=0; k<arOutBox.GetSize(); k++)
    {
		mapOutBox.Lookup(arOutBox[k],count);

		CRange range(Excel.GetRange(i,SHIPPING_WOFFSET,i+SHIPPING_HIGH,SHIPPING_WIDTH));
		range.Border();	
		
		Excel.SetRowWidth(i+1,TITLE_ROW_WIDTH);
		Excel.SetRowWidth(i+3,TITLE_ROW_WIDTH);
		Excel.SetRowWidth(i+5,TITLE_ROW_WIDTH);
		Excel.SetRowWidth(i+7,TITLE_ROW_WIDTH);
		Excel.SetRowWidth(i+9,TITLE_ROW_WIDTH);
		
		//  Shipping Mark
		range = Excel.GetRange(i,SHIPPING_WOFFSET,i,SHIPPING_WIDTH);
		range.Border();	
		range.Merge();	
		range = "Shipping Mark";
		range.SetHAlign(HAlignCenter);

		///////////////////////////////////
		//  CKD
		range = Excel.GetRange(i+2,SHIPPING_WOFFSET+TITLE1_COL_OFFSET,i+2,SHIPPING_WOFFSET+TITLE1_COL_OFFSET+1);
		range.Border();	
		range.Merge();
		range = "●";
		range.SetHAlign(HAlignCenter);
		range = Excel.GetRange(i+2,SHIPPING_WOFFSET+TITLE1_COL_OFFSET+2,i+2,SHIPPING_WOFFSET+TITLE1_COL_OFFSET+2+TITLE_WCELL1-2);
		range.Merge();	
		range = "CKD";

		//  ASSY
		range = Excel.GetRange(i+2,SHIPPING_WOFFSET+TITLE2_COL_OFFSET,i+2,SHIPPING_WOFFSET+TITLE2_COL_OFFSET+TITLE_WCELL);
		range.Merge();	
		range = "ASS'Y";

		//  Spare Part
		range = Excel.GetRange(i+2,SHIPPING_WOFFSET+TITLE3_COL_OFFSET,i+2,SHIPPING_WOFFSET+TITLE3_COL_OFFSET+TITLE_WCELL);
		range.Merge();	
		range = "Spare Part";

		///////////////////////////////////
		//  LCD/3240005210
		range = Excel.GetRange(i+4,SHIPPING_WOFFSET+TITLE1_COL_OFFSET,i+4,SHIPPING_WOFFSET+TITLE1_COL_OFFSET+TITLE_WCELL1);
		range.Merge();
		range = prop1->m_cs11;
		
		range = Excel.GetRange(i+4,SHIPPING_WOFFSET+TITLE2_COL_OFFSET,i+4,SHIPPING_WOFFSET+TITLE2_COL_OFFSET+TITLE_WCELL);
		range.Border();	
		range.Merge();	
		range = prop1->m_cs12;

		//  外箱号
		range = Excel.GetRange(i+4,SHIPPING_WOFFSET+TITLE3_COL_OFFSET,i+4,SHIPPING_WOFFSET+TITLE3_COL_OFFSET+TITLE_WCELL3_4);
		range.Merge();	
		range = prop1->m_cs13;

		///////////////////////////////////
		//  Supplier Name
		range = Excel.GetRange(i+6,SHIPPING_WOFFSET+TITLE1_COL_OFFSET,i+6,SHIPPING_WOFFSET+TITLE1_COL_OFFSET+TITLE_WCELL1);
		range.Merge();
		range = prop1->m_cs21;
		//  公司
		range = Excel.GetRange(i+6,SHIPPING_WOFFSET+TITLE2_COL_OFFSET,i+6,SHIPPING_WOFFSET+TITLE2_COL_OFFSET+TITLE_WCELL);
		range.Border();	
		range.Merge();	
		range = prop1->m_cs22;

		///////////////////////////////////
		//  Date
		range = Excel.GetRange(i+6,SHIPPING_WOFFSET+TITLE3_COL_OFFSET,i+6,SHIPPING_WOFFSET+TITLE3_COL_OFFSET+TITLE_WCELL3_4);
		range.Merge();	
		range.SetHAlign(HAlignLeft);
		range = prop1->m_cs23;
		//  日期
		range = Excel.GetRange(i+6,SHIPPING_WOFFSET+TITLE4_COL_OFFSET,i+6,SHIPPING_WOFFSET+TITLE4_COL_OFFSET+TITLE_WCELL3_4);
		range.Border();	
		range.Merge();	
		range.SetHAlign(HAlignLeft);
		range = prop1->m_cs24;

		///////////////////////////////////
		//  出口国
		range = Excel.GetRange(i+8,SHIPPING_WOFFSET+TITLE1_COL_OFFSET,i+8,SHIPPING_WOFFSET+TITLE1_COL_OFFSET+TITLE_WCELL1);
		range.Merge();
		range = prop1->m_cs31;

		range = Excel.GetRange(i+8,SHIPPING_WOFFSET+TITLE2_COL_OFFSET,i+8,SHIPPING_WOFFSET+TITLE2_COL_OFFSET+TITLE_WCELL);
		range.Border();	
		range.Merge();	
		range = prop1->m_cs32;

		//  生产国
		range = Excel.GetRange(i+8,SHIPPING_WOFFSET+TITLE3_COL_OFFSET,i+8,SHIPPING_WOFFSET+TITLE3_COL_OFFSET+TITLE_WCELL3_4);
		range.Merge();
		range = prop1->m_cs33;

		range = Excel.GetRange(i+8,SHIPPING_WOFFSET+TITLE4_COL_OFFSET,i+8,SHIPPING_WOFFSET+TITLE4_COL_OFFSET+TITLE_WCELL3_4);
		range.Border();	
		range.Merge();	
		range = prop1->m_cs34;

		//  PACKING DETAILS
		range = Excel.GetRange(i+10,SHIPPING_WOFFSET,i+10,SHIPPING_WIDTH);
		range.Border();	
		range.Merge();	
		range = "PACKING DETAILS";
		range.SetHAlign(HAlignCenter);
				
		Excel.SelectSheet(3);
		POSITION pos = mapInBox.GetStartPosition();
		CArrayClayer inBoxArray;
		CString strinbox;
		CLayers cValue;

		/////////////////////////////////////////////////////////////////
		//检索某个外箱的所有内箱 并按内箱排序
		while(pos)
		{
			mapInBox.GetNextAssoc(pos, strinbox, cValue);
			if(arOutBox[k] == cValue.strOutBox)
			{
				inBoxArray.Add(cValue);
			}
		}
		g_OrderMode = ORDER_INBOX;
		BubbleSort(inBoxArray, inBoxArray[0], inBoxArray.GetSize(), true);
		//////////////////////////////////////////////////////////////////

		int x = 1;
		switch(m_iSaveMode)
		{
		case 0:
			for(int j=0; j<inBoxArray.GetSize(); j++)
			{
				Excel.SelectSheet(3);
				cValue = inBoxArray[j];
				{
					Excel.InsertRow(i+SHIPPING_HIGH+x);
					//内箱号
					range = Excel.GetRange(i+SHIPPING_HIGH+x,CONTENT1_COL_OFFSET,i+SHIPPING_HIGH+x,CONTENT2_COL_OFFSET+1);
					range.Border(1,1);	
					range.Merge();
					if(cValue.strInBox.Find(INBOX_COMMON) != -1)
						range = INBOX_COMMON;
					else
						range = cValue.strInBox;
					//客户号
					range = Excel.GetRange(i+SHIPPING_HIGH+x,CONTENT2_COL_OFFSET+2,i+SHIPPING_HIGH+x,CONTENT3_COL_OFFSET-1);
					range.Border(1,1);	
					range.Merge();	
	
					//描述
					range = Excel.GetRange(i+SHIPPING_HIGH+x,CONTENT3_COL_OFFSET,i+SHIPPING_HIGH+x,CONTENT4_COL_OFFSET-1);
					range.Border(1,1);	
					range.Merge();	
					range = cValue.strDetail;
					//数量
					CString csTemp;
					csTemp.Format("%d",cValue.iQty);
					range = Excel.GetRange(i+SHIPPING_HIGH+x,CONTENT4_COL_OFFSET,i+SHIPPING_HIGH+x,CONTENT5_COL_OFFSET-1);
					range.Border(1,1);	
					range.Merge();	
					range = csTemp;
					//部品号
					range = Excel.GetRange(i+SHIPPING_HIGH+x,CONTENT5_COL_OFFSET,i+SHIPPING_HIGH+x,SHIPPING_WIDTH);
					range.Border(1,1);	
					range.Merge();	
					range = cValue.strPartNo;
					

					//外箱号
					range = Excel.GetRange(i+4,SHIPPING_WOFFSET+TITLE4_COL_OFFSET,i+4,SHIPPING_WOFFSET+TITLE4_COL_OFFSET+TITLE_WCELL3_4);
					range.Border();	
					range.Merge();	
					range = cValue.strOutBox;

					x++;
				}
			}
			break;
		case 1:
		default:
			for(int j=0; j<inBoxArray.GetSize(); j++)
			{
				Excel.SelectSheet(3);
				cValue = inBoxArray[j];
				{
					Excel.InsertRow(i+SHIPPING_HIGH+x);
					//内箱号
					range = Excel.GetRange(i+SHIPPING_HIGH+x,CONTENT1_COL_OFFSET,i+SHIPPING_HIGH+x,CONTENT2_COL_OFFSET-1);
					range.Border(1,1);	
					range.Merge();
					if(cValue.strInBox.Find(INBOX_COMMON) != -1)
						range = INBOX_COMMON;
					else
						range = cValue.strInBox;
					//客户号
					range = Excel.GetRange(i+SHIPPING_HIGH+x,CONTENT2_COL_OFFSET,i+SHIPPING_HIGH+x,CONTENT3_COL_OFFSET-1);
					range.Border(1,1);	
					range.Merge();	
					range = cValue.strUserCode;
					//描述
					range = Excel.GetRange(i+SHIPPING_HIGH+x,CONTENT3_COL_OFFSET,i+SHIPPING_HIGH+x,CONTENT4_COL_OFFSET-1);
					range.Border(1,1);	
					range.Merge();	
					range = cValue.strDetail;
					//数量
					CString csTemp;
					csTemp.Format("%d",cValue.iQty);
					range = Excel.GetRange(i+SHIPPING_HIGH+x,CONTENT4_COL_OFFSET,i+SHIPPING_HIGH+x,CONTENT5_COL_OFFSET-1);
					range.Border(1,1);	
					range.Merge();	
					range = csTemp;
					//部品号
					range = Excel.GetRange(i+SHIPPING_HIGH+x,CONTENT5_COL_OFFSET,i+SHIPPING_HIGH+x,SHIPPING_WIDTH);
					range.Border(1,1);	
					range.Merge();	
					range = cValue.strPartNo;
					

					//外箱号
					range = Excel.GetRange(i+4,SHIPPING_WOFFSET+TITLE4_COL_OFFSET,i+4,SHIPPING_WOFFSET+TITLE4_COL_OFFSET+TITLE_WCELL3_4);
					range.Border();	
					range.Merge();	
					range = cValue.strOutBox;

					x++;
				}
			}
			break;
		case 2:
			for(int j=0; j<inBoxArray.GetSize(); j++)
			{
				Excel.SelectSheet(3);
				cValue = inBoxArray[j];
				{
					Excel.InsertRow(i+SHIPPING_HIGH+x);
					//内箱号
					range = Excel.GetRange(i+SHIPPING_HIGH+x,CONTENT1_COL_OFFSET_2,i+SHIPPING_HIGH+x,CONTENT2_COL_OFFSET_2-1);
					range.Border(1,1);	
					range.Merge();
					if(cValue.strInBox.Find(INBOX_COMMON) != -1)
						range = INBOX_COMMON;
					else
						range = cValue.strInBox;

					//客户号
					range = Excel.GetRange(i+SHIPPING_HIGH+x,CONTENT2_COL_OFFSET_2,i+SHIPPING_HIGH+x,CONTENT3_COL_OFFSET_2-1);
					range.Border(1,1);	
					range.Merge();	
					range = cValue.strUserCode;

					//描述
					range = Excel.GetRange(i+SHIPPING_HIGH+x,CONTENT3_COL_OFFSET_2,i+SHIPPING_HIGH+x,CONTENT4_COL_OFFSET_2-1);
					range.Border(1,1);	
					range.Merge();	
					range = cValue.strDetail;
					//数量
					CString csTemp;
					csTemp.Format("%d",cValue.iQty);
					range = Excel.GetRange(i+SHIPPING_HIGH+x,CONTENT4_COL_OFFSET_2,i+SHIPPING_HIGH+x,CONTENT5_COL_OFFSET_2-1);
					range.Border(1,1);	
					range.Merge();	
					range = csTemp;

					//部品号
					range = Excel.GetRange(i+SHIPPING_HIGH+x,CONTENT5_COL_OFFSET_2,i+SHIPPING_HIGH+x,CONTENT6_COL_OFFSET_2-1);
					range.Border(1,1);	
					range.Merge();	
					range = cValue.strPartNo;

					range = Excel.GetRange(i+SHIPPING_HIGH+x,CONTENT6_COL_OFFSET_2,i+SHIPPING_HIGH+x,SHIPPING_WIDTH);
					range.Border(1,1);	
					range.Merge();
					{
						CLayers cBase;
						if(mapBaseList.Lookup(cValue.strPartNo, cBase))
							range = cBase.strFactoryNo;
					}
					

					//外箱号
					range = Excel.GetRange(i+4,SHIPPING_WOFFSET+TITLE4_COL_OFFSET,i+4,SHIPPING_WOFFSET+TITLE4_COL_OFFSET+TITLE_WCELL3_4);
					range.Border();	
					range.Merge();	
					range = cValue.strOutBox;

					x++;
				}
			}
			break;
		}
		range = Excel.GetRange(i+SHIPPING_HIGH+1,SHIPPING_WOFFSET,i+SHIPPING_HIGH+x,SHIPPING_WIDTH);
		range.SetHAlign(HAlignLeft);

		range = Excel.GetRange(i+SHIPPING_HIGH+1,SHIPPING_WOFFSET,i+SHIPPING_HIGH+1+count,SHIPPING_WIDTH);
		range.Border();
		
		i += (count + SHIPPING_HIGH + SHIPPING_HOFFSET);
	}
	
}


void CPageCkd::OnBtnLoadBaseList() 
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
	m_icolProductCnt = m_icolPart = m_irowPart = m_icolDetail = m_icolUserCode = m_icolFactoryNo = 0;
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
			else if(strcmp(pCell,TOULIAO_LIST_FACTORY_NO) == 0)
			{
				m_icolFactoryNo = j;
			}

			if(m_icolProductCnt && m_icolPart && m_irowPart && m_icolDetail && m_icolUserCode && m_icolNo && m_icolFactoryNo)
			{
				m_flgCheck = true;
				break;
			}
		}

		if(m_icolProductCnt && m_icolPart && m_irowPart && m_icolDetail && m_icolUserCode && m_icolNo && m_icolFactoryNo)
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
		//CString css;
		//css.Format("%d-%d-%d-%d-%d-%d-%d",m_icolProductCnt,m_icolPart,m_irowPart,m_icolDetail,m_icolUserCode,m_icolNo,m_icolFactoryNo);
		//MessageBox(css);
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
			//strCell.MakeUpper();
			ssk.strDetail = strCell;

			strCell = Excel.GetCell(i,m_icolNo).bstrVal;
			strCell.MakeUpper();
			ssk.strNo = strCell;

			ssk.iQty = Excel.GetCellValue(i,m_icolProductCnt);

			strCell = Excel.GetCell(i,m_icolFactoryNo).bstrVal;
			strCell.MakeUpper();
			ssk.strFactoryNo = strCell;

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

void CPageCkd::OnBtnCheck() 
{
	// TODO: Add your control notification handler code here
	if(mapBaseList.GetCount() == 0)
	{
		MessageBox1("请先载入基础清单!","校验错误",MB_ICONERROR);
		return ;
	}
	else if(mapInBox.GetCount() == 0)
	{
		MessageBox1("未录入任何数据,请先录入!","校验错误",MB_ICONERROR);
		return ;
	}

	// 根据录入数据 生成 "部品号-总数量" 的映射表
	 
	m_ListMain.DeleteAllItems();
	m_ListMain.AddItem("料号","部品号","记录数量","实际数量","");

	CString strRecordPart;
	CString csTemp, csRecordQty, csBaseQty,strSpeak;
	int iPartToal = 0,iListOrder = 0;
	POSITION pos = mapPartCnt.GetStartPosition();
	int iErrorCnt = 0;
	CLayers cBase;
    while(pos)
    {
        mapPartCnt.GetNextAssoc(pos, strRecordPart, iPartToal);
		if(mapBaseList.Lookup(strRecordPart, cBase))
		{
			if(iPartToal != cBase.iQty)
			{
				
				csTemp.Format("%d",++iListOrder);
				csRecordQty.Format("%d",iPartToal);
				csBaseQty.Format("%d",cBase.iQty);
				m_ListMain.AddItem(cBase.strNo,strRecordPart,csRecordQty,csBaseQty,"");

				csTemp.Format("%d",cBase.iQty);
				csTemp = cBase.strNo + "号料不符实际" + csTemp + ",录入" + csRecordQty + ",";
				strSpeak += csTemp;	
				//if(++iErrorCnt >=5)
				//{
					//break;
				//}
			}
		}
    }

	if(iListOrder == 0)
		MessageBox1("没有不符项","校验成功",MB_OK);
	else if(g_flgSpeak)
		MessageBox1(strSpeak,"校验",MB_ICONWARNING);
}

void CPageCkd::OnSetfocusEditUndo() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	m_csUndoInbox.Empty();
	UpdateData(FALSE);	
}

void CPageCkd::OnTimer(UINT nIDEvent) 
{
	// TODO: Add your message handler code here and/or call default

	switch(nIDEvent)
	{
	case 1:
		g_iSpeakCheckCnt = 0;
		KillTimer(1);
		break;

	default :
		break;
	}
	
	CPropertyPage::OnTimer(nIDEvent);
}

void CPageCkd::ReloadData(BOOL bActive)
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

		CPropSheet* pParent = (CPropSheet*) GetParent();
		CPageShipMark* prop1 = (CPageShipMark*)pParent->GetPage(pParent->GetPageCount()-1);
		//序列化文件
		CFile file;
		file.Open(strFileName, CFile::modeReadWrite);
		CArchive ar(&file, CArchive::load);
		if(bActive)
		{
			CString csTemp;
			ar>>csTemp>>csTemp>>csTemp>>csTemp
			>>csTemp>>csTemp>>csTemp>>csTemp
			>>csTemp>>csTemp>>csTemp>>csTemp;
		}
		else
		{
			ar>>prop1->m_cs11>>prop1->m_cs12>>prop1->m_cs13>>prop1->m_cs14
			>>prop1->m_cs21>>prop1->m_cs22>>prop1->m_cs23>>prop1->m_cs24
			>>prop1->m_cs31>>prop1->m_cs32>>prop1->m_cs33>>prop1->m_cs34;
		}
		m_LayerList.Serialize(ar);
		ar.Close();
		file.Close();

		POSITION pos = m_LayerList.GetHeadPosition();
		while (pos != NULL)
		{
			CLayers* pLayer = m_LayerList.GetNext(pos);
			int count = 0;
			g_iOrder++;
			mapInBox.SetAt(pLayer->strInBox, *pLayer);

			if(mapOutBox.Lookup(pLayer->strOutBox,count))
				mapOutBox.SetAt(pLayer->strOutBox, count+1);
			else
				mapOutBox.SetAt(pLayer->strOutBox, 1);
			cs.Format("%d",pLayer->iQty);
			m_ListMain.AddItem(pLayer->strOrder,pLayer->strPartNo,pLayer->strOutBox,pLayer->strInBox,cs);
		}
		m_flgTaojian = TRUE;
		m_ListMain.Sort(0,true);
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
				mapBaseList.SetAt(pLayer->strPartNo, *pLayer);
				cs.Format("%d",pLayer->iQty);
				m_ListMain.AddItem(pLayer->strNo,pLayer->strPartNo,pLayer->strUserCode,pLayer->strDetail,cs);	
			}
			m_ListMain.Sort(0,true);
			m_ListMain.EnsureVisible(g_iOrder-1,TRUE);
			MessageBox1("套件载入成功,基础清单载入成功!", "载入成功", MB_ICONINFORMATION);
		}
		else
		{
			MessageBox1("套件载入成功,基础清单未载入!", "载入", MB_ICONINFORMATION);
		}
		//==================================================================================================

		pos = mapInBox.GetStartPosition();
		CString strInBox;
		CLayers cV, cBase;
		int iTemp = 0;
		while(pos)
		{
			// 完善记录表 mapInBox 缺失的 "客户代码"和"部品描述"
			mapInBox.GetNextAssoc(pos, strInBox, cV);
			if(mapBaseList.Lookup(cV.strPartNo, cBase))
			{
				cV.strUserCode = cBase.strUserCode;
				cV.strDetail = cBase.strDetail;
				cV.strNo = cBase.strNo;
				cV.strFactoryNo = cBase.strFactoryNo;
				mapInBox.SetAt(strInBox,cV);
			}

			// 生成 "部品号-总数量" 的映射表
			if(mapPartCnt.Lookup(cV.strPartNo, iTemp))
			{
				iTemp += cV.iQty;
				mapPartCnt.SetAt(cV.strPartNo,iTemp);
			}
			else
				mapPartCnt.SetAt(cV.strPartNo,cV.iQty);
		}

		//CEdit* pEdit = (CEdit*)GetDlgItem(IDC_EDIT_ZCS);
		//pEdit->EnableWindow(FALSE);
		OnBtnReload();
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

	CEdit* pEdit = (CEdit*)GetDlgItem(IDC_EDIT_RECEIVE);
	pEdit->SetFocus();
}


void CPageCkd::PlayVoice(void) 
{
	// TODO: Add extra validation here
	CString path,path1,path2,path3;
    CString cs;
	cs = CExcel::GetAppPath() + "\\Voice\\";
	UpdateData(TRUE);
	if(m_unRatio<10)
	{
		path.Format("%d.wav",m_unRatio);
		path = cs + path;
		PlaySound(path, NULL, SND_FILENAME | SND_ASYNC);
	}
	else if(10<m_unRatio&&m_unRatio<100)
	{
		path.Format("%d.wav",m_unRatio/10);
		path1.Format("%d.wav",m_unRatio%10);
		path = cs + path;
		PlaySound(path, NULL, SND_FILENAME | SND_SYNC);
		PlaySound(cs+"10.wav", NULL, SND_FILENAME | SND_SYNC);
		if((m_unRatio%10)!=0)
			PlaySound(cs+path1, NULL, SND_FILENAME | SND_SYNC);
	}
	else if(m_unRatio == 10)
	{
		PlaySound(cs+"10.wav", NULL, SND_FILENAME | SND_SYNC);
	}
	else if(m_unRatio == 100)
	{
		PlaySound(cs+"1.wav", NULL, SND_FILENAME | SND_SYNC);
		PlaySound(cs+"100.wav", NULL, SND_FILENAME | SND_SYNC);
	}
	else if(100<m_unRatio&&m_unRatio<1000)
	{
		int j=m_unRatio%100;
		path.Format("%d.wav",m_unRatio/100);
		path1.Format("%d.wav",j/10);
		path2.Format("%d.wav",j%10);
		PlaySound(cs+path, NULL, SND_FILENAME | SND_SYNC);
		PlaySound(cs+"100.wav", NULL, SND_FILENAME | SND_SYNC);
		if(((j/10)!=0)&&((j%10))!=0)
		{
			PlaySound(cs+path1, NULL, SND_FILENAME | SND_SYNC);
			PlaySound(cs+"10.wav", NULL, SND_FILENAME | SND_SYNC);
			PlaySound(cs+path2, NULL, SND_FILENAME | SND_SYNC);
		}
		else if(((j/10)!=0)&&((j%10)==0))
		{
			PlaySound(cs+path1, NULL, SND_FILENAME | SND_SYNC);
			PlaySound(cs+"10.wav", NULL, SND_FILENAME | SND_SYNC);
		}
		else if(((j%10)!=0)&&(j/10)==0)
		{
			PlaySound(cs+"0.wav", NULL, SND_FILENAME | SND_SYNC);
			PlaySound(cs+path2, NULL, SND_FILENAME | SND_SYNC);
		}
	}
	else 
		MessageBox("请输入千以下的数字");
}



void CPageCkd::OnButton1() 
{
	// TODO: Add your control notification handler code here
	CString strFileName = CExcel::GetAppPath() + "\\Record List\\"  + "111.ckd";
	CFile file;
	file.Open(strFileName, CFile::modeReadWrite | CFile::modeCreate);
	CArchive ar(&file, CArchive::store);
	g_ssk.Clear();
	g_ssk.strOrder = "ss";
	mapInBox.SetAt("123",g_ssk);
	mapInBox.Serialize(ar);
	ar.Close();
	file.Close();
}
//CMapStringToOb
void CPageCkd::OnButton2() 
{
	// TODO: Add your control notification handler code here
	mapInBox.RemoveAll();
	CString strFileName = CExcel::GetAppPath() + "\\Record List\\"  + "111.ckd";
	CFile file;
	file.Open(strFileName, CFile::modeReadWrite);
	CArchive ar(&file, CArchive::load);
	mapInBox.Serialize(ar);
	ar.Close();
	file.Close();
	mapInBox.Lookup("123",g_ssk);
	MessageBox(g_ssk.strOrder);
}

