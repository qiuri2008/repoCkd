// Prop3.cpp : implementation file
//

#include "stdafx.h"
#include "CKD¼��.h"
#include "PageList.h"
#include "PropSheet.h"

#include "Global.h"
#include <afxtempl.h>
#include "excel2003.h"
#include "user.h"
#include "SaveDlg.h"
#include <bitset>

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

using std::bitset;
bitset<4> g_bCheck;

/////////////////////////////////////////////////////////////////////////////
// CPageList property page

IMPLEMENT_DYNCREATE(CPageList, CPropertyPage)

CPageList::CPageList() : CPropertyPage(CPageList::IDD)
{
	//{{AFX_DATA_INIT(CPageList)
	m_csReceive = _T("");
	//}}AFX_DATA_INIT
	m_listInit = false;
}

CPageList::~CPageList()
{
}

void CPageList::DoDataExchange(CDataExchange* pDX)
{
	CPropertyPage::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CPageList)
	DDX_Control(pDX, IDC_LIST2, m_ListMain);
	DDX_Text(pDX, IDC_EDIT_ADMIN, m_csReceive);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CPageList, CPropertyPage)
	//{{AFX_MSG_MAP(CPageList)
	ON_BN_CLICKED(IDC_BUTTON1, OnBtnImportBaseList)
	ON_BN_CLICKED(IDC_BUTTON2, OnBtnSaveBaseList)
	ON_BN_CLICKED(IDC_BTN_SAVE_BOM, OnBtnSaveBom)
	ON_BN_CLICKED(IDC_BTN_LOAD_BOM, OnBtnLoad)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CPageList message handlers

BOOL CPageList::OnSetActive() 
{
	// TODO: Add your specialized code here and/or call the base class
	

	DWORD dwStyle;
	if(!m_listInit)
	{
		dwStyle=m_ListMain.GetExtendedStyle();
		dwStyle|=LVS_EX_FULLROWSELECT;
		dwStyle|=LVS_EX_GRIDLINES;

		(void)m_ListMain.SetExtendedStyle( dwStyle );///����ѡ��ģʽ//LVS_EX_FULLROWSELECT
		m_ListMain.SetHeadings("���,60;��Ʒ��,110;����,170"); ///������ͷ��Ϣ 
		m_ListMain.LoadColumnInfo(); 
		m_listInit = true;


		CString strFileName = CExcel::GetAppPath() + "\\Record List\\" + "_BomList" + ".ckd";
		if(IsFileExist(strFileName,FALSE)==TRUE)
		{
			mapBom.RemoveAll();
			m_LayerList.RemoveAll();
			m_ListMain.DeleteAllItems();
			g_iOrder = 0;
			//���л��ļ�
			CFile file;
			file.Open(strFileName, CFile::modeReadWrite);
			CArchive ar(&file, CArchive::load);
			m_LayerList.Serialize(ar);
			ar.Close();
			file.Close();

			POSITION pos = m_LayerList.GetHeadPosition();
			while (pos != NULL)
			{
				g_iOrder++;
				CLayers* pLayer = m_LayerList.GetNext(pos);
				mapBom.SetAt(pLayer->strPartNo, *pLayer);
				m_ListMain.AddItem(pLayer->strNo,pLayer->strPartNo,pLayer->strDetail);	
			}
			m_ListMain.Sort(0,true);
			if(g_iOrder)
				m_ListMain.EnsureVisible(g_iOrder-1,TRUE);
			g_bCheck.set(STYLE_BOM_LOAD);
			//MessageBox("����嵥�Ѿ�����!", "����嵥");
		}
		else
		{
			MessageBox("�׼�����ɹ�,����嵥δ����!", "����", MB_ICONWARNING);
			g_bCheck.reset();
		}

	((CButton*)GetDlgItem(IDC_BTN_LOAD_BOM))->ShowWindow(SW_HIDE);
	((CButton*)GetDlgItem(IDC_BTN_SAVE_BOM))->ShowWindow(SW_HIDE);

	}
	
	return CPropertyPage::OnSetActive();
}

void CPageList::CalcWindowRect(LPRECT lpClientRect, UINT nAdjustType) 
{
	// TODO: Add your specialized code here and/or call the base class
	
	CPropertyPage::CalcWindowRect(lpClientRect, nAdjustType);
}


void CPageList::OnBtnImportBaseList() 
{
	if(!g_bCheck[STYLE_BOM_LOAD])
	{
		MessageBox("δ�ҵ�����嵥,������������嵥!", "����", MB_ICONERROR);
		return ;
	}
	CFileDialog dlg(true,"*.xls","",OFN_HIDEREADONLY,"Excel�ļ�(*.xls)|*.xls");
	//===========================================================
	//���ض�Ŀ¼
	CString strFileName = CExcel::GetAppPath() + "\\Base List";
	dlg.m_ofn.lpstrInitialDir = strFileName;
	//===========================================================
	if(dlg.DoModal()==IDOK)
		g_strOpenFile=dlg.GetPathName();

	if(g_strOpenFile.Find(".xls")<0)        //û�ж�ȡ��ֱ�ӷ���
		return;

	unsigned int i,j;
	CString strCell;
	CExcel Excel;
	Excel.AddNewFile(g_strOpenFile);				// ��һ���ļ�
	Excel.SetVisible(false);						// ���ÿɼ�
	Excel.SelectSheet(1);				        // �������1
	m_usedRow = Excel.GetUsedRowCount();
	m_usedCol = Excel.GetUsedColCount();
	Excel.SelectSheet(1);				        // ��������RANGE
	m_icolProductCnt = m_icolPart = m_irowPart = m_icolDetail = m_icolUserCode = m_icolNo = 0;
	m_flgCheck = false;
	
	//////////////////////////////////////
	//��ȡ��Ʒ�� λ�� ���� ��װ��λ��
	for(i=1;i<=m_usedRow;i++)
	{
		for(j=1;j<=m_usedCol;j++)
		{
			strCell = Excel.GetCell(i,j).bstrVal;
			strCell = Excel.DeleteBlackSpace(strCell);  //ɾ���ո�
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

			//if(m_icolProductCnt && m_icolPart && m_irowPart && m_icolDetail && m_icolUserCode)
			if(m_icolPart && m_irowPart && m_icolDetail)
			{
				m_flgCheck = true;
				break;
			}
		}

		if(m_icolPart && m_irowPart && m_icolDetail)
		{
				m_flgCheck = true;
				break;
		}
	}

	/////////////////////////////////////////////////////
	//����Ƿ�Ϊ��׼K3�嵥
	if(!m_flgCheck)  
	{
		Excel.Save(true);
		MessageBox("�뵼����ȷ�Ļ����嵥!!!", "����", MB_ICONERROR);
		return ;
	}

	CLayers ssk;
	mapBaseList.RemoveAll();
	/////////////////////////////////////////////////////	
	for(i=m_irowPart+1;i<=m_usedRow;i++)
	{
		strCell = Excel.GetCell(i,m_icolPart).bstrVal; 
		strCell = Excel.DeleteBlackSpace(strCell);
		if(strCell.GetLength() >=13)
		{
			///////////////////////////////////////////////////
			//д���ϣ��
			strCell.MakeUpper();
			ssk.strPartNo = strCell;

			if(m_icolUserCode)
			{
				strCell = Excel.GetCell(i,m_icolUserCode).bstrVal;
				strCell.MakeUpper();
				ssk.strUserCode = strCell;
			}

			if(m_icolDetail)
			{
			strCell = Excel.GetCell(i,m_icolDetail).bstrVal;
			strCell.MakeUpper();
			ssk.strDetail = strCell;
			}

			if(m_icolNo)
			{
				strCell = Excel.GetCell(i,m_icolNo).bstrVal;
				strCell.MakeUpper();
				ssk.strNo = strCell;
			}

			if(m_icolProductCnt)
			{
				ssk.iQty = Excel.GetCellValue(i,m_icolProductCnt);
			}

			mapBaseList.SetAt(ssk.strPartNo,ssk);
		}
	}
	
	m_ListMain.DeleteAllItems();
	m_LayerList.RemoveAll();
	POSITION pos = mapBaseList.GetStartPosition();
	CString strKey;
	g_iOrder = 0;
    while(pos)
    {
		g_iOrder++;
        mapBaseList.GetNextAssoc(pos, strKey, ssk);
		m_ListMain.AddItem(ssk.strNo,strKey,ssk.strDetail);	

		//���ɱ����嵥����
		CLayers* pLayersItem = new CLayers(ssk);
		m_LayerList.AddTail(pLayersItem);
    }
	m_ListMain.Sort(0,true);
	m_ListMain.EnsureVisible(g_iOrder-1,TRUE);
	g_bCheck.set(STYLE_BASE_LOAD);

	CString strFileName1 = CExcel::GetAppPath() + "\\Record List\\" + "���BOM�嵥.xls";
	if(IsFileExist(strFileName1,FALSE)==TRUE)
		DeleteFile(strFileName1);
    Excel.InsertCol(1);
	Excel.SelectSheet(1);				        // �������1
	Excel.SetCell(m_irowPart, 1, "No.");
	Excel.SaveAs(strFileName1);
	MessageBox("�����嵥����ɹ�!!!","����", MB_ICONINFORMATION);

}

void CPageList::OnBtnSaveBaseList() 
{
	if(!g_bCheck[STYLE_BOM_LOAD])
	{
		MessageBox("������������嵥!","����",MB_ICONERROR);
		return;
	}
	else if(!g_bCheck[STYLE_BASE_LOAD])
	{
		MessageBox("������������嵥!","����",MB_ICONERROR);
		return;
	}

	CString strFileName = CExcel::GetAppPath() + "\\Record List\\" + "���BOM�嵥.xls";

	unsigned int i,j;
	CString strCell;
	CExcel Excel;
	Excel.AddNewFile(strFileName);				// ��һ���ļ�
	Excel.SetVisible(false);						// ���ÿɼ�
	Excel.SelectSheet(1);				        // �������1
	m_usedRow = Excel.GetUsedRowCount();
	m_usedCol = Excel.GetUsedColCount();
	Excel.SelectSheet(1);				        // ��������RANGE
	m_icolProductCnt = m_icolPart = m_irowPart = m_icolDetail = m_icolUserCode = m_icolNo = 0;
	m_flgCheck = false;
	
	//////////////////////////////////////
	//��ȡ��Ʒ�� λ�� ���� ��װ��λ��
	for(i=1;i<=m_usedRow;i++)
	{
		for(j=1;j<=m_usedCol;j++)
		{
			strCell = Excel.GetCell(i,j).bstrVal;
			strCell = Excel.DeleteBlackSpace(strCell);  //ɾ���ո�
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

			if(m_icolPart && m_irowPart && m_icolDetail)
			{
				m_flgCheck = true;
				break;
			}
		}

		if(m_icolPart && m_irowPart && m_icolDetail)
		{
				m_flgCheck = true;
				break;
		}
	}

	/////////////////////////////////////////////////////
	//����Ƿ�Ϊ��׼K3�嵥
	if(!m_flgCheck)  
	{
		Excel.Save(true);
		MessageBox("�뵼����ȷ�Ļ����嵥!!!", "����", MB_ICONERROR);
		return ;
	}

	CLayers ssk,sskBak,cBase;
	mapBaseList.RemoveAll();
	mapBomAdd.RemoveAll();
	m_ListMain.DeleteAllItems();
	/////////////////////////////////////////////////////	
	for(i=m_irowPart+1;i<=m_usedRow;i++)
	{
		strCell = Excel.GetCell(i,m_icolPart).bstrVal; 
		strCell = Excel.DeleteBlackSpace(strCell);
		if(strCell.GetLength() >= 13)
		{
			///////////////////////////////////////////////////
			//д���ϣ��
			strCell.MakeUpper();
			ssk.strPartNo = strCell;
			Excel.SetCell(i, m_icolPart, ssk.strPartNo);

			if(m_icolUserCode)
			{
				strCell = Excel.GetCell(i,m_icolUserCode).bstrVal;
				strCell.MakeUpper();
				ssk.strUserCode = strCell;
			}

			if(m_icolDetail)
			{
				strCell = Excel.GetCell(i,m_icolDetail).bstrVal;
				strCell.MakeUpper();
				ssk.strDetail = strCell;
			}

			if(mapBom.Lookup(ssk.strPartNo,cBase))
			{
				ssk.strNo = cBase.strNo;
				Excel.SetCell(i, 1, ssk.strNo);
			}
			else
			{
				CString cs;
				static noEixtCnt = 1;
				cs.Format("%d",mapBom.GetCount() + noEixtCnt);
				sskBak = ssk;
				sskBak.strNo = cs;
				mapBomAdd.SetAt(ssk.strPartNo,sskBak);
				noEixtCnt++;
				
				ssk.strNo = "";
				g_iOrder++;
				m_ListMain.AddItem(ssk.strNo,ssk.strPartNo,ssk.strDetail);
				m_flgNoExist = true;
			}

			if(m_icolProductCnt)
				ssk.iQty = Excel.GetCellValue(i,m_icolProductCnt);
			mapBaseList.SetAt(ssk.strPartNo,ssk);
		}
	}

	if(m_flgNoExist)
	{
		m_ListMain.Sort(0,true);
		m_ListMain.EnsureVisible(g_iOrder-1,TRUE);
		CString cs;
		cs.Format("%d�Ʒ��-��������嵥��!!!",mapBomAdd.GetCount());
		MessageBox(cs,"����",MB_ICONERROR);
		m_flgNoExist = false;
	}
	else
	{
		m_ListMain.DeleteAllItems();
		m_LayerList.RemoveAll();
		POSITION pos = mapBaseList.GetStartPosition();
		CString strKey;
		g_iOrder = 0;
		while(pos)
		{
			g_iOrder++;
			mapBaseList.GetNextAssoc(pos, strKey, ssk);
			m_ListMain.AddItem(ssk.strNo,strKey,ssk.strDetail);	

			//���ɱ����嵥����
			CLayers* pLayersItem = new CLayers(ssk);
			m_LayerList.AddTail(pLayersItem);
		}
		m_ListMain.Sort(0,true);
		m_ListMain.EnsureVisible(g_iOrder-1,TRUE);
		MessageBox("����ɹ�!!!", "����ɹ�", MB_ICONINFORMATION);
	}
	Excel.SaveAs(strFileName);

}

template <typename T, typename U> void BubbleSort(T& collection, U element, int count, bool ascend = true)
{
	for (int i = 0; i < count-1; i++)
	for (int j = 0; j < count-1-i; j++)
		if (ascend)
		{
		// ����
			if (collection[j] > collection[j+1])
			{
				U temp = collection[j];
				collection[j] = collection[j+1];
				collection[j+1] = temp;
			}
		}
		else
		{
		// ����
			if (collection[j] < collection[j+1])
			{
			U temp = collection[j];
			collection[j] = collection[j+1];
			collection[j+1] = temp;
			}
		}
}

void CPageList::OnBtnLoad() 
{
	CFileDialog dlg(true,"*.xls","",OFN_HIDEREADONLY,"Excel�ļ�(*.xls)|*.xls");
	//===========================================================
	//���ض�Ŀ¼
	CString strFileName = CExcel::GetAppPath();
	dlg.m_ofn.lpstrInitialDir = strFileName;
	//===========================================================
	if(dlg.DoModal()==IDOK)
		g_strOpenFile=dlg.GetPathName();

	if(g_strOpenFile.Find(".xls")<0)        //û�ж�ȡ��ֱ�ӷ���
		return;

	unsigned int i,j;
	CString strCell;
	CExcel Excel;
	Excel.AddNewFile(g_strOpenFile);				// ��һ���ļ�
	Excel.SetVisible(false);						// ���ÿɼ�
	Excel.SelectSheet(1);				        // �������1
	m_usedRow = Excel.GetUsedRowCount();
	m_usedCol = Excel.GetUsedColCount();
	Excel.SelectSheet(1);				        // ��������RANGE
	m_icolProductCnt = m_icolPart = m_irowPart = m_icolDetail = m_icolUserCode = 0;
	m_flgCheck = false;
	
	//////////////////////////////////////
	//��ȡ��Ʒ�� λ�� ���� ��װ��λ��
	for(i=1;i<=m_usedRow;i++)
	{
		for(j=1;j<=m_usedCol;j++)
		{
			strCell = Excel.GetCell(i,j).bstrVal;
			strCell = Excel.DeleteBlackSpace(strCell);  //ɾ���ո�
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
			else if(strcmp(pCell,"���") == 0)
			{
				m_icolNo = j;
			}

			if(m_icolPart && m_irowPart && m_icolDetail && m_icolNo)
			{
				m_flgCheck = true;
				break;
			}
		}

		if(m_icolPart && m_irowPart && m_icolDetail && m_icolNo)
		{
				m_flgCheck = true;
				break;
		}
	}

	/////////////////////////////////////////////////////
	//���
	if(!m_flgCheck)  
	{
		Excel.Save(true);
		MessageBox("�뵼����ȷ������嵥!!!", "����", MB_ICONERROR);
		return ;
	}

	CLayers ssk;
	mapBom.RemoveAll();
		
	/////////////////////////////////////////////////////	
	for(i=m_irowPart+1;i<=m_usedRow;i++)
	{
		strCell = Excel.GetCell(i,m_icolPart).bstrVal; 
		strCell = Excel.DeleteBlackSpace(strCell);
		if(strCell.GetLength() >= 13)
		{
			///////////////////////////////////////////////////
			//д���ϣ��
			strCell.MakeUpper();
			ssk.strPartNo = strCell;

			strCell = Excel.GetCell(i,m_icolDetail).bstrVal;
			strCell.MakeUpper();
			ssk.strDetail = strCell;

			strCell = Excel.GetCell(i,m_icolNo).bstrVal;
			strCell.MakeUpper();
			ssk.strNo = strCell;

			mapBom.SetAt(ssk.strPartNo,ssk);
		}
	}
	
	m_ListMain.DeleteAllItems();
	m_LayerList.RemoveAll();
	POSITION pos = mapBom.GetStartPosition();
	CString strKey;
	g_iOrder = 0;
    while(pos)
    {
		g_iOrder++;
        mapBom.GetNextAssoc(pos, strKey, ssk);
		m_ListMain.AddItem(ssk.strNo,strKey,ssk.strDetail);	

		//���ɱ����嵥����
		CLayers* pLayersItem = new CLayers(ssk);
		m_LayerList.AddTail(pLayersItem);
    }
	m_ListMain.Sort(0,true);
	m_ListMain.EnsureVisible(g_iOrder-1,TRUE);

	//�����嵥
	CFile file;
	strFileName = CExcel::GetAppPath() + "\\Record List\\" + "_BomList" + ".ckd";
	file.Open(strFileName, CFile::modeCreate|CFile::modeReadWrite);
	CArchive ar(&file, CArchive::store);
	m_LayerList.Serialize(ar);
	ar.Close();
	file.Close();

	g_bCheck.set(STYLE_BOM_LOAD);

	MessageBox("����嵥����ɹ�!", "����ɹ�", MB_ICONINFORMATION);
	
}


void CPageList::OnBtnSaveBom() 
{
	CFileDialog dlg(FALSE, "xls", "", OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, "Excel�ļ�(*.xls)|*.xls" );  
	//===========================================================
	//���浽ָ��Ŀ¼
	CString strFileName = CExcel::GetAppPath();
	dlg.m_ofn.lpstrInitialDir = strFileName;
	//===========================================================
	if(dlg.DoModal()==IDOK)
		g_strSaveFile=dlg.GetPathName();

	if(g_strSaveFile.Find(".xls")<0)        //û�б�����ֱ�ӷ���
		return;

	CExcel Excel;
	Excel.AddNewFile();				            // �½�һ���ļ�
	Excel.SetVisible(FALSE);					// ���ò��ɼ�
	Excel.SelectSheet(1);
	Excel.ActiveSheet().SetName("����嵥");

	Excel.SetCell(1,1,"���");
	Excel.SetCell(1,2,"��Ʒ��");
	Excel.SetCell(1,3,"����");

	CString strKey;
	CLayers cValue;
	int i = 2;
	CArrayClayer NoArray;

	POSITION pos = mapBomAdd.GetStartPosition();
	while(pos)
	{
		mapBomAdd.GetNextAssoc(pos, strKey, cValue);
		mapBom.SetAt(strKey,cValue);
	}
	
	//================================================================
	//���������
	pos = mapBom.GetStartPosition();
	m_LayerList.RemoveAll();	
	while(pos)
	{
		mapBom.GetNextAssoc(pos, strKey, cValue);
		NoArray.Add(cValue);

		//���ɱ����嵥����
		CLayers* pLayersItem = new CLayers(cValue);
		m_LayerList.AddTail(pLayersItem);
	}
	g_OrderMode = ORDER_NO;
	BubbleSort(NoArray, NoArray[0], NoArray.GetSize(), true);
	//================================================================
	 for(int j=0; j<NoArray.GetSize(); j++)
     {
		cValue = NoArray[j];
		Excel.SetCell(i,1, cValue.strNo);
		Excel.SetCell(i,2, cValue.strPartNo);
		Excel.SetCell(i,3,cValue.strDetail);

		i++;
     }

	Excel.SetColAutoFit(1);	
	Excel.SetColAutoFit(2);	
	Excel.SetColAutoFit(3);	
	Excel.SaveAs(g_strSaveFile);
	

	//if(m_flgNoExist)
	{
		//�����嵥
		CFile file;
		CString strFileName = CExcel::GetAppPath() + "\\Record List\\" + "_BomList" + ".ckd";
		file.Open(strFileName, CFile::modeCreate|CFile::modeReadWrite);
		CArchive ar(&file, CArchive::store);
		m_LayerList.Serialize(ar);
		ar.Close();
		file.Close();
		m_flgNoExist = false;
	}
	MessageBox("����ɹ�!","����",MB_ICONINFORMATION);
}


BOOL CPageList::PreTranslateMessage(MSG* pMsg) 
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
				case IDC_EDIT_ADMIN: 
					UpdateData(TRUE);
					if(m_csReceive == "1147")
					{
						((CButton*)GetDlgItem(IDC_BTN_LOAD_BOM))->ShowWindow(SW_SHOW);
						((CButton*)GetDlgItem(IDC_BTN_SAVE_BOM))->ShowWindow(SW_SHOW);
						CString strFileName1 = CExcel::GetAppPath() + "\\Record List\\" + "���BOM�嵥.xls";
						if(IsFileExist(strFileName1,FALSE)==TRUE)
							DeleteFile(strFileName1);
					}
					else if(m_csReceive == "   ")
					{
						((CButton*)GetDlgItem(IDC_BTN_LOAD_BOM))->ShowWindow(SW_HIDE);
						((CButton*)GetDlgItem(IDC_BTN_SAVE_BOM))->ShowWindow(SW_HIDE);
					}
					else
					{
						((CButton*)GetDlgItem(IDC_BTN_LOAD_BOM))->ShowWindow(SW_HIDE);
						((CButton*)GetDlgItem(IDC_BTN_SAVE_BOM))->ShowWindow(SW_HIDE);
						MessageBox("�������,����������!!!", "��½", MB_ICONERROR);
					}
					m_csReceive.Empty();
					UpdateData(FALSE);
					return true;
					break;
			}
		}
	}
	
	return CPropertyPage::PreTranslateMessage(pMsg);
	
}
