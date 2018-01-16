// Prop1.cpp : implementation file
//

#include "stdafx.h"
#include "CKD录入.h"
#include "PageShipMark.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif
#include "excel2003.h"
#include "Global.h"
#include <afxtempl.h>
#include "user.h"
#include "PropSheet.h"
/////////////////////////////////////////////////////////////////////////////
// CPageShipMark property page

CMapInBox mapInBox1;  
CMapCS mapOutBox1;    //  外箱号 - 内箱数量

IMPLEMENT_DYNCREATE(CPageShipMark, CPropertyPage)

CPageShipMark::CPageShipMark() : CPropertyPage(CPageShipMark::IDD)
{
	//{{AFX_DATA_INIT(CPageShipMark)
	m_cs11 = _T("VIL PO NO");
	m_cs12 = _T("LCD/3240005210");
	m_cs13 = _T("Carton Box No");
	m_cs14 = _T("C023");
	m_cs21 = _T("Supplier Name");
	m_cs22 = _T("Kitking");
	m_cs23 = _T("Date");
	m_cs24 = _T("2013-08-05");
	m_cs31 = _T("Country");
	m_cs32 = _T("India");
	m_cs33 = _T("Made in");
	m_cs34 = _T("China");
	m_nRadio1 = 0;
	m_bSpeak = TRUE;
	//}}AFX_DATA_INIT
}

CPageShipMark::~CPageShipMark()
{
}

void CPageShipMark::DoDataExchange(CDataExchange* pDX)
{
	CPropertyPage::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CPageShipMark)
	DDX_Text(pDX, IDC_EDIT_11, m_cs11);
	DDX_Text(pDX, IDC_EDIT_12, m_cs12);
	DDX_Text(pDX, IDC_EDIT_13, m_cs13);
	DDX_Text(pDX, IDC_EDIT_14, m_cs14);
	DDX_Text(pDX, IDC_EDIT_21, m_cs21);
	DDX_Text(pDX, IDC_EDIT_22, m_cs22);
	DDX_Text(pDX, IDC_EDIT_23, m_cs23);
	DDX_Text(pDX, IDC_EDIT_24, m_cs24);
	DDX_Text(pDX, IDC_EDIT_31, m_cs31);
	DDX_Text(pDX, IDC_EDIT_32, m_cs32);
	DDX_Text(pDX, IDC_EDIT_33, m_cs33);
	DDX_Text(pDX, IDC_EDIT1_34, m_cs34);
	DDX_Radio(pDX, IDC_RADIO1, m_nRadio1);
	DDX_Check(pDX, IDC_CHECK1, m_bSpeak);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CPageShipMark, CPropertyPage)
	//{{AFX_MSG_MAP(CPageShipMark)
	ON_BN_CLICKED(IDC_BUTTON1, OnBtnLoadFile)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CPageShipMark message handlers

void CPageShipMark::OnBtnLoadFile() 
{
	// TODO: Add your control notification handler code here
	CFileDialog dlg(true,"*.xls","",OFN_HIDEREADONLY,"Excel文件(*.xls)|*.xls");
	if(dlg.DoModal()==IDOK)
		g_strOpenFile=dlg.GetPathName();

	if(g_strOpenFile.Find(".xls")<0)        //没有读取则直接返回
		return;

	unsigned int i,j;
	CString strCell,csTemp;
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
			else if(strcmp(pCell,TOULIAO_LIST_INBOX) == 0)
			{
				m_icolInBox = j;
			}
			else if(strcmp(pCell,TOULIAO_LIST_OUTBOX) == 0)
			{
				m_icolOutBox = j;
			}

			if(m_icolProductCnt && m_icolPart && m_irowPart && m_icolDetail && m_icolUserCode && m_icolInBox && m_icolOutBox)
			{
				m_flgCheck = true;
				break;
			}
		}

		if(m_icolProductCnt && m_icolPart && m_irowPart && m_icolDetail && m_icolUserCode && m_icolInBox && m_icolOutBox)
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

	for(i=m_irowPart+1;i<=m_usedRow;i++)
	{
		// 填写部品号
		strCell = Excel.GetCell(i,m_icolPart).bstrVal;
		//////////////////////////////////
		if(strCell.GetLength() >= 5)
		{
			for(unsigned int k=i+1;k<=m_usedRow;k++)
			{
				csTemp = Excel.GetCell(k,m_icolPart).bstrVal;
				if(csTemp.GetLength() >= 5)
					break;
			}

			for(j=i+1;j<k;j++)
			{
				Excel.SetCell(j,m_icolPart,strCell);
			}
		}
		//////////////////

		// 填写客户号
		strCell = Excel.GetCell(i,m_icolUserCode).bstrVal;
		//////////////////////////////////
		if(strCell.GetLength() >= 5)
		{
			for(unsigned int k=i+1;k<=m_usedRow;k++)
			{
				csTemp = Excel.GetCell(k,m_icolUserCode).bstrVal;
				if(csTemp.GetLength() >= 5)
					break;
			}

			for(j=i+1;j<k;j++)
			{
				Excel.SetCell(j,m_icolUserCode,strCell);
			}
		}
		//////////////////

		// 填写描述
		strCell = Excel.GetCell(i,m_icolDetail).bstrVal;
		//////////////////////////////////
		if(strCell.GetLength() >= 5)
		{
			for(unsigned int k=i+1;k<=m_usedRow;k++)
			{
				csTemp = Excel.GetCell(k,m_icolDetail).bstrVal;
				if(csTemp.GetLength() >= 5)
					break;
			}

			for(j=i+1;j<k;j++)
			{
				Excel.SetCell(j,m_icolDetail,strCell);
			}
		}
		//////////////////

		// 填写外箱
		strCell = Excel.GetCell(i,m_icolOutBox).bstrVal;
		//////////////////////////////////
		if(strCell.GetLength() >= 2)
		{
			for(unsigned int k=i+1;k<=m_usedRow;k++)
			{
				csTemp = Excel.GetCell(k,m_icolOutBox).bstrVal;
				if(csTemp.GetLength() >= 2)
					break;
			}

			for(j=i+1;j<k;j++)
			{
				Excel.SetCell(j,m_icolOutBox,strCell);
			}
		}
		//////////////////
			
	}

	mapInBox1.RemoveAll();
	mapOutBox1.RemoveAll();

	for(i=m_irowPart+1;i<=m_usedRow;i++)
	{
		CLayers cl;
		cl.strInBox = Excel.GetCell(i,m_icolInBox).bstrVal;
		cl.strOutBox = Excel.GetCell(i,m_icolOutBox).bstrVal;
		cl.strUserCode = Excel.GetCell(i,m_icolUserCode).bstrVal;
		cl.strDetail = Excel.GetCell(i,m_icolDetail).bstrVal;
		cl.strPartNo = Excel.GetCell(i,m_icolPart).bstrVal;
		cl.iQty = Excel.GetCellValue(i,m_icolProductCnt);
		int count = 0;
		mapInBox1.SetAt(cl.strInBox, cl);

		if(mapOutBox1.Lookup(cl.strOutBox,count))
			mapOutBox1.SetAt(cl.strOutBox, count+1);
		else
			mapOutBox1.SetAt(cl.strOutBox, 1);
	}

	Excel.Save(false);

	CExcel Excel1;
	Excel1.AddNewFile();				            // 新建一个文件
	Excel1.SetVisible(FALSE);					// 设置不可见

	SaveShippingMark1(Excel1);
    g_strSaveFile = CExcel::GetAppPath() + "\\唛头.xls";
	Excel1.SaveAs(g_strSaveFile);
	MessageBox("保存成功");
	
}

void BubbleSort1(CStringArray& collection, CString element, int count, bool ascend = true)
{
	for (int i = 0; i < count-1; i++)
	for (int j = 0; j < count-1-i; j++)
		if (ascend)
		{
		// 升序
			if (collection[j] > collection[j+1])
			{
				CString temp = collection[j];
				collection[j] = collection[j+1];
				collection[j+1] = temp;
			}
		}
		else
		{
		// 降序
			if (collection[j] < collection[j+1])
			{
			CString temp = collection[j];
			collection[j] = collection[j+1];
			collection[j+1] = temp;
			}
		}
}

void CPageShipMark::SaveShippingMark1(CExcel &Excel) 
{
	CString strCell;
	Excel.SelectSheet(2);				        // 激活工作簿1
	Excel.ActiveSheet().SetName("唛头");
	UpdateData(TRUE);
	CPropSheet* pParent = (CPropSheet*) GetParent();//首先获得属性页父窗口指针
	CPageShipMark* prop1 = (CPageShipMark*)pParent->GetPage(pParent->GetPageCount()-1);
	//MessageBox(prop1->m_cs11);//调用该窗口类中的函数

	//////////////////////////////////////////////////////////////////
	//设置表头
	//Excel.InsertRow(1);
	for(int j=1;j<=SHIPPING_WIDTH;j++)
		Excel.SetColWidth(j,1);

	CString strKey;
	int count = 0;
	int i =2;
	CStringArray arOutBox;
	POSITION pos = mapOutBox1.GetStartPosition();
	while(pos)
    {
		mapOutBox1.GetNextAssoc(pos, strKey, count);
		arOutBox.Add(strKey);
	}

	BubbleSort1(arOutBox, arOutBox[0], arOutBox.GetSize(), true);

    for(int k=0; k<arOutBox.GetSize(); k++)
    {
		mapOutBox1.Lookup(arOutBox[k],count);

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
		range = Excel.GetRange(i+2,SHIPPING_WOFFSET+TITLE1_COL_OFFSET+2,i+2,SHIPPING_WOFFSET+TITLE1_COL_OFFSET+2+TITLE_WCELL);
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
		range = Excel.GetRange(i+4,SHIPPING_WOFFSET+TITLE1_COL_OFFSET,i+4,SHIPPING_WOFFSET+TITLE1_COL_OFFSET+TITLE_WCELL);
		range.Merge();
		range = prop1->m_cs11;
		
		range = Excel.GetRange(i+4,SHIPPING_WOFFSET+TITLE2_COL_OFFSET,i+4,SHIPPING_WOFFSET+TITLE2_COL_OFFSET+TITLE_WCELL);
		range.Border();	
		range.Merge();	
		range = prop1->m_cs12;

		//  外箱号
		range = Excel.GetRange(i+4,SHIPPING_WOFFSET+TITLE3_COL_OFFSET,i+4,SHIPPING_WOFFSET+TITLE3_COL_OFFSET+TITLE_WCELL);
		range.Merge();	
		range = prop1->m_cs13;

		///////////////////////////////////
		//  Supplier Name
		range = Excel.GetRange(i+6,SHIPPING_WOFFSET+TITLE1_COL_OFFSET,i+6,SHIPPING_WOFFSET+TITLE1_COL_OFFSET+TITLE_WCELL);
		range.Merge();
		range = prop1->m_cs21;
		//  公司
		range = Excel.GetRange(i+6,SHIPPING_WOFFSET+TITLE2_COL_OFFSET,i+6,SHIPPING_WOFFSET+TITLE2_COL_OFFSET+TITLE_WCELL);
		range.Border();	
		range.Merge();	
		range = prop1->m_cs22;

		///////////////////////////////////
		//  Date
		range = Excel.GetRange(i+6,SHIPPING_WOFFSET+TITLE3_COL_OFFSET,i+6,SHIPPING_WOFFSET+TITLE3_COL_OFFSET+TITLE_WCELL);
		range.Merge();	
		range.SetHAlign(HAlignLeft);
		range = prop1->m_cs23;
		//  日期
		range = Excel.GetRange(i+6,SHIPPING_WOFFSET+TITLE4_COL_OFFSET,i+6,SHIPPING_WOFFSET+TITLE4_COL_OFFSET+TITLE_WCELL);
		range.Border();	
		range.Merge();	
		range.SetHAlign(HAlignLeft);
		range = prop1->m_cs24;

		///////////////////////////////////
		//  出口国
		range = Excel.GetRange(i+8,SHIPPING_WOFFSET+TITLE1_COL_OFFSET,i+8,SHIPPING_WOFFSET+TITLE1_COL_OFFSET+TITLE_WCELL);
		range.Merge();
		range = prop1->m_cs31;

		range = Excel.GetRange(i+8,SHIPPING_WOFFSET+TITLE2_COL_OFFSET,i+8,SHIPPING_WOFFSET+TITLE2_COL_OFFSET+TITLE_WCELL);
		range.Border();	
		range.Merge();	
		range = prop1->m_cs32;

		//  生产国
		range = Excel.GetRange(i+8,SHIPPING_WOFFSET+TITLE3_COL_OFFSET,i+8,SHIPPING_WOFFSET+TITLE3_COL_OFFSET+TITLE_WCELL);
		range.Merge();
		range = prop1->m_cs33;

		range = Excel.GetRange(i+8,SHIPPING_WOFFSET+TITLE4_COL_OFFSET,i+8,SHIPPING_WOFFSET+TITLE4_COL_OFFSET+TITLE_WCELL);
		range.Border();	
		range.Merge();	
		range = prop1->m_cs34;

		//  PACKING DETAILS
		range = Excel.GetRange(i+10,SHIPPING_WOFFSET,i+10,SHIPPING_WIDTH);
		range.Border();	
		range.Merge();	
		range = "PACKING DETAILS";
		range.SetHAlign(HAlignCenter);
		
		Excel.SelectSheet(2);
		POSITION pos = mapInBox1.GetStartPosition();
		CString strinbox;
		CLayers cValue;
		int x = 1;
		while(pos)
		{
			Excel.SelectSheet(2);
			mapInBox1.GetNextAssoc(pos, strinbox, cValue);
			if(arOutBox[k] == cValue.strOutBox)
			{
				Excel.InsertRow(i+SHIPPING_HIGH+x);
				//内箱号
				range = Excel.GetRange(i+SHIPPING_HIGH+x,CONTENT1_COL_OFFSET,i+SHIPPING_HIGH+x,CONTENT2_COL_OFFSET-1);
				range.Border(1,1);	
				range.Merge();	
				range = strinbox;
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
				range = Excel.GetRange(i+4,SHIPPING_WOFFSET+TITLE4_COL_OFFSET,i+4,SHIPPING_WOFFSET+TITLE4_COL_OFFSET+TITLE_WCELL);
				range.Border();	
				range.Merge();	
				range = cValue.strOutBox;

				x++;
			}
		}
		range = Excel.GetRange(i+SHIPPING_HIGH+1,SHIPPING_WOFFSET,i+SHIPPING_HIGH+x,SHIPPING_WIDTH);
		range.SetHAlign(HAlignLeft);

		range = Excel.GetRange(i+SHIPPING_HIGH+1,SHIPPING_WOFFSET,i+SHIPPING_HIGH+1+count,SHIPPING_WIDTH);
		range.Border();
		
		i += (count + SHIPPING_HIGH + SHIPPING_HOFFSET);
	}
}

