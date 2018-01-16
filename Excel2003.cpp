
#include "stdafx.h"
#include "excel2003.h"

/************************************
  REVISION LOG ENTRY
  Revision By: abao++
  Revised on 2006-8-11 8:13:15
  Comments: ...
 ************************************/


/////////////////////////////////////////////////////////////////////////////////////////
// ��Ԫ��������
CRange::CRange(Range& range)
{
	rg=range;
}
void CRange::Merge()
{
	rg.Merge(COleVariant((short)1));
}
CRange& CRange::operator=(const CString s)
{
	rg.SetValue2(COleVariant(s));
	return *this;
}
CRange& CRange::operator=(const char* str)
{
	rg.SetValue2(COleVariant(str));
	return *this;
}
CRange& CRange::operator=(Range& range)
{
	rg=range;
	return *this;
}

int CRange::Border(short mode,long BoderWidth,long ColorIndex, VARIANT color)
{
	rg.BorderAround(COleVariant((short)mode),(long)BoderWidth,(long)ColorIndex,color);	
	return 1;
}

int CRange::Border1(short mode,long BoderWidth,long ColorIndex, VARIANT color)
{
	rg.BorderAround(COleVariant((short)mode),(long)BoderWidth,(long)ColorIndex,color);	
	return 1;
}

int CRange::SetHAlign(RangeHAlignment mode)
{
	rg.SetHorizontalAlignment(COleVariant((short)mode));
	return 1;
}
int CRange::SetVAlign(RangeVAlignment mode)
{
	rg.SetVerticalAlignment(COleVariant((short)mode));
	return 1;
}

/////////////////////////////////////////////////////////////////////////////////////////
// EXCEL������

//=============================================================================
// ���캯��, ����EXCEL����, ����ȡ���еĹ�����(�ļ�)
CExcel::CExcel()
{
	if(!App.CreateDispatch("Excel.Application",NULL)) 
	{ 
		AfxMessageBox("����Excel����ʧ��!"); 
		exit(1); 
	}
	
	workbooks.AttachDispatch(App.GetWorkbooks(),true);
	this->SetVisible(false);
}

//=============================================================================
// ��������, �رյ�ǰ�ļ�,
CExcel::~CExcel()
{
	//---------------------------------------------------------------
	// �ر������Դ
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	workbook.Close(covOptional,covOptional,covOptional);	// �رյ�ǰ�ļ�
	workbooks.Close();										// �رջ�ȡ�Ĺ�����
	App.Quit();												// �˳������ķ���

	//---------------------------------------------------------------
	// �ͷŰ󶨵Ľӿ�
	shapes.ReleaseDispatch();					// �ͷ�ͼƬ�ӿ�
	range.ReleaseDispatch();					// �ͷŵ�Ԫ��ӿ�
	sheet.ReleaseDispatch();					// �ͷŹ�����ӿ�
	sheets.ReleaseDispatch(); 					// �ͷŹ������Ͻӿ�
	workbook.ReleaseDispatch();					// �ͷŹ������ӿ�
	workbooks.ReleaseDispatch();				// �ͷŹ��������Ͻӿ�
	App.ReleaseDispatch();						// �ͷŷ���ӿ�
}

//=============================================================================
// ���ĳ���ļ���EXCEL������, ���ָ���ļ���,��򿪸��ļ�
// �����ָ���ļ���, ���½�һ���ļ�
void CExcel::AddNewFile(const CString& ExtPath)
{
	if(ExtPath.IsEmpty())
	{
		COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
		workbook.AttachDispatch(workbooks.Add(covOptional));	
	}
	else
	{
		workbook.AttachDispatch(workbooks.Add(_variant_t(ExtPath)));			
	}
	sheets.AttachDispatch(workbook.GetWorksheets(),true);
	SelectSheet(1);			// ѡ���һ��������, �԰��صĵ�Ԫ���
}

//=============================================================================
// ����ļ��Ƿ����
BOOL CExcel::IsFileExist(CString strFn, BOOL bDir)
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

//=============================================================================
// ��ȡ��װĿ¼
CString CExcel::GetAppPath()
{
	char lpFileName[MAX_PATH];
	GetModuleFileName(AfxGetInstanceHandle(),lpFileName,MAX_PATH);

	CString strFileName = lpFileName;
	int nIndex = strFileName.ReverseFind ('\\');
	
	CString strPath;

	if (nIndex > 0)
		strPath = strFileName.Left (nIndex);
	else
		strPath = "";
	return strPath;
}

//=============================================================================
// ���ݹ��������, ѡ�иù�����
_Worksheet& CExcel::SelectSheet(CString& SheetName)
{
	sheet.AttachDispatch(sheets.GetItem(_variant_t(SheetName.AllocSysString())),true);
	range.AttachDispatch(sheet.GetCells(),true);
	shapes.AttachDispatch( sheet.GetShapes() );		// ��ȡ��״����
	return sheet;
}

//=============================================================================
// ���ݹ���������, ѡ�иù�����
_Worksheet& CExcel::SelectSheet(int index)//ѡ��һ����֪�����ı�
{
	sheet.AttachDispatch(sheets.GetItem(_variant_t((long)index)));
	sheet.Activate();                               //�趨������Ϊ�������
	range.AttachDispatch(sheet.GetCells(),true);
	//shapes.AttachDispatch(sheet.GetShapes());		// ��ȡ��״����
	return sheet;
}

//=============================================================================
// ɾ��ָ��������
void CExcel::DeleteSheet(int index)
{
	sheet.AttachDispatch(sheets.GetItem(COleVariant((short)index)));
	sheet.Delete();
}

//=============================================================================
// 
Range& CExcel::ActiveSheetRange()
{
	range.AttachDispatch(sheet.GetCells(),true);
	return range;
}

//=============================================================================
// ��ȡָ�����еĵ�Ԫ������
VARIANT CExcel::GetCell(int row, int col)
{
	Range rg = range.GetItem(_variant_t((long)row),_variant_t((long)col)).pdispVal;
	return rg.GetText();
}

//=============================================================================
// ��ȡָ�����еĵ�Ԫ����ֵ
int CExcel::GetCellValue(int row, int col)
{
	Range rg;
	CString str;
	COleVariant vResult;
	rg= range.GetItem(_variant_t((long)row),_variant_t((long)col)).pdispVal;
	
		vResult.lVal=0;
		vResult=rg.GetText();
		if(vResult.vt == VT_BSTR)
		{
			str = vResult.bstrVal;
			return(atoi(str));
		}
		else if(vResult.vt == VT_I2)
		{
			return(vResult.iVal);
		}
		else if(vResult.vt == VT_I4)
		{
			return(vResult.lVal);
		}
		else if(vResult.vt == VT_R8)
		{
			return((int)(vResult.dblVal));  
		}
		else if(vResult.vt == VT_EMPTY)
		{
			return 0;
		}
		else
		return 0;		
}

//=============================================================================
// ��ȡָ�����еĵ�Ԫ����ֵ
double CExcel::GetCellValueFloat(int row, int col)
{
	Range rg;
	CString str;
	COleVariant vResult;
	rg= range.GetItem(_variant_t((long)row),_variant_t((long)col)).pdispVal;
	
		vResult.lVal=0;
		vResult=rg.GetText();
		if(vResult.vt == VT_BSTR)
		{
			str = vResult.bstrVal;
			return(atof(str));
		}  
		else if(vResult.vt == VT_I2)
		{
			return(vResult.iVal);
		}
		else if(vResult.vt == VT_I4)
		{
			return((double)vResult.lVal);
		}
		else if(vResult.vt == VT_R8)
		{
			return((double)(vResult.dblVal));  
		}
		else if(vResult.vt == VT_EMPTY)
		{
			return 0;
		}
		else
		return 0;		
}

//=============================================================================
// ��ȡλ�Ÿ�����ɾ������ո����ַ�����󷵻���','Ϊ�����CString
CString CExcel::DeleteBlackSpace(CString sCell,int * inRow)
{
	int flg = 0;
	char *s = (LPTSTR)(LPCTSTR)sCell;
	char *p = s,*q = s;
	*inRow = 0;
	
	for(;*s;s++)
	{
		if(*s==' '|| *s==','||*s=='\t'|| *s=='\n')
			flg = (flg == 2) ? 1 : flg;
		else
		{
			if(flg == 1)
			{
				*q++ = ',';
				*inRow = *inRow+1;
			}
			*q++ = *s;
			flg = 2;
		}
	}
	*q = '\0';
	sCell = p;
	return sCell;
}

//=============================================================================
// ��ȡλ�Ÿ�����ɾ������ո����ַ�����󷵻���','Ϊ�����CString
CString CExcel::DeleteBlackSpace(CString sCell)
{
	int flg = 0;
	char *s = (LPTSTR)(LPCTSTR)sCell;
	char *p = s,*q = s;
	
	for(;*s;s++)
	{
		if(*s==' '|| *s==','||*s=='\t'|| *s=='\n')
			flg = (flg == 2) ? 1 : flg;
		else
		{
			if(flg == 1)
			{
				*q++ = ',';
			}
			*q++ = *s;
			flg = 2;
		}
	}
	*q = '\0';
	sCell = p;
	return sCell;
}
//=============================================================================
// ��ȡCString ǰiPart��λ�ţ�����λ��֮��','���
CString CExcel::GetPartInRow(CString strCell,int iPart)
{
	CString strCellTemp,cy;
	int i,iSpace,j=0;
	for(i=1;i<=iPart;i++)
	{
		iSpace = strCell.Find(',');
		cy = strCell.Left(iSpace);
		strCellTemp += cy;
		strCellTemp += ",";
		
		
		strCell = strCell.Right(strCell.GetLength() - iSpace - 1);
	}
	strCellTemp = DeleteBlackSpace(strCellTemp,&j);

	return strCellTemp;
}

//=============================================================================
// 
CString CExcel::GetOtherPartInRow(CString strCell,int iPart)
{
	CString strCellTemp,cy;
	int i,iSpace,j=0;
	for(i=1;i<=iPart;i++)
	{
		iSpace = strCell.Find(',');
		cy = strCell.Left(iSpace);
		strCellTemp += cy;
		strCellTemp += ",";
		
		
		strCell = strCell.Right(strCell.GetLength() - iSpace - 1);
	}
	strCell = DeleteBlackSpace(strCell,&j);
	
	return strCell;
}
//=============================================================================
// ������','Ϊ�����λ�ţ����� λ��+' '=7��CString
CString CExcel::InsetThreeSpace(CString strCell)
{
	CString cy,strCellTemp;
	int i,iSpace = 0;
	
	while(strCell.Find(',')>0)
	{
		iSpace = strCell.Find(',');
		cy = strCell.Left(iSpace);
		strCell = strCell.Right(strCell.GetLength() - iSpace - 1);

		strCellTemp += cy;

	if(iSpace<7)
		for(i=0;i<7-iSpace;i++)
		strCellTemp += " ";
	}
	strCellTemp += strCell;
	return strCellTemp;
}

//=============================================================================
// ɾ��ָ����
void CExcel::DeleteRow(int row)
{
	Range m_range;
	typedef enum 
	{
	   xlShiftToLeft = -4159,
	   xlShiftUp = -4162
	} XlDeleteShiftDirection;
	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem(_variant_t((long)row),_variant_t((long)1)).pdispVal);
	m_range.AttachDispatch(m_range.GetEntireRow());
	m_range.Delete(COleVariant((short)xlShiftUp));
	//m_range.SetRowHeight
	// �� 1,1��Ԫ���ý���
	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem(COleVariant((long)1),COleVariant((long)1)).pdispVal); 
	//m_range.Activate();
	m_range.ReleaseDispatch();
}

//=============================================================================
// ����ָ����
void CExcel::InsertRow(int row)
{
	Range m_range;
	typedef enum 
	{
	   xlShiftToLeft = -4159,
	   xlShiftUp = -4162
	} XlDeleteShiftDirection;
	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem(_variant_t((long)row),_variant_t((long)1)).pdispVal);
	m_range.AttachDispatch(m_range.GetEntireRow());
	m_range.Insert(vtMissing,vtMissing); 

	// �� 1,1��Ԫ���ý���
	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem(COleVariant((long)1),COleVariant((long)1)).pdispVal); 
	//m_range.Activate();
	m_range.ReleaseDispatch();
}

//=============================================================================
// ����ָ����
void CExcel::InsertCol(int col)
{
	Range m_range;
	COleVariant     covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
	typedef enum 
	{
	   xlShiftToLeft = -4159,
	   xlShiftUp = -4162
	} XlDeleteShiftDirection;
	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem( covOptional ,COleVariant((long)col)).pdispVal); 
	m_range.AttachDispatch(m_range.GetEntireColumn());
	m_range.Insert(vtMissing,vtMissing); 

	// �� 1,1��Ԫ���ý���
	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem(COleVariant((long)1),COleVariant((long)1)).pdispVal); 
	//m_range.Activate();
	m_range.ReleaseDispatch();
}

//=============================================================================
// ɾ��ĳ��
void CExcel::DeleteCol(int iCol)
{
	Range m_range;
	COleVariant     covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
	typedef enum 
	{
	   xlShiftToLeft = -4159,
	   xlShiftUp = -4162
	} XlDeleteShiftDirection;
	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem( covOptional ,COleVariant((long)iCol)).pdispVal); 
	m_range.AttachDispatch(m_range.GetEntireColumn());
	m_range.Delete(COleVariant((short)xlShiftToLeft));
	// �� 1,1��Ԫ���ý���
	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem(COleVariant((long)1),COleVariant((long)1)).pdispVal); 
	//m_range.Activate();
	m_range.ReleaseDispatch();
}
//=============================================================================
// �п�����Ӧ
void CExcel::SetColAutoFit(int iCol)
{
	Range m_range;
	COleVariant     covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR);

	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem( covOptional ,COleVariant((long)iCol)).pdispVal); 
	m_range.AttachDispatch(m_range.GetEntireColumn());
	m_range.AutoFit();
	// �� 1,1��Ԫ���ý���
	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem(COleVariant((long)1),COleVariant((long)1)).pdispVal); 
	//m_range.Activate();
	m_range.ReleaseDispatch();
}

//=============================================================================
// �����п�
void CExcel::SetColWidth(int iCol, float width)
{
	Range m_range;
	COleVariant     covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR);

	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem( covOptional ,COleVariant((long)iCol)).pdispVal); 
	m_range.AttachDispatch(m_range.GetEntireColumn());
	m_range.SetColumnWidth(_variant_t(width)); 
	// �� 1,1��Ԫ���ý���
	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem(COleVariant((long)1),COleVariant((long)1)).pdispVal); 
	//m_range.Activate();
	m_range.ReleaseDispatch();
}

//=============================================================================
// �����п�
void CExcel::SetRowWidth(int row, float width)
{
	Range m_range;
	COleVariant     covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR);

	m_range.AttachDispatch( sheet.GetCells() );   
	//m_range.AttachDispatch(m_range.GetItem( covOptional ,COleVariant((long)iCol)).pdispVal); 
	m_range.AttachDispatch(m_range.GetItem(_variant_t((long)row),_variant_t((long)1)).pdispVal);
	//m_range.AttachDispatch(m_range.GetEntireColumn());
	m_range.AttachDispatch(m_range.GetEntireRow());
	m_range.SetRowHeight(_variant_t(width)); 
	// �� 1,1��Ԫ���ý���
	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem(COleVariant((long)1),COleVariant((long)1)).pdispVal); 
	//m_range.Activate();
	m_range.ReleaseDispatch();
}

//=============================================================================
// ��ȡ��ʹ����
int CExcel::GetUsedRowCount()
{
	range.AttachDispatch(sheet.GetUsedRange());//Range�Ƿ�Χ����˼
	range.AttachDispatch(range.GetRows());//�õ����е���
	return (range.GetCount());    //�Ѿ�ʹ�õ�����
}

//=============================================================================
// ��ȡ��ʹ����
int CExcel::GetUsedColCount()
{
	range.AttachDispatch(sheet.GetUsedRange());//Range�Ƿ�Χ����˼
	range.AttachDispatch(range.GetColumns());//�õ����е���
	return(range.GetCount());//��ʹ�õ�����
}

//=============================================================================
// �����������Ԫ������
void CExcel::ClearUsedRange()
{
	range.AttachDispatch(sheet.GetUsedRange());//Range�Ƿ�Χ����˼
	range.Clear();
}

//=============================================================================
// ��ָ�����еĵ�Ԫ����д���ַ���
void CExcel::SetCell(int row,int col,CString &str)
{
	range.SetItem(_variant_t((long)row),_variant_t((long)col),_variant_t(str));
}

//=============================================================================
// ��ָ�����еĵ�Ԫ����д���ַ���
void CExcel::SetCell(int row,int col,char* str)
{
	SetCell(row,col,CString(str));
}

//=============================================================================
// ��ָ�����еĵ�Ԫ����д����������
void CExcel::SetCell(int row,int col,long lv)
{
	CString t;
	t.Format("%ld",lv);
	SetCell(row,col,t);
}

//=============================================================================
// ��ָ�����еĵ�Ԫ����д�븡��������, ���һ������Ϊ����С������λ��
void CExcel::SetCell(int row,int col,double dv,int n)
{
	CString t;
	CString format;
	format.Format("%%.%dlf",n);
	t.Format(format,dv);
	SetCell(row,col,t);
}

//=============================================================================
// �ѵ�ǰ�ļ����Ϊָ�����ļ���
int CExcel::SaveAs(CString FileName)
{
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if(IsFileExist(FileName,FALSE)==TRUE)
		DeleteFile(FileName);
	this->workbook.SaveAs(COleVariant(FileName),covOptional,covOptional, covOptional,covOptional,covOptional,0,
	covOptional,covOptional,covOptional,covOptional,covOptional);  

	return 1;
}

//=============================================================================
// �ѵ�ǰ�ļ����Ϊָ�����ļ���
void CExcel::Save(BOOL bNewValue)
{
	workbook.SetSaved(bNewValue);
}

//=============================================================================
// �ѵ�ǰ��������Ϊָ�����Ƶ��¹�����
void CExcel::CopySheet(_Worksheet &sht)
{
	sheet.Copy(vtMissing,_variant_t(sht));
}

//=============================================================================
// ��ȡrange
Range& CExcel::GetRange(CString RangeStart,CString RangeEnd)
{
	range=sheet.GetRange(COleVariant(RangeStart),COleVariant(RangeEnd));	
	return range;
}

//=============================================================================
// ��ȡrange ����ΪA1:A2ģʽ
Range& CExcel::GetRange(CString RangeStr)
{
	int pos=RangeStr.Find(':');
	if(pos>0)
	{
		CString a,b;
		a=RangeStr.Left(pos);
		b=RangeStr.Right(RangeStr.GetLength()-pos-1);
		return GetRange(a,b);
	}
	else
	{
		return GetRange(RangeStr,RangeStr);
	}
}

//=============================================================================
//��ȡrange ��ʼ��/�У�������/��
Range& CExcel::GetRange(int startRow, int startCol, int endRow, int endCol)	
{
	CString RangeStr,csTemp,csStart,csEnd;
	#define  ALPHABET_CNT  26
	unsigned char assiLetter[ALPHABET_CNT] = {65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90};

	if(startCol>ALPHABET_CNT)
	for(int i=0; i<startCol/ALPHABET_CNT; i++)
	{
		csTemp.Format("%c",assiLetter[i]);
		csStart = csStart + csTemp;
	}
	csTemp.Format("%c",assiLetter[(startCol-1)%ALPHABET_CNT]);
	csStart = csStart + csTemp; 

	csTemp.Format("%d",startRow);
	csStart = csStart + csTemp;

	if(endCol>ALPHABET_CNT)
	for(int j=0; j<endCol/ALPHABET_CNT; j++)
	{
		csTemp.Format("%c",assiLetter[j]);
		csEnd = csEnd + csTemp;
	}
	csTemp.Format("%c",assiLetter[(endCol-1)%ALPHABET_CNT]);
	csEnd = csEnd + csTemp;

	csTemp.Format("%d",endRow);
	csEnd = csEnd + csTemp;

	RangeStr = csStart + ":" + csEnd;
	
	//RangeStr.Format("%c%d:%c%d",assiLetter[(startCol-1)%ALPHABET_CNT],startRow,assiLetter[(endCol-1)%ALPHABET_CNT],endRow);


	int pos=RangeStr.Find(':');
	if(pos>0)
	{
		CString a,b;
		a=RangeStr.Left(pos);
		b=RangeStr.Right(RangeStr.GetLength()-pos-1);
		return GetRange(a,b);
	}
	else
	{
		return GetRange(RangeStr,RangeStr);
	}
}

//=============================================================================
// �ϲ�Range
void CExcel::MergeRange(CString RangeStr)
{
	GetRange(RangeStr).Merge(COleVariant(long(1)));
}

//=============================================================================
// �������
_Worksheet& CExcel::ActiveSheet()
{
	sheet=workbook.GetActiveSheet();
	return sheet;
}

//=============================================================================
// ����ͼƬ
// �����еĸ߶ȺͿ��ָ����ͼƬ��,��ͼƬ���ź�ĸ߶�,�������Ϊ-1��ʹ��ԭͼƬ�Ĵ�С
ShapeRange& CExcel::AddPicture(LPCTSTR Filename, float Left, float Top, float Width, float Height)
{
	sharpRange = shapes.AddPicture(Filename,0, 1,Left,Top,Width,Height); 
	return sharpRange;
}
