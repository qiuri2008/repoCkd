
#include "stdafx.h"
#include "excel2003.h"

/************************************
  REVISION LOG ENTRY
  Revision By: abao++
  Revised on 2006-8-11 8:13:15
  Comments: ...
 ************************************/


/////////////////////////////////////////////////////////////////////////////////////////
// 单元格区域类
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
// EXCEL操作类

//=============================================================================
// 构造函数, 启动EXCEL服务, 并获取所有的工作簿(文件)
CExcel::CExcel()
{
	if(!App.CreateDispatch("Excel.Application",NULL)) 
	{ 
		AfxMessageBox("创建Excel服务失败!"); 
		exit(1); 
	}
	
	workbooks.AttachDispatch(App.GetWorkbooks(),true);
	this->SetVisible(false);
}

//=============================================================================
// 析构函数, 关闭当前文件,
CExcel::~CExcel()
{
	//---------------------------------------------------------------
	// 关闭相关资源
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	workbook.Close(covOptional,covOptional,covOptional);	// 关闭当前文件
	workbooks.Close();										// 关闭获取的工作簿
	App.Quit();												// 退出创建的服务

	//---------------------------------------------------------------
	// 释放绑定的接口
	shapes.ReleaseDispatch();					// 释放图片接口
	range.ReleaseDispatch();					// 释放单元格接口
	sheet.ReleaseDispatch();					// 释放工作表接口
	sheets.ReleaseDispatch(); 					// 释放工作表集合接口
	workbook.ReleaseDispatch();					// 释放工作簿接口
	workbooks.ReleaseDispatch();				// 释放工作簿集合接口
	App.ReleaseDispatch();						// 释放服务接口
}

//=============================================================================
// 添加某个文件到EXCEL服务中, 如果指定文件名,则打开该文件
// 如果不指定文件名, 则新建一个文件
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
	SelectSheet(1);			// 选择第一个工作表, 以邦定相关的单元格等
}

//=============================================================================
// 检查文件是否存在
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
// 获取安装目录
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
// 根据工作表表名, 选中该工作表
_Worksheet& CExcel::SelectSheet(CString& SheetName)
{
	sheet.AttachDispatch(sheets.GetItem(_variant_t(SheetName.AllocSysString())),true);
	range.AttachDispatch(sheet.GetCells(),true);
	shapes.AttachDispatch( sheet.GetShapes() );		// 获取形状集合
	return sheet;
}

//=============================================================================
// 根据工作表索引, 选中该工作表
_Worksheet& CExcel::SelectSheet(int index)//选择一个已知表名的表
{
	sheet.AttachDispatch(sheets.GetItem(_variant_t((long)index)));
	sheet.Activate();                               //设定工作表为活动工作表
	range.AttachDispatch(sheet.GetCells(),true);
	//shapes.AttachDispatch(sheet.GetShapes());		// 获取形状集合
	return sheet;
}

//=============================================================================
// 删除指定工作表
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
// 读取指定行列的单元格内容
VARIANT CExcel::GetCell(int row, int col)
{
	Range rg = range.GetItem(_variant_t((long)row),_variant_t((long)col)).pdispVal;
	return rg.GetText();
}

//=============================================================================
// 读取指定行列的单元格数值
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
// 读取指定行列的单元格数值
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
// 获取位号个数并删除多余空格类字符，最后返回以','为间隔的CString
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
// 获取位号个数并删除多余空格类字符，最后返回以','为间隔的CString
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
// 获取CString 前iPart个位号，并且位号之间','间隔
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
// 处理以','为间隔的位号，返回 位号+' '=7的CString
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
// 删除指定行
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
	// 让 1,1单元格获得焦点
	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem(COleVariant((long)1),COleVariant((long)1)).pdispVal); 
	//m_range.Activate();
	m_range.ReleaseDispatch();
}

//=============================================================================
// 插入指定行
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

	// 让 1,1单元格获得焦点
	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem(COleVariant((long)1),COleVariant((long)1)).pdispVal); 
	//m_range.Activate();
	m_range.ReleaseDispatch();
}

//=============================================================================
// 插入指定列
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

	// 让 1,1单元格获得焦点
	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem(COleVariant((long)1),COleVariant((long)1)).pdispVal); 
	//m_range.Activate();
	m_range.ReleaseDispatch();
}

//=============================================================================
// 删除某列
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
	// 让 1,1单元格获得焦点
	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem(COleVariant((long)1),COleVariant((long)1)).pdispVal); 
	//m_range.Activate();
	m_range.ReleaseDispatch();
}
//=============================================================================
// 列宽自适应
void CExcel::SetColAutoFit(int iCol)
{
	Range m_range;
	COleVariant     covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR);

	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem( covOptional ,COleVariant((long)iCol)).pdispVal); 
	m_range.AttachDispatch(m_range.GetEntireColumn());
	m_range.AutoFit();
	// 让 1,1单元格获得焦点
	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem(COleVariant((long)1),COleVariant((long)1)).pdispVal); 
	//m_range.Activate();
	m_range.ReleaseDispatch();
}

//=============================================================================
// 设置列宽
void CExcel::SetColWidth(int iCol, float width)
{
	Range m_range;
	COleVariant     covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR);

	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem( covOptional ,COleVariant((long)iCol)).pdispVal); 
	m_range.AttachDispatch(m_range.GetEntireColumn());
	m_range.SetColumnWidth(_variant_t(width)); 
	// 让 1,1单元格获得焦点
	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem(COleVariant((long)1),COleVariant((long)1)).pdispVal); 
	//m_range.Activate();
	m_range.ReleaseDispatch();
}

//=============================================================================
// 设置行宽
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
	// 让 1,1单元格获得焦点
	m_range.AttachDispatch( sheet.GetCells() );   
	m_range.AttachDispatch(m_range.GetItem(COleVariant((long)1),COleVariant((long)1)).pdispVal); 
	//m_range.Activate();
	m_range.ReleaseDispatch();
}

//=============================================================================
// 获取已使用行
int CExcel::GetUsedRowCount()
{
	range.AttachDispatch(sheet.GetUsedRange());//Range是范围的意思
	range.AttachDispatch(range.GetRows());//得到所有的行
	return (range.GetCount());    //已经使用的行数
}

//=============================================================================
// 获取已使用列
int CExcel::GetUsedColCount()
{
	range.AttachDispatch(sheet.GetUsedRange());//Range是范围的意思
	range.AttachDispatch(range.GetColumns());//得到所有的列
	return(range.GetCount());//已使用的列数
}

//=============================================================================
// 清除已用区域单元格内容
void CExcel::ClearUsedRange()
{
	range.AttachDispatch(sheet.GetUsedRange());//Range是范围的意思
	range.Clear();
}

//=============================================================================
// 往指定行列的单元格中写入字符串
void CExcel::SetCell(int row,int col,CString &str)
{
	range.SetItem(_variant_t((long)row),_variant_t((long)col),_variant_t(str));
}

//=============================================================================
// 往指定行列的单元格中写入字符串
void CExcel::SetCell(int row,int col,char* str)
{
	SetCell(row,col,CString(str));
}

//=============================================================================
// 往指定行列的单元格中写入整型数据
void CExcel::SetCell(int row,int col,long lv)
{
	CString t;
	t.Format("%ld",lv);
	SetCell(row,col,t);
}

//=============================================================================
// 往指定行列的单元格中写入浮点型数据, 最后一个参数为保留小数点后的位数
void CExcel::SetCell(int row,int col,double dv,int n)
{
	CString t;
	CString format;
	format.Format("%%.%dlf",n);
	t.Format(format,dv);
	SetCell(row,col,t);
}

//=============================================================================
// 把当前文件另存为指定的文件名
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
// 把当前文件另存为指定的文件名
void CExcel::Save(BOOL bNewValue)
{
	workbook.SetSaved(bNewValue);
}

//=============================================================================
// 把当前工作表复制为指定名称的新工作表
void CExcel::CopySheet(_Worksheet &sht)
{
	sheet.Copy(vtMissing,_variant_t(sht));
}

//=============================================================================
// 获取range
Range& CExcel::GetRange(CString RangeStart,CString RangeEnd)
{
	range=sheet.GetRange(COleVariant(RangeStart),COleVariant(RangeEnd));	
	return range;
}

//=============================================================================
// 获取range 参数为A1:A2模式
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
//获取range 起始行/列，结束行/列
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
// 合并Range
void CExcel::MergeRange(CString RangeStr)
{
	GetRange(RangeStr).Merge(COleVariant(long(1)));
}

//=============================================================================
// 激活工作簿
_Worksheet& CExcel::ActiveSheet()
{
	sheet=workbook.GetActiveSheet();
	return sheet;
}

//=============================================================================
// 插入图片
// 参数中的高度和宽度指插入图片后,此图片缩放后的高度,如果参数为-1将使用原图片的大小
ShapeRange& CExcel::AddPicture(LPCTSTR Filename, float Left, float Top, float Width, float Height)
{
	sharpRange = shapes.AddPicture(Filename,0, 1,Left,Top,Width,Height); 
	return sharpRange;
}
