#ifndef _abaoExcel_h_

/******************************************************************************
  描述: EXCEL操作类: 
		1,写入
		2,读取
		3,单元格合并
		4,单元格格式操作
		5,插入图片
******************************************************************************/

#define _abaoExcel_h_

#include "excel.h"
#include <comdef.h>


enum RangeHAlignment{HAlignDefault=1,HAlignCenter=-4108,HAlignLeft=-4131,HAlignRight=-4152};
enum RangeVAlignment{VAlignDefault=2,VAlignCenter=-4108,VAlignTop=-4160,VAlignBottom=-4107};

///////////////////////////////////////////////////////////////////////////////////////////
// EXCEL单元格操作类
class CRange
{
	Range rg;							//参数用range
public:
	CRange(Range& range);				//从一个range构造
	CRange& operator=(const CString s);	//填入一个CString
	CRange& operator=(const char* str);	//填入char*
	CRange& operator=(Range& range);	//赋值另一个range

	void Merge();		//合并
	//设置边框，参数还不是很清楚
	int Border(short mode=1,long BoderWidth=3,long ColorIndex=1,
		VARIANT color=COleVariant((long)DISP_E_PARAMNOTFOUND,VT_ERROR));

	int Border1(short mode=1,long BoderWidth=1,long ColorIndex=1,
		VARIANT color=COleVariant((long)DISP_E_PARAMNOTFOUND,VT_ERROR));

	//设置水平对齐方式
	int SetHAlign(RangeHAlignment mode=HAlignDefault);

	//设置竖直对齐
	int SetVAlign(RangeVAlignment mode=VAlignDefault);
};

///////////////////////////////////////////////////////////////////////////////////////////
// EXCEL操作类
class CExcel
{
public:
	CExcel();
	~CExcel();

	//=========================================================================
	// 
	_Application	App;			// EXCEL 应用程序
	Workbooks		workbooks;		// 工作簿集合,相当于所有的EXCEL文件
	_Workbook		workbook;		// 某个工作簿,相当于某个EXCEL文件
	Worksheets		sheets;			// 工作表集合,相当于一个EXCEL文件内的所有工作表
	_Worksheet		sheet;			// 某个工作表 ,相当于一个文件的某个工作表
	Range			range;			// 单元格区域,range
	Shapes			shapes;			// 所有图片的集合
	ShapeRange		sharpRange;		// 某个图片


	//=========================================================================
	// EXCEL程序相关操作
	int SetVisible(bool visible) {App.SetVisible(visible);return 1;}//设置为可见以及隐藏
	int SaveAs(CString FileName);		//保存到文件名
	void Save(BOOL bNewValue);            //保存文件
	BOOL IsFileExist(CString strFn, BOOL bDir);   //检查文件是否存在
	static CString GetAppPath();                 //获取安装目录

	//=========================================================================
	// 工作簿相关操作
	_Worksheet& ActiveSheet();		//当前活动的sheet,在SelectSheet 后改变
	void CopySheet(_Worksheet &sht);			//复制一个sheet	
	void AddNewFile(const CString& ExtPath=CString(""));	//从一个模版构造

	//=========================================================================
	// 选择工作表
	_Worksheet& SelectSheet(CString& SheetName);	//选择一个已知表名的表
	_Worksheet& SelectSheet(char* SheetName)
		{return SelectSheet(CString(SheetName));};	//选择一个已知表名的表
	_Worksheet& SelectSheet(int index);				//选择一个已知索引的表
	void DeleteSheet(int index);            // 删除指定工作表

	//=========================================================================
	// 行/列操作
	void DeleteRow(int row);			        // 删除指定行
	void DeleteCol(int iCol);					//删除指定列
	int GetUsedRowCount(void);			        //获取已使用行
	int GetUsedColCount(void);			         //获取已使用列
	void ClearUsedRange();
	void InsertRow(int row);                   //插入指定行
	void InsertCol(int col);                   //插入指定列
	void SetColAutoFit(int iCol);              // 列宽自适应
	void SetColWidth(int iCol,float width);    // 设置列宽
	void SetRowWidth(int iCol,float width);    // 设置行宽

	//=========================================================================
	// 填写单元格
	void SetCell(int row,int col,CString &str);		//指定行列的单元格填入值
	void SetCell(int row,int col,char* str);		//指定行列的单元格填入值
	void SetCell(int row,int col,long lv);			//指定行列填入long值
	void SetCell(int row,int col,double dv,int n=6);//指定行列填入浮点值，并截取为指定的小数位
	VARIANT GetCell(int row, int col);
	int GetCellValue(int row, int col);
	double GetCellValueFloat(int row, int col);

	//=========================================================================
	// 单元格区域操作
	Range& ActiveSheetRange();								//当前的range,在使用GetRange后改变
	Range& GetRange(CString RangeStart,CString RangeEnd);	//获取range,
	Range& GetRange(CString RangeStr);						//获取range A1:A2模式
	Range& GetRange(int startRow, int startCol, int endRow, int endCol);	//获取range 起始行/列，结束行/列
	void MergeRange(CString RangeStr);						//合并Range
	CString DeleteBlackSpace(CString sCell,int * inRow);                //获取位号个数并删除多余空格
	CString GetPartInRow(CString strCell,int iPart);
	CString GetOtherPartInRow(CString strCell,int iPart);
	CString InsetThreeSpace(CString strCell);
	static CString DeleteBlackSpace(CString sCell);    

	//=========================================================================
	// 图片操作
	ShapeRange& AddPicture(LPCTSTR Filename, float Left, float Top,
		float Width, float Height);		// 插入图片, 如果高度和宽度为-1,则使用图片原始大小

};
#endif