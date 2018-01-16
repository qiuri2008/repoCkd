#ifndef _abaoExcel_h_

/******************************************************************************
  ����: EXCEL������: 
		1,д��
		2,��ȡ
		3,��Ԫ��ϲ�
		4,��Ԫ���ʽ����
		5,����ͼƬ
******************************************************************************/

#define _abaoExcel_h_

#include "excel.h"
#include <comdef.h>


enum RangeHAlignment{HAlignDefault=1,HAlignCenter=-4108,HAlignLeft=-4131,HAlignRight=-4152};
enum RangeVAlignment{VAlignDefault=2,VAlignCenter=-4108,VAlignTop=-4160,VAlignBottom=-4107};

///////////////////////////////////////////////////////////////////////////////////////////
// EXCEL��Ԫ�������
class CRange
{
	Range rg;							//������range
public:
	CRange(Range& range);				//��һ��range����
	CRange& operator=(const CString s);	//����һ��CString
	CRange& operator=(const char* str);	//����char*
	CRange& operator=(Range& range);	//��ֵ��һ��range

	void Merge();		//�ϲ�
	//���ñ߿򣬲��������Ǻ����
	int Border(short mode=1,long BoderWidth=3,long ColorIndex=1,
		VARIANT color=COleVariant((long)DISP_E_PARAMNOTFOUND,VT_ERROR));

	int Border1(short mode=1,long BoderWidth=1,long ColorIndex=1,
		VARIANT color=COleVariant((long)DISP_E_PARAMNOTFOUND,VT_ERROR));

	//����ˮƽ���뷽ʽ
	int SetHAlign(RangeHAlignment mode=HAlignDefault);

	//������ֱ����
	int SetVAlign(RangeVAlignment mode=VAlignDefault);
};

///////////////////////////////////////////////////////////////////////////////////////////
// EXCEL������
class CExcel
{
public:
	CExcel();
	~CExcel();

	//=========================================================================
	// 
	_Application	App;			// EXCEL Ӧ�ó���
	Workbooks		workbooks;		// ����������,�൱�����е�EXCEL�ļ�
	_Workbook		workbook;		// ĳ��������,�൱��ĳ��EXCEL�ļ�
	Worksheets		sheets;			// ��������,�൱��һ��EXCEL�ļ��ڵ����й�����
	_Worksheet		sheet;			// ĳ�������� ,�൱��һ���ļ���ĳ��������
	Range			range;			// ��Ԫ������,range
	Shapes			shapes;			// ����ͼƬ�ļ���
	ShapeRange		sharpRange;		// ĳ��ͼƬ


	//=========================================================================
	// EXCEL������ز���
	int SetVisible(bool visible) {App.SetVisible(visible);return 1;}//����Ϊ�ɼ��Լ�����
	int SaveAs(CString FileName);		//���浽�ļ���
	void Save(BOOL bNewValue);            //�����ļ�
	BOOL IsFileExist(CString strFn, BOOL bDir);   //����ļ��Ƿ����
	static CString GetAppPath();                 //��ȡ��װĿ¼

	//=========================================================================
	// ��������ز���
	_Worksheet& ActiveSheet();		//��ǰ���sheet,��SelectSheet ��ı�
	void CopySheet(_Worksheet &sht);			//����һ��sheet	
	void AddNewFile(const CString& ExtPath=CString(""));	//��һ��ģ�湹��

	//=========================================================================
	// ѡ������
	_Worksheet& SelectSheet(CString& SheetName);	//ѡ��һ����֪�����ı�
	_Worksheet& SelectSheet(char* SheetName)
		{return SelectSheet(CString(SheetName));};	//ѡ��һ����֪�����ı�
	_Worksheet& SelectSheet(int index);				//ѡ��һ����֪�����ı�
	void DeleteSheet(int index);            // ɾ��ָ��������

	//=========================================================================
	// ��/�в���
	void DeleteRow(int row);			        // ɾ��ָ����
	void DeleteCol(int iCol);					//ɾ��ָ����
	int GetUsedRowCount(void);			        //��ȡ��ʹ����
	int GetUsedColCount(void);			         //��ȡ��ʹ����
	void ClearUsedRange();
	void InsertRow(int row);                   //����ָ����
	void InsertCol(int col);                   //����ָ����
	void SetColAutoFit(int iCol);              // �п�����Ӧ
	void SetColWidth(int iCol,float width);    // �����п�
	void SetRowWidth(int iCol,float width);    // �����п�

	//=========================================================================
	// ��д��Ԫ��
	void SetCell(int row,int col,CString &str);		//ָ�����еĵ�Ԫ������ֵ
	void SetCell(int row,int col,char* str);		//ָ�����еĵ�Ԫ������ֵ
	void SetCell(int row,int col,long lv);			//ָ����������longֵ
	void SetCell(int row,int col,double dv,int n=6);//ָ���������븡��ֵ������ȡΪָ����С��λ
	VARIANT GetCell(int row, int col);
	int GetCellValue(int row, int col);
	double GetCellValueFloat(int row, int col);

	//=========================================================================
	// ��Ԫ���������
	Range& ActiveSheetRange();								//��ǰ��range,��ʹ��GetRange��ı�
	Range& GetRange(CString RangeStart,CString RangeEnd);	//��ȡrange,
	Range& GetRange(CString RangeStr);						//��ȡrange A1:A2ģʽ
	Range& GetRange(int startRow, int startCol, int endRow, int endCol);	//��ȡrange ��ʼ��/�У�������/��
	void MergeRange(CString RangeStr);						//�ϲ�Range
	CString DeleteBlackSpace(CString sCell,int * inRow);                //��ȡλ�Ÿ�����ɾ������ո�
	CString GetPartInRow(CString strCell,int iPart);
	CString GetOtherPartInRow(CString strCell,int iPart);
	CString InsetThreeSpace(CString strCell);
	static CString DeleteBlackSpace(CString sCell);    

	//=========================================================================
	// ͼƬ����
	ShapeRange& AddPicture(LPCTSTR Filename, float Left, float Top,
		float Width, float Height);		// ����ͼƬ, ����߶ȺͿ��Ϊ-1,��ʹ��ͼƬԭʼ��С

};
#endif