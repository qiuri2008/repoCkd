/******************************************************************************
//定义并声明全局变量
*******************************************************************************/
#pragma once

#include <afxtempl.h>
#include "User.h"


#define  USE_SPEECH_DLL
#include "LaneSpeech.h"
#pragma  comment (lib,"Speech_Re.lib")

#define ENABLE_TTS

#ifdef _GLOBAL_C_
  #define _GLOBALDEC_
#else
  #define _GLOBALDEC_ extern
#endif

#define SPEAK_BUFFSIZE	4096
_GLOBALDEC_ CTTS  tts;
_GLOBALDEC_ BOOL	g_ttsCreate;
_GLOBALDEC_ wchar_t  g_wcRead[SPEAK_BUFFSIZE];

typedef CMap <CString, LPCTSTR, CLayers, CLayers &> CMapInBox;  
typedef CMap <CString, LPCTSTR, int, int > CMapCS;
typedef CMap <CString, LPCTSTR, int, int> CMapPartCnt; 
typedef CArray<CLayers,CLayers &> CArrayClayer;
typedef CMap <CString, LPCTSTR, CArrayClayer, CArrayClayer &> CMapOutBox; 

//===================================================================
//测试
//===================================================================
typedef CMap <CString, LPCTSTR, CStringList *, CStringList *> CMapTest;
_GLOBALDEC_ CMapTest mapTest;
//===================================================================

_GLOBALDEC_ CMapInBox mapInBox;								//内箱清单 内箱号为键值
_GLOBALDEC_ CMapInBox mapBaseList;							//基础清单 部品号为键值
_GLOBALDEC_ CMapInBox mapWeightList;						//重量清单 部品号为键值
_GLOBALDEC_ CMapCS mapOutBox;								//外箱-内箱数量清单 内箱数量的仅作为唛头边框
_GLOBALDEC_ CTypedPtrList<CObList,CLayers*> m_LayerList;	//保存清单 
_GLOBALDEC_ CMapPartCnt mapPartCnt;							//部品号-数量清单  从内箱清单中生产 表示某个物料的录入数量
_GLOBALDEC_ CLayers g_ssk;									//数据单元	每条记录的实体
_GLOBALDEC_ CMapStringToString mapOutboxPallet;				//外箱-托盘  外箱为键值
_GLOBALDEC_ CMapCS mapPalletCnt;							//托盘-外箱数量  托盘为键值

_GLOBALDEC_ CMapStringToOb mapCheck;						//部品号-供应商

_GLOBALDEC_ CMapInBox mapBom,mapBomAdd;						//序号 部品号  
_GLOBALDEC_ BOOL m_flgNoExist;								//检查序号清单

_GLOBALDEC_ CString g_strSaveFile;
_GLOBALDEC_ CString g_strOpenFile;
_GLOBALDEC_ unsigned int m_usedRow;
_GLOBALDEC_ unsigned int m_usedCol;
_GLOBALDEC_ int g_iOrder;       
_GLOBALDEC_ int g_iReOrder;  
_GLOBALDEC_ CString g_csTemp;   
  

_GLOBALDEC_ int m_icolProductCnt;			//生产数量所在列
_GLOBALDEC_ int m_icolPart;					//部品号所在列
_GLOBALDEC_ int m_irowPart;					//部品号所在行
_GLOBALDEC_ int m_icolUserCode;				//客户号所在列          
_GLOBALDEC_ int m_icolDetail;				//描述所在列  
_GLOBALDEC_ int m_icolNo;					//料号所在列   
_GLOBALDEC_ int m_icolWeight;				//单位重量所在列   
_GLOBALDEC_ int m_icolRadio;				//加乘系数所在列   
_GLOBALDEC_ int m_icolUnitPrice;			//单价所在列
_GLOBALDEC_ int m_icolFactoryNo;			//工厂部品号所在列   

_GLOBALDEC_ int m_icolInBox;          
_GLOBALDEC_ int m_icolOutBox;  
   
_GLOBALDEC_ BOOL m_flgCheck;				//打开文件是否正确标志   
_GLOBALDEC_ BOOL m_flgTaojian;				//指定套件标记
_GLOBALDEC_ BOOL g_flgSpeak;				//指定语音提示标记
_GLOBALDEC_ int g_iSpeakCheckCnt ;			//录入两次内箱，触发校验

#define TOULIAO_LIST_PART					"部品号"
#define TOULIAO_LIST_PRODUCT_CNT			"总数量"
#define TOULIAO_LIST_DETAIL					"描述"
#define TOULIAO_LIST_USER					"客户品号"
#define TOULIAO_LIST_INBOX					"内箱号"
#define TOULIAO_LIST_OUTBOX					"外箱箱号"
#define TOULIAO_LIST_NO						"No."
#define TOULIAO_LIST_FACTORY_NO				"工厂部品号"


#define  INBOX_COMMON						"CKD"
#define  SHIPPING_WIDTH						51
#define  SHIPPING_HIGH						10
#define  SHIPPING_HOFFSET					4
#define  SHIPPING_WOFFSET					2

#define  TITLE1_COL_OFFSET					1
#define  TITLE2_COL_OFFSET					11
#define  TITLE3_COL_OFFSET					25
#define  TITLE4_COL_OFFSET					35
#define  TITLE_WCELL1						8
#define  TITLE_WCELL						9
#define  TITLE_WCELL3_4						8

#define  TITLE_ROW_WIDTH					10 

#define  CONTENT1_COL_OFFSET  SHIPPING_WOFFSET
#define  CONTENT2_COL_OFFSET  SHIPPING_WOFFSET + 4
#define  CONTENT3_COL_OFFSET  SHIPPING_WOFFSET + 11
#define  CONTENT4_COL_OFFSET  SHIPPING_WOFFSET + 34
#define  CONTENT5_COL_OFFSET  SHIPPING_WOFFSET + 39

#define  CONTENT1_COL_OFFSET_2  SHIPPING_WOFFSET				//4 內箱起始
#define  CONTENT2_COL_OFFSET_2  CONTENT1_COL_OFFSET_2 + 4		//7 SAP 
#define  CONTENT3_COL_OFFSET_2  CONTENT2_COL_OFFSET_2 + 7		//20 描述
#define  CONTENT4_COL_OFFSET_2  CONTENT3_COL_OFFSET_2 + 19		//4 数量
#define  CONTENT5_COL_OFFSET_2  CONTENT4_COL_OFFSET_2 + 4		//9 部品号
#define  CONTENT6_COL_OFFSET_2  CONTENT5_COL_OFFSET_2 + 9		//7 工厂


#define  PROPPAGE_SETTING_INDEX		4

typedef enum  _ReceiveStyle_
{
	STYLE_PARTNO,
	STYLE_INBOX,
	STYLE_OUTBOX,
	STYLE_QTY,
	STYLE_END
}ReceiveStyle;

typedef enum
{
	STYLE_PART,
	STYLE_PROVIDER,
	STYLE_END_CHECK,
}CheckReceiveStyle;

typedef enum  _BomStyle_
{
	STYLE_BOM_LOAD,
	STYLE_BOM_SAVE,
	STYLE_BASE_LOAD,
	STYLE_BASE_SAVE,
	STYLE_CHECK_END
}BomStyle;

typedef enum _PalletStyle_
{
	STYLE_PALLET,
	STYLE_PALLET_OUTBOX,
	STYLE_PALLET_END
}PalletStyle;

extern void Speak(CString m_sText);
extern void MessageBox1(CString strText,CString strCaption,UINT nType);
extern BOOL IsFileExist(CString strFn, BOOL bDir);
extern void InitConfigFile(BOOL read);

extern BOOL g_bUseDefine;
extern unsigned int g_iDisplayWith;
extern unsigned int g_iDisplayHigh;
extern unsigned int g_iSaveMode;







