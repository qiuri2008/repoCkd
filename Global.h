/******************************************************************************
//���岢����ȫ�ֱ���
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
//����
//===================================================================
typedef CMap <CString, LPCTSTR, CStringList *, CStringList *> CMapTest;
_GLOBALDEC_ CMapTest mapTest;
//===================================================================

_GLOBALDEC_ CMapInBox mapInBox;								//�����嵥 �����Ϊ��ֵ
_GLOBALDEC_ CMapInBox mapBaseList;							//�����嵥 ��Ʒ��Ϊ��ֵ
_GLOBALDEC_ CMapInBox mapWeightList;						//�����嵥 ��Ʒ��Ϊ��ֵ
_GLOBALDEC_ CMapCS mapOutBox;								//����-���������嵥 ���������Ľ���Ϊ��ͷ�߿�
_GLOBALDEC_ CTypedPtrList<CObList,CLayers*> m_LayerList;	//�����嵥 
_GLOBALDEC_ CMapPartCnt mapPartCnt;							//��Ʒ��-�����嵥  �������嵥������ ��ʾĳ�����ϵ�¼������
_GLOBALDEC_ CLayers g_ssk;									//���ݵ�Ԫ	ÿ����¼��ʵ��
_GLOBALDEC_ CMapStringToString mapOutboxPallet;				//����-����  ����Ϊ��ֵ
_GLOBALDEC_ CMapCS mapPalletCnt;							//����-��������  ����Ϊ��ֵ

_GLOBALDEC_ CMapStringToOb mapCheck;						//��Ʒ��-��Ӧ��

_GLOBALDEC_ CMapInBox mapBom,mapBomAdd;						//��� ��Ʒ��  
_GLOBALDEC_ BOOL m_flgNoExist;								//�������嵥

_GLOBALDEC_ CString g_strSaveFile;
_GLOBALDEC_ CString g_strOpenFile;
_GLOBALDEC_ unsigned int m_usedRow;
_GLOBALDEC_ unsigned int m_usedCol;
_GLOBALDEC_ int g_iOrder;       
_GLOBALDEC_ int g_iReOrder;  
_GLOBALDEC_ CString g_csTemp;   
  

_GLOBALDEC_ int m_icolProductCnt;			//��������������
_GLOBALDEC_ int m_icolPart;					//��Ʒ��������
_GLOBALDEC_ int m_irowPart;					//��Ʒ��������
_GLOBALDEC_ int m_icolUserCode;				//�ͻ���������          
_GLOBALDEC_ int m_icolDetail;				//����������  
_GLOBALDEC_ int m_icolNo;					//�Ϻ�������   
_GLOBALDEC_ int m_icolWeight;				//��λ����������   
_GLOBALDEC_ int m_icolRadio;				//�ӳ�ϵ��������   
_GLOBALDEC_ int m_icolUnitPrice;			//����������
_GLOBALDEC_ int m_icolFactoryNo;			//������Ʒ��������   

_GLOBALDEC_ int m_icolInBox;          
_GLOBALDEC_ int m_icolOutBox;  
   
_GLOBALDEC_ BOOL m_flgCheck;				//���ļ��Ƿ���ȷ��־   
_GLOBALDEC_ BOOL m_flgTaojian;				//ָ���׼����
_GLOBALDEC_ BOOL g_flgSpeak;				//ָ��������ʾ���
_GLOBALDEC_ int g_iSpeakCheckCnt ;			//¼���������䣬����У��

#define TOULIAO_LIST_PART					"��Ʒ��"
#define TOULIAO_LIST_PRODUCT_CNT			"������"
#define TOULIAO_LIST_DETAIL					"����"
#define TOULIAO_LIST_USER					"�ͻ�Ʒ��"
#define TOULIAO_LIST_INBOX					"�����"
#define TOULIAO_LIST_OUTBOX					"�������"
#define TOULIAO_LIST_NO						"No."
#define TOULIAO_LIST_FACTORY_NO				"������Ʒ��"


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

#define  CONTENT1_COL_OFFSET_2  SHIPPING_WOFFSET				//4 ������ʼ
#define  CONTENT2_COL_OFFSET_2  CONTENT1_COL_OFFSET_2 + 4		//7 SAP 
#define  CONTENT3_COL_OFFSET_2  CONTENT2_COL_OFFSET_2 + 7		//20 ����
#define  CONTENT4_COL_OFFSET_2  CONTENT3_COL_OFFSET_2 + 19		//4 ����
#define  CONTENT5_COL_OFFSET_2  CONTENT4_COL_OFFSET_2 + 4		//9 ��Ʒ��
#define  CONTENT6_COL_OFFSET_2  CONTENT5_COL_OFFSET_2 + 9		//7 ����


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







