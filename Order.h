
/******************************************************************************
//定义并声明全局变量
*******************************************************************************/
#pragma once

#include <afxtempl.h>

#ifdef _GLOBAL_C_
  #define _GLOBALDEC__
#else
  #define _GLOBALDEC__ extern
#endif


typedef enum _SortOrderMode
{
	ORDER_NON,
	ORDER_INBOX,
	ORDER_OUTBOX,
	ORDER_DATE,
	ORDER_NO,
	ORDER_PALLET,
	ORDER_MAX
}SortOrderMode;

_GLOBALDEC__ SortOrderMode g_OrderMode; 



