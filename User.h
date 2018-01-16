#ifndef _USER_CLASS_
#define _USER_CLASS_

#include "Order.h"

class CLayers : public CObject
{
public:
	CString strPartNo;			//部品号
	CString strOutBox;			//外箱号
	CString strInBox;			//内箱号
	CString strDetail;			//描述
	CString strUserCode;		//客户号
	unsigned int iQty;			//数量
	CString strOrder;			//录入日期
	CString strNo;				//序号
	double iUnitWeight;			//单位重量  单位为g
	double iRadio;				//加乘系数
	CString strPallet;			//托盘
	double iUnitPrice;			//单价
	CString strFactoryNo;		//工厂部品号

	//解決序列化兼容問題預留字段
	CString strUser1;			//預留字串1
	CString strUser2;			//預留字串2
	CString strUser3;			//預留字串3
	CString strUser4;			//預留字串4
	unsigned int iUser1;		//預留整形1
	unsigned int iUser2;		//預留整形2
	unsigned int iUser3;		//預留整形3
	unsigned int iUser4;		//預留整形4

protected:
        DECLARE_SERIAL(CLayers)

public:
        virtual void Serialize(CArchive& ar);
		CLayers():iQty(0),iUser1(0),iUser2(0),iUser3(0),iUser4(0){}
		void Clear(void);
		CLayers(const CLayers &test)
		{
			strPartNo = test.strPartNo;
			strOutBox = test.strOutBox;
			strInBox = test.strInBox;
			strDetail = test.strDetail;
			strUserCode = test.strUserCode;
			iQty = test.iQty;
			strOrder = test.strOrder;	
			strNo = test.strNo;
			iUnitWeight = test.iUnitWeight;
			iRadio = test.iRadio;
			strPallet = test.strPallet;
			iUnitPrice = test.iUnitPrice;
			strFactoryNo = test.strFactoryNo;

			strUser1 = test.strUser1;
			strUser2 = test.strUser2;
			strUser3 = test.strUser3;
			strUser4 = test.strUser4;
			iUser1 = test.iUser1;
			iUser2 = test.iUser2;
			iUser3 = test.iUser3;
			iUser4 = test.iUser4;

		}
		CLayers& operator=(const CLayers &test)
		{
			strPartNo = test.strPartNo;
			strOutBox = test.strOutBox;
			strInBox = test.strInBox;
			strDetail = test.strDetail;
			strUserCode = test.strUserCode;
			iQty = test.iQty;
			strOrder = test.strOrder;
			strNo = test.strNo;
			iUnitWeight = test.iUnitWeight;
			iRadio = test.iRadio;
			strPallet = test.strPallet;
			iUnitPrice = test.iUnitPrice;
			strFactoryNo = test.strFactoryNo;

			strUser1 = test.strUser1;
			strUser2 = test.strUser2;
			strUser3 = test.strUser3;
			strUser4 = test.strUser4;
			iUser1 = test.iUser1;
			iUser2 = test.iUser2;
			iUser3 = test.iUser3;
			iUser4 = test.iUser4;
			return *this;
		}

		//==================================
		//重载操作符 按不同类型排序
		bool operator<(const CLayers& std2)
		{
			
			switch(g_OrderMode)
			{
			case ORDER_INBOX:
				{
					if(strInBox.GetLength() < std2.strInBox.GetLength())
						return false;
					else if(strInBox.GetLength() > std2.strInBox.GetLength())
						return true;
					else
						return strInBox < std2.strInBox;
				}
				break;
			case ORDER_OUTBOX:
				{
					if(strOutBox.GetLength() < std2.strOutBox.GetLength())
						return false;
					else if(strOutBox.GetLength() > std2.strOutBox.GetLength())
						return true;
					else
						return strOutBox < std2.strOutBox;
				}
				break;
			case ORDER_DATE:
				return strOrder < std2.strOrder;
				break;
			case ORDER_PALLET:
				if(strPallet == std2.strPallet)
				{
					if(strOutBox.GetLength() < std2.strOutBox.GetLength())
						return false;
					else if(strOutBox.GetLength() > std2.strOutBox.GetLength())
						return true;
					else
						return strOutBox < std2.strOutBox;
				}
				else
					return strPallet < std2.strPallet;
				break;
			case ORDER_NO:
				{
					if(strNo.GetLength() < std2.strNo.GetLength())
						return false;
					else if(strNo.GetLength() > std2.strNo.GetLength())
						return true;
					else
						return strNo < std2.strNo;
				}
			default:
				return true;
				break;
			}
			
			
		}

		bool operator>(const CLayers& std2)
		{
			switch(g_OrderMode)
			{
			case ORDER_INBOX:
				{
					if(strInBox.GetLength() < std2.strInBox.GetLength())
						return false;
					else if(strInBox.GetLength() > std2.strInBox.GetLength())
						return true;
					else
						return strInBox > std2.strInBox;
				}
				break;
			case ORDER_OUTBOX:
				{
					if(strOutBox.GetLength() < std2.strOutBox.GetLength())
						return false;
					else if(strOutBox.GetLength() > std2.strOutBox.GetLength())
						return true;
					else
						return strOutBox > std2.strOutBox;
				}
				break;
			case ORDER_DATE:
				return strOrder > std2.strOrder;
				break;
			case ORDER_PALLET:
				if(strPallet == std2.strPallet)
				{
					if(strOutBox.GetLength() < std2.strOutBox.GetLength())
						return false;
					else if(strOutBox.GetLength() > std2.strOutBox.GetLength())
						return true;
					else
						return strOutBox > std2.strOutBox;
				}
				else
					return strPallet > std2.strPallet;
				break;
			case ORDER_NO:
				{
					if(strNo.GetLength() < std2.strNo.GetLength())
						return false;
					else if(strNo.GetLength() > std2.strNo.GetLength())
						return true;
					else
						return strNo > std2.strNo;
				}
			default:
				return true;
				break;
			}
		}
		//==================================


		//===================================
		//按日期排序只能重载其它操作符 
		//><已经被重载 用于生产箱单
		bool operator+(const CLayers& std2)
		{
			return strOrder < std2.strOrder;
		}

		bool operator-(const CLayers& std2)
		{
			return strOrder > std2.strOrder;
		}
		//====================================

};
#endif