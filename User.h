#ifndef _USER_CLASS_
#define _USER_CLASS_

#include "Order.h"

class CLayers : public CObject
{
public:
	CString strPartNo;			//��Ʒ��
	CString strOutBox;			//�����
	CString strInBox;			//�����
	CString strDetail;			//����
	CString strUserCode;		//�ͻ���
	unsigned int iQty;			//����
	CString strOrder;			//¼������
	CString strNo;				//���
	double iUnitWeight;			//��λ����  ��λΪg
	double iRadio;				//�ӳ�ϵ��
	CString strPallet;			//����
	double iUnitPrice;			//����
	CString strFactoryNo;		//������Ʒ��

	//��Q���л����݆��}�A���ֶ�
	CString strUser1;			//�A���ִ�1
	CString strUser2;			//�A���ִ�2
	CString strUser3;			//�A���ִ�3
	CString strUser4;			//�A���ִ�4
	unsigned int iUser1;		//�A������1
	unsigned int iUser2;		//�A������2
	unsigned int iUser3;		//�A������3
	unsigned int iUser4;		//�A������4

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
		//���ز����� ����ͬ��������
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
		//����������ֻ���������������� 
		//><�Ѿ������� ���������䵥
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