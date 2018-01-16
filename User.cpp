

#include "stdafx.h"
#include "User.h"

IMPLEMENT_SERIAL(CLayers, CObject, 1)

void CLayers::Clear(void)
{	
	strPartNo.Empty();
	strOutBox.Empty();
	iQty = 0;
	strOrder.Empty();
	iUnitWeight = 0;
	iRadio = 0;
	strPallet.Empty();
	iUnitPrice = 0;
	strFactoryNo.Empty();

	strUser1.Empty();
	strUser2.Empty();
	strUser3.Empty();
	strUser4.Empty();
	iUser1 = 0;
	iUser2 = 0;
	iUser3 = 0;
	iUser4 = 0;
}

void CLayers::Serialize(CArchive& ar)
{
	//CObject::Serialize(ar);//save is here
	if (ar.IsStoring())
	{
		ar <<strPartNo<<strOutBox<<strInBox<<strDetail<<strUserCode<<iQty<<strOrder<<strNo<<iUnitWeight<<iRadio<<strPallet<<iUnitPrice<<strFactoryNo<<strUser1<<strUser2<<strUser3<<strUser4<<iUser1<<iUser2<<iUser3<<iUser4;
	}
	else
	{
		ar >>strPartNo>>strOutBox>>strInBox>>strDetail>>strUserCode>>iQty>>strOrder>>strNo>>iUnitWeight>>iRadio>>strPallet>>iUnitPrice>>strFactoryNo>>strUser1>>strUser2>>strUser3>>strUser4>>iUser1>>iUser2>>iUser3>>iUser4;
	}
}
