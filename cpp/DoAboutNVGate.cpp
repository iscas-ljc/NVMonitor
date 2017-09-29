#include "stdafx.h"
#include "DoAboutNVGate.h"
#include  <fstream>
#include <iostream>

#pragma warning(disable : 4996)

using namespace std;
SOCKET hSockClient1;
int CDoAboutNVGate::nHead;

CDoAboutNVGate::CDoAboutNVGate(void)
{
	bDataFromNVGATE = NULL;
	m_nCour1 = 0;
	nHead = 0;
	NVGATEhead[10] = 0;
}

CDoAboutNVGate::~CDoAboutNVGate(void)
{
	;
}

BOOL CDoAboutNVGate::ConnectNVGate(const CString addr)
{
	WSADATA ws = { 0 };
	WSAStartup(MAKEWORD(2, 2), &ws);
	hSockClient1 = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);

	if (hSockClient1 == INVALID_SOCKET)
	{
		return FALSE;
	}

	SOCKADDR_IN addrSvr = { 0 };
	addrSvr.sin_family = AF_INET;
	addrSvr.sin_port = htons(3000);

	CStringA temp(addr);
	const char *cAddr = temp.GetString();
	addrSvr.sin_addr.S_un.S_addr = inet_addr(cAddr);

	if (SOCKET_ERROR == connect(hSockClient1, (SOCKADDR*)&addrSvr, sizeof(addrSvr)))
	{
		int nErrorCode = WSAGetLastError();
		CString str;
		str.Format(_T("ErrorNumber:%d"), nErrorCode);
		CString strError = _T("ConnectNVGateError!");
		strError += str;
		MessageBoxExW(NULL, strError, NULL, MB_OK, MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US));
		return FALSE;
	}
	return TRUE;
}

int CDoAboutNVGate::SendDataToNVGATE(CString str_data)
{
	if (str_data[str_data.GetLength() - 1] != '\n')
	{
		str_data += "\n";
	}

	CStringA temp(str_data);
	const char* cAddr = temp.GetString();

	int n = send(hSockClient1, cAddr, strlen(cAddr), 0);

	if (SOCKET_ERROR == n)
	{
		int ErrorCode = 0;
		ErrorCode = WSAGetLastError();
		TRACE("Send Error. ErrorCode = %d.", ErrorCode);
		// 		closesocket(hSockClient1);
		// 		WSACleanup();
	}

	return GetDataFromNVGATE(NULL);
}

int CDoAboutNVGate::GetDataFromNVGATE(char *str_data)
{
	int n = recv(hSockClient1, NVGATEhead, 10, 0);

	if (SOCKET_ERROR == n)
	{
		int ErrorCode = 0;
		ErrorCode = WSAGetLastError();
		TRACE("Receive Error. ErrorCode = %d.", ErrorCode);
		return -2;
	}

	if (NVGATEhead[0] == '0')
	{
		nHead = atoi(NVGATEhead);

		if (bDataFromNVGATE != NULL)
		{
			delete[]bDataFromNVGATE;
			bDataFromNVGATE = NULL;
		}

		if (nHead)
		{
			bDataFromNVGATE = new char[nHead];
		}
	}
	else
	{
		CString strError = _T("GetErrorMessage ");
		strError += NVGATEhead;
		return -1;
	}

	if (nHead>0)
	{
		m_nCour1 = 0;
		int total = 0;
		while (total != nHead)
		{
			int ret = recv(hSockClient1, (char*)bDataFromNVGATE + total, nHead - total, 0);
			total += ret;
		}
	}

	return 0;
}

int CDoAboutNVGate::GetSettingValue(CString settings)
{
	return SendDataToNVGATE(_T("GetSettingValue ") + settings);
}

void CDoAboutNVGate::MakeString(char cEnfOfString, CString & szString)
{
	szString = bDataFromNVGATE + m_nCour1;
	m_nCour1 += szString.GetLength() + 1;
	return;
}

void CDoAboutNVGate::MakeInteger(char* bDataFromNVGATE, CString& SettingValue)
{
	short *a = (short*)(bDataFromNVGATE + m_nCour1);
	SettingValue.Format(_T("%d"), *a);
}

void CDoAboutNVGate::MakeFloat(char* bDataFromNVGATE, CString& SettingValue)
{
	float *a = (float*)(bDataFromNVGATE + m_nCour1);

	if (*a >= 1)
	{ //之前为直接强转为int的代码，会把小数丢失
		int b = *a;
		SettingValue.Format(_T("%d"), b);
	}
	else if (*a < 1)
	{//为了维护原来的代码，此处
		SettingValue.Format(_T("%f"), *a);
	}

	return;
}
