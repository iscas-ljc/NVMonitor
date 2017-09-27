#pragma once
class CDoAboutNVGate
{
public:
	CDoAboutNVGate();
	~CDoAboutNVGate();
	static BOOL ConnectNVGate(const CString addr);
	int SendDataToNVGATE(CString str_data);
	int GetDataFromNVGATE(char *str_data);
	int GetSettingValue(CString settings);
	void MakeString(char cEnfOfString, CString & szString);
	void MakeInteger(char *bDataFromNVGATE, CString& SettingValue);
	void MakeFloat(char *bDataFromNVGATE, CString& SettingValue);
	void CharToStr(char* buf, int len, CString &str);

public:
	char * bDataFromNVGATE;
	char NVGATEhead[10];
	static int nHead;
	int m_nCour1;

};

