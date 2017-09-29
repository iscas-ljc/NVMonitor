#pragma once

#include <vector>
#include "ExcelHelper.h"


class CDoAboutExcel
{
public:
	CDoAboutExcel();
	~CDoAboutExcel();

private:
	CTime time;
	int year;
	int month;
	int day;
	
	// �����Ĺ������͹�����
	CString m_ExcelBook;
	CString m_Sheet;

	// Excel����·��(Ĭ���ڵ�ǰ������)
	CString m_FilePath;
	// OLE����Excel�ļ������
	CExcelHelper m_ExcelHelper;

private:
	// �ж��ļ��Ƿ����
	BOOL IsFileExist(const char* m_FileName);
	// ��ȡ��ǰ����·��
	CString GetCurrentAppPath();
	//��ȡ��ǰ�������͹���������
	void GetCurWorkSpace();
	// ��ȡ��ǰʱ��
	void GetCurTime();
	// ����Excel��ͷ
	void SetExcelFormTitle();
	
public:
	// д��Excel���
	void ExportToExcel(CString SN, CString Date, vector<CString> NG, vector<CString> VecRms, vector<CString> VecPeck);
};

