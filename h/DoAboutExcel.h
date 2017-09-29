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
	
	// 操作的工作簿和工作表
	CString m_ExcelBook;
	CString m_Sheet;

	// Excel保存路径(默认在当前程序下)
	CString m_FilePath;
	// OLE操作Excel文件类对象
	CExcelHelper m_ExcelHelper;

private:
	// 判断文件是否存在
	BOOL IsFileExist(const char* m_FileName);
	// 获取当前程序路径
	CString GetCurrentAppPath();
	//获取当前工作簿和工作表名字
	void GetCurWorkSpace();
	// 获取当前时间
	void GetCurTime();
	// 设置Excel表头
	void SetExcelFormTitle();
	
public:
	// 写入Excel表格
	void ExportToExcel(CString SN, CString Date, vector<CString> NG, vector<CString> VecRms, vector<CString> VecPeck);
};

