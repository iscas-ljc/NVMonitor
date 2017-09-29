#include "stdafx.h"
#include "DoAboutExcel.h"


CDoAboutExcel::CDoAboutExcel()
{
	m_FilePath = GetCurrentAppPath();
}


CDoAboutExcel::~CDoAboutExcel()
{
}


BOOL CDoAboutExcel::IsFileExist(const char * szFileName)
{
	CFileStatus stat;
	if (CFile::GetStatus((LPCTSTR)szFileName, stat)) { return TRUE; }
	else { return FALSE; }
}

CString CDoAboutExcel::GetCurrentAppPath()
{
	HMODULE module = GetModuleHandle(0);
	TCHAR pFileName[MAX_PATH];
	GetModuleFileName(module, pFileName, MAX_PATH);

	CString appFullPath(pFileName);
	int nPos = appFullPath.ReverseFind('\\');
	if (nPos < 0)
	{
		return TEXT("");
	}
	else
	{
		return appFullPath.Left(nPos);
	}
}

void CDoAboutExcel::GetCurWorkSpace()
{
	// 生成使用的工作簿和工作表
	m_ExcelBook.Format(TEXT("%d-%d%s"), year, month, _T(".xlsx"));
	m_ExcelBook = m_FilePath + TEXT("\\") + m_ExcelBook;
	m_Sheet.Format(TEXT("%d-%d-%d"), year, month, day);
}


void CDoAboutExcel::SetExcelFormTitle()
{
	std::vector<CString> vecTitleText;
	vecTitleText.push_back(TEXT("NO"));
	vecTitleText.push_back(TEXT("SN"));
	vecTitleText.push_back(TEXT("DAQ Date+Time"));

	int row = 1;	// 第一行设置为标题
	int col = 1;
	CString csCellName;

	_tstring cellStart = TEXT("A1");
	_tstring cellEnd = TEXT("BM2");

	// 设置标题行字体
	long size = 15;
	m_ExcelHelper.SetCellFontFormat(cellStart, cellEnd, size);

	for (UINT i = 0; i < vecTitleText.size(); i++)
	{
		m_ExcelHelper.SetCellValue(row, i + 1, xlHAlignCenter, _variant_t(vecTitleText.at(i)));
		m_ExcelHelper.SetMergeCell(_variant_t(m_ExcelHelper.IndexToString(row, i + 1)), _variant_t(m_ExcelHelper.IndexToString(row + 1, i + 1)));
	}

	m_ExcelHelper.SetCellValue(row + 1, 4, xlHAlignCenter, _variant_t("FFT"));
	m_ExcelHelper.SetCellValue(row + 1, 5, xlHAlignCenter, _variant_t("SPL"));
	m_ExcelHelper.SetCellValue(row, 4, xlHAlignCenter, _variant_t("ON/NG State"));
	m_ExcelHelper.SetMergeCell(_variant_t(m_ExcelHelper.IndexToString(row, 4)), _variant_t(m_ExcelHelper.IndexToString(row, 5)));

	int startIndex = 6;
	for (UINT i = 1; i <= 10; i++)
	{
		csCellName.Format(TEXT("%s%d/%s%d"), _T("RMS"), i, _T("Peak"), i);
		col = startIndex + (i - 1) * 3;
		m_ExcelHelper.SetCellValue(row + 1, col, xlHAlignCenter, _variant_t("Start time"));
		m_ExcelHelper.SetCellValue(row + 1, col + 1, xlHAlignCenter, _variant_t("Stop time"));
		m_ExcelHelper.SetCellValue(row + 1, col + 2, xlHAlignCenter, _variant_t(csCellName));
		m_ExcelHelper.SetCellValue(row, col, xlHAlignCenter, _variant_t(csCellName));
		m_ExcelHelper.SetMergeCell(_variant_t(m_ExcelHelper.IndexToString(row, col)), _variant_t(m_ExcelHelper.IndexToString(row, col + 2)));
	}

	startIndex = 36;
	for (UINT i = 1; i <= 10; i++)
	{
		csCellName.Format(TEXT("%s%d"), _T("Peak"), i);
		col = startIndex + (i - 1) * 3;
		m_ExcelHelper.SetCellValue(row + 1, col, xlHAlignCenter, _variant_t("Start frequency"));
		m_ExcelHelper.SetCellValue(row + 1, col + 1, xlHAlignCenter, _variant_t("Stop frequency"));
		m_ExcelHelper.SetCellValue(row + 1, col + 2, xlHAlignCenter, _variant_t(csCellName));
		m_ExcelHelper.SetCellValue(row, col, xlHAlignCenter, _variant_t(csCellName));
		m_ExcelHelper.SetMergeCell(_variant_t(m_ExcelHelper.IndexToString(row, col)), _variant_t(m_ExcelHelper.IndexToString(row, col + 2)));
	}
}


void CDoAboutExcel::GetCurTime()
{
	time = CTime::GetCurrentTime();
	year = time.GetYear();
	month = time.GetMonth();
	day = time.GetDay();
}


void CDoAboutExcel::ExportToExcel(CString SN, CString Date, vector<CString> NG, vector<CString> VecRms, vector<CString> VecPeck)
{
	if (!m_ExcelHelper.CreateExcelApplication())
	{
		AfxMessageBox(TEXT("无法创建Excel应用！"));

		// 如果已经无法创建Excel应用，保持m_btnExportLogFile为禁用状态
		return;
	}

	GetCurTime();
	GetCurWorkSpace();

	BOOL m_IsExist = IsFileExist((const char *)m_ExcelBook.GetBuffer());
	if (m_IsExist)
	{
		m_ExcelHelper.OpenWorkBook(m_ExcelBook);
	}
	else {
		m_ExcelHelper.AddNewWorkBook();
	}

	m_ExcelHelper.OpenWorkSheet(m_Sheet);

	UINT row = m_ExcelHelper.GetUsedRowsCount();
	if (row == 1)
	{
		SetExcelFormTitle();
		row = 2;
	}
	row = row + 1;

	CString m_NO;
	m_NO.Format(TEXT("%d"), row - 2);

	std::vector<CString> vecTest;

	vecTest.push_back(m_NO);
	vecTest.push_back(SN);
	vecTest.push_back(Date);
	
	for (UINT col = 1; col <= 65; col++)
	{
		m_ExcelHelper.SetCellValue(row, col, xlHAlignCenter, _variant_t(""));
	}
	
	for (UINT col = 1; col <= 3; col++)
	{
		m_ExcelHelper.SetCellValue(row, col, xlHAlignCenter, _variant_t(vecTest.at(col - 1)));
	}

	for (UINT col = 4; col <= 5; col++)
	{
		m_ExcelHelper.SetCellValue(row, col, xlHAlignCenter, _variant_t(NG.at(col - 4)));
	}


	for (UINT col = 6; col < 6 + VecRms.size(); col++)
	{
		m_ExcelHelper.SetCellValue(row, col, xlHAlignCenter, _variant_t(VecRms.at(col - 6)));
	}

	for (UINT col = 36; col < 36 + VecPeck.size(); col++)
	{
		m_ExcelHelper.SetCellValue(row, col, xlHAlignCenter, _variant_t(VecPeck.at(col - 36)));
	}

	m_ExcelHelper.SetColumnAutoFit();
	m_ExcelHelper.SetHorizontalAlignment(xlHAlignCenter);
	m_ExcelHelper.SetVerticalAlignment(xlVAlignCenter);

	try
	{
		// 保存Excel文件
		if (m_IsExist)
		{
			m_ExcelHelper.WorkBookSave();
		}
		else {
			m_ExcelHelper.WorkBookSaveAs(m_ExcelBook.GetBuffer());
		}
	}
	catch (const std::exception&)
	{
		AfxMessageBox(TEXT("保存失败！ 请检查该文件是否已被其他程序占用！"));
	}

	m_ExcelHelper.DestroyExcelApplication();

}
