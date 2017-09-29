#include "stdafx.h"
#include "ExcelHelper.h"


CExcelHelper::CExcelHelper()
{
	m_strExcelFile = _T("");
	m_strWorkSheet = _T("");
	CoInitialize(NULL);
}


CExcelHelper::~CExcelHelper()
{
	CloseWorkBook();
	DestroyExcelApplication();
	// 关闭COM库
	CoUninitialize();
}


////////////////////////////////////////////////////////////////////////  
///Function:    CreateExcelApplication  
///Description: 创建excel应用服务   
////////////////////////////////////////////////////////////////////////
BOOL CExcelHelper::CreateExcelApplication()
{
	CoInitialize(NULL);
	//创建Excel服务
	if (!m_oExcelApp.CreateDispatch(TEXT("Excel.Application"), NULL))
	{
		return FALSE;
	}
	COleVariant covOptional(DISP_E_PARAMNOTFOUND, VT_ERROR);

	// 打开工作薄
	m_oWorkBooks.AttachDispatch(m_oExcelApp.get_Workbooks(), TRUE);
	m_oWorkBook.AttachDispatch(m_oWorkBooks.Add(covOptional), TRUE);

	

	return TRUE;
}


////////////////////////////////////////////////////////////////////////  
///Function:    DestroyExcelApplication  
///Description: 销毁excel服务   
////////////////////////////////////////////////////////////////////////
void CExcelHelper::DestroyExcelApplication()
{
	// 释放对象
	m_oCurrRange.ReleaseDispatch();
	m_oWorkSheet.ReleaseDispatch();
	m_oWorkSheets.ReleaseDispatch();
	m_oWorkBook.ReleaseDispatch();
	m_oWorkBooks.ReleaseDispatch();

	// Quit必须在m_ExlApp释放之前，否则程序结束后还会有一个Excel进程驻留在内存中，而且程序重复运行的时候会出错
	m_oExcelApp.Quit();
	m_oExcelApp.ReleaseDispatch();


}


////////////////////////////////////////////////////////////////////////  
///Function:    GetCellValue  
///Description: 得到的单元格中的值  
///Call:        IndexToString() 从(x,y)坐标形式转化为“A1”格式字符串  
///Input:       int row 单元格所在行  
///             int col 单元格所在列  
///Return:      CString 单元格中的值  
////////////////////////////////////////////////////////////////////////
CString CExcelHelper::GetCellValue(int row, int col)
{
	m_oCurrRange = m_oWorkSheet.get_Range(COleVariant(IndexToString(row, col)), COleVariant(IndexToString(row, col)));
	COleVariant rValue;
	rValue = COleVariant(m_oCurrRange.get_Value2());
	rValue.ChangeType(VT_BSTR);
	return CString(rValue.bstrVal);
}


////////////////////////////////////////////////////////////////////////  
///Function:    SetCellValue  
///Description: 修改单元格内的值  
///Call:        IndexToString() 从(x,y)坐标形式转化为“A1”格式字符串  
///Input:       int row 单元格所在行  
///             int col 单元格所在列  
///             int Align       对齐方式默认为居中  
//////////////////////////////////////////////////////////////////////// 
void CExcelHelper::SetCellValue(int row, int col, int Align, VARIANT & value)
{
	m_oCurrRange = m_oWorkSheet.get_Range(COleVariant(IndexToString(row, col)), COleVariant(IndexToString(row, col)));
	m_oCurrRange.put_Value2(value);
	m_oCurrRange.AttachDispatch((m_oCurrRange.get_Item(COleVariant(long(1)), COleVariant(long(1)))).pdispVal);
	m_oCurrRange.put_HorizontalAlignment(COleVariant((short)Align));

	//设置外边框
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	//LineStyle=线型 Weight=线宽 ColorIndex=线的颜色(-4105为自动)  
	m_oCurrRange.BorderAround(COleVariant((long)xlContinuous), xlThin, xlColorIndexAutomatic
		, COleVariant(long(0x000000)), covOptional);
}


////////////////////////////////////////////////////////////////////////  
///Function:    SetCellText  
///Description: 设置单元格文本  
///Input:       _tstring cellPos 单元格所在位置 
///Input:       _tstring cellText 设置的文本
//////////////////////////////////////////////////////////////////////// 
void CExcelHelper::SetCellText(_tstring cellPos, _tstring cellText)
{
	//获取单元格区域
	m_oCurrRange.AttachDispatch(m_oWorkSheet.get_Range(COleVariant(cellPos.c_str()),
		COleVariant(cellPos.c_str())));

	m_oCurrRange.put_Value2(COleVariant(cellText.c_str()));

	//设置外边框
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	//LineStyle=线型 Weight=线宽 ColorIndex=线的颜色(-4105为自动)  
	m_oCurrRange.BorderAround(COleVariant((long)xlContinuous), xlThin, xlColorIndexAutomatic
		, COleVariant(long(0x000000)), covOptional);
}


////////////////////////////////////////////////////////////////////////  
///Function:    SetMergeCell  
///Description: 合并单元格  
///Input:       VARIANT &cellStart, VARIANT &cellEnd 单元格所在起始位置和结束位置  
//////////////////////////////////////////////////////////////////////// 
void CExcelHelper::SetMergeCell(VARIANT &cellStart, VARIANT &cellEnd)
{
	//加载要合并的单元格  
	m_oCurrRange = m_oWorkSheet.get_Range(cellStart, cellEnd);

	//合并单元格
	m_oCurrRange.Merge(_variant_t((long)0));

	//设置外边框
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	//LineStyle=线型 Weight=线宽 ColorIndex=线的颜色(-4105为自动)  
	m_oCurrRange.BorderAround(COleVariant((long)xlContinuous), xlThin, xlColorIndexAutomatic
		, COleVariant(long(0x000000)), covOptional);

}


////////////////////////////////////////////////////////////////////////  
///Function:    SetCellBackgroundColor  
///Description: 设置单元格背景颜色  
///Input:       long color 线条颜色
////////////////////////////////////////////////////////////////////////
void CExcelHelper::SetCellBackgroundColor(_tstring cellStart, _tstring cellEnd, long color)
{
	Cnterior oInterior;

	//获取单元格区域
	m_oCurrRange.AttachDispatch(m_oWorkSheet.get_Range(COleVariant(cellStart.c_str()),
		COleVariant(cellEnd.c_str())));

	//设置单元格背景色
	oInterior.AttachDispatch(m_oCurrRange.get_Interior(), TRUE);
	oInterior.put_Color(COleVariant(color));

	oInterior.ReleaseDispatch();
}


////////////////////////////////////////////////////////////////////////  
///Function:    SetCellFontFormat  
///Description: 设置单元格字体  
///Input:       _tstring cellStart, _tstring cellEnd 设置的单元格始末位置
///Input:       long size 字体大小
///Input:       bool bold 线条粗细
///Input:       long color 线条颜色
////////////////////////////////////////////////////////////////////////
void CExcelHelper::SetCellFontFormat(_tstring cellStart, _tstring cellEnd
	, long size, bool bold, long color)
{
	CFont0 oFont;

	//获取单元格区域
	m_oCurrRange.AttachDispatch(m_oWorkSheet.get_Range(COleVariant(cellStart.c_str()),
		COleVariant(cellEnd.c_str())));

	oFont.AttachDispatch(m_oCurrRange.get_Font());
	// 设置字体是否粗体
	oFont.put_Bold(COleVariant((short)bold));
	// 设置字体大小
	oFont.put_Size(COleVariant(size));
	// 设置字体颜色
	oFont.put_Color(COleVariant(color));

	oFont.ReleaseDispatch();
}


////////////////////////////////////////////////////////////////////////  
///Function:    SetCellBorderAround  
///Description: 设置边框  
///Input:       _tstring cellStart, _tstring cellEnd 设置的单元格始末位置
///Input:       XlLineStyle lineStyle 线条样式
///Input:       XlBorderWeight borderWeight 线条粗细
///Input:       long color 线条颜色
////////////////////////////////////////////////////////////////////////
void CExcelHelper::SetCellBorderAround(_tstring cellStart, _tstring cellEnd, XlLineStyle lineStyle
	, XlBorderWeight borderWeight, long color)
{
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	//获取单元格区域
	m_oCurrRange.AttachDispatch(m_oWorkSheet.get_Range(COleVariant(cellStart.c_str()),
		COleVariant(cellEnd.c_str())));

	//设置外边框
	m_oCurrRange.BorderAround(COleVariant((long)lineStyle), borderWeight, xlColorIndexAutomatic
		, COleVariant(color), covOptional);
}


////////////////////////////////////////////////////////////////////////  
///Function:    SetColumnAutoFit  
///Description: 自动适应列宽  
////////////////////////////////////////////////////////////////////////
void CExcelHelper::SetColumnAutoFit()
{
	// 获得使用的区域Range
	m_oCurrRange.AttachDispatch(m_oWorkSheet.get_UsedRange(), TRUE);

	m_oCurrRange = m_oCurrRange.get_EntireColumn();
	m_oCurrRange.AutoFit();
}


////////////////////////////////////////////////////////////////////////  
///Function:    SetHorizontalAlignment  
///Description: 设置水平对齐  
///Input:       XlVAlign vAlign 对齐方式  
//////////////////////////////////////////////////////////////////////// 
void CExcelHelper::SetHorizontalAlignment(XlHAlign hAlign)
{
	// 获得使用的区域Range
	m_oCurrRange.AttachDispatch(m_oWorkSheet.get_UsedRange(), TRUE);

	// 设置水平对齐
	m_oCurrRange.put_HorizontalAlignment(COleVariant((short)hAlign));
}


////////////////////////////////////////////////////////////////////////  
///Function:    SetVerticalAlignment  
///Description: 设置垂直对齐  
///Input:       XlVAlign vAlign 对齐方式  
//////////////////////////////////////////////////////////////////////// 
void CExcelHelper::SetVerticalAlignment(XlVAlign vAlign)
{
	// 获得使用的区域Range
	m_oCurrRange.AttachDispatch(m_oWorkSheet.get_UsedRange(), TRUE);

	// 设置垂直对齐
	m_oCurrRange.put_VerticalAlignment(COleVariant((short)vAlign));
}


////////////////////////////////////////////////////////////////////////  
///Function:    SetRowHeight  
///Description: 设置行高  
///Call:        IndexToString() 从(x,y)坐标形式转化为“A1”格式字符串  
///Input:       int row 单元格所在行  
//////////////////////////////////////////////////////////////////////// 
void CExcelHelper::SetRowHeight(int row, CString height)
{
	int col = 1;
	m_oCurrRange = m_oWorkSheet.get_Range(COleVariant(IndexToString(row, col)), COleVariant(IndexToString(row, col)));
	m_oCurrRange.put_RowHeight(COleVariant(height));
}


////////////////////////////////////////////////////////////////////////  
///Function:    GetRowHeight  
///Description: 设置行高  
///Call:        IndexToString() 从(x,y)坐标形式转化为“A1”格式字符串  
///Input:       int row 要设置行高的行  
///             CString 宽值  
//////////////////////////////////////////////////////////////////////// 
CString CExcelHelper::GetRowHeight(int row)
{
	int col = 1;
	m_oCurrRange = m_oWorkSheet.get_Range(COleVariant(IndexToString(row, col)), COleVariant(IndexToString(row, col)));
	VARIANT height = m_oCurrRange.get_RowHeight();
	CString strheight;
	strheight.Format(CString((LPCSTR)(_bstr_t)(_variant_t)height));
	return strheight;
}

////////////////////////////////////////////////////////////////////////  
///Function:    SetColumnWidth  
///Description: 设置列宽  
///Call:        IndexToString() 从(x,y)坐标形式转化为“A1”格式字符串  
///Input:       int col 要设置列宽的列  
///             CString 宽值  
////////////////////////////////////////////////////////////////////////  
void CExcelHelper::SetColumnWidth(int col, CString width)
{
	int row = 1;
	m_oCurrRange = m_oWorkSheet.get_Range(COleVariant(IndexToString(row, col)), COleVariant(IndexToString(row, col)));
	m_oCurrRange.put_ColumnWidth(COleVariant(width));
}


////////////////////////////////////////////////////////////////////////  
///Function:    GetColumnWidth  
///Description: 得到列宽 
///Call:        IndexToString() 从(x,y)坐标形式转化为“A1”格式字符串  
///Input:       int col 单元格所在列 
//////////////////////////////////////////////////////////////////////// 
CString CExcelHelper::GetColumnWidth(int col)
{
	int row = 1;
	m_oCurrRange = m_oWorkSheet.get_Range(COleVariant(IndexToString(row, col)), COleVariant(IndexToString(row, col)));
	VARIANT width = m_oCurrRange.get_ColumnWidth();
	CString strwidth;
	strwidth.Format(CString((LPCSTR)(_bstr_t)(_variant_t)width));
	return strwidth;
}


////////////////////////////////////////////////////////////////////////  
///Function:    IndexToString  
///Description: 得到的单元格在EXCEL中的定位名称字符串  
///Input:       int row 单元格所在行  
///             int col 单元格所在列  
///Return:      CString 单元格在EXCEL中的定位名称字符串 
////////////////////////////////////////////////////////////////////////  
CString CExcelHelper::IndexToString(int row, int col)
{
	CString strResult;
	if (col > 26)
	{
		strResult.Format(_T("%c%c%d"), 'A' + (col - 1) / 26 - 1, 'A' + (col - 1) % 26, row);
	}
	else
	{
		strResult.Format(_T("%c%d"), 'A' + (col - 1) % 26, row);
	}
	return strResult;
}


////////////////////////////////////////////////////////////////////////  
///Function:    LastLineIndex  
///Description: 得到表格总第一个空行的索引  
///Return:      int 空行的索引号  
////////////////////////////////////////////////////////////////////////  
int CExcelHelper::LastLineIndex()
{
	int i, j, flag = 0;
	CString str;
	for (i = 1;; i++)
	{
		flag = 0;
		//粗略统计，认为前列都没有数据即为空行  
		for (j = 1; j <= 5; j++)
		{
			str.Format(_T("%s"), this->GetCellValue(i, j).Trim());
			if (str.Compare(_T("")) != 0)
			{
				flag = 1;
				break;
			}
		}
		if (flag == 0)
			return i;
	}
}


BOOL CExcelHelper::OpenWorkBook(CString strExcelFile)
{
	LPDISPATCH lpDisp = NULL;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	CloseWorkBook();
	m_oWorkBooks.AttachDispatch(m_oExcelApp.get_Workbooks(), TRUE); //没有这条语句，下面打开文件返回失败。
																	// 打开文件
	lpDisp = m_oWorkBooks.Open(strExcelFile,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covOptional);
	// 获得活动的WorkBook( 工作簿 )
	
	m_oWorkBook.AttachDispatch(lpDisp, TRUE);

	// 获得活动的WorkSheet( 工作表 )
	// m_oWorkSheet.AttachDispatch(m_oWorkBook.GetActiveSheet(), TRUE);

	m_strExcelFile = strExcelFile;
	return TRUE;
}

BOOL CExcelHelper::AddNewWorkBook()
{
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	//获得所有工作表  
	m_oWorkBooks.AttachDispatch(m_oExcelApp.get_Workbooks());
	//新建工作表，必须指出的是valTemp必须是赋值之后的！！！！，否则不能通过  
	m_oWorkBook.AttachDispatch(m_oWorkBooks.Add(covOptional));
	return TRUE;
}




BOOL CExcelHelper::OpenWorkSheet(CString strWorkSheet)
{
	CString strMsg;
	long i;

	CloseWorkSheet();

	// 打开工作表集，查找工作表strSheet
	m_oWorkSheets.AttachDispatch(m_oWorkBook.get_Worksheets(), TRUE);
	long lSheetNUM = m_oWorkSheets.get_Count();
	for (i = 1; i <= lSheetNUM; i++)
	{
		m_oWorkSheet.AttachDispatch(m_oWorkSheets.get_Item(COleVariant((short)i)), TRUE);
		if (strWorkSheet == m_oWorkSheet.get_Name())
			break;
	}

	if (i > lSheetNUM)
	{
		//strMsg.Format(_T("在%s中找不到名为%s的表"), m_strExcelFile, strWorkSheet);
		//WarningMessageBox(strMsg);
		COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
		m_oWorkSheet = m_oWorkSheets.Add(covOptional, covOptional, COleVariant((short)1), covOptional);
		m_oWorkSheet.put_Name(strWorkSheet);
	}
#ifdef _DEBUG
	else
	{
		strMsg.Format(_T("%s是%s第%d个表"), strWorkSheet, m_strExcelFile, i);
		//DebugMessageBox(strMsg);
	}
#endif

	m_strWorkSheet = strWorkSheet;
	return TRUE;
}



BOOL CExcelHelper::OpenWorkSheet(CString strExcelFile, CString strWorkSheet)
{
	if (!OpenWorkBook(strExcelFile))
		return FALSE;

	return OpenWorkSheet(strWorkSheet);
}


void CExcelHelper::CloseWorkBook()
{
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	if (!m_strExcelFile.IsEmpty())
	{
		// 关闭
		m_oWorkBook.put_Saved(TRUE); // 将Workbook的保存状态设置为已保存，即不让系统提示是否人工保存
		m_oWorkBook.Close(covOptional, COleVariant(m_strExcelFile), covOptional);
		m_oWorkBooks.Close();
		// 释放
		m_oWorkBook.ReleaseDispatch();
		m_oWorkBooks.ReleaseDispatch();
	}

	m_strExcelFile = _T("");

	CloseWorkSheet();
}


void CExcelHelper::CloseWorkSheet()
{
	if (!m_strWorkSheet.IsEmpty())
	{
		m_oCurrRange.ReleaseDispatch();
		m_oWorkSheet.ReleaseDispatch();
		m_oWorkSheets.ReleaseDispatch();
	}

	m_strWorkSheet = _T("");
}


void CExcelHelper::WorkBookSave()
{
	m_oExcelApp.put_AlertBeforeOverwriting(false);
	m_oExcelApp.put_DisplayAlerts(false);
	m_oWorkBook.Save();
	m_oWorkBook.put_Saved(TRUE); // 将Workbook的保存状态设置为已保存
}


BOOL CExcelHelper::WorkBookSaveAs(CString strSaveAsFile, XlFileFormat fileFormat)
{
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	CString strMsg;

#ifdef _DEBUG
	strMsg.Format(_T("excel %s save as %s"),
		m_strExcelFile, strSaveAsFile);
	DebugMessageBox(strMsg);
#endif

	// 删除strSaveAsFile文件，以免excel另存为strSaveAsFile，磁盘中已存在同名文件
	::DeleteFile(strSaveAsFile);

	// excel另存为strSaveAsFile
	m_oWorkBook.SaveAs(COleVariant(strSaveAsFile), COleVariant((short)fileFormat), covOptional,
		covOptional, covOptional, covOptional, 0,
		covOptional, covOptional, covOptional, covOptional, covOptional);

	return TRUE;
}


void CExcelHelper::WorkSheetSaveAs(CString strSaveAsFile, XlFileFormat fileFormat)
{
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	CString strMsg;

#ifdef _DEBUG
	strMsg.Format(_T("excel %s sheet %s save as %s"),
		m_strExcelFile, m_oWorkSheet.get_Name(), strSaveAsFile);
	DebugMessageBox(strMsg);
#endif

	// 删除strSaveAsFile文件，以免excel另存为strSaveAsFile文件，磁盘中已存在同名文件
	::DeleteFile(strSaveAsFile);

	// excel另存为strSaveAsFile
	m_oWorkSheet.SaveAs(strSaveAsFile, COleVariant((short)fileFormat), covOptional,
		covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional);
}


long CExcelHelper::GetUsedRowsCount()
{
	// 获得使用的区域Range( 区域 )
	m_oCurrRange.AttachDispatch(m_oWorkSheet.get_UsedRange(), TRUE);

	// 获得使用的行数
	m_oCurrRange.AttachDispatch(m_oCurrRange.get_Rows(), TRUE);
	return m_oCurrRange.get_Count();
}


long CExcelHelper::GetUsedColumnsCount()
{
	// 获得使用的区域Range( 区域 )
	m_oCurrRange.AttachDispatch(m_oWorkSheet.get_UsedRange(), TRUE);

	// 获得使用的列数
	m_oCurrRange.AttachDispatch(m_oCurrRange.get_Columns(), TRUE);
	return m_oCurrRange.get_Count();
}


BOOL CExcelHelper::GetRowString(UINT nRow, CStringArray & array)
{
	CRange oCurCell;
	CString strMsg;

	// 获得使用的列数
	long lUsedColumnNum = GetUsedColumnsCount();

	COleVariant covRow((long)nRow);
	CString strHeader;

	// 遍历列头
#ifdef _DEBUG
	strMsg.Format(_T("row %d:"), nRow);
#endif
	m_oCurrRange.AttachDispatch(m_oWorkSheet.get_UsedRange(), TRUE);
	for (long i = 1; i <= lUsedColumnNum; i++)
	{
		if (0 != nRow)
		{
			//获取单元格数据
			oCurCell.AttachDispatch(m_oCurrRange.get_Item(covRow, COleVariant(i)).pdispVal, TRUE);
			strHeader = (oCurCell.get_Text()).bstrVal;
		}

		//保存单元格数据
		array.Add(strHeader);
#ifdef _DEBUG
		strMsg += _T(" ") + strHeader;
#endif
	}

#ifdef _DEBUG    
	strHeader.Format(_T("size%d"), array.GetSize());
	strMsg += _T(" ") + strHeader;
	DebugMessageBox(strMsg);
#endif

	return TRUE;
}


void CExcelHelper::InsertRow(UINT nRow, CStringArray & array)
{
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	long size = array.GetSize();
	CString strCell, strData;
	char cColumn = 'A';
	CFont0 oFont;

	if (nRow < 1)
	{
		WarningMessageBox(_T("%s: 行号必须大于0"));
		return;
	}

#ifdef _DEBUG
	CString strMsg;
	strMsg.Format(_T("row %d:"), nRow);
#endif

	for (long i = 1; i <= size; i++)
	{
		//获取每1列第1个单元格
		strCell.Format(TEXT("%c%d"), cColumn++, nRow);
		m_oCurrRange.AttachDispatch(m_oWorkSheet.get_Range(COleVariant(strCell),
			covOptional), TRUE);

		//当前位置插入单元格
		m_oCurrRange.Insert(COleVariant((short)xlShiftDown), covOptional);
		//设置单元格内容
		m_oCurrRange.AttachDispatch(m_oWorkSheet.get_Range(COleVariant(strCell),
			covOptional), TRUE);
		strData = array.GetAt(i - 1);
		m_oCurrRange.put_Value2(COleVariant(strData));
#ifdef _DEBUG
		strMsg += _T(" ") + strData;
#endif
		//设置单元格字体加粗
		oFont.AttachDispatch(m_oCurrRange.get_Font(), TRUE);
		oFont.put_Bold(COleVariant((short)TRUE));

		//自动列宽
		m_oCurrRange.AttachDispatch(m_oCurrRange.get_EntireColumn(), TRUE);
		m_oCurrRange.AutoFit();
	}

#ifdef _DEBUG
	DebugMessageBox(strMsg);
#endif

	oFont.ReleaseDispatch();
}


void CExcelHelper::SetRangeBackgroundColor(CString strStart, CString strEnd, long nColorIndex)
{
	Cnterior oInterior;

	//获取单元格区域
	m_oCurrRange.AttachDispatch(m_oWorkSheet.get_Range(COleVariant(strStart),
		COleVariant(strEnd)), TRUE);

	//设置单元格背景色
	oInterior.AttachDispatch(m_oCurrRange.get_Interior(), TRUE);
	oInterior.put_ColorIndex(COleVariant(nColorIndex));

	oInterior.ReleaseDispatch();
}


void CExcelHelper::SetRangeBorders(CString strStart, CString strEnd, XlLineStyle lineStyle, XlBorderWeight borderWeight)
{
	m_oCurrRange.AttachDispatch(m_oWorkSheet.get_Range(COleVariant(strStart),
		COleVariant(strEnd)), TRUE);

	// 设置区域内所有单元格的边框 
	CBorders oBorders;
	oBorders.AttachDispatch(m_oCurrRange.get_Borders(), TRUE);
	oBorders.put_ColorIndex(COleVariant((long)1));            // 线的颜色(-4105为自动, 1为黑色) 
	oBorders.put_LineStyle(COleVariant((long)lineStyle));
	oBorders.put_Weight(COleVariant((long)borderWeight));
	oBorders.ReleaseDispatch();
}


void CExcelHelper::SetRangeBorderAround(CString strStart, CString strEnd, XlLineStyle lineStyle, XlBorderWeight borderWeight)
{
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	m_oCurrRange.AttachDispatch(m_oWorkSheet.get_Range(COleVariant(strStart),
		COleVariant(strEnd)), TRUE);

	//设置外边框
	//线的颜色(-4105为自动, 1为黑色) 
	m_oCurrRange.BorderAround(COleVariant((long)lineStyle), borderWeight, -4105, COleVariant(), covOptional);
}


void CExcelHelper::SetAlignment(CString strStart, CString strEnd, XlHAlign hAlign, XlVAlign vAlign)
{
	// 获得使用的区域Range( 区域 )
	m_oCurrRange.AttachDispatch(m_oWorkSheet.get_Range(COleVariant(strStart),
		COleVariant(strEnd)), TRUE);
	// 设置水平对齐
	m_oCurrRange.put_HorizontalAlignment(COleVariant((short)hAlign));
	// 设置垂直对齐
	m_oCurrRange.put_VerticalAlignment(COleVariant((short)vAlign));
}


////////////////////////////////////////////////////////////////////////  
///Function:    SaveAsExcel  
///Description: 保存excel文件 
///input        deleteFile 删除同名函数
//////////////////////////////////////////////////////////////////////// 
BOOL CExcelHelper::WorkBookSaveAs(_tstring excelFileParh, bool deleteFile)
{
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	m_oExcelApp.put_AlertBeforeOverwriting(false);
	m_oExcelApp.put_DisplayAlerts(false);

	if (deleteFile)
	{
		// 删除excelFileParh文件，以免excel另存为excelFileParh，磁盘中已存在同名文件
		::DeleteFile(excelFileParh.c_str());
	}
	
	m_oWorkBook.SaveAs(COleVariant(excelFileParh.c_str()), covOptional,
		covOptional, covOptional,
		covOptional, covOptional, (long)0,
		covOptional, covOptional, covOptional,
		covOptional, covOptional);

	return TRUE;
}