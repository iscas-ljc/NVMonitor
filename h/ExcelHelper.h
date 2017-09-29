#pragma once


#include "CApplication.h"
#include "CFont0.h"
#include "CRange.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"
#include "Cnterior.h"
#include "CBorders.h"

#include <string>

using namespace std;

#ifdef _UNICODE
#define _tstring std::wstring
#else
#define _tstring std::string
#endif



#define InfoMessageBox(strMsg)	\
AfxMessageBox(strMsg, MB_OK | MB_ICONINFORMATION)
#define DebugMessageBox(strMsg)	\
AfxMessageBox(strMsg, MB_OK | MB_ICONINFORMATION)
#define WarningMessageBox(strMsg)	\
AfxMessageBox(strMsg, MB_OK | MB_ICONWARNING)
#define ErrorMessageBox(strMsg)	\
AfxMessageBox(strMsg, MB_OK | MB_ICONERROR)


//Excel文件格式
//http://msdn.microsoft.com/zh-cn/library/office/ff198017.aspx
typedef enum
{
	xlCSV = 6,				   //csv
	xlTextWindows = 20,        //Windows 文本
	xlTextMSDOS = 21,          //MSDOS 文本
	xlUnicodeText = 42,        //Unicode 文本
	xlExcel9795 = 43,          //Excel9795
	xlWorkbookNormal = -4143,  //常规工作簿
	xlExcel12 = 50,            //Excel 12
	xlWorkbookDefault = 51,    //默认工作簿
} XlFileFormat;


//水平对齐方式
//http://msdn.microsoft.com/zh-cn/library/office/ff840772.aspx
typedef enum
{
	xlHAlignCenter = -4108,					//居中对齐
	xlHAlignCenterAcrossSelection = 7,		//跨列居中
	xlHAlignDistributed = -4117,			//分散对齐
	xlHAlignFill = 5,						//填充
	xlHAlignGeneral = 1,					//按数据类型对齐
	xlHAlignJustify = -4130,				//两端对齐
	xlHAlignLeft = -4131,					//左对齐
	xlHAlignRight = -4152,					//右对齐
} XlHAlign;

//垂直对齐方式
//http://msdn.microsoft.com/zh-cn/library/office/ff835305.aspx
typedef enum
{
	xlVAlignBottom = -4107,				//靠下 
	xlVAlignCenter = -4108,				//居中对齐 
	xlVAlignDistributed = -4117,        //分散对齐 
	xlVAlignJustify = -4130,			//两端对齐 
	xlVAlignTop = -4160,				//靠上 
} XlVAlign;

//插入时单元格的移动方向
//http://msdn.microsoft.com/zh-cn/library/office/ff837618.aspx
typedef enum
{
	xlShiftDown = -4121,        //向下移动单元格
	xlShiftToRight = -4161,        //向右移动单元格
} XlInsertShiftDirection;

//边框的线条样式
//http://msdn.microsoft.com/zh-cn/library/office/ff821622.aspx
typedef enum
{
	xlContinuous = 1,					//实线
	xlDash = -4115,						//虚线
	xlDashDot = 4,						//点划相间线
	xlDashDotDot = 5,					//划线后跟两个点
	xlDot = -4118,						//点式线
	xlDouble = -4119,					//双线
	xlLineStyleNone = -4142,			//无线条
	xlSlantDashDot = 13,				//倾斜的划线
} XlLineStyle;

//边框的粗细
//http://msdn.microsoft.com/zh-cn/library/office/ff197515.aspx
typedef enum
{
	xlHairline = 1,						//细线(最细的边框)
	xlMedium = -4138,					//中等
	xlThick = 4,						//粗(最宽的边框)
	xlThin = 2,							//细
} XlBorderWeight;

//指定所选功能（如边框、字体或填充）的颜色
//https://msdn.microsoft.com/ZH-CN/library/office/ff838258.aspx
//https://msdn.microsoft.com/en-us/library/cc296089.aspx
typedef enum
{
	xlColorIndexAutomatic = -4105,		//自动配色
	xlColorIndexNone = -4142,			//无颜色
} XlColorIndex;


class CExcelHelper
{
public:
	CExcelHelper();
	~CExcelHelper();
private:
	CString            m_strExcelFile;
	CString            m_strWorkSheet;

	CApplication      m_oExcelApp;	   // Excel程序
	CWorksheet        m_oWorkSheet;    // 工作表
	CWorkbook         m_oWorkBook;     // 工作簿
	CWorkbooks        m_oWorkBooks;    // 工作簿集合
	CWorksheets       m_oWorkSheets;   // 工作表集合
	CRange            m_oCurrRange;    // 使用区域

public:
	//创建Excel服务
	BOOL CreateExcelApplication();

	//销毁Excel服务
	void DestroyExcelApplication();

	//得到单元格中的值
	CString GetCellValue(int row, int col);

	//修改单元格中的值
	void SetCellValue(int row, int col, int Align, VARIANT & value);

	//设置指定单元格文本
	void SetCellText(_tstring cellPos, _tstring cellText);

	//设置指定单元格底色
	void SetCellBackgroundColor(_tstring cellStart, _tstring cellEnd, long color = 0xFFFFFF);

	//设置指定单元格字体格式
	void SetCellFontFormat(_tstring cellStart, _tstring cellEnd, long size, bool bold = true, long color = 0x000000);

	//设置单元格外边框
	void SetCellBorderAround(_tstring cellStart, _tstring cellEnd, XlLineStyle lineStyle = xlContinuous,
		XlBorderWeight borderWeight = xlThin, long color = 0x000000);

	//合并单元格
	void SetMergeCell(VARIANT & cellStart, VARIANT & cellEnd);

	//设置已使用列自适应宽度
	void SetColumnAutoFit();

	//设置已使用单元格的水平对齐方式
	void SetHorizontalAlignment(XlHAlign hAlign);

	//设置已使用单元格的垂直对齐方式
	void SetVerticalAlignment(XlVAlign vAlign);

	//设置行高
	void SetRowHeight(int row, CString height);

	//设置行高
	CString GetRowHeight(int row);

	//设置列宽
	void SetColumnWidth(int col, CString width);

	// 设置列宽
	CString GetColumnWidth(int col);

	// Converts (row,col) indices to an Excel-style A1:C1 string in Excel  
	CString IndexToString(int row, int col);

	//得到表格总第一个空行的索引
	int LastLineIndex();



	//打开指定工作薄
	BOOL OpenWorkBook(CString strExcelFile);
	
	//添加新的工作簿
	BOOL AddNewWorkBook();

	//在已打开的工作薄中，打开指定工作表
	BOOL OpenWorkSheet(CString strWorkSheet);

	//打开指定工作薄和工作表
	BOOL OpenWorkSheet(CString strExcelFile, CString strWorkSheet);

	//关闭当前打开的工作薄
	void CloseWorkBook();

	//关闭工作表
	void CloseWorkSheet();


	//将工作薄另存为指定格式
	BOOL WorkBookSaveAs(CString strSaveAsFile, XlFileFormat fileFormat);

	//将工作表另存为指定格式
	void WorkSheetSaveAs(CString strSaveAsFile, XlFileFormat fileFormat);

	//获取当前工作表已使用的行数
	long GetUsedRowsCount();

	//获取当前工作表已使用的列数
	long GetUsedColumnsCount();

	//获取当前工作表中一行数据
	BOOL GetRowString(UINT nRow, CStringArray &array);

	//在指定行插入数据
	void InsertRow(UINT nRow, CStringArray &array);

	//设置指定单元格颜色
	void SetRangeBackgroundColor(CString strStart, CString strEnd, long nColorIndex);

	//设置单元格边框
	void SetRangeBorders(CString strStart, CString strEnd, XlLineStyle lineStyle,
		XlBorderWeight borderWeight);

	//设置单元格外边框
	void SetRangeBorderAround(CString strStart, CString strEnd, XlLineStyle lineStyle,
		XlBorderWeight borderWeight);

	//设置所有单元格的水平对齐和垂直对齐
	void SetAlignment(CString strStart, CString strEnd, XlHAlign hAlign, XlVAlign vAlign);

	//保存工作薄
	void WorkBookSave();

	//另存工作薄
	BOOL WorkBookSaveAs(_tstring excelFileParh, bool deleteFile = true);
};

