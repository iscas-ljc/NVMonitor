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


//Excel�ļ���ʽ
//http://msdn.microsoft.com/zh-cn/library/office/ff198017.aspx
typedef enum
{
	xlCSV = 6,				   //csv
	xlTextWindows = 20,        //Windows �ı�
	xlTextMSDOS = 21,          //MSDOS �ı�
	xlUnicodeText = 42,        //Unicode �ı�
	xlExcel9795 = 43,          //Excel9795
	xlWorkbookNormal = -4143,  //���湤����
	xlExcel12 = 50,            //Excel 12
	xlWorkbookDefault = 51,    //Ĭ�Ϲ�����
} XlFileFormat;


//ˮƽ���뷽ʽ
//http://msdn.microsoft.com/zh-cn/library/office/ff840772.aspx
typedef enum
{
	xlHAlignCenter = -4108,					//���ж���
	xlHAlignCenterAcrossSelection = 7,		//���о���
	xlHAlignDistributed = -4117,			//��ɢ����
	xlHAlignFill = 5,						//���
	xlHAlignGeneral = 1,					//���������Ͷ���
	xlHAlignJustify = -4130,				//���˶���
	xlHAlignLeft = -4131,					//�����
	xlHAlignRight = -4152,					//�Ҷ���
} XlHAlign;

//��ֱ���뷽ʽ
//http://msdn.microsoft.com/zh-cn/library/office/ff835305.aspx
typedef enum
{
	xlVAlignBottom = -4107,				//���� 
	xlVAlignCenter = -4108,				//���ж��� 
	xlVAlignDistributed = -4117,        //��ɢ���� 
	xlVAlignJustify = -4130,			//���˶��� 
	xlVAlignTop = -4160,				//���� 
} XlVAlign;

//����ʱ��Ԫ����ƶ�����
//http://msdn.microsoft.com/zh-cn/library/office/ff837618.aspx
typedef enum
{
	xlShiftDown = -4121,        //�����ƶ���Ԫ��
	xlShiftToRight = -4161,        //�����ƶ���Ԫ��
} XlInsertShiftDirection;

//�߿��������ʽ
//http://msdn.microsoft.com/zh-cn/library/office/ff821622.aspx
typedef enum
{
	xlContinuous = 1,					//ʵ��
	xlDash = -4115,						//����
	xlDashDot = 4,						//�㻮�����
	xlDashDotDot = 5,					//���ߺ��������
	xlDot = -4118,						//��ʽ��
	xlDouble = -4119,					//˫��
	xlLineStyleNone = -4142,			//������
	xlSlantDashDot = 13,				//��б�Ļ���
} XlLineStyle;

//�߿�Ĵ�ϸ
//http://msdn.microsoft.com/zh-cn/library/office/ff197515.aspx
typedef enum
{
	xlHairline = 1,						//ϸ��(��ϸ�ı߿�)
	xlMedium = -4138,					//�е�
	xlThick = 4,						//��(���ı߿�)
	xlThin = 2,							//ϸ
} XlBorderWeight;

//ָ����ѡ���ܣ���߿��������䣩����ɫ
//https://msdn.microsoft.com/ZH-CN/library/office/ff838258.aspx
//https://msdn.microsoft.com/en-us/library/cc296089.aspx
typedef enum
{
	xlColorIndexAutomatic = -4105,		//�Զ���ɫ
	xlColorIndexNone = -4142,			//����ɫ
} XlColorIndex;


class CExcelHelper
{
public:
	CExcelHelper();
	~CExcelHelper();
private:
	CString            m_strExcelFile;
	CString            m_strWorkSheet;

	CApplication      m_oExcelApp;	   // Excel����
	CWorksheet        m_oWorkSheet;    // ������
	CWorkbook         m_oWorkBook;     // ������
	CWorkbooks        m_oWorkBooks;    // ����������
	CWorksheets       m_oWorkSheets;   // ��������
	CRange            m_oCurrRange;    // ʹ������

public:
	//����Excel����
	BOOL CreateExcelApplication();

	//����Excel����
	void DestroyExcelApplication();

	//�õ���Ԫ���е�ֵ
	CString GetCellValue(int row, int col);

	//�޸ĵ�Ԫ���е�ֵ
	void SetCellValue(int row, int col, int Align, VARIANT & value);

	//����ָ����Ԫ���ı�
	void SetCellText(_tstring cellPos, _tstring cellText);

	//����ָ����Ԫ���ɫ
	void SetCellBackgroundColor(_tstring cellStart, _tstring cellEnd, long color = 0xFFFFFF);

	//����ָ����Ԫ�������ʽ
	void SetCellFontFormat(_tstring cellStart, _tstring cellEnd, long size, bool bold = true, long color = 0x000000);

	//���õ�Ԫ����߿�
	void SetCellBorderAround(_tstring cellStart, _tstring cellEnd, XlLineStyle lineStyle = xlContinuous,
		XlBorderWeight borderWeight = xlThin, long color = 0x000000);

	//�ϲ���Ԫ��
	void SetMergeCell(VARIANT & cellStart, VARIANT & cellEnd);

	//������ʹ��������Ӧ���
	void SetColumnAutoFit();

	//������ʹ�õ�Ԫ���ˮƽ���뷽ʽ
	void SetHorizontalAlignment(XlHAlign hAlign);

	//������ʹ�õ�Ԫ��Ĵ�ֱ���뷽ʽ
	void SetVerticalAlignment(XlVAlign vAlign);

	//�����и�
	void SetRowHeight(int row, CString height);

	//�����и�
	CString GetRowHeight(int row);

	//�����п�
	void SetColumnWidth(int col, CString width);

	// �����п�
	CString GetColumnWidth(int col);

	// Converts (row,col) indices to an Excel-style A1:C1 string in Excel  
	CString IndexToString(int row, int col);

	//�õ�����ܵ�һ�����е�����
	int LastLineIndex();



	//��ָ��������
	BOOL OpenWorkBook(CString strExcelFile);
	
	//����µĹ�����
	BOOL AddNewWorkBook();

	//���Ѵ򿪵Ĺ������У���ָ��������
	BOOL OpenWorkSheet(CString strWorkSheet);

	//��ָ���������͹�����
	BOOL OpenWorkSheet(CString strExcelFile, CString strWorkSheet);

	//�رյ�ǰ�򿪵Ĺ�����
	void CloseWorkBook();

	//�رչ�����
	void CloseWorkSheet();


	//�����������Ϊָ����ʽ
	BOOL WorkBookSaveAs(CString strSaveAsFile, XlFileFormat fileFormat);

	//�����������Ϊָ����ʽ
	void WorkSheetSaveAs(CString strSaveAsFile, XlFileFormat fileFormat);

	//��ȡ��ǰ��������ʹ�õ�����
	long GetUsedRowsCount();

	//��ȡ��ǰ��������ʹ�õ�����
	long GetUsedColumnsCount();

	//��ȡ��ǰ��������һ������
	BOOL GetRowString(UINT nRow, CStringArray &array);

	//��ָ���в�������
	void InsertRow(UINT nRow, CStringArray &array);

	//����ָ����Ԫ����ɫ
	void SetRangeBackgroundColor(CString strStart, CString strEnd, long nColorIndex);

	//���õ�Ԫ��߿�
	void SetRangeBorders(CString strStart, CString strEnd, XlLineStyle lineStyle,
		XlBorderWeight borderWeight);

	//���õ�Ԫ����߿�
	void SetRangeBorderAround(CString strStart, CString strEnd, XlLineStyle lineStyle,
		XlBorderWeight borderWeight);

	//�������е�Ԫ���ˮƽ����ʹ�ֱ����
	void SetAlignment(CString strStart, CString strEnd, XlHAlign hAlign, XlVAlign vAlign);

	//���湤����
	void WorkBookSave();

	//��湤����
	BOOL WorkBookSaveAs(_tstring excelFileParh, bool deleteFile = true);
};

