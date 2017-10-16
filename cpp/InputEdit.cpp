// InputEdit.cpp : implementation file
//

#include "stdafx.h"
#include "NVMonitor.h"
#include "InputEdit.h"


// CInputEdit

IMPLEMENT_DYNAMIC(CInputEdit, CEdit)

CInputEdit::CInputEdit()
{

}

CInputEdit::~CInputEdit()
{
}


BEGIN_MESSAGE_MAP(CInputEdit, CEdit)
	ON_WM_CHAR()
END_MESSAGE_MAP()



// CInputEdit message handlers




void CInputEdit::OnChar(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	// TODO: Add your message handler code here and/or call default

	// 保证小数点最多只能出现一次
	if (nChar == '.')
	{
		CString str;
		// 获取原来编辑框中的字符串
		GetWindowText(str);
		//若原来的字符串中已经有一个小数点,则不将小数点输入,保证了最多只能输入一个小数点
		if (str.Find('.') != -1)
		{
		}
		// 否则就输入这个小数点
		else
		{
			CEdit::OnChar(nChar, nRepCnt, nFlags);
		}
	}
	// 保证负号只能出现一次,并且只能出现在第一个字符
	else if (nChar == '-')
	{
		CString str;
		GetWindowText(str);
		// 还没有输入任何字符串
		if (str.IsEmpty())
		{
			CEdit::OnChar(nChar, nRepCnt, nFlags);
		}
		else
		{
			int nSource, nDestination;
			GetSel(nSource, nDestination);
			// 此时选择了全部的内容
			if (nSource == 0 && nDestination == str.GetLength())
			{
				CEdit::OnChar(nChar, nRepCnt, nFlags);
			}
			else
			{
			}
		}
	}
	// 除了小数点和负号,还允许输入数字,Backspace,Delete
	else if ((nChar >= '0' && nChar <= '9') || (nChar == 0x08) || (nChar == 0x10))
	{
		CEdit::OnChar(nChar, nRepCnt, nFlags);
	}
	// 其它的键,都不响应
	else
	{
	}
}
