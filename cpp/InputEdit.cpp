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

	// ��֤С�������ֻ�ܳ���һ��
	if (nChar == '.')
	{
		CString str;
		// ��ȡԭ���༭���е��ַ���
		GetWindowText(str);
		//��ԭ�����ַ������Ѿ���һ��С����,�򲻽�С��������,��֤�����ֻ������һ��С����
		if (str.Find('.') != -1)
		{
		}
		// ������������С����
		else
		{
			CEdit::OnChar(nChar, nRepCnt, nFlags);
		}
	}
	// ��֤����ֻ�ܳ���һ��,����ֻ�ܳ����ڵ�һ���ַ�
	else if (nChar == '-')
	{
		CString str;
		GetWindowText(str);
		// ��û�������κ��ַ���
		if (str.IsEmpty())
		{
			CEdit::OnChar(nChar, nRepCnt, nFlags);
		}
		else
		{
			int nSource, nDestination;
			GetSel(nSource, nDestination);
			// ��ʱѡ����ȫ��������
			if (nSource == 0 && nDestination == str.GetLength())
			{
				CEdit::OnChar(nChar, nRepCnt, nFlags);
			}
			else
			{
			}
		}
	}
	// ����С����͸���,��������������,Backspace,Delete
	else if ((nChar >= '0' && nChar <= '9') || (nChar == 0x08) || (nChar == 0x10))
	{
		CEdit::OnChar(nChar, nRepCnt, nFlags);
	}
	// �����ļ�,������Ӧ
	else
	{
	}
}
