#pragma once


// CInputEdit

class CInputEdit : public CEdit
{
	DECLARE_DYNAMIC(CInputEdit)

public:
	CInputEdit();
	virtual ~CInputEdit();

protected:
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnChar(UINT nChar, UINT nRepCnt, UINT nFlags);
public:
	int row;
	int column;
};


