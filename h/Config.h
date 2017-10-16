#pragma once

#include "vector"
using namespace std;

typedef struct _Section {
	double start;
	double end;
} Section;

class CConfig
{
public:
	CConfig();
	~CConfig();

public:
	vector<Section> m_VecFConfigs;
	vector<Section> m_VecTConfigs;

	void Convert2(Section section, double ancher, double step, int &index_start, int &index_end);
	void SaveConfigs();
	void ReadConfigs();
	void Write(CString strFileName, vector<Section> vecConfigs);
	void Read(CString strFileName, vector<Section> &vecConfigs);
};

