#include "stdafx.h"
#include "Config.h"


CConfig::CConfig()
{
}


CConfig::~CConfig()
{
}

void CConfig::Convert2(Section section, double dAncher, double dStep, int &index_start, int &index_end)
{
	index_start = (section.start - dAncher + dStep - 0.001) / dStep;
	index_end = (section.end - dAncher) / dStep + 1;
}

void CConfig::SaveConfigs()
{
	Write(_T("d:\\LocalDatas.txt"), m_VecFConfigs);
	Write(_T("d:\\LocalDatas_1.txt"), m_VecTConfigs);
}

void CConfig::ReadConfigs()
{
	Read(_T("d:\\LocalDatas.txt"), m_VecFConfigs);
	Read(_T("d:\\LocalDatas_1.txt"),m_VecTConfigs);
}

void CConfig::Write(CString strFileName, vector<Section> vecConfigs)
{
	CFile file(strFileName, CFile::modeCreate |
		CFile::modeWrite);
	CArchive ar(&file, CArchive::store);
	for (int i = 0; i < vecConfigs.size(); i++) {
		ar.Write(&vecConfigs.at(i), sizeof(Section));
	}
}

void CConfig::Read(CString strFileName, vector<Section> &vecConfigs)
{
	vecConfigs.clear();
	try
	{
		CFile file(strFileName, CFile::modeRead); //创建要读取的文件
		CArchive ar(&file, CArchive::load);//创建CArchive对象，标识是load
		Section section;
		vecConfigs.clear();
		DWORD dwBytesRemaining = file.GetLength();
		while (dwBytesRemaining) {
			UINT nBytesRead = ar.Read(&section, sizeof(Section));
			vecConfigs.push_back(section);
			dwBytesRemaining -= nBytesRead;
		}
	}
	catch (CFileException* e) {
		e->ReportError();
		e->Delete();
	}
}
