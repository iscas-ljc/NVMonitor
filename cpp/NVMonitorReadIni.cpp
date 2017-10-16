#include "stdafx.h"
#include "NVMonitor.h"
#include "NVMonitorReadIni.h"

// CNVMonitorReadIni

IMPLEMENT_DYNAMIC(CNVMonitorReadIni, CWnd)

CNVMonitorReadIni::CNVMonitorReadIni() 
{
	mDataSrcRPM.strMinLebal = RPM_XMIN_LEBAL;
	mDataSrcRPM.strMaxLebal = RPM_XMAX_LEBAL;
	mDataSrcRPM.strStepLebal = RPM_XSTEP_LEBAL;
	mDataSrcRPM.strXLabel = RPM_X_LABEL;
	mDataSrcRPM.strYLebal = RPM_Y_LABEL;
}
CNVMonitorReadIni::~CNVMonitorReadIni() {}


BEGIN_MESSAGE_MAP(CNVMonitorReadIni, CWnd)
END_MESSAGE_MAP()


BOOL CNVMonitorReadIni::IsFileExist(const char* szFileName)
{
	CFileStatus stat;
	if (CFile::GetStatus((LPCTSTR)szFileName, stat)) { return TRUE; }
	else { return FALSE; }
}

void CNVMonitorReadIni::ReadIniString()
{
}




void CNVMonitorReadIni::ReadIniString(DataSrc &Data, CString filename, CString Section)
{
	//准备文件路径
	char Path[256];
	GetCurrentDirectory(256, (LPWSTR)Path);
	CString strFilePath(Path);
	strFilePath.Format(_T("%s\\threshold\\%s"), Path, filename);
	if (!IsFileExist((const char *)strFilePath.GetBuffer()))
	{
		CString strGetDataError("IniFile Do Not Exist!");
		strGetDataError.Format(_T("%s"), strGetDataError);
		MessageBoxEx(NULL, strGetDataError, _T("Warning"), MB_OK | MB_ICONINFORMATION,
			MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US));
		exit(0);
	}
	CString strKey; char buf[256]; double dNum;

	//获取X轴坐标单位
	::GetPrivateProfileString(Section, Data.strXLabel, NULL,
		Data.strXUnit.GetBuffer(256), 256, strFilePath);

	//获取Y轴坐标单位
	::GetPrivateProfileString(Section, Data.strYLebal, NULL,
		Data.strYUnit.GetBuffer(256), 256, strFilePath);

	//获取X轴初值
	::GetPrivateProfileString(Section, Data.strMinLebal, NULL,
		(LPWSTR)buf, 256, strFilePath);
	dNum = _wtof((const wchar_t*)buf);
	Data.dXMin = dNum;

	//获取X轴终值
	::GetPrivateProfileString(Section, Data.strMaxLebal, NULL,
		(LPWSTR)buf, 256, strFilePath);
	dNum = _wtof((const wchar_t*)buf);
	Data.dXMax = dNum;

	//获取X轴步长
	::GetPrivateProfileString(Section, Data.strStepLebal, NULL,
		(LPWSTR)buf, 256, strFilePath);
	dNum = _wtof((const wchar_t*)buf);
	Data.dXStep = dNum;

	//循环获取X Y轴的值
	Data.mBasePointVec.clear();
	CPlanarPoint point;
	for (double i = Data.dXMin; i <= Data.dXMax; i = i + Data.dXStep)
	{
		strKey.Format(_T("%.3lf"), i);
		::GetPrivateProfileString(Section, strKey, NULL,
			(LPWSTR)buf, 256, strFilePath);
		
		dNum = _wtof((const wchar_t*)buf);
		point.dX = i;
		point.dY = dNum;
		Data.mBasePointVec.push_back(point);
	}
}




Property CNVMonitorReadIni::GetProperty(CPlanarPoint point)
{
	Property property;
	int index = (point.dX / 1000);
	CPlanarPoint pointL = mDataSrcRPM.mBasePointVec.at(index);
	CPlanarPoint pointR = mDataSrcRPM.mBasePointVec.at(index + 1);
	double slope = (pointR.dY - pointL.dY) / (pointR.dX - pointL.dX);
	double b = pointR.dY - slope * pointR.dX;
	property.slope = slope;
	property.b = b;
	return property;
}

BOOL CNVMonitorReadIni::ComparePoint(CPlanarPoint point)
{
	Property property = GetProperty(point);
	double dTemp = property.slope * point.dX + property.b;
	if (dTemp < point.dY) {
		return FALSE;
	}
	else {
		return TRUE;
	}
}

BOOL CNVMonitorReadIni::SearchThresholdFiles()
{
	//准备文件路径
	char Path[256];
	GetCurrentDirectory(256, (LPWSTR)Path);
	CString strFilePath(Path);
	strFilePath.Format(_T("%s\\threshold\\"), Path);
	if (!IsFileExist((const char *)strFilePath.GetBuffer()))
	{
		CString strGetDataError("IniDirectory Do Not Exist!");
		strGetDataError.Format(_T("%s"), strGetDataError);
		MessageBoxEx(NULL, strGetDataError, _T("Warning"), MB_OK | MB_ICONINFORMATION,
			MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US));
		exit(0);
	}
	strFilePath.Format(_T("%s\\threshold\\*.ini"), Path);
	
	WIN32_FIND_DATA fd;
	HANDLE hFind = ::FindFirstFile(strFilePath, &fd);
	if (hFind != INVALID_HANDLE_VALUE)
	{
		do
		{
			if (!(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
			{
				Thresholdfiles.push_back(fd.cFileName);
			}
		} while (::FindNextFile(hFind, &fd));
		::FindClose(hFind);
	}
	return 0;
}

