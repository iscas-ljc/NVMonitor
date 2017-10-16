#pragma once

#include "vector"
#include "PlanarPoint.h"
using namespace std;


#define RPM_XMAX_LEBAL		 _T("X_MAX");
#define RPM_XMIN_LEBAL		 _T("X_MIN")
#define RPM_XSTEP_LEBAL      _T("X_STEP")
#define RPM_X_LABEL			_T("Angular velocity")
#define RPM_Y_LABEL			_T("Acoustic pressure")

#define ThresholdDirectory	_T("threshold")

typedef struct {
	double slope;
	double b;
} Property;

class CNVMonitorReadIni : public CWnd
{
	DECLARE_DYNAMIC(CNVMonitorReadIni)

public:
	CNVMonitorReadIni();
	virtual ~CNVMonitorReadIni();

protected:
	DECLARE_MESSAGE_MAP()
public:
	//判断文件是否存在
	BOOL IsFileExist(const char* szFileName);
	//读取文件内容
	void ReadIniString();
	
	//定义结构体存储数据
	typedef struct{
		//X轴的初值
		CString strMinLebal;
		double dXMin;
		//X轴终值
		CString strMaxLebal;
		double dXMax;
		//X轴步长lebal
		CString strStepLebal;
		double dXStep;
		//X轴lebal、单位
		CString strXLabel;
		CString strXUnit;
		//Y轴lebal、单位
		CString strYLebal;
		CString strYUnit;
		//X轴 Y轴数据
		vector<CPlanarPoint> mBasePointVec;
	}DataSrc;

	DataSrc mDataSrcRPM;
	vector<CString> Thresholdfiles;
	void ReadIniString(DataSrc &Data, CString filename, CString Section);
	//定义结构体存储Section的数目
	struct {
		int Count;
	}Threshold_Mode;
	
	Property GetProperty(CPlanarPoint point);
	BOOL ComparePoint(CPlanarPoint point);
	BOOL SearchThresholdFiles();
};

