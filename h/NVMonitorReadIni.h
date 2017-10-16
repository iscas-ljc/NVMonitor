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
	//�ж��ļ��Ƿ����
	BOOL IsFileExist(const char* szFileName);
	//��ȡ�ļ�����
	void ReadIniString();
	
	//����ṹ��洢����
	typedef struct{
		//X��ĳ�ֵ
		CString strMinLebal;
		double dXMin;
		//X����ֵ
		CString strMaxLebal;
		double dXMax;
		//X�Ჽ��lebal
		CString strStepLebal;
		double dXStep;
		//X��lebal����λ
		CString strXLabel;
		CString strXUnit;
		//Y��lebal����λ
		CString strYLebal;
		CString strYUnit;
		//X�� Y������
		vector<CPlanarPoint> mBasePointVec;
	}DataSrc;

	DataSrc mDataSrcRPM;
	vector<CString> Thresholdfiles;
	void ReadIniString(DataSrc &Data, CString filename, CString Section);
	//����ṹ��洢Section����Ŀ
	struct {
		int Count;
	}Threshold_Mode;
	
	Property GetProperty(CPlanarPoint point);
	BOOL ComparePoint(CPlanarPoint point);
	BOOL SearchThresholdFiles();
};

