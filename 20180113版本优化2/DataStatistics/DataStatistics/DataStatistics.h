
// DataStatistics.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CDataStatisticsApp: 
// �йش����ʵ�֣������ DataStatistics.cpp
//

class CDataStatisticsApp : public CWinApp
{
public:
	CDataStatisticsApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CDataStatisticsApp theApp;