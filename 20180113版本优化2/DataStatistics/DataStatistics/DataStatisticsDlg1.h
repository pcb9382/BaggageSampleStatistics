#pragma once
#include "resource.h"
#include "DataStatisticsDlg.h"

// DataStatisticsDlg1 对话框

class DataStatisticsDlg1 : public CDialogEx
{
	DECLARE_DYNAMIC(DataStatisticsDlg1)

public:
	DataStatisticsDlg1(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~DataStatisticsDlg1();

// 对话框数据
	enum { IDD = IDD_DIALOG1 };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	BOOL OnInitDialog();

	//存放由父对话框中传过来的CString数据
	CString Category1;                                   //行李的类别
	CString TrayMark1;                                   //行李的托盘
	CString Number1;                                     //行李的件数

	//单选按钮的成员变量
	int  m_BaggageColor;                                  //行李的颜色                            
	int  m_PasteStick;                                    //是否具有可粘贴区域
	bool m_BaggageDiscrimination;                         //行李判别是否正确
	bool Exit;

	//传递给父窗口的参数
	CString BaggageColor;                                 //行李的颜色
	CString PasteStick;                                   //行李的贴条区域
	CString ReviseCategory;                               //修改后的行李种类
	CString ReviseTrayMark;                               //修改后的有无托盘
	CString ReviseNumber;                                 //修改后的件数
public:
	afx_msg void OnBnClickedOk();                         //行李判别正确的函数
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButtonexit();
};
