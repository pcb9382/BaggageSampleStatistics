
// DataStatisticsDlg.h : 头文件
//

#pragma once
#include <afxdisp.h>
#include "Resource.h"
#include "CApplication.h"
#include "CRange.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"
#include "CShapes.h"
#include "Cnterior.h"
#include "CShapeRange.h"
#include "resource.h"
#include "CvvImage.h"
#include <Windows.h>
#include <stdio.h>
#include <iostream>
#include <string.h>
#include <opencv2\opencv.hpp>  
#include "afxwin.h"
#include "windows.h"
#include "DataStatisticsDlg1.h"

using namespace std;
using namespace cv;

// CDataStatisticsDlg 对话框
class CDataStatisticsDlg : public CDialogEx
{
// 构造
public:
	CDataStatisticsDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_DATASTATISTICS_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	
	CString strFilePathTxt1;                            //用于加载txt处理文件的地址
	String strFilePathTxt;                              //把txt地址由CString转换成String
	vector<CString> vecResult;                          //用于存储读取txt中内容向量
	vector<char> FileName;                              //用于获取txt文件中每一行中的数据文件名
	CEdit m_edit1;
	Mat frame;                                          //存放读取的图像                                    
	CString strNumFile;                                 //用于存放每一行的内容
	char DecisionCode;                                  //决策代码
	char CategoryCode;                                  //类别代码
	char TrayMarkCode;                                  //托盘标志
	char NumberCode;                                    //件数代码
	String  strFilePathPicture;                         //相应的图像的地址
	
	CString LaserFileName;                              //用于获得激光的文件名
	bool GlobalVariable;                                //用于判断是否输出在编辑框中
	
	//这里用于向子窗口中传递数据
	CString Category;                                   //行李的类别
	CString TrayMark;                                   //行李的托盘
	CString Number;                                     //行李的件数

	//这里接收由子窗口传递来的数据
	CString BaggageColor1;                              //行李的颜色
	CString PasteStick1;                                //行李的贴条区域
	BOOL    BaggageDiscrimination;                      //用来判断行李判别是否正确
	CString ReviseCategory1;                            //修改后的行李种类
	CString ReviseTrayMark1;                            //修改后的有无托盘
	CString ReviseNumber1;                              //修改后的件数
	BOOL    Exit1;                                      //用于判断是否突出函数的标志

   //这里存放关于EXCEL表格的一些相关类函数
	CWorkbooks books;                                   //EXCEL工作簿集合
	CWorkbook book;                                     //EXCEL工作簿
	CWorksheets sheets;                                 //EXCEL工作表集合
	CWorksheet sheet;                                   //EXCEL工作表
	CApplication app;                                   //显示文档
	Cnterior it;                                        //颜色
	CShapes  shapes;
	CShapeRange sRange;
	CRange range, cols, rows;

public:

	
	void ShowImage(Mat Img, UINT ID);                   //用于显示图像函数
	afx_msg void OnBnClickedButton3();
	int BaggageTypeJudgement(char ch1,char ch2,char ch3);//判断行李的种类、件数以及有无托盘
	void MatchingPicture(CString strFile, vector<char>FileName);//根据激光文件名称和图像名称进行匹配函数
	afx_msg void OnClickedButtonExit();                 //退出程序
	//参数为：写入的内容（从txt中读入的一行数据），行李的颜色（人的判断），行李的贴条区域，用来判断行李判别是否正确（错误标红）
	//修改后的行李种类,修改后的有无托盘,修改后的件数
	void SaveDataToEXCEL(CString saveFilePath,CString WriteContents, CString BaggageColor1, CString PasteStick1, BOOL BaggageDiscrimination,
		                  CString ReviseCategory1, CString ReviseTrayMark1, CString ReviseNumber1);
};
