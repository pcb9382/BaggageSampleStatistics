
// DataStatisticsDlg.h : ͷ�ļ�
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

// CDataStatisticsDlg �Ի���
class CDataStatisticsDlg : public CDialogEx
{
// ����
public:
	CDataStatisticsDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_DATASTATISTICS_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	
	CString strFilePathTxt1;                            //���ڼ���txt�����ļ��ĵ�ַ
	String strFilePathTxt;                              //��txt��ַ��CStringת����String
	vector<CString> vecResult;                          //���ڴ洢��ȡtxt����������
	vector<char> FileName;                              //���ڻ�ȡtxt�ļ���ÿһ���е������ļ���
	CEdit m_edit1;
	Mat frame;                                          //��Ŷ�ȡ��ͼ��                                    
	CString strNumFile;                                 //���ڴ��ÿһ�е�����
	char DecisionCode;                                  //���ߴ���
	char CategoryCode;                                  //������
	char TrayMarkCode;                                  //���̱�־
	char NumberCode;                                    //��������
	String  strFilePathPicture;                         //��Ӧ��ͼ��ĵ�ַ
	
	CString LaserFileName;                              //���ڻ�ü�����ļ���
	bool GlobalVariable;                                //�����ж��Ƿ�����ڱ༭����
	
	//�����������Ӵ����д�������
	CString Category;                                   //��������
	CString TrayMark;                                   //���������
	CString Number;                                     //����ļ���

	//����������Ӵ��ڴ�����������
	CString BaggageColor1;                              //�������ɫ
	CString PasteStick1;                                //�������������
	BOOL    BaggageDiscrimination;                      //�����ж������б��Ƿ���ȷ
	CString ReviseCategory1;                            //�޸ĺ����������
	CString ReviseTrayMark1;                            //�޸ĺ����������
	CString ReviseNumber1;                              //�޸ĺ�ļ���
	BOOL    Exit1;                                      //�����ж��Ƿ�ͻ�������ı�־

   //�����Ź���EXCEL����һЩ����ຯ��
	CWorkbooks books;                                   //EXCEL����������
	CWorkbook book;                                     //EXCEL������
	CWorksheets sheets;                                 //EXCEL��������
	CWorksheet sheet;                                   //EXCEL������
	CApplication app;                                   //��ʾ�ĵ�
	Cnterior it;                                        //��ɫ
	CShapes  shapes;
	CShapeRange sRange;
	CRange range, cols, rows;

public:

	
	void ShowImage(Mat Img, UINT ID);                   //������ʾͼ����
	afx_msg void OnBnClickedButton3();
	int BaggageTypeJudgement(char ch1,char ch2,char ch3);//�ж���������ࡢ�����Լ���������
	void MatchingPicture(CString strFile, vector<char>FileName);//���ݼ����ļ����ƺ�ͼ�����ƽ���ƥ�亯��
	afx_msg void OnClickedButtonExit();                 //�˳�����
	//����Ϊ��д������ݣ���txt�ж����һ�����ݣ����������ɫ���˵��жϣ���������������������ж������б��Ƿ���ȷ�������죩
	//�޸ĺ����������,�޸ĺ����������,�޸ĺ�ļ���
	void SaveDataToEXCEL(CString saveFilePath,CString WriteContents, CString BaggageColor1, CString PasteStick1, BOOL BaggageDiscrimination,
		                  CString ReviseCategory1, CString ReviseTrayMark1, CString ReviseNumber1);
};
