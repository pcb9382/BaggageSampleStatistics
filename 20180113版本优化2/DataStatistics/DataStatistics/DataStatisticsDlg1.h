#pragma once
#include "resource.h"
#include "DataStatisticsDlg.h"

// DataStatisticsDlg1 �Ի���

class DataStatisticsDlg1 : public CDialogEx
{
	DECLARE_DYNAMIC(DataStatisticsDlg1)

public:
	DataStatisticsDlg1(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~DataStatisticsDlg1();

// �Ի�������
	enum { IDD = IDD_DIALOG1 };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
public:
	BOOL OnInitDialog();

	//����ɸ��Ի����д�������CString����
	CString Category1;                                   //��������
	CString TrayMark1;                                   //���������
	CString Number1;                                     //����ļ���

	//��ѡ��ť�ĳ�Ա����
	int  m_BaggageColor;                                  //�������ɫ                            
	int  m_PasteStick;                                    //�Ƿ���п�ճ������
	bool m_BaggageDiscrimination;                         //�����б��Ƿ���ȷ
	bool Exit;

	//���ݸ������ڵĲ���
	CString BaggageColor;                                 //�������ɫ
	CString PasteStick;                                   //�������������
	CString ReviseCategory;                               //�޸ĺ����������
	CString ReviseTrayMark;                               //�޸ĺ����������
	CString ReviseNumber;                                 //�޸ĺ�ļ���
public:
	afx_msg void OnBnClickedOk();                         //�����б���ȷ�ĺ���
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButtonexit();
};
