// DataStatisticsDlg1.cpp : 实现文件
//

#include "stdafx.h"
#include "DataStatistics.h"
#include "DataStatisticsDlg1.h"
#include "afxdialogex.h"
#include "DataStatisticsDlg.h"

// DataStatisticsDlg1 对话框

IMPLEMENT_DYNAMIC(DataStatisticsDlg1, CDialogEx)

DataStatisticsDlg1::DataStatisticsDlg1(CWnd* pParent /*=NULL*/)
	: CDialogEx(DataStatisticsDlg1::IDD, pParent)
	, m_BaggageColor(3)                                          
	, m_PasteStick(0)                                            
{

}

DataStatisticsDlg1::~DataStatisticsDlg1()
{
}

void DataStatisticsDlg1::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Radio(pDX, IDC_RADIO1, m_BaggageColor);
	DDX_Radio(pDX, IDC_RADIO7, m_PasteStick);
	
}


BEGIN_MESSAGE_MAP(DataStatisticsDlg1, CDialogEx)
	ON_BN_CLICKED(IDOK, &DataStatisticsDlg1::OnBnClickedOk)
	ON_BN_CLICKED(IDC_BUTTON1, &DataStatisticsDlg1::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTONExit, &DataStatisticsDlg1::OnBnClickedButtonexit)
END_MESSAGE_MAP()


// DataStatisticsDlg1 消息处理程序


BOOL DataStatisticsDlg1::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	CDataStatisticsDlg *pFatherDlg;
	pFatherDlg = (CDataStatisticsDlg*)GetParent();


	Category1 = pFatherDlg->Category;                                   //行李的类别
	TrayMark1 = pFatherDlg->TrayMark;                                   //行李的托盘
	Number1 = pFatherDlg->Number;                                       //行李的件数
	GetDlgItem(IDC_EDIT1)->SetWindowText(Category1);
	GetDlgItem(IDC_EDIT2)->SetWindowText(TrayMark1);
	GetDlgItem(IDC_EDIT3)->SetWindowText(Number1);


	return TRUE;
}

//行李判断正确
void DataStatisticsDlg1::OnBnClickedOk()
{
	// TODO:  在此添加控件通知处理程序代码
	UpdateData(TRUE);
	switch (m_BaggageColor)
	{
	case 0 :
		BaggageColor = "白色";
	    break;
	case 1:
		BaggageColor = "灰色";
		break;
	case 2:
		BaggageColor = "红色";
		break;
	case 3:
		BaggageColor = "黑色";
		break;
	case 4:
		BaggageColor = "银色";
		break;
	case 5:
		BaggageColor = "黄色";
		break;
	default:
		break;
	}

	switch (m_PasteStick)
	{

	case 0:
		PasteStick = "有可粘贴区域";
		break;
	case 1:
		PasteStick = "无可粘贴区域";
		break;
	default:
		break;
	}

	m_BaggageDiscrimination = TRUE;       //判别正确
	CDataStatisticsDlg *pFatherDlg;
	pFatherDlg = (CDataStatisticsDlg*)GetParent();
	//向父窗口传递函数
	pFatherDlg->BaggageColor1 = BaggageColor;
	pFatherDlg->PasteStick1 = PasteStick;
	pFatherDlg->BaggageDiscrimination = m_BaggageDiscrimination;
	CDialogEx::OnOK();
}

//行李判断错误，需要修改参数的函数
void DataStatisticsDlg1::OnBnClickedButton1()
{
	//如果错误的话，就会填写正确的种类，件数以及有无托盘
	UpdateData(TRUE);
	GetDlgItem(IDC_EDIT1)->GetWindowText(ReviseCategory);
	GetDlgItem(IDC_EDIT2)->GetWindowText(ReviseTrayMark);
	GetDlgItem(IDC_EDIT3)->GetWindowText(ReviseNumber);
	// TODO:  在此添加控件通知处理程序代码
	
	switch (m_BaggageColor)
	{
	
	case 0:
		BaggageColor = "白色";
		break;
	case 1:
		BaggageColor = "灰色";
		break;
	case 2:
		BaggageColor = "红色";
		break;
	case 3:
		BaggageColor = "黑色";
		break;
	case 4:
		BaggageColor = "银色";
		break;
	case 5:
		BaggageColor = "黄色";
		break;
	default:
		MessageBox("请选择颜色！");
		break;
	}

	switch (m_PasteStick)
	{

	case 0:
		PasteStick = "有可粘贴区域";
		break;
	case 1:
		PasteStick = "无可粘贴区域";
		break;
	default:
		MessageBox("请选择是否有黏贴区域！");
		break;
	}

	m_BaggageDiscrimination = FALSE;       //判别正确
	CDataStatisticsDlg *pFatherDlg;
	pFatherDlg = (CDataStatisticsDlg*)GetParent();
	
	//向父窗口传递函数
	pFatherDlg->BaggageColor1 = BaggageColor;
	pFatherDlg->PasteStick1 = PasteStick;
	pFatherDlg->BaggageDiscrimination = m_BaggageDiscrimination;
	pFatherDlg->ReviseCategory1 = ReviseCategory;                               //修改后的行李种类
	pFatherDlg->ReviseTrayMark1 = ReviseTrayMark;                               //修改后的有无托盘
	pFatherDlg->ReviseNumber1 = ReviseNumber;                                   //修改后的件数
	CDialogEx::OnOK();

}

//退出程序
void DataStatisticsDlg1::OnBnClickedButtonexit()
{
	// TODO:  在此添加控件通知处理程序代码
	Exit = TRUE;
	CDataStatisticsDlg *pFatherDlg;
	pFatherDlg = (CDataStatisticsDlg*)GetParent();
	pFatherDlg->Exit1 = Exit;
	Exit = FALSE;
	CDialogEx::OnOK();
}
