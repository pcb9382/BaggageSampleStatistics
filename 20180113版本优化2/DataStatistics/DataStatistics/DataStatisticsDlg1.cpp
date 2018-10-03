// DataStatisticsDlg1.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "DataStatistics.h"
#include "DataStatisticsDlg1.h"
#include "afxdialogex.h"
#include "DataStatisticsDlg.h"

// DataStatisticsDlg1 �Ի���

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


// DataStatisticsDlg1 ��Ϣ�������


BOOL DataStatisticsDlg1::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	CDataStatisticsDlg *pFatherDlg;
	pFatherDlg = (CDataStatisticsDlg*)GetParent();


	Category1 = pFatherDlg->Category;                                   //��������
	TrayMark1 = pFatherDlg->TrayMark;                                   //���������
	Number1 = pFatherDlg->Number;                                       //����ļ���
	GetDlgItem(IDC_EDIT1)->SetWindowText(Category1);
	GetDlgItem(IDC_EDIT2)->SetWindowText(TrayMark1);
	GetDlgItem(IDC_EDIT3)->SetWindowText(Number1);


	return TRUE;
}

//�����ж���ȷ
void DataStatisticsDlg1::OnBnClickedOk()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	UpdateData(TRUE);
	switch (m_BaggageColor)
	{
	case 0 :
		BaggageColor = "��ɫ";
	    break;
	case 1:
		BaggageColor = "��ɫ";
		break;
	case 2:
		BaggageColor = "��ɫ";
		break;
	case 3:
		BaggageColor = "��ɫ";
		break;
	case 4:
		BaggageColor = "��ɫ";
		break;
	case 5:
		BaggageColor = "��ɫ";
		break;
	default:
		break;
	}

	switch (m_PasteStick)
	{

	case 0:
		PasteStick = "�п�ճ������";
		break;
	case 1:
		PasteStick = "�޿�ճ������";
		break;
	default:
		break;
	}

	m_BaggageDiscrimination = TRUE;       //�б���ȷ
	CDataStatisticsDlg *pFatherDlg;
	pFatherDlg = (CDataStatisticsDlg*)GetParent();
	//�򸸴��ڴ��ݺ���
	pFatherDlg->BaggageColor1 = BaggageColor;
	pFatherDlg->PasteStick1 = PasteStick;
	pFatherDlg->BaggageDiscrimination = m_BaggageDiscrimination;
	CDialogEx::OnOK();
}

//�����жϴ�����Ҫ�޸Ĳ����ĺ���
void DataStatisticsDlg1::OnBnClickedButton1()
{
	//�������Ļ����ͻ���д��ȷ�����࣬�����Լ���������
	UpdateData(TRUE);
	GetDlgItem(IDC_EDIT1)->GetWindowText(ReviseCategory);
	GetDlgItem(IDC_EDIT2)->GetWindowText(ReviseTrayMark);
	GetDlgItem(IDC_EDIT3)->GetWindowText(ReviseNumber);
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	
	switch (m_BaggageColor)
	{
	
	case 0:
		BaggageColor = "��ɫ";
		break;
	case 1:
		BaggageColor = "��ɫ";
		break;
	case 2:
		BaggageColor = "��ɫ";
		break;
	case 3:
		BaggageColor = "��ɫ";
		break;
	case 4:
		BaggageColor = "��ɫ";
		break;
	case 5:
		BaggageColor = "��ɫ";
		break;
	default:
		MessageBox("��ѡ����ɫ��");
		break;
	}

	switch (m_PasteStick)
	{

	case 0:
		PasteStick = "�п�ճ������";
		break;
	case 1:
		PasteStick = "�޿�ճ������";
		break;
	default:
		MessageBox("��ѡ���Ƿ����������");
		break;
	}

	m_BaggageDiscrimination = FALSE;       //�б���ȷ
	CDataStatisticsDlg *pFatherDlg;
	pFatherDlg = (CDataStatisticsDlg*)GetParent();
	
	//�򸸴��ڴ��ݺ���
	pFatherDlg->BaggageColor1 = BaggageColor;
	pFatherDlg->PasteStick1 = PasteStick;
	pFatherDlg->BaggageDiscrimination = m_BaggageDiscrimination;
	pFatherDlg->ReviseCategory1 = ReviseCategory;                               //�޸ĺ����������
	pFatherDlg->ReviseTrayMark1 = ReviseTrayMark;                               //�޸ĺ����������
	pFatherDlg->ReviseNumber1 = ReviseNumber;                                   //�޸ĺ�ļ���
	CDialogEx::OnOK();

}

//�˳�����
void DataStatisticsDlg1::OnBnClickedButtonexit()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	Exit = TRUE;
	CDataStatisticsDlg *pFatherDlg;
	pFatherDlg = (CDataStatisticsDlg*)GetParent();
	pFatherDlg->Exit1 = Exit;
	Exit = FALSE;
	CDialogEx::OnOK();
}
