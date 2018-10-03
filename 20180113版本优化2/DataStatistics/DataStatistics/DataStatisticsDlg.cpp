
// DataStatisticsDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "DataStatistics.h"
#include "DataStatisticsDlg.h"
#include "afxdialogex.h"
#include "Resource.h"
#include "resource.h"
#include <iostream>
#include <string.h>
#include "windows.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#define MAX_FILE_LEN 250                                          //�����ȡ�ļ�һ�е�����ַ���
// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CDataStatisticsDlg �Ի���



CDataStatisticsDlg::CDataStatisticsDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CDataStatisticsDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CDataStatisticsDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	//DDX_Control(pDX, IDC_LIST1, m_List1);
	//DDX_Control(pDX, IDC_EDIT1, m_edit1);
}

BEGIN_MESSAGE_MAP(CDataStatisticsDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON3, &CDataStatisticsDlg::OnBnClickedButton3)
	ON_BN_CLICKED(IDC_BUTTON1, &CDataStatisticsDlg::OnClickedButtonExit)
END_MESSAGE_MAP()


// CDataStatisticsDlg ��Ϣ�������

BOOL CDataStatisticsDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO:  �ڴ���Ӷ���ĳ�ʼ������
	FileName.clear();
    frame = 0;	
	Exit1 = FALSE;
	ReviseCategory1="";
	ReviseTrayMark1="";
    ReviseNumber1="";
	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CDataStatisticsDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CDataStatisticsDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CDataStatisticsDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

//������PictureBox������ʾͼƬ
void CDataStatisticsDlg::ShowImage(Mat Img, UINT ID)
{
	//��Matת����IplImage*����ʽ�����ڿؼ�����ʾ
	
	IplImage* pBinary = &IplImage(Img);                                                          //ǳ����
	IplImage *img = cvCloneImage(pBinary);                                                       //��ǳ����ת��Ϊ���

	CDC* pDC = GetDlgItem(ID)->GetDC();
	HDC hDC = pDC->GetSafeHdc();
	CRect rect;
	GetDlgItem(ID)->GetClientRect(&rect);
	SetRect(rect, rect.left, rect.top, rect.right, rect.bottom);
	CvvImage cimg;
	cimg.CopyOf(img);
	cimg.DrawToHDC(hDC, &rect);
	ReleaseDC(pDC);
}


//������Ҫ�����txt�����ļ�
void CDataStatisticsDlg::OnBnClickedButton3()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	
	
	
	//����ͼƬ��·��
	CFileDialog dlg(TRUE, NULL, NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, "All Files|*.*|txt|*.txt||");
	INT_PTR nRet = dlg.DoModal();
	if (nRet != IDOK)
	{
		return;
	}

	strFilePathTxt1 = dlg.GetPathName();


	if (!PathFileExists(strFilePathTxt1))
	{
		MessageBox("�ļ������ڣ�");
		return;
	}
	strFilePathTxt = strFilePathTxt1.GetBuffer(0);

	//��ȡ���ļ�������
	if (!PathFileExists(strFilePathTxt1))
	{
		return ;
	}

	CStdioFile file;
	//�ж��Ƿ���ȷ��txt�ļ�
	if (!file.Open(strFilePathTxt1, CFile::modeRead))
	{
		return ;
	}

	//���ڴ洢��ȡ�����vector����
	
	//����һ����ȡ�ı���
	CString strValue = _T("");

	

	//��ȡtxt�е����е����ݵ�����vector��
	while (file.ReadString(strValue))
	{
		vecResult.push_back(strValue);
	}
	//�ر�txt�ļ�
	file.Close();
   

    //��ʼ���������ı����ݣ��ҳ�ԭʼ�ļ�������
	int VectorSize = vecResult.size();             //��ȡ�����Ĵ�С
	

	//���￪ʼ��EXCEL���

	int nPos = strFilePathTxt1.ReverseFind(_T('\\'));
	CString strFile1 = strFilePathTxt1.Left(nPos + 1);
	CString m_saveFilePath1 = strFile1;


	CString m_saveFilePath = strFile1 + "����ͳ��5.xlsx";

	LPDISPATCH lpDisp;
	COleVariant  covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if (!app.CreateDispatch("Excel.Application"))
	{
		AfxMessageBox("�޷���ExcelӦ��", MB_OK | MB_ICONWARNING);
		return;
	}
	books.AttachDispatch(app.get_Workbooks());
	//���ļ�  
	lpDisp = books.Open(m_saveFilePath, covOptional, covOptional, covOptional, covOptional, covOptional
		, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional);
	//�õ��õ���ͼworkbook  
	book.AttachDispatch(lpDisp);
	//�õ�worksheets  
	sheets.AttachDispatch(book.get_Worksheets());
	//�õ�sheet  
	lpDisp = book.get_ActiveSheet();
	sheet.AttachDispatch(lpDisp);



	int length;
	int length1;
	int length2;
	for (int i = 0; i < VectorSize; i++)
	{  
		
		//�ж��Ƿ�ΪSick�ļ�
		if (vecResult[i][24] != 'S')
		{
			MessageBox("�ļ�����Sick�ļ�");
			continue;                           //��������ѭ����������һ���ļ����ж�
		}
		else
		{
			strNumFile=vecResult[i];           //��ÿһ�е����ݷŵ�strNumFile��
		}

		
//***********************����־λ********************************************************

		//�жϾ��ߴ���
			if (length=strNumFile.Find("���̱�־"))
			{
				DecisionCode = strNumFile[length-3];
			}

	   //�ж�������
			if (length = strNumFile.Find("�ڷŽǶ�"))
			{
				CategoryCode = strNumFile[length - 3];

			}
		//�ж����̱�־
			if (length = strNumFile.Find("����"))
			{
				TrayMarkCode = strNumFile[length - 3];

			}
	   //�жϼ�������
			if (length = strNumFile.Find("ƥ�䵽���̵ĵ���"))
			{  
				NumberCode = strNumFile[length - 3];

			}
			//������ߴ���Ϊ0����˵�����ʧ�ܣ�������һ�еĶ���
			if (DecisionCode == '0')
			{
				//MessageBox("���ʧ��");
				continue;
			}
      

//***********************����־λ������********************************************************



//***********************�Ա�־λ�����жϲ����ԭʼ���ݵ��ļ����Լ������ļ���ʱ��******************

         


  //�ҵ�ÿһ��ԭʼ�ļ��е�ʱ�䣬�����FileName��
		for (int j = 40; j < 46; j++)
		{
			char Name1 = vecResult[i][j];
            FileName.push_back(Name1);         //�õ��ļ���ʱ��	

		}
  //��ȡһ���е������ļ���
		length1 = strNumFile.Find("S");
		length2 = strNumFile.Find("���ߴ���"); 
        LaserFileName = strNumFile.Mid(length1, (length2 - length1 - 1));
		
  //�����������������ļ����Լ���ƥ����ļ���
		CString INum;    
		INum.Format(_T("%d"), i+1);
		
		GetDlgItem(IDC_EDIT5)->SetWindowText("���ڴ����" + INum + "�ļ�");
//***********************�Ա�־λ�����жϲ����ԭʼ���ݵ��ļ����Լ������ļ���ʱ��******************		
		GetDlgItem(IDC_EDIT6)->SetWindowText(LaserFileName);
		
		//ƥ����Ӧ��ͼƬ
		BaggageTypeJudgement(CategoryCode, TrayMarkCode, NumberCode);
		UpdateWindow();
	   
		MatchingPicture(strFilePathTxt1, FileName);
		

		//��ȡ�༭�����ַ������͸��Ӵ���
		GetDlgItem(IDC_EDIT1)->GetWindowText(Category);
		GetDlgItem(IDC_EDIT2)->GetWindowText(TrayMark);
		GetDlgItem(IDC_EDIT3)->GetWindowText(Number);

		DataStatisticsDlg1 dlg;                             //�����ӶԻ���
		dlg.DoModal();
		
		//�ж��Ƿ�Ҫ�˳�����
		if (Exit1 == TRUE)
		{
			
			break;
		}

		//���ñ������ݵ�EXCEL���ļ�
		SaveDataToEXCEL(strFilePathTxt1,vecResult[i], BaggageColor1, PasteStick1, BaggageDiscrimination,
			            ReviseCategory1, ReviseTrayMark1, ReviseNumber1);
	
		//һЩ����ÿһ��ѭ����Ҫ������ʼ����Ҫ��Ȼ�����������
	    FileName.clear();                                   //ʱ����������
		frame = 0;                                          //ͼ������
		ReviseCategory1 = "";                               //�޸ĺ����������
		ReviseTrayMark1 = "";                               //�޸ĺ����������
		ReviseNumber1 = "";                                 //�޸ĺ�ļ���
		Exit1 = FALSE;
	}

	//�ͷ�EXCEL���Ķ���      
	range.ReleaseDispatch();
	sheet.ReleaseDispatch();
	sheets.ReleaseDispatch();
	book.ReleaseDispatch();
	books.ReleaseDispatch();
	app.Quit();
	app.ReleaseDispatch();
}

//�����ж��������ࡢ�Ƿ��������Լ�����
int CDataStatisticsDlg::BaggageTypeJudgement(char ch1, char ch2, char ch3)
{
	
	//aa1.Format(_T("����"), );
	//�����ж�������
	switch (ch1)
	{
	case '1' :
		
		GetDlgItem(IDC_EDIT1)->SetWindowText("������");
		
	    break;
	case '0':
		GetDlgItem(IDC_EDIT1)->SetWindowText(" ");
		break;
	case '3':
		GetDlgItem(IDC_EDIT1)->SetWindowText("���");
		break;
	default:
		GetDlgItem(IDC_EDIT1)->SetWindowText("�����ʧ��");
		break;
	}
	
	//�����ж����̱�־
	switch (ch2)
	{
	case '0':
		GetDlgItem(IDC_EDIT2)->SetWindowText("��");
		Sleep(100);
		break;
	default:
		GetDlgItem(IDC_EDIT2)->SetWindowText("��");
		break;
	}
	Sleep(100);
	//�����Ƿ�������
	switch (ch3)
	{
	case '0':
		GetDlgItem(IDC_EDIT3)->SetWindowText("�������ʧ��");
		
		break;
	case '1':
		GetDlgItem(IDC_EDIT3)->SetWindowText("����");
		Sleep(100);
		break;
	default:
		GetDlgItem(IDC_EDIT3)->SetWindowText("���");
		break;
	}
	Sleep(100);
	return 1;
}

//���ݵ�ַ��ƥ����Ӧ��ͼƬ,��һ�������ַ���ڶ�����������ʱ��
void CDataStatisticsDlg::MatchingPicture(CString strFile, vector<char>FileName)
{   

	int H = (FileName[0] - '0') * 10 + (FileName[1] - '0');
	int M = (FileName[2] - '0') * 10 + (FileName[3] - '0');
	int S = (FileName[4] - '0') * 10 + (FileName[5] - '0');
	double Times = (H * 60 + M) * 60 + S;                    //��λs

	int nPos = strFile.ReverseFind(_T('\\'));
	CString strFile1 = strFile.Left(nPos + 1);
	CString path_suffix = strFile1;
	path_suffix += L"\\*.*";
	CFileFind finder;                                        //����һ���ļ���Ŀ¼
	BOOL bWorking = finder.FindFile(path_suffix);            //bWorking�ж�Ŀ���ļ����µ��ļ��Ƿ�ִ�����
	while (bWorking)
	{
		bWorking = finder.FindNextFile();
		if (finder.IsDots())                                 //�жϸú��������жϲ��ҵ��ļ������Ƿ��ǡ�.������..������0��ʾ�ǣ�0��ʾ����
			                                               //������ִ����FindNextFile()��ú�������ִ�гɹ�
		{
			continue;
		}
		else
		{
			String strFileName = (LPCTSTR)finder.GetFileName();
			char s_FileName[1024];
			strcpy(s_FileName, strFileName.c_str());
			if (s_FileName[0] == 'S'&&s_FileName[1] == 'B'&&s_FileName[2] == 'T'&&s_FileName[11]=='B')//�Ƿ�ΪBasler���������
			{
				int H1 = (s_FileName[17] - '0') * 10 + (s_FileName[18] - '0');
				int M1 = (s_FileName[19] - '0') * 10 + (s_FileName[20] - '0');
				int S1 = (s_FileName[21] - '0') * 10 + (s_FileName[22] - '0');
				double Times1 =  (H1 * 60 + M1) * 60 + S1;                    //��λs
				if (Times1 - Times <= 5 && Times1 - Times>=0)
				{
					
					strFilePathPicture = "";
					String str12;
					str12 = strFile1.GetBuffer(0);
					strFilePathPicture = str12 + strFileName;
					frame = imread(strFilePathPicture);  //��ȡ��ǰ֡
					ShowImage(frame, IDC_STATIC_PIC1);                                                        //��Picture�ؼ�����ʾ��ǰ�ɼ���ͼ��
					waitKey(30);
					CString str(strFileName.c_str());
					GetDlgItem(IDC_EDIT4)->SetWindowText(str);
					//ShowImage();
				}
			}
		}

	}
}




//�˳�����
void CDataStatisticsDlg::OnClickedButtonExit()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	if (MessageBox("ȷ���˳�", "", MB_YESNO | MB_ICONQUESTION) != IDYES)
	{

		return;
	}
	CDialog::OnOK();
	AfxGetMainWnd()->PostMessage(WM_CLOSE, 0, 0);
}

//�������ݵ�EXCEL
void CDataStatisticsDlg::SaveDataToEXCEL(CString saveFilePath, CString WriteContents, CString BaggageColor1, CString PasteStick1, BOOL BaggageDiscrimination,
	CString ReviseCategory1, CString ReviseTrayMark1, CString ReviseNumber1)
{
	

	

	

	//����EXCELӦ��
	



	//ȡ���û���  �����ʹ�ù�������
	CRange userRange;
	userRange.AttachDispatch(sheet.get_UsedRange());
	//�õ��û���������  
	range.AttachDispatch(userRange.get_Rows());
	long rowNum = range.get_Count();
	//�õ��û���������  
	range.AttachDispatch(userRange.get_Columns());
	long colNum = range.get_Count();
	//�õ��û����Ŀ�ʼ�кͿ�ʼ��  
	long startRow = userRange.get_Row();
	long startCol = userRange.get_Column();

	//�õ���ʼ��EXCEL��д����к���
	long startWriteRow = startRow + rowNum;                   //��
	long startWriteCol = startCol + colNum;                   //��

	//�Զ������к�����
	
	
	/*
	cols = range.get_EntireColumn();
	cols.AutoFit();
	rows = range.get_EntireRow();
	rows.AutoFit();

   */
//*********��txt��һ�������ļ�����ȡ��Ϣ��д�뵽EXCEL��ȥ********************************************

	//*********���ҵ��ļ���******************************
	CString writeLocationA;                                                              //д���EXCEL��A��Ԫ��λ��
	int sequence1 = strNumFile.Find("S");
	int sequence2 = strNumFile.Find("���ߴ���");
	LaserFileName = strNumFile.Mid(sequence1, (sequence2 - sequence1 - 1));
    writeLocationA.Format("A%d", startWriteRow);                                         //����к�A��
	range.put_ColumnWidth(COleVariant("40"));
	range = sheet.get_Range(COleVariant(writeLocationA), COleVariant(writeLocationA));   //�õ�����
	range.put_Value2(COleVariant(LaserFileName));                                        //д������
	

//*********��EXCEL�е�B���в�����Ӧ��ͼƬ********************************************
	CString writeLocationB;
	shapes = sheet.get_Shapes();
	
	writeLocationB.Format("B%d", startWriteRow);
	
    range = sheet.get_Range(COleVariant(writeLocationB), COleVariant(writeLocationB));
	LPCSTR strFilePathPicture1 = strFilePathPicture.c_str();
	range.put_ColumnWidth(COleVariant("20"));
	range.put_RowHeight(COleVariant("70"));
	shapes.AddPicture(strFilePathPicture1, false, true, (float)range.get_Left().dblVal, (float)range.get_Top().dblVal, (float)range.get_Width().dblVal, (float)range.get_Height().dblVal);
	sRange = shapes.get_Range(_variant_t(long(1)));
	sRange.put_Height(float(70));
	sRange.put_Width(float(20));



	//************���ߴ���********************************
    //��дC��Ԫ��Ϊ�ַ��������ߴ��롱
	CString writeLocationC;                                                              //д���EXCEL��C��Ԫ��λ��
	CString writeLocationD;                                                              //д���EXCEL��D��Ԫ��λ��
	CString DecisionCode1;
	writeLocationC.Format("C%d", startWriteRow);                                         //����к�C��
	range = sheet.get_Range(COleVariant(writeLocationC), COleVariant(writeLocationC));   //�õ�����
	range.put_ColumnWidth(COleVariant("8"));
	range.put_Value2(COleVariant("���ߴ���"));

	//��дD��Ԫ��Ϊ���ߴ���
	writeLocationD.Format("D%d", startWriteRow);                                         //����к�C��
	range = sheet.get_Range(COleVariant(writeLocationD), COleVariant(writeLocationD));   //�õ�����
	range.put_ColumnWidth(COleVariant("5"));
	if (int length = strNumFile.Find("���̱�־"))
	{
		DecisionCode1 = strNumFile[length - 3];
	}
	range.put_Value2(COleVariant(DecisionCode1));

   //****************���̱�־********************************
	//��дE��λΪ�ַ��������̱�־��
	CString writeLocationE;                                                              //д���EXCEL��C��Ԫ��λ��
	CString writeLocationF;                                                              //д���EXCEL��D��Ԫ��λ��
	CString TrayMarkCode1;
	writeLocationE.Format("E%d", startWriteRow);                                         //����к�C��
	range = sheet.get_Range(COleVariant(writeLocationE), COleVariant(writeLocationE));   //�õ�����
	range.put_ColumnWidth(COleVariant("8"));
	range.put_Value2(COleVariant("���̱�־"));

	//��дF��λ�����̱�־
	writeLocationF.Format("F%d", startWriteRow);                                         //����к�C��
	range = sheet.get_Range(COleVariant(writeLocationF), COleVariant(writeLocationF));   //�õ�����
	range.put_ColumnWidth(COleVariant("5"));
	if (int length = strNumFile.Find("����"))
	{
		TrayMarkCode1 = strNumFile[length - 3];

	}
	range.put_Value2(COleVariant(TrayMarkCode1));

   //****************����********************************
   // ��дG��λΪ�ַ��������̱�־��
	CString writeLocationG;                                                              //д���EXCEL��C��Ԫ��λ��
	CString writeLocationH;                                                              //д���EXCEL��D��Ԫ��λ��
	CString NumberCode1;
	writeLocationG.Format("G%d", startWriteRow);                                         //����к�C��
	range = sheet.get_Range(COleVariant(writeLocationG), COleVariant(writeLocationG));   //�õ�����
	range.put_ColumnWidth(COleVariant("8"));
	range.put_Value2(COleVariant("����"));
	
	//��дF��λ�ļ�������
	//�жϼ�������
	writeLocationH.Format("H%d", startWriteRow);                                         //����к�C��
	range = sheet.get_Range(COleVariant(writeLocationH), COleVariant(writeLocationH));   //�õ�����
	range.put_ColumnWidth(COleVariant("5"));
	if (int length = strNumFile.Find("ƥ�䵽���̵ĵ���"))
	{
		NumberCode1 = strNumFile[length - 3];

	}
	range.put_Value2(COleVariant(NumberCode1));
	
	////****************ƥ�䵽�ĵ���********************************
	CString writeLocationI;                                                              //д���EXCEL��C��Ԫ��λ��
	CString writeLocationJ;                                                              //д���EXCEL��D��Ԫ��λ��
	CString MatchingPoint;
	writeLocationI.Format("I%d", startWriteRow);                                         //����к�C��
	range = sheet.get_Range(COleVariant(writeLocationI), COleVariant(writeLocationI));   //�õ�����
	range.put_ColumnWidth(COleVariant("16"));
	range.put_Value2(COleVariant("ƥ�䵽���̵ĵ���"));

	//��дJ��λ����
	writeLocationJ.Format("J%d", startWriteRow);                                         //����к�C��
	range = sheet.get_Range(COleVariant(writeLocationJ), COleVariant(writeLocationJ));   //�õ�����
	range.put_ColumnWidth(COleVariant("10"));
	
	if (int length = strNumFile.Find("������"))
	{


		int length1 = strNumFile.Find("ƥ�䵽���̵ĵ���");
		for (int k = (length1 + 18); k < (length - 2); k++)
		{
			
			if ((strNumFile[k] - '0') >= 0 && (strNumFile[k] - '0') <= 9)
				MatchingPoint = MatchingPoint + strNumFile[k];

		}


	}
	range.put_Value2(COleVariant(MatchingPoint));

	//****************������********************************
	// ��дG��λΪ�ַ����������롱
	CString writeLocationK;                                                              //д���EXCEL��C��Ԫ��λ��
	CString writeLocationL;                                                              //д���EXCEL��D��Ԫ��λ��
	CString CategoryCode1;
	writeLocationK.Format("K%d", startWriteRow);                                         //����к�C��
	range = sheet.get_Range(COleVariant(writeLocationK), COleVariant(writeLocationK));   //�õ�����
	range.put_ColumnWidth(COleVariant("8"));
	range.put_Value2(COleVariant("������"));
	
	//��дJ��λ��������
	//�ж�������
	writeLocationL.Format("L%d", startWriteRow);                                         //����к�C��
	range = sheet.get_Range(COleVariant(writeLocationL), COleVariant(writeLocationL));   //�õ�����
	range.put_ColumnWidth(COleVariant("5"));
	if (int length = strNumFile.Find("�ڷŽǶ�"))
	{
		CategoryCode1 = strNumFile[length - 3];

	}
	range.put_Value2(COleVariant(CategoryCode1));


	//****************�ڷŽǶ�********************************
	//��дM��λ������
	CString writeLocationM;                                                              //д���EXCEL��C��Ԫ��λ��
	CString writeLocationN;                                                              //д���EXCEL��D��Ԫ��λ��
	CString PlacementAngle;
	writeLocationM.Format("M%d", startWriteRow);                                         //����к�C��
	range = sheet.get_Range(COleVariant(writeLocationM), COleVariant(writeLocationM));   //�õ�����
	range.put_ColumnWidth(COleVariant("8"));
	range.put_Value2(COleVariant("�ڷŽǶ�"));

	//��дN��λ������
	writeLocationN.Format("N%d", startWriteRow);                                         //����к�C��
	range = sheet.get_Range(COleVariant(writeLocationN), COleVariant(writeLocationN));   //�õ�����
	range.put_ColumnWidth(COleVariant("5"));
	if (int length = strNumFile.Find("��������"))
	{

		
		int length1 = strNumFile.Find("�ڷŽǶ�");
		for (int k = (length1 + 10); k < (length - 2); k++)
		{

			if ((strNumFile[k] - '0') >= 0 && (strNumFile[k] - '0') <= 9)
				PlacementAngle = PlacementAngle + strNumFile[k];
		}
	}
	range.put_Value2(COleVariant(PlacementAngle));


   //****************��������********************************
	//��дO��λ������
	CString writeLocationO;                                                              //д���EXCEL��C��Ԫ��λ��
	CString writeLocationP;                                                              //д���EXCEL��D��Ԫ��λ��
	CString ReturnType;
	writeLocationO.Format("O%d", startWriteRow);                                         //����к�C��
	range = sheet.get_Range(COleVariant(writeLocationO), COleVariant(writeLocationO));   //�õ�����
	range.put_ColumnWidth(COleVariant("8"));
	range.put_Value2(COleVariant("�ڷŽǶ�"));
	//��дP��λ����
	writeLocationP.Format("P%d", startWriteRow);                                         //����к�C��
	range = sheet.get_Range(COleVariant(writeLocationP), COleVariant(writeLocationP));   //�õ�����
	range.put_ColumnWidth(COleVariant("5"));
	int len3 = 0;
	len3 = strNumFile.Find("��������");
	ReturnType = strNumFile.Mid((len3 + 10));
	range.put_Value2(COleVariant(ReturnType));

	

	//****************��ɫ���Ƿ��п�ճ������********************************
	// ��дK��λΪ�ַ�������ɫ��
	CString writeLocationQ;                                                              //д���EXCEL��C��Ԫ��λ��
	CString writeLocationR;                                                              //д���EXCEL��D��Ԫ��λ��
	CString writeLocationS;                                                              //д���EXCEL��C��Ԫ��λ��
	writeLocationQ.Format("Q%d", startWriteRow);                                         //����к�C��
	range = sheet.get_Range(COleVariant(writeLocationQ), COleVariant(writeLocationQ));   //�õ�����
	range.put_ColumnWidth(COleVariant("8"));
	range.put_Value2(COleVariant(BaggageColor1));

	//д���Ƿ��п�ճ������
	writeLocationR.Format("R%d", startWriteRow);                                         //����к�C��
	range = sheet.get_Range(COleVariant(writeLocationR), COleVariant(writeLocationR));   //�õ�����
	range.put_ColumnWidth(COleVariant("10"));
	range.put_Value2(COleVariant("��ճ������"));

	//�ж��Ƿ��п�ճ������
	writeLocationS.Format("S%d", startWriteRow);                                         //����к�C��
	range = sheet.get_Range(COleVariant(writeLocationS), COleVariant(writeLocationS));   //�õ�����
	range.put_ColumnWidth(COleVariant("8"));
	if (PasteStick1 == "�п�ճ������")
	{
		range.put_Value2(COleVariant("1"));
	}
	else
	{
		range.put_Value2(COleVariant("0"));
	}

	
	
 
	//****************����������ͣ����̺ͼ���********************************
	CString writeLocationV;                                                              //д���EXCEL��C��Ԫ��λ��
	CString writeLocationT;                                                              //д���EXCEL��D��Ԫ��λ��
	CString writeLocationU;                                                              //д���EXCEL��D��Ԫ��λ��
	
	//�ж��Ƿ�������޸ģ�����������޸ģ�����޸ĺõ�������д�뵽EXCEL��ȥ
	if (ReviseCategory1 != "")                                                        
	{
		writeLocationV.Format("V%d", startWriteRow);                                         //����к�C��
		range = sheet.get_Range(COleVariant(writeLocationV), COleVariant(writeLocationV));   //�õ�����
		range.put_ColumnWidth(COleVariant("8"));
		range.put_Value2(COleVariant(ReviseCategory1));

	}
	if (ReviseTrayMark1 != "")
	{
		writeLocationT.Format("T%d", startWriteRow);                                         //����к�C��
		range = sheet.get_Range(COleVariant(writeLocationT), COleVariant(writeLocationT));   //�õ�����
		range.put_ColumnWidth(COleVariant("8"));
		range.put_Value2(COleVariant(ReviseTrayMark1));
    }
	if (ReviseNumber1 != "")
	{
		writeLocationU.Format("U%d", startWriteRow);                                         //����к�C��
		range = sheet.get_Range(COleVariant(writeLocationU), COleVariant(writeLocationU));   //�õ�����
		range.put_ColumnWidth(COleVariant("8"));
		range.put_Value2(COleVariant(ReviseNumber1));
	}



	//****************�ж��Ƿ���󣬴���Ļ��ú�ɫ��־��һ��********************************
	if (BaggageDiscrimination == FALSE)
	{
		CString writeLocation1;
		CString writeLocation2;
		writeLocation1.Format("A%d", startWriteRow);
		writeLocation2.Format("Z%d", startWriteRow);
		range.AttachDispatch(sheet.get_Range(COleVariant(writeLocation1), COleVariant(writeLocation2)));
		range.put_ColumnWidth(COleVariant("10"));
		it.AttachDispatch(range.get_Interior());
		it.put_Color(COleVariant(long(255)));

	}

	


	book.Save();
	
	
	
}
