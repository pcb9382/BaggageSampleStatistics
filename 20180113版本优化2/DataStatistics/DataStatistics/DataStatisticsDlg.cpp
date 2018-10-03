
// DataStatisticsDlg.cpp : 实现文件
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

#define MAX_FILE_LEN 250                                          //定义读取文件一行的最大字符数
// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
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


// CDataStatisticsDlg 对话框



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


// CDataStatisticsDlg 消息处理程序

BOOL CDataStatisticsDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
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

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO:  在此添加额外的初始化代码
	FileName.clear();
    frame = 0;	
	Exit1 = FALSE;
	ReviseCategory1="";
	ReviseTrayMark1="";
    ReviseNumber1="";
	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
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

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CDataStatisticsDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CDataStatisticsDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

//用于在PictureBox框中显示图片
void CDataStatisticsDlg::ShowImage(Mat Img, UINT ID)
{
	//把Mat转换成IplImage*的形式用于在控件中显示
	
	IplImage* pBinary = &IplImage(Img);                                                          //浅拷贝
	IplImage *img = cvCloneImage(pBinary);                                                       //由浅拷贝转换为深拷贝

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


//读入需要处理的txt数据文件
void CDataStatisticsDlg::OnBnClickedButton3()
{
	// TODO:  在此添加控件通知处理程序代码
	
	
	
	//加载图片的路径
	CFileDialog dlg(TRUE, NULL, NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, "All Files|*.*|txt|*.txt||");
	INT_PTR nRet = dlg.DoModal();
	if (nRet != IDOK)
	{
		return;
	}

	strFilePathTxt1 = dlg.GetPathName();


	if (!PathFileExists(strFilePathTxt1))
	{
		MessageBox("文件不存在！");
		return;
	}
	strFilePathTxt = strFilePathTxt1.GetBuffer(0);

	//读取的文件不存在
	if (!PathFileExists(strFilePathTxt1))
	{
		return ;
	}

	CStdioFile file;
	//判断是否正确打开txt文件
	if (!file.Open(strFilePathTxt1, CFile::modeRead))
	{
		return ;
	}

	//用于存储读取结果的vector变量
	
	//创建一个读取的变量
	CString strValue = _T("");

	

	//读取txt中的所有的内容到向量vector中
	while (file.ReadString(strValue))
	{
		vecResult.push_back(strValue);
	}
	//关闭txt文件
	file.Close();
   

    //开始处理读入的文本内容，找出原始文件的数据
	int VectorSize = vecResult.size();             //获取向量的大小
	

	//这里开始打开EXCEL表格

	int nPos = strFilePathTxt1.ReverseFind(_T('\\'));
	CString strFile1 = strFilePathTxt1.Left(nPos + 1);
	CString m_saveFilePath1 = strFile1;


	CString m_saveFilePath = strFile1 + "数据统计5.xlsx";

	LPDISPATCH lpDisp;
	COleVariant  covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if (!app.CreateDispatch("Excel.Application"))
	{
		AfxMessageBox("无法打开Excel应用", MB_OK | MB_ICONWARNING);
		return;
	}
	books.AttachDispatch(app.get_Workbooks());
	//打开文件  
	lpDisp = books.Open(m_saveFilePath, covOptional, covOptional, covOptional, covOptional, covOptional
		, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional);
	//得到得到视图workbook  
	book.AttachDispatch(lpDisp);
	//得到worksheets  
	sheets.AttachDispatch(book.get_Worksheets());
	//得到sheet  
	lpDisp = book.get_ActiveSheet();
	sheet.AttachDispatch(lpDisp);



	int length;
	int length1;
	int length2;
	for (int i = 0; i < VectorSize; i++)
	{  
		
		//判断是否为Sick文件
		if (vecResult[i][24] != 'S')
		{
			MessageBox("文件不是Sick文件");
			continue;                           //跳出本次循环，进行下一行文件的判断
		}
		else
		{
			strNumFile=vecResult[i];           //把每一行的内容放到strNumFile中
		}

		
//***********************检测标志位********************************************************

		//判断决策代码
			if (length=strNumFile.Find("托盘标志"))
			{
				DecisionCode = strNumFile[length-3];
			}

	   //判断类别代码
			if (length = strNumFile.Find("摆放角度"))
			{
				CategoryCode = strNumFile[length - 3];

			}
		//判断托盘标志
			if (length = strNumFile.Find("件数"))
			{
				TrayMarkCode = strNumFile[length - 3];

			}
	   //判断件数代码
			if (length = strNumFile.Find("匹配到托盘的点数"))
			{  
				NumberCode = strNumFile[length - 3];

			}
			//如果决策代码为0，则说明检测失败，进行下一行的读入
			if (DecisionCode == '0')
			{
				//MessageBox("检测失败");
				continue;
			}
      

//***********************检测标志位检测完毕********************************************************



//***********************对标志位进行判断并获得原始数据的文件名以及创建文件的时间******************

         


  //找到每一个原始文件中的时间，存放在FileName中
		for (int j = 40; j < 46; j++)
		{
			char Name1 = vecResult[i][j];
            FileName.push_back(Name1);         //得到文件的时间	

		}
  //提取一行中的数据文件名
		length1 = strNumFile.Find("S");
		length2 = strNumFile.Find("决策代码"); 
        LaserFileName = strNumFile.Mid(length1, (length2 - length1 - 1));
		
  //输出处理个数，激光文件名以及相匹配的文件名
		CString INum;    
		INum.Format(_T("%d"), i+1);
		
		GetDlgItem(IDC_EDIT5)->SetWindowText("正在处理第" + INum + "文件");
//***********************对标志位进行判断并获得原始数据的文件名以及创建文件的时间******************		
		GetDlgItem(IDC_EDIT6)->SetWindowText(LaserFileName);
		
		//匹配相应的图片
		BaggageTypeJudgement(CategoryCode, TrayMarkCode, NumberCode);
		UpdateWindow();
	   
		MatchingPicture(strFilePathTxt1, FileName);
		

		//获取编辑框中字符串发送给子窗口
		GetDlgItem(IDC_EDIT1)->GetWindowText(Category);
		GetDlgItem(IDC_EDIT2)->GetWindowText(TrayMark);
		GetDlgItem(IDC_EDIT3)->GetWindowText(Number);

		DataStatisticsDlg1 dlg;                             //建立子对话框
		dlg.DoModal();
		
		//判断是否要退出程序
		if (Exit1 == TRUE)
		{
			
			break;
		}

		//调用保存数据到EXCEL的文件
		SaveDataToEXCEL(strFilePathTxt1,vecResult[i], BaggageColor1, PasteStick1, BaggageDiscrimination,
			            ReviseCategory1, ReviseTrayMark1, ReviseNumber1);
	
		//一些变量每一次循环需要把它初始化，要不然程序会有问题
	    FileName.clear();                                   //时间向量清零
		frame = 0;                                          //图像清零
		ReviseCategory1 = "";                               //修改后的行李种类
		ReviseTrayMark1 = "";                               //修改后的有无托盘
		ReviseNumber1 = "";                                 //修改后的件数
		Exit1 = FALSE;
	}

	//释放EXCEL表格的对象      
	range.ReleaseDispatch();
	sheet.ReleaseDispatch();
	sheets.ReleaseDispatch();
	book.ReleaseDispatch();
	books.ReleaseDispatch();
	app.Quit();
	app.ReleaseDispatch();
}

//用于判断行李种类、是否有托盘以及件数
int CDataStatisticsDlg::BaggageTypeJudgement(char ch1, char ch2, char ch3)
{
	
	//aa1.Format(_T("单件"), );
	//用于判断类别代码
	switch (ch1)
	{
	case '1' :
		
		GetDlgItem(IDC_EDIT1)->SetWindowText("拉杆箱");
		
	    break;
	case '0':
		GetDlgItem(IDC_EDIT1)->SetWindowText(" ");
		break;
	case '3':
		GetDlgItem(IDC_EDIT1)->SetWindowText("软包");
		break;
	default:
		GetDlgItem(IDC_EDIT1)->SetWindowText("类别检测失败");
		break;
	}
	
	//用于判断托盘标志
	switch (ch2)
	{
	case '0':
		GetDlgItem(IDC_EDIT2)->SetWindowText("无");
		Sleep(100);
		break;
	default:
		GetDlgItem(IDC_EDIT2)->SetWindowText("有");
		break;
	}
	Sleep(100);
	//用于是否多件行李
	switch (ch3)
	{
	case '0':
		GetDlgItem(IDC_EDIT3)->SetWindowText("件数检测失败");
		
		break;
	case '1':
		GetDlgItem(IDC_EDIT3)->SetWindowText("单件");
		Sleep(100);
		break;
	default:
		GetDlgItem(IDC_EDIT3)->SetWindowText("多件");
		break;
	}
	Sleep(100);
	return 1;
}

//根据地址，匹配相应的图片,第一个输入地址，第二个参数输入时间
void CDataStatisticsDlg::MatchingPicture(CString strFile, vector<char>FileName)
{   

	int H = (FileName[0] - '0') * 10 + (FileName[1] - '0');
	int M = (FileName[2] - '0') * 10 + (FileName[3] - '0');
	int S = (FileName[4] - '0') * 10 + (FileName[5] - '0');
	double Times = (H * 60 + M) * 60 + S;                    //单位s

	int nPos = strFile.ReverseFind(_T('\\'));
	CString strFile1 = strFile.Left(nPos + 1);
	CString path_suffix = strFile1;
	path_suffix += L"\\*.*";
	CFileFind finder;                                        //遍历一个文件的目录
	BOOL bWorking = finder.FindFile(path_suffix);            //bWorking判断目标文件夹下的文件是否执行完毕
	while (bWorking)
	{
		bWorking = finder.FindNextFile();
		if (finder.IsDots())                                 //判断该函数用来判断查找的文件属性是否是“.”，“..”，非0表示是，0表示不是
			                                               //必须在执行了FindNextFile()后该函数才能执行成功
		{
			continue;
		}
		else
		{
			String strFileName = (LPCTSTR)finder.GetFileName();
			char s_FileName[1024];
			strcpy(s_FileName, strFileName.c_str());
			if (s_FileName[0] == 'S'&&s_FileName[1] == 'B'&&s_FileName[2] == 'T'&&s_FileName[11]=='B')//是否为Basler相机的数据
			{
				int H1 = (s_FileName[17] - '0') * 10 + (s_FileName[18] - '0');
				int M1 = (s_FileName[19] - '0') * 10 + (s_FileName[20] - '0');
				int S1 = (s_FileName[21] - '0') * 10 + (s_FileName[22] - '0');
				double Times1 =  (H1 * 60 + M1) * 60 + S1;                    //单位s
				if (Times1 - Times <= 5 && Times1 - Times>=0)
				{
					
					strFilePathPicture = "";
					String str12;
					str12 = strFile1.GetBuffer(0);
					strFilePathPicture = str12 + strFileName;
					frame = imread(strFilePathPicture);  //读取当前帧
					ShowImage(frame, IDC_STATIC_PIC1);                                                        //在Picture控件中显示当前采集的图像
					waitKey(30);
					CString str(strFileName.c_str());
					GetDlgItem(IDC_EDIT4)->SetWindowText(str);
					//ShowImage();
				}
			}
		}

	}
}




//退出程序
void CDataStatisticsDlg::OnClickedButtonExit()
{
	// TODO:  在此添加控件通知处理程序代码
	if (MessageBox("确认退出", "", MB_YESNO | MB_ICONQUESTION) != IDYES)
	{

		return;
	}
	CDialog::OnOK();
	AfxGetMainWnd()->PostMessage(WM_CLOSE, 0, 0);
}

//保存数据到EXCEL
void CDataStatisticsDlg::SaveDataToEXCEL(CString saveFilePath, CString WriteContents, CString BaggageColor1, CString PasteStick1, BOOL BaggageDiscrimination,
	CString ReviseCategory1, CString ReviseTrayMark1, CString ReviseNumber1)
{
	

	

	

	//创建EXCEL应用
	



	//取得用户区  获得已使用过得区域
	CRange userRange;
	userRange.AttachDispatch(sheet.get_UsedRange());
	//得到用户区的行数  
	range.AttachDispatch(userRange.get_Rows());
	long rowNum = range.get_Count();
	//得到用户区的列数  
	range.AttachDispatch(userRange.get_Columns());
	long colNum = range.get_Count();
	//得到用户区的开始行和开始列  
	long startRow = userRange.get_Row();
	long startCol = userRange.get_Column();

	//得到开始向EXCEL中写入的行和列
	long startWriteRow = startRow + rowNum;                   //行
	long startWriteCol = startCol + colNum;                   //列

	//自动调整行和列列
	
	
	/*
	cols = range.get_EntireColumn();
	cols.AutoFit();
	rows = range.get_EntireRow();
	rows.AutoFit();

   */
//*********从txt的一行数据文件中提取信息，写入到EXCEL中去********************************************

	//*********先找到文件名******************************
	CString writeLocationA;                                                              //写入的EXCEL的A单元格位置
	int sequence1 = strNumFile.Find("S");
	int sequence2 = strNumFile.Find("决策代码");
	LaserFileName = strNumFile.Mid(sequence1, (sequence2 - sequence1 - 1));
    writeLocationA.Format("A%d", startWriteRow);                                         //获得行号A几
	range.put_ColumnWidth(COleVariant("40"));
	range = sheet.get_Range(COleVariant(writeLocationA), COleVariant(writeLocationA));   //得到坐标
	range.put_Value2(COleVariant(LaserFileName));                                        //写入数据
	

//*********向EXCEL中的B列中插入相应的图片********************************************
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



	//************决策代码********************************
    //先写C单元格为字符串“决策代码”
	CString writeLocationC;                                                              //写入的EXCEL的C单元格位置
	CString writeLocationD;                                                              //写入的EXCEL的D单元格位置
	CString DecisionCode1;
	writeLocationC.Format("C%d", startWriteRow);                                         //获得行号C几
	range = sheet.get_Range(COleVariant(writeLocationC), COleVariant(writeLocationC));   //得到坐标
	range.put_ColumnWidth(COleVariant("8"));
	range.put_Value2(COleVariant("决策代码"));

	//再写D单元格为决策代码
	writeLocationD.Format("D%d", startWriteRow);                                         //获得行号C几
	range = sheet.get_Range(COleVariant(writeLocationD), COleVariant(writeLocationD));   //得到坐标
	range.put_ColumnWidth(COleVariant("5"));
	if (int length = strNumFile.Find("托盘标志"))
	{
		DecisionCode1 = strNumFile[length - 3];
	}
	range.put_Value2(COleVariant(DecisionCode1));

   //****************托盘标志********************************
	//先写E单位为字符串“托盘标志”
	CString writeLocationE;                                                              //写入的EXCEL的C单元格位置
	CString writeLocationF;                                                              //写入的EXCEL的D单元格位置
	CString TrayMarkCode1;
	writeLocationE.Format("E%d", startWriteRow);                                         //获得行号C几
	range = sheet.get_Range(COleVariant(writeLocationE), COleVariant(writeLocationE));   //得到坐标
	range.put_ColumnWidth(COleVariant("8"));
	range.put_Value2(COleVariant("托盘标志"));

	//再写F单位的托盘标志
	writeLocationF.Format("F%d", startWriteRow);                                         //获得行号C几
	range = sheet.get_Range(COleVariant(writeLocationF), COleVariant(writeLocationF));   //得到坐标
	range.put_ColumnWidth(COleVariant("5"));
	if (int length = strNumFile.Find("件数"))
	{
		TrayMarkCode1 = strNumFile[length - 3];

	}
	range.put_Value2(COleVariant(TrayMarkCode1));

   //****************件数********************************
   // 先写G单位为字符串“托盘标志”
	CString writeLocationG;                                                              //写入的EXCEL的C单元格位置
	CString writeLocationH;                                                              //写入的EXCEL的D单元格位置
	CString NumberCode1;
	writeLocationG.Format("G%d", startWriteRow);                                         //获得行号C几
	range = sheet.get_Range(COleVariant(writeLocationG), COleVariant(writeLocationG));   //得到坐标
	range.put_ColumnWidth(COleVariant("8"));
	range.put_Value2(COleVariant("件数"));
	
	//再写F单位的件数代码
	//判断件数代码
	writeLocationH.Format("H%d", startWriteRow);                                         //获得行号C几
	range = sheet.get_Range(COleVariant(writeLocationH), COleVariant(writeLocationH));   //得到坐标
	range.put_ColumnWidth(COleVariant("5"));
	if (int length = strNumFile.Find("匹配到托盘的点数"))
	{
		NumberCode1 = strNumFile[length - 3];

	}
	range.put_Value2(COleVariant(NumberCode1));
	
	////****************匹配到的点数********************************
	CString writeLocationI;                                                              //写入的EXCEL的C单元格位置
	CString writeLocationJ;                                                              //写入的EXCEL的D单元格位置
	CString MatchingPoint;
	writeLocationI.Format("I%d", startWriteRow);                                         //获得行号C几
	range = sheet.get_Range(COleVariant(writeLocationI), COleVariant(writeLocationI));   //得到坐标
	range.put_ColumnWidth(COleVariant("16"));
	range.put_Value2(COleVariant("匹配到托盘的点数"));

	//再写J单位内容
	writeLocationJ.Format("J%d", startWriteRow);                                         //获得行号C几
	range = sheet.get_Range(COleVariant(writeLocationJ), COleVariant(writeLocationJ));   //得到坐标
	range.put_ColumnWidth(COleVariant("10"));
	
	if (int length = strNumFile.Find("类别代码"))
	{


		int length1 = strNumFile.Find("匹配到托盘的点数");
		for (int k = (length1 + 18); k < (length - 2); k++)
		{
			
			if ((strNumFile[k] - '0') >= 0 && (strNumFile[k] - '0') <= 9)
				MatchingPoint = MatchingPoint + strNumFile[k];

		}


	}
	range.put_Value2(COleVariant(MatchingPoint));

	//****************类别代码********************************
	// 先写G单位为字符串“类别代码”
	CString writeLocationK;                                                              //写入的EXCEL的C单元格位置
	CString writeLocationL;                                                              //写入的EXCEL的D单元格位置
	CString CategoryCode1;
	writeLocationK.Format("K%d", startWriteRow);                                         //获得行号C几
	range = sheet.get_Range(COleVariant(writeLocationK), COleVariant(writeLocationK));   //得到坐标
	range.put_ColumnWidth(COleVariant("8"));
	range.put_Value2(COleVariant("类别代码"));
	
	//再写J单位的类别代码
	//判断类别代码
	writeLocationL.Format("L%d", startWriteRow);                                         //获得行号C几
	range = sheet.get_Range(COleVariant(writeLocationL), COleVariant(writeLocationL));   //得到坐标
	range.put_ColumnWidth(COleVariant("5"));
	if (int length = strNumFile.Find("摆放角度"))
	{
		CategoryCode1 = strNumFile[length - 3];

	}
	range.put_Value2(COleVariant(CategoryCode1));


	//****************摆放角度********************************
	//先写M单位的数据
	CString writeLocationM;                                                              //写入的EXCEL的C单元格位置
	CString writeLocationN;                                                              //写入的EXCEL的D单元格位置
	CString PlacementAngle;
	writeLocationM.Format("M%d", startWriteRow);                                         //获得行号C几
	range = sheet.get_Range(COleVariant(writeLocationM), COleVariant(writeLocationM));   //得到坐标
	range.put_ColumnWidth(COleVariant("8"));
	range.put_Value2(COleVariant("摆放角度"));

	//再写N单位的数据
	writeLocationN.Format("N%d", startWriteRow);                                         //获得行号C几
	range = sheet.get_Range(COleVariant(writeLocationN), COleVariant(writeLocationN));   //得到坐标
	range.put_ColumnWidth(COleVariant("5"));
	if (int length = strNumFile.Find("返回类型"))
	{

		
		int length1 = strNumFile.Find("摆放角度");
		for (int k = (length1 + 10); k < (length - 2); k++)
		{

			if ((strNumFile[k] - '0') >= 0 && (strNumFile[k] - '0') <= 9)
				PlacementAngle = PlacementAngle + strNumFile[k];
		}
	}
	range.put_Value2(COleVariant(PlacementAngle));


   //****************返回类型********************************
	//先写O单位的数据
	CString writeLocationO;                                                              //写入的EXCEL的C单元格位置
	CString writeLocationP;                                                              //写入的EXCEL的D单元格位置
	CString ReturnType;
	writeLocationO.Format("O%d", startWriteRow);                                         //获得行号C几
	range = sheet.get_Range(COleVariant(writeLocationO), COleVariant(writeLocationO));   //得到坐标
	range.put_ColumnWidth(COleVariant("8"));
	range.put_Value2(COleVariant("摆放角度"));
	//再写P单位数据
	writeLocationP.Format("P%d", startWriteRow);                                         //获得行号C几
	range = sheet.get_Range(COleVariant(writeLocationP), COleVariant(writeLocationP));   //得到坐标
	range.put_ColumnWidth(COleVariant("5"));
	int len3 = 0;
	len3 = strNumFile.Find("返回类型");
	ReturnType = strNumFile.Mid((len3 + 10));
	range.put_Value2(COleVariant(ReturnType));

	

	//****************颜色和是否有可粘贴区域********************************
	// 先写K单位为字符串“颜色”
	CString writeLocationQ;                                                              //写入的EXCEL的C单元格位置
	CString writeLocationR;                                                              //写入的EXCEL的D单元格位置
	CString writeLocationS;                                                              //写入的EXCEL的C单元格位置
	writeLocationQ.Format("Q%d", startWriteRow);                                         //获得行号C几
	range = sheet.get_Range(COleVariant(writeLocationQ), COleVariant(writeLocationQ));   //得到坐标
	range.put_ColumnWidth(COleVariant("8"));
	range.put_Value2(COleVariant(BaggageColor1));

	//写入是否有可粘贴区域
	writeLocationR.Format("R%d", startWriteRow);                                         //获得行号C几
	range = sheet.get_Range(COleVariant(writeLocationR), COleVariant(writeLocationR));   //得到坐标
	range.put_ColumnWidth(COleVariant("10"));
	range.put_Value2(COleVariant("可粘贴区域"));

	//判断是否有可粘贴区域
	writeLocationS.Format("S%d", startWriteRow);                                         //获得行号C几
	range = sheet.get_Range(COleVariant(writeLocationS), COleVariant(writeLocationS));   //得到坐标
	range.put_ColumnWidth(COleVariant("8"));
	if (PasteStick1 == "有可粘贴区域")
	{
		range.put_Value2(COleVariant("1"));
	}
	else
	{
		range.put_Value2(COleVariant("0"));
	}

	
	
 
	//****************改正后的类型，托盘和件数********************************
	CString writeLocationV;                                                              //写入的EXCEL的C单元格位置
	CString writeLocationT;                                                              //写入的EXCEL的D单元格位置
	CString writeLocationU;                                                              //写入的EXCEL的D单元格位置
	
	//判断是否进行了修改，如果进行了修改，则把修改好的新类型写入到EXCEL中去
	if (ReviseCategory1 != "")                                                        
	{
		writeLocationV.Format("V%d", startWriteRow);                                         //获得行号C几
		range = sheet.get_Range(COleVariant(writeLocationV), COleVariant(writeLocationV));   //得到坐标
		range.put_ColumnWidth(COleVariant("8"));
		range.put_Value2(COleVariant(ReviseCategory1));

	}
	if (ReviseTrayMark1 != "")
	{
		writeLocationT.Format("T%d", startWriteRow);                                         //获得行号C几
		range = sheet.get_Range(COleVariant(writeLocationT), COleVariant(writeLocationT));   //得到坐标
		range.put_ColumnWidth(COleVariant("8"));
		range.put_Value2(COleVariant(ReviseTrayMark1));
    }
	if (ReviseNumber1 != "")
	{
		writeLocationU.Format("U%d", startWriteRow);                                         //获得行号C几
		range = sheet.get_Range(COleVariant(writeLocationU), COleVariant(writeLocationU));   //得到坐标
		range.put_ColumnWidth(COleVariant("8"));
		range.put_Value2(COleVariant(ReviseNumber1));
	}



	//****************判断是否错误，错误的换用红色标志这一行********************************
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
