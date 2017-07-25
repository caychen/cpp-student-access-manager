// StudentAccessDlg.cpp : implementation file
//

#include "stdafx.h"
#include "StudentAccess.h"
#include "StudentAccessDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

	// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CStudentAccessDlg dialog




CStudentAccessDlg::CStudentAccessDlg(CWnd* pParent /*=NULL*/)
: CDialog(CStudentAccessDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CStudentAccessDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, m_dataList);
}

BEGIN_MESSAGE_MAP(CStudentAccessDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BTN_ADD, &CStudentAccessDlg::OnBnClickedBtnAdd)
	ON_BN_CLICKED(IDC_BTN_DELETE, &CStudentAccessDlg::OnBnClickedBtnDelete)
	ON_BN_CLICKED(IDC_BTN_TOTXTFILES, &CStudentAccessDlg::OnToTxtFiles)
	ON_BN_CLICKED(IDC_BTN_TOEXCEL, &CStudentAccessDlg::OnToExcel)
	ON_BN_CLICKED(IDC_BTN_ABOUT, &CStudentAccessDlg::OnBnClickedBtnAbout)
	ON_WM_DROPFILES()
	ON_BN_CLICKED(IDC_BTN_TODATABASE, &CStudentAccessDlg::OnToDataBase)
	ON_BN_CLICKED(IDC_BTN_CLEAR, &CStudentAccessDlg::OnBnClickedBtnClear)
	ON_WM_CTLCOLOR()
END_MESSAGE_MAP()


// CStudentAccessDlg message handlers

BOOL CStudentAccessDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here
	::SendMessage(m_dataList.m_hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE,// | LVM_GETSELECTEDCOLUMN,
		LVS_EX_FULLROWSELECT, LVS_EX_FULLROWSELECT);
	m_dataList.InsertColumn(0,_T("编号"),LVCFMT_LEFT,150);
	m_dataList.InsertColumn(1,_T("姓名"),LVCFMT_LEFT,150);
	m_dataList.InsertColumn(2,_T("性别"),LVCFMT_LEFT,150);
	m_dataList.InsertColumn(3,_T("学号"),LVCFMT_LEFT,150);
	m_dataList.InsertColumn(4,_T("年龄"),LVCFMT_LEFT,150);
	CheckRadioButton(IDC_RADIO1,IDC_RADIO2,IDC_RADIO1);

	static CFont font;
	font.CreatePointFont(180, _T("隶书"));
	//font
	GetDlgItem(IDC_HINT)->SetFont(&font);

	EnableControls(FALSE);

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CStudentAccessDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CStudentAccessDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CStudentAccessDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CStudentAccessDlg::OnBnClickedBtnAdd()
{
	// TODO: Add your control notification handler code here
	CString strName,strSex;
	UINT id,age;
	UINT uCount;
	TCHAR szBuf[50];

	GetDlgItemText(IDC_EDIT_NAME,strName);
	id = GetDlgItemInt(IDC_EDIT_ID);
	age = GetDlgItemInt(IDC_EDIT_AGE);

	if(strName.IsEmpty())
	{
		MessageBox(_T("姓名为空，请重新输入姓名"));
		GetDlgItem(IDC_EDIT_NAME)->SetFocus();
		return;
	}
	else if(!id)
	{
		MessageBox(_T("学号为空，请重新输入学号"));	
		GetDlgItem(IDC_EDIT_ID)->SetFocus();
		return;
	}
	else if(age < 0 || age > 100)
	{
		MessageBox(_T("年龄不符要求，请重新输入年龄:"));
		GetDlgItem(IDC_EDIT_AGE)->SetFocus();
		return;
	}
	else
	{
		if(IDC_RADIO1 == GetCheckedRadioButton(IDC_RADIO1,IDC_RADIO2))
			strSex = "Male";
		else
			strSex = "Famale";

		uCount = m_dataList.GetItemCount();

		//编号
		wsprintf(szBuf,_T("%d"),uCount);
		m_dataList.InsertItem(uCount,szBuf);

		//姓名
		m_dataList.SetItemText(uCount,1,strName);

		//性别
		m_dataList.SetItemText(uCount,2,strSex);

		//学号
		wsprintf(szBuf,_T("%d"),id);
		m_dataList.SetItemText(uCount,3,szBuf);

		//年龄
		wsprintf(szBuf,_T("%d"),age);
		m_dataList.SetItemText(uCount,4,szBuf);

		m_dataList.GetItemCount() > 0 ? EnableControls() : EnableControls(FALSE);
	}
}

void CStudentAccessDlg::OnBnClickedBtnDelete()
{
	// TODO: Add your control notification handler code here
	POSITION ps = m_dataList.GetFirstSelectedItemPosition();
	UINT index;
	TCHAR szBuf[50];

	if(ps == NULL)
	{
		MessageBox(_T("请选中项之后再删除"));	
		return;
	}
	else
	{
		index = m_dataList.GetNextSelectedItem(ps);

		for(int i = index + 1;i < m_dataList.GetItemCount();++i)
		{
			wsprintf(szBuf,_T("%d"),i - 1);
			m_dataList.SetItemText(i,0,szBuf);
		}
		m_dataList.DeleteItem(index);

		m_dataList.GetItemCount() > 0 ? EnableControls() : EnableControls(FALSE);
	}
}

void CStudentAccessDlg::OnToTxtFiles()
{
	// TODO: Add your control notification handler code here
	CString strFilePath;
	CString strName,strSex,strNum,strAge;
	CString strBuf;
	char *p;

	CFileDialog fileDlg(FALSE);
	fileDlg.m_ofn.lpstrFilter = _T("文本文件(*.txt)\0*.txt");
	fileDlg.m_ofn.lpstrDefExt = _T("txt");

	if(IDOK == fileDlg.DoModal())
	{
		strFilePath = fileDlg.GetPathName();
		CFile file;
		file.Open(strFilePath,CFile::modeCreate | CFile::modeWrite);

		for(int i = 0;i < m_dataList.GetItemCount();++i)
		{
			strName = m_dataList.GetItemText(i,1);
			strSex = m_dataList.GetItemText(i,2);
			strNum = m_dataList.GetItemText(i,3);
			strAge = m_dataList.GetItemText(i,4);

			strBuf.Format(_T("%s %s %s %s\r\n"),
				strName.GetBuffer(),strSex.GetBuffer(),strNum.GetBuffer(),strAge.GetBuffer());
			USES_CONVERSION;
			p = T2A(strBuf);

			file.Write(p,(UINT)strlen(p) * sizeof(char));

			strName.ReleaseBuffer();
			strSex.ReleaseBuffer();
			strNum.ReleaseBuffer();
			strAge.ReleaseBuffer();
		}
		p = NULL;
		file.Close();
	}
}

void CStudentAccessDlg::OnToExcel()
{
	// TODO: Add your control notification handler code here
	CString strFilePath;

	CFileDialog fileDlg(FALSE);
	fileDlg.m_ofn.lpstrFilter = _T("Excel2003 文件(*.xls)\0*.xls\0Excel2007 文件(*.xlsx)\0*.xlsx");
	fileDlg.m_ofn.lpstrDefExt = _T("xls");
	if(IDOK == fileDlg.DoModal())
	{
		strFilePath = fileDlg.GetPathName();
		CreateExcelFiles(strFilePath);
	}
}

void CStudentAccessDlg::EnableControls(BOOL bEnable/*= TRUE*/)
{
	GetDlgItem(IDC_BTN_TOTXTFILES)->EnableWindow(bEnable);
	GetDlgItem(IDC_BTN_TODATABASE)->EnableWindow(bEnable);
	GetDlgItem(IDC_BTN_TOEXCEL)->EnableWindow(bEnable);
	GetDlgItem(IDC_BTN_CLEAR)->EnableWindow(bEnable);
	GetDlgItem(IDC_BTN_DELETE)->EnableWindow(bEnable);
}

void CStudentAccessDlg::CreateExcelFiles(CString FilePath)
{
	if(PathFileExists(FilePath))
	{
		DeleteFile(FilePath);
	}

	if(!PathFileExists(FilePath))
	{
		CFile file;
		file.Open(FilePath,CFile::modeCreate);//清空
		file.Close();
	}

	_Application ExcelApp;    
	Workbooks ExcelBooks;
	_Workbook ExcelBook;
	Worksheets ExcelSheets;
	_Worksheet ExcelSheet;
	Range ExcelRange;
	Font ft;

	Range UnitRge;
	CString CellName;
	CString CellName1;

	CString str;
	CString strName;
	CString strSex;
	CString strNum;
	CString strAge;

	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR);

	if(!ExcelApp.CreateDispatch(_T("Excel.Application"),NULL))
	{
		MessageBox(_T("无法创建Excel应用"));
		return; 
	}

	ExcelBooks = ExcelApp.GetWorkbooks();

	ExcelBook = ExcelBooks.Open(FilePath,
		covOptional, covOptional, covOptional, covOptional, 
		covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional);

	ExcelSheets = ExcelBook.GetWorksheets(); 

	ExcelSheet = ExcelSheets.GetItem(COleVariant((short)1));

	ExcelSheet.SetName(_T("Excel1"));
	ExcelRange.Clear();

	//设置齐方式为水平垂直居中 
	//水平对齐：默认＝1,居中＝-4108,左＝-4131,右＝-4152 
	//垂直对齐：默认＝2,居中＝-4108,左＝-4160,右＝-4107 

	UpdateData();

	//设置标题
	ExcelRange.AttachDispatch(ExcelSheet.GetRange(COleVariant(_T("A1")),COleVariant(_T("J1"))));
	ExcelRange.Merge(COleVariant((long)0));//合并单元格
	ExcelRange.SetItem(COleVariant((long)1),COleVariant((long)1),COleVariant(_T("学生信息管理小系统")));
	ft = ExcelRange.GetFont();
	ft.SetName(COleVariant(_T("宋体")));//设置标题字体类型
	ft.SetBold(COleVariant((long)1));//设置标题为粗体
	ft.SetSize(COleVariant((long)22));//标题字体大小22号
	ft.SetUnderline(COleVariant((long)2));
	ExcelRange.SetHorizontalAlignment(COleVariant((long)-4108));
	ExcelRange.SetVerticalAlignment(COleVariant((long)-4108));

	//五个表头
	//编号
	ExcelRange.AttachDispatch(ExcelSheet.GetRange(COleVariant(_T("A2")),COleVariant(_T("B2"))));
	ExcelRange.Merge(COleVariant((long)0));//合并单元格
	ExcelRange.SetItem(COleVariant((long)1),COleVariant((long)1),COleVariant(_T("编号")));
	ft = ExcelRange.GetFont();
	ft.SetName(COleVariant(_T("宋体")));//设置标题字体类型
	ft.SetBold(COleVariant((long)1));//设置标题为粗体
	ft.SetSize(COleVariant((long)16));//标题字体大小16号
	ExcelRange.SetHorizontalAlignment(COleVariant((long)-4108));
	ExcelRange.SetVerticalAlignment(COleVariant((long)-4108));

	//姓名
	ExcelRange.AttachDispatch(ExcelSheet.GetRange(COleVariant(_T("C2")),COleVariant(_T("D2"))));
	ExcelRange.Merge(COleVariant((long)0));//合并单元格
	ExcelRange.SetItem(COleVariant((long)1),COleVariant((long)1),COleVariant(_T("姓名")));
	ft = ExcelRange.GetFont();
	ft.SetName(COleVariant(_T("宋体")));//设置标题字体类型
	ft.SetBold(COleVariant((long)1));//设置标题为粗体
	ft.SetSize(COleVariant((long)16));//标题字体大小16号
	ExcelRange.SetHorizontalAlignment(COleVariant((long)-4108));
	ExcelRange.SetVerticalAlignment(COleVariant((long)-4108));

	//性别
	ExcelRange.AttachDispatch(ExcelSheet.GetRange(COleVariant(_T("E2")),COleVariant(_T("F2"))));
	ExcelRange.Merge(COleVariant((long)0));//合并单元格
	ExcelRange.SetItem(COleVariant((long)1),COleVariant((long)1),COleVariant(_T("性别")));
	ft = ExcelRange.GetFont();
	ft.SetName(COleVariant(_T("宋体")));//设置标题字体类型
	ft.SetBold(COleVariant((long)1));//设置标题为粗体
	ft.SetSize(COleVariant((long)16));//标题字体大小16号
	ExcelRange.SetHorizontalAlignment(COleVariant((long)-4108));
	ExcelRange.SetVerticalAlignment(COleVariant((long)-4108));

	//学号
	ExcelRange.AttachDispatch(ExcelSheet.GetRange(COleVariant(_T("G2")),COleVariant(_T("H2"))));
	ExcelRange.Merge(COleVariant((long)0));//合并单元格
	ExcelRange.SetItem(COleVariant((long)1),COleVariant((long)1),COleVariant(_T("学号")));
	ft = ExcelRange.GetFont();
	ft.SetName(COleVariant(_T("宋体")));//设置标题字体类型
	ft.SetBold(COleVariant((long)1));//设置标题为粗体
	ft.SetSize(COleVariant((long)16));//标题字体大小16号
	ExcelRange.SetHorizontalAlignment(COleVariant((long)-4108));
	ExcelRange.SetVerticalAlignment(COleVariant((long)-4108));

	//年龄
	ExcelRange.AttachDispatch(ExcelSheet.GetRange(COleVariant(_T("I2")),COleVariant(_T("J2"))));
	ExcelRange.Merge(COleVariant((long)0));//合并单元格
	ExcelRange.SetItem(COleVariant((long)1),COleVariant((long)1),COleVariant(_T("年龄")));
	ft = ExcelRange.GetFont();
	ft.SetName(COleVariant(_T("宋体")));//设置标题字体类型
	ft.SetBold(COleVariant((long)1));//设置标题为粗体
	ft.SetSize(COleVariant((long)16));//标题字体大小16号
	ExcelRange.SetHorizontalAlignment(COleVariant((long)-4108));
	ExcelRange.SetVerticalAlignment(COleVariant((long)-4108));

	int dataCount = m_dataList.GetItemCount();

	//表格合并单元格
	for(int i = 0;i < dataCount;++i)
	{
		CellName.Format(_T("A%d"),i + 3);
		CellName1.Format(_T("B%d"),i + 3);
		ExcelRange.AttachDispatch(ExcelSheet.GetRange(COleVariant(CellName),COleVariant(CellName1)));
		ExcelRange.Merge(COleVariant((long)0));//合并单元格
		ft = ExcelRange.GetFont();
		ft.SetName(COleVariant(_T("宋体")));//设置标题字体类型
		ft.SetSize(COleVariant((long)12));//标题字体大小16号
		ExcelRange.SetHorizontalAlignment(COleVariant((long)-4108));
		ExcelRange.SetVerticalAlignment(COleVariant((long)-4108));
	}

	for(int i = 0;i < dataCount;++i)
	{
		CellName.Format(_T("C%d"),i + 3);
		CellName1.Format(_T("D%d"),i + 3);
		ExcelRange.AttachDispatch(ExcelSheet.GetRange(COleVariant(CellName),COleVariant(CellName1)));
		ExcelRange.Merge(COleVariant((long)0));//合并单元格
		ft = ExcelRange.GetFont();
		ft.SetName(COleVariant(_T("宋体")));//设置标题字体类型
		ft.SetSize(COleVariant((long)12));//标题字体大小16号
		ExcelRange.SetHorizontalAlignment(COleVariant((long)-4108));
		ExcelRange.SetVerticalAlignment(COleVariant((long)-4108));
	}

	for(int i = 0;i < dataCount;++i)
	{
		CellName.Format(_T("E%d"),i + 3);
		CellName1.Format(_T("F%d"),i + 3);
		ExcelRange.AttachDispatch(ExcelSheet.GetRange(COleVariant(CellName),COleVariant(CellName1)));
		ExcelRange.Merge(COleVariant((long)0));//合并单元格
		ft = ExcelRange.GetFont();
		ft.SetName(COleVariant(_T("宋体")));//设置标题字体类型
		ft.SetSize(COleVariant((long)12));//标题字体大小16号
		ExcelRange.SetHorizontalAlignment(COleVariant((long)-4108));
		ExcelRange.SetVerticalAlignment(COleVariant((long)-4108));
	}

	for(int i = 0;i < dataCount;++i)
	{
		CellName.Format(_T("G%d"),i + 3);
		CellName1.Format(_T("H%d"),i + 3);
		ExcelRange.AttachDispatch(ExcelSheet.GetRange(COleVariant(CellName),COleVariant(CellName1)));
		ExcelRange.Merge(COleVariant((long)0));//合并单元格
		ft = ExcelRange.GetFont();
		ft.SetName(COleVariant(_T("宋体")));//设置标题字体类型
		ft.SetSize(COleVariant((long)12));//标题字体大小16号
		ExcelRange.SetHorizontalAlignment(COleVariant((long)-4108));
		ExcelRange.SetVerticalAlignment(COleVariant((long)-4108));
	}

	for(int i = 0;i < dataCount;++i)
	{
		CellName.Format(_T("I%d"),i + 3);
		CellName1.Format(_T("J%d"),i + 3);
		ExcelRange.AttachDispatch(ExcelSheet.GetRange(COleVariant(CellName),COleVariant(CellName1)));
		ExcelRange.Merge(COleVariant((long)0));//合并单元格
		ft = ExcelRange.GetFont();
		ft.SetName(COleVariant(_T("宋体")));//设置标题字体类型
		ft.SetSize(COleVariant((long)12));//标题字体大小16号
		ExcelRange.SetHorizontalAlignment(COleVariant((long)-4108));
		ExcelRange.SetVerticalAlignment(COleVariant((long)-4108));
	}

	//设置边框
	for(int item = 2;item <= dataCount + 2;++item)
	{
		for(int column = 1;column <= 'J' - 'A' + 1;++column)
		{
			CellName.Format(_T("%c%d"),column + 64,item);
			UnitRge.AttachDispatch(ExcelSheet.GetRange(COleVariant(CellName),COleVariant(CellName)));

			//LineStyle=线型 Weight=线宽 ColorIndex=线的颜色(-4105为自动) 
			UnitRge.BorderAround(COleVariant((long)1),(long)2,(long)-4105,COleVariant((long)1));//设置边框 
		}
	}

	//开始存入数据
	for(int i = 0;i < dataCount;++i)
	{
		//编号
		str.Format(_T("%d"),i);
		CellName.Format(_T("A%d"),i + 3);
		ExcelRange.AttachDispatch(ExcelSheet.GetRange(COleVariant(CellName),COleVariant(CellName)));
		ExcelRange.SetValue(COleVariant(str));

		//姓名
		strName = m_dataList.GetItemText(i,1);
		CellName.Format(_T("C%d"),i + 3);
		ExcelRange.AttachDispatch(ExcelSheet.GetRange(COleVariant(CellName),COleVariant(CellName)));
		ExcelRange.SetValue(COleVariant(strName));

		//性别
		strSex = m_dataList.GetItemText(i,2);
		CellName.Format(_T("E%d"),i + 3);
		ExcelRange.AttachDispatch(ExcelSheet.GetRange(COleVariant(CellName),COleVariant(CellName)));
		ExcelRange.SetValue(COleVariant(strSex));

		//学号
		strNum = m_dataList.GetItemText(i,3);
		CellName.Format(_T("G%d"),i + 3);
		ExcelRange.AttachDispatch(ExcelSheet.GetRange(COleVariant(CellName),COleVariant(CellName)));
		ExcelRange.SetValue(COleVariant(strNum));

		//年龄
		strAge = m_dataList.GetItemText(i,4);
		CellName.Format(_T("I%d"),i + 3);
		ExcelRange.AttachDispatch(ExcelSheet.GetRange(COleVariant(CellName),COleVariant(CellName)));
		ExcelRange.SetValue(COleVariant(strAge));
	}

	ExcelApp.m_bAutoRelease = TRUE;
	ExcelBook.SetSaved(TRUE);

	ExcelApp.SetVisible(FALSE);
	//ExcelBook.Save();
	ExcelBook.SaveAs(COleVariant(FilePath),covOptional,covOptional,
		covOptional,covOptional,covOptional,0,covOptional,covOptional,covOptional,covOptional);
	ExcelApp.SetUserControl(TRUE);

	ExcelRange.ReleaseDispatch();
	ExcelBook.ReleaseDispatch();
	ExcelBooks.ReleaseDispatch();
	ExcelSheet.ReleaseDispatch();
	ExcelSheets.ReleaseDispatch();

	ExcelApp.ReleaseDispatch();
	ExcelApp.Quit();
}

void CStudentAccessDlg::OnBnClickedBtnAbout()
{
	// TODO: Add your control notification handler code here
	CAboutDlg dlg;
	dlg.DoModal();
}

void CStudentAccessDlg::OnDropFiles(HDROP hDropInfo)
{
	// TODO: Add your message handler code here and/or call default
	UINT FileCounts;
	wchar_t filePath[200];

	FileCounts = DragQueryFile(hDropInfo, 0xFFFFFFFF, NULL, 0);
	if(FileCounts)
	{        
		for(UINT i = 0; i < FileCounts; i++)
		{
			DragQueryFile(hDropInfo, i, filePath, sizeof(filePath));
			ShowRecord(filePath);
		}
	}

	DragFinish(hDropInfo);

	CDialog::OnDropFiles(hDropInfo);
}

void CStudentAccessDlg::ShowRecord(CString FilePath)
{
	int pos = FilePath.ReverseFind('.');
	CString FileExt = FilePath.Mid(pos + 1);
	//MessageBox(FileExt);


	if (FileExt == "txt")
	{
		if(LoadFromTxtFiles(FilePath))
			EnableControls();
	}
	else if (FileExt == "xls" || FileExt == "xlsx")
	{

	}
	else if (FileExt == "mdb")
	{

	}
	else
		MessageBox(_T("本程序只支持txt、xls、xlsx、mdb格式的文件！"));
}

void CStudentAccessDlg::OnToDataBase()
{
	// TODO: Add your control notification handler code here
	CString strFilePath;

	CFileDialog fileDlg(FALSE);
	fileDlg.m_ofn.lpstrFilter = _T("Access数据库文件(*.mdb)\0*.mdb");
	fileDlg.m_ofn.lpstrDefExt = _T("mdb");

	if(IDOK == fileDlg.DoModal())
	{
		strFilePath = fileDlg.GetPathName();

		if(!PathFileExists(strFilePath))
		{
			if(!CreateMdbFile(strFilePath))
			{
				return ;
			}
		}
		CreateTable(strFilePath);
		InsertDataToMdb(strFilePath);
	}
}

BOOL CStudentAccessDlg::CreateMdbFile(CString MdbFilePath)
{
	CString strMdbName = _T("Provider='Microsoft.Jet.OLEDB.4.0';Data Source = ") + MdbFilePath;

	try
	{
		HRESULT hr = S_OK;
		ADOX::_CatalogPtr pCatalog = NULL;

		hr = pCatalog.CreateInstance(__uuidof(ADOX::Catalog));
		if(FAILED(hr))
		{
			_com_issue_error(hr);
		}
		else
			pCatalog->Create(_bstr_t(strMdbName));

	}
	catch (_com_error &e)
	{
		CString strErro = CString(_T("创建ACCEESS数据库出错: ")) 
			+ (LPCSTR)e.Description();
		AfxMessageBox(strErro);
		return FALSE;
	}
	return TRUE;
}

BOOL CStudentAccessDlg::CreateTable(CString mdbFilePath)
{
	HRESULT hr;
	ADODB::_ConnectionPtr  pConnection;
	CString strDBName = _T("Provider='Microsoft.Jet.OLEDB.4.0';Data Source=");

	try
	{
		pConnection.CreateInstance(_T("ADODB.Connection"));

		strDBName = strDBName + mdbFilePath;
		hr = pConnection->Open(_bstr_t(strDBName),"","",ADODB::adModeUnknown);
		if(FAILED(hr))
		{
			_com_issue_error(hr);
		}

		CString sqlStr;
		sqlStr.Format(_T("CREATE TABLE Student(编号 COUNTER(0,1),姓名 Varchar(20),性别 Varchar(10),学号 INT,年龄 INT)"));//NUM INT,AGE INT)"));

		//sqlStr.Format(_T("CREATE TABLE Students(StudentID Char(7) not null,StudentName nVarchar(32) not null)"));
		hr = pConnection->Execute(_bstr_t(sqlStr),NULL,ADODB::adCmdText);
		if(FAILED(hr))
		{
			_com_issue_error(hr);
		}
		pConnection->Close();

	}
	catch(_com_error & e)
	{
		CString strErro = CString(_T("创建数据库表出错: ")) 
			+ (LPCSTR)e.Description();

		AfxMessageBox(strErro);
		return FALSE;
	}

	return TRUE;
}

void CStudentAccessDlg::InsertDataToMdb(CString mdbFilePath)
{
	ADODB::_RecordsetPtr pRecordSet;
	ADODB::_ConnectionPtr  pConnection;
	CString strSql;

	CString strName,strSex,strNum,strAge;

	CString strDBName = _T("Provider='Microsoft.Jet.OLEDB.4.0';Data Source=");

	HRESULT hr;
	try
	{
		pConnection.CreateInstance(_T("ADODB.Connection"));

		strDBName = strDBName + mdbFilePath;
		hr = pConnection->Open(_bstr_t(strDBName),"","",ADODB::adModeUnknown);
		if(FAILED(hr))
		{
			_com_issue_error(hr);
		}

		hr = pRecordSet.CreateInstance(_T("ADODB.Recordset"));
		if(FAILED(hr))
		{
			_com_issue_error(hr);
		}

		strSql = _T("SELECT * FROM Student");
		hr = pRecordSet->Open(_bstr_t(strSql),pConnection.GetInterfacePtr(),ADODB::adOpenStatic,ADODB::adLockOptimistic,ADODB::adCmdText);
		if(FAILED(hr))
		{
			_com_issue_error(hr);
		}

		for(int i = 0;i < m_dataList.GetItemCount();++i)
		{

			pRecordSet->AddNew();

			strName = m_dataList.GetItemText(i,1);
			strSex = m_dataList.GetItemText(i,2);
			strNum = m_dataList.GetItemText(i,3);
			strAge = m_dataList.GetItemText(i,4);

			pRecordSet->PutCollect(_T("姓名"),_bstr_t(strName));
			pRecordSet->PutCollect(_T("性别"),_bstr_t(strSex));
			pRecordSet->PutCollect(_T("学号"),_bstr_t(strNum));
			pRecordSet->PutCollect(_T("年龄"),_bstr_t(strAge));
			pRecordSet->Update();

		}
	}
	catch (_com_error& e)
	{
		CString strErro = CString(_T("插入数据出错: ")) 
			+ (LPCSTR)e.Description();

		AfxMessageBox(strErro);
	}

	pConnection->Close();
	pConnection.Release();
	pRecordSet.Release();
}

BOOL CStudentAccessDlg::LoadFromTxtFiles(CString FilePath)
{
	m_dataList.DeleteAllItems();
	CString strItem, strSubItem;
	TCHAR szBuf[50];

	UINT uItems = 0;
	int pos;

	CStdioFile file;
	file.Open(FilePath, CFile::modeRead);

	while (file.ReadString(strItem))
	{
		//编号
		wsprintf(szBuf, _T("%d"), uItems);
		m_dataList.InsertItem(uItems, szBuf);

		//姓名
		pos = strItem.Find(' ');
		strSubItem = (strItem.Left(pos));
		m_dataList.SetItemText(uItems, 1, strSubItem);
		strItem = strItem.Mid(pos + 1);

		//性别
		pos = strItem.Find(' ');
		strSubItem = strItem.Left(pos);
		m_dataList.SetItemText(uItems, 2, strSubItem);
		strItem = strItem.Mid(pos + 1);

		//学号
		pos = strItem.Find(' ');
		strSubItem = strItem.Left(pos);
		m_dataList.SetItemText(uItems, 3, strSubItem);
		strItem = strItem.Mid(pos + 1);

		//年龄
		m_dataList.SetItemText(uItems, 4, strItem);
		++uItems;
	}
	file.Close();
	if(uItems > 0)
		return TRUE;
	else 
		return FALSE;
}
void CStudentAccessDlg::OnBnClickedBtnClear()
{
	// TODO: Add your control notification handler code here
	if(m_dataList.GetItemCount() > 0)
	{
		m_dataList.DeleteAllItems();
		EnableControls(FALSE);
	}
}
HBRUSH CStudentAccessDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialog::OnCtlColor(pDC, pWnd, nCtlColor);

	// TODO:  Change any attributes of the DC here

	// TODO:  Return a different brush if the default is not desired
	if(IDC_HINT == pWnd->GetDlgCtrlID())
	{
		pDC->SetTextColor(RGB(255,0,0));
		//pDC->SetBkColor(RGB(0,0,255));
	}
	return hbr;
}