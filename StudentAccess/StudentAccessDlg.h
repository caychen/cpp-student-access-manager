// StudentAccessDlg.h : header file
//

#pragma once
#include "afxcmn.h"
#include "excel9.h"
#include <shlwapi.h>

#import "msado15.dll"  rename("EOF","adoEOF") 
#import "msadox.dll"  rename("EOF","adoXEOF")


// CStudentAccessDlg dialog
class CStudentAccessDlg : public CDialog
{
// Construction
public:
	CStudentAccessDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_STUDENTACCESS_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	CListCtrl m_dataList;

	afx_msg void OnBnClickedBtnAdd();
	afx_msg void OnBnClickedBtnDelete();
	afx_msg void OnToTxtFiles();
	afx_msg void OnToExcel();
	afx_msg void OnBnClickedBtnAbout();
	afx_msg void OnDropFiles(HDROP hDropInfo);
	afx_msg void OnToDataBase();

	void EnableControls(BOOL bEnable = TRUE);
	void CreateExcelFiles(CString FilePath);
	void ShowRecord(CString FilePath);
public:
	BOOL CreateMdbFile(CString MdbFilePath);

	
public:
	BOOL CreateTable(CString mdbFilePath);
public:
	void InsertDataToMdb(CString mdbFilePath);
	BOOL LoadFromTxtFiles(CString FilePath);
public:
	afx_msg void OnBnClickedBtnClear();
public:
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);

};
