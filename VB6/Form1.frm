VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "IExplorerBrowser Custom Folder"
   ClientHeight    =   5445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8385
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   559
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Results"
      Height          =   3960
      Left            =   75
      TabIndex        =   6
      Top             =   1365
      Width           =   8265
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run"
      Height          =   360
      Left            =   45
      TabIndex        =   4
      Top             =   975
      Width           =   990
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Text            =   "*.docx;*.pdf;*.jpg"
      Top             =   405
      Width           =   1560
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1950
      TabIndex        =   1
      Top             =   405
      Width           =   2790
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "When ready, click 'Run' and the results will be displayed as a single folder."
      Height          =   195
      Left            =   30
      TabIndex        =   5
      Top             =   705
      Width           =   5745
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "in"
      Height          =   195
      Left            =   1665
      TabIndex        =   3
      Top             =   405
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First, set up a search to gather a custom file set:"
      Height          =   195
      Left            =   75
      TabIndex        =   0
      Top             =   135
      Width           =   3675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As oleexp.RECT) As Long
Private Declare Function ILCreateFromPathW Lib "shell32" (ByVal pwszPath As Long) As Long
Private Declare Function PathMatchSpecW Lib "shlwapi" (ByVal pszFileParam As Long, ByVal pszSpec As Long) As Long
Private Declare Function SHGetKnownFolderPath Lib "shell32" (rfid As Any, ByVal dwFlags As Long, ByVal hToken As Long, ppszPath As Long) As Long
Private Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
Private Declare Function vbaObjSetAddRef Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (ByRef objDest As Object, ByVal pObject As Long) As Long

Private pEBrowse As ExplorerBrowser
Private pResFolder As IResultsFolder
Private lprf As Long

Private nMatch As Long

Private mSpec As String
 
Private Sub Command1_Click()
If (pResFolder Is Nothing) Then
    Debug.Print "Error->No resultsfolder"
    Exit Sub
End If
If nMatch Then pResFolder.RemoveAll
nMatch = 0
mSpec = Text2.Text
DoSearch
End Sub

Private Sub Form_Load()
Dim lpDocs As Long
SHGetKnownFolderPath FOLDERID_Documents, 0&, 0&, lpDocs
Text1.Text = LPWSTRtoStr(lpDocs)

Set pEBrowse = New ExplorerBrowser
Dim lFlag As EXPLORER_BROWSER_OPTIONS
Dim tFS As FOLDERSETTINGS
Dim rcFrame As oleexp.RECT
Dim rcEB As oleexp.RECT
Dim pfv As IFolderView2
Dim pColMgr As IColumnManager

tFS.ViewMode = FVM_DETAILS
tFS.fFlags = FWF_AUTOARRANGE Or FWF_NOWEBVIEW

lFlag = EBO_NAVIGATEONCE

GetClientRect Frame1.hWnd, rcFrame
With rcEB
    .Left = 4
    .Top = 15
    .Right = rcFrame.Right - 4
    .Bottom = rcFrame.Bottom - 4
End With

pEBrowse.Initialize Frame1.hWnd, rcEB, tFS
pEBrowse.SetOptions lFlag
pEBrowse.FillFromObject Nothing, EBF_NODROPTARGET
pEBrowse.GetCurrentView IID_IFolderView2, pfv
If (pfv Is Nothing) = False Then
'    'OPTIONAL
'    'Customize which columns appear: fill uCol with however many columns (PROPERTYKEY's from mPKEY)
'    ' you want, then be sure to change the second argument in SetColumns to the number of keys
'    Dim uCol() As PROPERTYKEY
'    ReDim uCol(2)
'    Set pColMgr = pfv
'    uCol(0) = PKEY_ItemNameDisplay
'    uCol(1) = PKEY_ItemFolderPathDisplay
'    uCol(2) = PKEY_Image_Dimensions
'    pColMgr.SetColumns uCol(0), 3&
    pfv.GetFolder IID_IResultsFolder, lprf
    If lprf Then
        vbaObjSetAddRef pResFolder, lprf
    Else
        Debug.Print "Error->No RF pointer"
    End If
Else
    Debug.Print "Error->No folderview"
End If
 
End Sub
Private Function LPWSTRtoStr(lPtr As Long, Optional ByVal fFree As Boolean = True) As String
SysReAllocString VarPtr(LPWSTRtoStr), lPtr
If fFree Then
    Call CoTaskMemFree(lPtr)
End If
End Function

Private Sub DoSearch()

Dim psi As IShellItem
Dim piesi As IEnumShellItems
Dim isia As IShellItemArray
Dim pidl As Long
Dim pFile As IShellItem
Dim lpName As Long, lpFolder As Long
Dim sName As String, sFolder As String
Dim sDisp As String
Dim pcl As Long
Dim sTarget As String
Dim sStart As String
Dim lAtr As SFGAO_Flags

pidl = ILCreateFromPathW(StrPtr(Text1.Text))
SHCreateItemFromIDList pidl, IID_IShellItem, psi
psi.BindToHandler 0&, BHID_EnumItems, IID_IEnumShellItems, piesi

Do While piesi.Next(1&, pFile, pcl) = S_OK
    pFile.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpFolder
    pFile.GetAttributes SFGAO_FOLDER Or SFGAO_DROPTARGET Or SFGAO_STREAM, lAtr
    If ((lAtr And SFGAO_FOLDER) = SFGAO_FOLDER) And ((lAtr And SFGAO_STREAM) = 0&) Then 'folder but not zip
        If (lAtr And SFGAO_DROPTARGET) = SFGAO_DROPTARGET Then
            ScanDeep pFile
        End If
    Else
        pFile.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpName
        sName = LPWSTRtoStr(lpName)
        sDisp = Right$(sName, Len(sName) - InStrRev(sName, "\"))
        If PathMatchSpecW(StrPtr(sDisp), StrPtr(mSpec)) Then
            Debug.Print "Found match: " & sName
            pResFolder.AddItem pFile
        End If
    End If
Loop
Call CoTaskMemFree(pidl)

End Sub
Private Sub ScanDeep(psiLoc As IShellItem)
'for recursive scan
Dim psi As IShellItem
Dim piesi As IEnumShellItems
Dim pFile As IShellItem
Dim lpName As Long
Dim sName As String
Dim sDisp As String
Dim pcl As Long
Dim sTarget As String
Dim lAtr As SFGAO_Flags


psiLoc.BindToHandler 0&, BHID_EnumItems, IID_IEnumShellItems, piesi
Do While piesi.Next(1&, pFile, pcl) = S_OK
    pFile.GetAttributes SFGAO_FOLDER Or SFGAO_DROPTARGET Or SFGAO_STREAM, lAtr
    If ((lAtr And SFGAO_FOLDER) = SFGAO_FOLDER) And ((lAtr And SFGAO_STREAM) = 0&) Then 'folder but not zip
        If (lAtr And SFGAO_DROPTARGET) = SFGAO_DROPTARGET Then
            ScanDeep pFile
        End If
    Else
        pFile.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpName
        sName = LPWSTRtoStr(lpName)
        sDisp = Right$(sName, Len(sName) - InStrRev(sName, "\"))
        If PathMatchSpecW(StrPtr(sDisp), StrPtr(mSpec)) Then
            Debug.Print "Found match: " & sName
            pResFolder.AddItem pFile
        End If
    End If
Loop
End Sub

Private Sub Form_Resize()
Frame1.Width = Me.ScaleWidth - 10
Frame1.Height = Me.ScaleHeight - 100
Dim rcFrame As oleexp.RECT
Dim lp As Long
GetClientRect Frame1.hWnd, rcFrame
If (pEBrowse Is Nothing = False) Then
    pEBrowse.SetRect lp, 4, 15, rcFrame.Right - 4, rcFrame.Bottom - 4
End If
End Sub
